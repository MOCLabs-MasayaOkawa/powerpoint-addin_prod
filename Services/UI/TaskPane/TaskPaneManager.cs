using Microsoft.Office.Tools;
using NLog;
using PowerPointEfficiencyAddin.Services.Core;
using PowerPointEfficiencyAddin.Services.Core.Alignment;
using PowerPointEfficiencyAddin.Services.Core.Shape;
using PowerPointEfficiencyAddin.Services.Core.Text;
using PowerPointEfficiencyAddin.Services.Infrastructure.MultiInstance;
using PowerPointEfficiencyAddin.UI;
using PowerPointEfficiencyAddin.Utils;
using System;
using System.Windows.Forms;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PowerPointEfficiencyAddin.Services.UI.TaskPane
{
    /// <summary>
    /// カスタムタスクペイン管理クラス
    /// </summary>
    public class TaskPaneManager : IDisposable
    {
        private static readonly Logger logger = LogManager.GetCurrentClassLogger();
        private readonly IApplicationProvider applicationProvider;

        private CustomTaskPane taskPane;
        private CustomTaskPaneUserControl taskPaneControl;

        // DI対応サービス参照
        private PowerToolService powerToolService;
        private AlignmentService alignmentService;
        private TextFormatService textFormatService;
        private ShapeService shapeService;

        private bool isDisposed = false;
        private bool isToggling = false; // 連続実行防止フラグ

        /// <summary>
        /// タスクペインが表示中かどうか
        /// </summary>
        public bool IsVisible => taskPane?.Visible ?? false;

        /// <summary>
        /// オブジェクトが破棄済みかどうか
        /// </summary>
        public bool IsDisposed => isDisposed;

        // DI対応コンストラクタ
        public TaskPaneManager(IApplicationProvider applicationProvider)
        {
            this.applicationProvider = applicationProvider ?? throw new ArgumentNullException(nameof(applicationProvider));
            logger.Debug("TaskPaneManager initialized with DI application provider");
        }

        // 既存コンストラクタ（後方互換性維持）
        public TaskPaneManager() : this(new DefaultApplicationProvider())
        {
            logger.Debug("TaskPaneManager initialized with default application provider");
        }

        /// <summary>
        /// タスクペインの幅
        /// </summary>
        public int Width
        {
            get
            {
                try
                {
                    return taskPane?.Width ?? 280;
                }
                catch (Exception ex)
                {
                    logger.Warn(ex, "Failed to get TaskPane width");
                    return 280; // デフォルト値
                }
            }
            set
            {
                if (taskPane != null)
                {
                    try
                    {
                        // 幅は200-600pxの範囲で制限
                        var newWidth = Math.Max(200, Math.Min(600, value));
                        taskPane.Width = newWidth;
                        logger.Debug($"TaskPane width set to {newWidth}");
                    }
                    catch (Exception ex)
                    {
                        logger.Warn(ex, $"Failed to set TaskPane width to {value}");
                    }
                }
            }
        }

        /// <summary>
        /// タスクペインの高さ（左/右ドック時は取得のみ、設定不可）
        /// </summary>
        public int Height
        {
            get
            {
                try
                {
                    return taskPane?.Height ?? 600;
                }
                catch (Exception ex)
                {
                    logger.Warn(ex, "Failed to get TaskPane height");
                    return 600; // デフォルト値
                }
            }
            set
            {
                if (taskPane != null)
                {
                    try
                    {
                        // ドック位置を確認してから高さ設定を試行
                        var dockPosition = taskPane.DockPosition;

                        if (dockPosition == Microsoft.Office.Core.MsoCTPDockPosition.msoCTPDockPositionTop ||
                            dockPosition == Microsoft.Office.Core.MsoCTPDockPosition.msoCTPDockPositionBottom)
                        {
                            // 上/下ドックの場合のみ高さ設定可能
                            var newHeight = Math.Max(300, value);
                            taskPane.Height = newHeight;
                            logger.Debug($"TaskPane height set to {newHeight}");
                        }
                        else
                        {
                            // 左/右ドックの場合は高さ設定不可
                            logger.Debug($"TaskPane height setting skipped for dock position: {dockPosition}");
                        }
                    }
                    catch (Exception ex)
                    {
                        logger.Warn(ex, $"Failed to set TaskPane height to {value}");
                    }
                }
            }
        }

        /// <summary>
        /// DI対応サービス注入（商用機能）
        /// </summary>
        public void InjectAllServices(PowerToolService powerToolService, AlignmentService alignmentService,
            TextFormatService textFormatService, ShapeService shapeService)
        {
            this.powerToolService = powerToolService ?? throw new ArgumentNullException(nameof(powerToolService));
            this.alignmentService = alignmentService ?? throw new ArgumentNullException(nameof(alignmentService));
            this.textFormatService = textFormatService ?? throw new ArgumentNullException(nameof(textFormatService));
            this.shapeService = shapeService ?? throw new ArgumentNullException(nameof(shapeService));

            logger.Info("All DI services injected into TaskPaneManager");

            // TaskPaneControlがある場合はサービス注入
            if (taskPaneControl != null)
            {
                taskPaneControl.InjectAllServices(powerToolService, alignmentService, textFormatService, shapeService);
            }
        }

        /// <summary>
        /// タスクペインを初期化します
        /// </summary>
        public void Initialize()
        {
            try
            {
                if (taskPane != null)
                {
                    logger.Warn("TaskPane already initialized");
                    return;
                }

                logger.Info("Initializing TaskPane with DI support");

                // UserControl作成（DI対応）
                taskPaneControl = new CustomTaskPaneUserControl();

                // サービス注入（利用可能な場合）
                if (powerToolService != null && alignmentService != null)
                {
                    taskPaneControl.InjectAllServices(powerToolService, alignmentService, textFormatService, shapeService);
                }

                // CustomTaskPane作成（ウィンドウ固有）
                var currentWindow = GetCurrentActiveWindow();
                if (currentWindow != null)
                {
                    taskPane = Globals.ThisAddIn.CustomTaskPanes.Add(taskPaneControl, "PowerPoint効率化ツール（商用版）", currentWindow);
                    logger.Info($"TaskPane created for specific window: {currentWindow.Caption}");
                }
                else
                {
                    // フォールバック：従来通り
                    taskPane = Globals.ThisAddIn.CustomTaskPanes.Add(taskPaneControl, "PowerPoint効率化ツール（商用版）");
                    logger.Warn("TaskPane created without window binding (fallback mode)");
                }

                // 基本設定
                taskPane.Visible = false;
                taskPane.DockPosition = Microsoft.Office.Core.MsoCTPDockPosition.msoCTPDockPositionLeft;
                taskPane.DockPositionRestrict = Microsoft.Office.Core.MsoCTPDockPositionRestrict.msoCTPDockPositionRestrictNone;

                // サイズ設定
                System.Threading.Thread.Sleep(200);
                try
                {
                    taskPane.Width = 280;
                }
                catch (Exception sizeEx)
                {
                    logger.Warn(sizeEx, "Failed to set initial TaskPane width");
                }

                // イベントハンドラ設定
                SetupEventHandlers();

                logger.Info("TaskPane with DI support initialized successfully");
            }
            catch (Exception ex)
            {
                logger.Error(ex, "Failed to initialize TaskPane with DI");
                CleanupPartialInitialization();
                throw;
            }
        }

        /// <summary>
        /// タスクペインコントロールを取得
        /// </summary>
        public CustomTaskPaneUserControl TaskPaneControl
        {
            get
            {
                if (taskPaneControl != null && !taskPaneControl.IsDisposed)
                {
                    return taskPaneControl;
                }
                return null;
            }
        }

        /// <summary>
        /// 現在アクティブなPowerPointウィンドウを取得
        /// </summary>
        private PowerPoint.DocumentWindow GetCurrentActiveWindow()
        {
            try
            {
                var currentApp = applicationProvider.GetCurrentApplication();
                if (currentApp?.ActiveWindow != null)
                {
                    logger.Debug($"Active window found: {currentApp.ActiveWindow.Caption}");
                    return currentApp.ActiveWindow;
                }

                logger.Debug("No active window found");
                return null;
            }
            catch (Exception ex)
            {
                logger.Warn(ex, "Failed to get current active window");
                return null;
            }
        }

        /// <summary>
        /// 部分的な初期化状態をクリーンアップします
        /// </summary>
        private void CleanupPartialInitialization()
        {
            try
            {
                logger.Debug("Cleaning up partial initialization");

                if (taskPane != null)
                {
                    try
                    {
                        Globals.ThisAddIn.CustomTaskPanes.Remove(taskPane);
                    }
                    catch (Exception ex)
                    {
                        logger.Warn(ex, "Failed to remove partially initialized TaskPane");
                    }
                    taskPane = null;
                }

                if (taskPaneControl != null && !taskPaneControl.IsDisposed)
                {
                    try
                    {
                        taskPaneControl.Dispose();
                    }
                    catch (Exception ex)
                    {
                        logger.Warn(ex, "Failed to dispose partially initialized TaskPaneControl");
                    }
                    taskPaneControl = null;
                }

                logger.Debug("Partial initialization cleanup completed");
            }
            catch (Exception ex)
            {
                logger.Error(ex, "Error during partial initialization cleanup");
            }
        }

        /// <summary>
        /// タスクペインの追加設定を行います（オプション）
        /// </summary>
        private void ConfigureTaskPane()
        {
            try
            {
                logger.Debug("Applying additional TaskPane configuration");

                // 追加の設定があればここで実行
                // 現在は基本設定のみなので、特に追加設定なし

                logger.Debug("Additional TaskPane configuration completed");
            }
            catch (Exception ex)
            {
                logger.Error(ex, "Failed to apply additional TaskPane configuration");
                throw;
            }
        }

        /// <summary>
        /// イベントハンドラを設定します
        /// </summary>
        private void SetupEventHandlers()
        {
            if (taskPane != null)
            {
                taskPane.VisibleChanged += TaskPane_VisibleChanged;
                taskPane.DockPositionChanged += TaskPane_DockPositionChanged;
            }

            logger.Debug("TaskPane event handlers setup completed");
        }

        /// <summary>
        /// 表示状態変更イベントハンドラ
        /// </summary>
        private void TaskPane_VisibleChanged(object sender, EventArgs e)
        {
            try
            {
                logger.Info($"TaskPane visibility changed: {IsVisible}");

                // 表示状態を保存
                SaveTaskPaneState();

                // リボンUIの状態更新（必要に応じて）
                UpdateRibbonState();
            }
            catch (Exception ex)
            {
                logger.Error(ex, "Error in TaskPane_VisibleChanged");
            }
        }

        /// <summary>
        /// ドック位置変更イベントハンドラ
        /// </summary>
        private void TaskPane_DockPositionChanged(object sender, EventArgs e)
        {
            try
            {
                var dockPosition = taskPane.DockPosition;
                logger.Info($"TaskPane dock position changed: {dockPosition}");

                // ドック位置に応じてサイズを再設定
                ConfigureTaskPaneSizeForDockPosition(dockPosition);

                // 位置状態を保存
                SaveTaskPaneState();
            }
            catch (Exception ex)
            {
                logger.Error(ex, "Error in TaskPane_DockPositionChanged");
            }
        }

        /// <summary>
        /// ドック位置に応じてタスクペインのサイズを設定します
        /// </summary>
        /// <param name="dockPosition">ドック位置</param>
        private void ConfigureTaskPaneSizeForDockPosition(Microsoft.Office.Core.MsoCTPDockPosition dockPosition)
        {
            try
            {
                switch (dockPosition)
                {
                    case Microsoft.Office.Core.MsoCTPDockPosition.msoCTPDockPositionLeft:
                    case Microsoft.Office.Core.MsoCTPDockPosition.msoCTPDockPositionRight:
                        // 左/右ドック：幅のみ設定可能
                        try
                        {
                            taskPane.Width = 280;
                            logger.Debug($"Set width for {dockPosition} dock position: {taskPane.Width}");
                        }
                        catch (Exception ex)
                        {
                            logger.Warn(ex, $"Failed to set width for {dockPosition} position");
                        }
                        break;

                    case Microsoft.Office.Core.MsoCTPDockPosition.msoCTPDockPositionTop:
                    case Microsoft.Office.Core.MsoCTPDockPosition.msoCTPDockPositionBottom:
                        // 上/下ドック：高さのみ設定可能
                        try
                        {
                            taskPane.Height = 200;
                            logger.Debug($"Set height for {dockPosition} dock position: {taskPane.Height}");
                        }
                        catch (Exception ex)
                        {
                            logger.Warn(ex, $"Failed to set height for {dockPosition} position");
                        }
                        break;

                    case Microsoft.Office.Core.MsoCTPDockPosition.msoCTPDockPositionFloating:
                        // フローティング：幅と高さ両方設定可能
                        try
                        {
                            taskPane.Width = 280;
                            taskPane.Height = 600;
                            logger.Debug($"Set size for floating position: {taskPane.Width}x{taskPane.Height}");
                        }
                        catch (Exception ex)
                        {
                            logger.Warn(ex, "Failed to set size for floating position");
                        }
                        break;

                    default:
                        logger.Debug($"Unknown dock position: {dockPosition}");
                        break;
                }
            }
            catch (Exception ex)
            {
                logger.Error(ex, $"Error configuring size for dock position: {dockPosition}");
            }
        }

        /// <summary>
        /// タスクペインの表示/非表示を切り替えます（UIスレッド対応版）
        /// </summary>
        public void ToggleVisibility()
        {
            try
            {
                // 連続実行防止
                if (isToggling)
                {
                    logger.Warn("ToggleVisibility is already in progress, skipping");
                    return;
                }

                // UIスレッドで実行されているかチェック
                if (taskPaneControl != null && taskPaneControl.InvokeRequired)
                {
                    logger.Info("ToggleVisibility: Not on UI thread, invoking on UI thread");

                    taskPaneControl.Invoke(new Action(ToggleVisibilityInternal));
                }
                else
                {
                    logger.Info("ToggleVisibility: Already on UI thread, executing directly");
                    ToggleVisibilityInternal();
                }
            }
            catch (Exception ex)
            {
                logger.Error(ex, "Failed to toggle TaskPane visibility");
                ErrorHandler.ExecuteSafely(() =>
                {
                    throw new InvalidOperationException("効率化ペインの表示切り替えに失敗しました。");
                }, "ペイン表示切り替え");
            }
        }

        /// <summary>
        /// タスクペイン表示切り替えの実際の処理（UIスレッド専用）
        /// </summary>
        private void ToggleVisibilityInternal()
        {
            try
            {
                isToggling = true;

                if (taskPane == null)
                {
                    logger.Warn("TaskPane not initialized, initializing now");
                    Initialize();
                }

                var currentVisibility = taskPane.Visible;
                var newVisibility = !currentVisibility;

                logger.Info($"ToggleVisibilityInternal: {currentVisibility} → {newVisibility}");

                // 表示状態を変更
                if (newVisibility)
                {
                    // 表示前に追加設定を実行
                    try
                    {
                        var dockPosition = taskPane.DockPosition;
                        ConfigureTaskPaneSizeForDockPosition(dockPosition);
                        logger.Debug($"Pre-show configuration completed for dock position: {dockPosition}");
                    }
                    catch (Exception configEx)
                    {
                        logger.Warn(configEx, "Failed to configure before show, continuing");
                    }

                    taskPane.Visible = true;
                    logger.Info($"TaskPane shown successfully on UI thread. Final visibility: {taskPane.Visible}");
                }
                else
                {
                    taskPane.Visible = false;
                    logger.Info($"TaskPane hidden successfully on UI thread. Final visibility: {taskPane.Visible}");
                }

                // UI更新を強制
                ForceTaskPaneRefresh();

                // 少し待機してから状態確認
                System.Threading.Thread.Sleep(50);
                var finalVisibility = taskPane.Visible;
                logger.Info($"Final TaskPane visibility after toggle: {finalVisibility}");

                if (finalVisibility != newVisibility)
                {
                    logger.Warn($"TaskPane visibility mismatch. Expected: {newVisibility}, Actual: {finalVisibility}");

                    // 再試行
                    logger.Info("Retrying visibility change...");
                    taskPane.Visible = newVisibility;
                    System.Threading.Thread.Sleep(100);

                    var retryVisibility = taskPane.Visible;
                    logger.Info($"Retry result: {retryVisibility}");
                }
            }
            catch (Exception ex)
            {
                logger.Error(ex, "Failed to toggle TaskPane visibility internally");
                throw;
            }
            finally
            {
                isToggling = false;
            }
        }

        /// <summary>
        /// タスクペインのUI更新を強制します
        /// </summary>
        private void ForceTaskPaneRefresh()
        {
            try
            {
                logger.Debug("Forcing TaskPane UI refresh");

                // タスクペインコントロールの更新
                if (taskPaneControl != null)
                {
                    taskPaneControl.Refresh();
                    taskPaneControl.Update();
                    logger.Debug("TaskPaneControl refreshed");
                }

                // Office CustomTaskPaneの更新
                if (taskPane != null)
                {
                    // タスクペインのサイズを微調整して再描画を促す
                    var currentWidth = taskPane.Width;
                    taskPane.Width = currentWidth + 1;
                    System.Threading.Thread.Sleep(10);
                    taskPane.Width = currentWidth;

                    logger.Debug("TaskPane size adjustment completed");
                }

                // Windows Formsの更新
                Application.DoEvents();

                logger.Debug("TaskPane UI refresh completed");
            }
            catch (Exception ex)
            {
                logger.Error(ex, "Failed to force TaskPane UI refresh");
            }
        }

        /// <summary>
        /// タスクペインを表示します
        /// </summary>
        public void Show()
        {
            try
            {
                if (taskPane == null)
                {
                    Initialize();
                }

                if (!IsVisible)
                {
                    logger.Info("Showing TaskPane");

                    // 初回表示時にサイズを再設定（初期化時に失敗していた場合の対策）
                    try
                    {
                        var dockPosition = taskPane.DockPosition;
                        ConfigureTaskPaneSizeForDockPosition(dockPosition);
                        logger.Debug($"Size configured for show: {dockPosition}");
                    }
                    catch (Exception sizeEx)
                    {
                        logger.Warn(sizeEx, "Failed to configure size on show, continuing anyway");
                    }

                    taskPane.Visible = true;

                    // 表示状態確認
                    System.Threading.Thread.Sleep(50);
                    var actualVisibility = taskPane.Visible;
                    logger.Info($"TaskPane show completed. Actual visibility: {actualVisibility}");

                    if (!actualVisibility)
                    {
                        logger.Warn("TaskPane failed to show properly");
                    }
                }
                else
                {
                    logger.Debug("TaskPane is already visible");
                }
            }
            catch (Exception ex)
            {
                logger.Error(ex, "Failed to show TaskPane");
                throw;
            }
        }

        /// <summary>
        /// タスクペインを非表示にします
        /// </summary>
        public void Hide()
        {
            try
            {
                if (taskPane != null && IsVisible)
                {
                    logger.Info("Hiding TaskPane");
                    taskPane.Visible = false;

                    // 非表示状態確認
                    System.Threading.Thread.Sleep(50);
                    var actualVisibility = taskPane.Visible;
                    logger.Info($"TaskPane hide completed. Actual visibility: {actualVisibility}");

                    if (actualVisibility)
                    {
                        logger.Warn("TaskPane failed to hide properly");
                    }
                }
                else
                {
                    logger.Debug("TaskPane is already hidden or not initialized");
                }
            }
            catch (Exception ex)
            {
                logger.Error(ex, "Failed to hide TaskPane");
                throw;
            }
        }

        /// <summary>
        /// タスクペインの状態を保存します
        /// </summary>
        private void SaveTaskPaneState()
        {
            try
            {
                if (taskPane == null) return;

                // 現在の状態を取得（エラーに対応）
                var isVisible = IsVisible;
                var currentWidth = Width;   // 既に安全な取得ロジック実装済み
                var currentHeight = Height; // 既に安全な取得ロジック実装済み
                var dockPosition = taskPane.DockPosition;

                // レジストリまたは設定ファイルに保存
                // 今回は簡易実装としてログ出力のみ
                logger.Debug($"TaskPane state - Visible: {isVisible}, " +
                           $"Position: {dockPosition}, " +
                           $"Size: {currentWidth}x{currentHeight}");

                // 実際の実装では Microsoft.Office.Tools.Settings等を使用
                // Properties.Settings.Default.TaskPaneVisible = isVisible;
                // Properties.Settings.Default.TaskPaneWidth = currentWidth;
                // Properties.Settings.Default.TaskPaneDockPosition = (int)dockPosition;
                // Properties.Settings.Default.Save();
            }
            catch (Exception ex)
            {
                logger.Warn(ex, "Failed to save TaskPane state");
            }
        }

        /// <summary>
        /// タスクペインの状態を復元します
        /// </summary>
        public void RestoreTaskPaneState()
        {
            try
            {
                // 保存された状態を復元
                // 今回は簡易実装としてデフォルト値を使用
                logger.Debug("Restoring TaskPane state to defaults");

                // 実際の実装例（設定ファイルから復元）
                // var wasVisible = Properties.Settings.Default.TaskPaneVisible;
                // var savedWidth = Properties.Settings.Default.TaskPaneWidth;
                // var savedDockPosition = Properties.Settings.Default.TaskPaneDockPosition;

                // if (wasVisible)
                // {
                //     Show();
                //     
                //     // ドック位置に応じて適切なサイズを設定
                //     try 
                //     {
                //         if (taskPane.DockPosition == msoCTPDockPositionLeft || 
                //             taskPane.DockPosition == msoCTPDockPositionRight)
                //         {
                //             Width = savedWidth;
                //         }
                //     }
                //     catch (Exception ex)
                //     {
                //         logger.Warn(ex, "Failed to restore TaskPane size");
                //     }
                // }

                logger.Debug("TaskPane state restoration completed");
            }
            catch (Exception ex)
            {
                logger.Warn(ex, "Failed to restore TaskPane state");
            }
        }

        /// <summary>
        /// リボンUIの状態を更新します
        /// </summary>
        private void UpdateRibbonState()
        {
            try
            {
                // リボンの表示切り替えボタンの状態を更新
                // 今回の実装ではリボンを最小化するため、この処理は簡素化
                logger.Debug("Ribbon state update (placeholder)");
            }
            catch (Exception ex)
            {
                logger.Warn(ex, "Failed to update ribbon state");
            }
        }

        /// <summary>
        /// 健全性チェックを実行します
        /// </summary>
        public bool PerformHealthCheck()
        {
            try
            {
                if (taskPane == null)
                {
                    logger.Warn("TaskPane is null");
                    return false;
                }

                if (taskPaneControl == null)
                {
                    logger.Warn("TaskPaneControl is null");
                    return false;
                }

                // UserControlが正常に機能しているかチェック
                if (taskPaneControl.IsDisposed)
                {
                    logger.Warn("TaskPaneControl is disposed");
                    return false;
                }

                logger.Debug("TaskPane health check passed");
                return true;
            }
            catch (Exception ex)
            {
                logger.Error(ex, "TaskPane health check failed");
                return false;
            }
        }

        /// <summary>
        /// タスクペインを再初期化します
        /// </summary>
        public void Reinitialize()
        {
            try
            {
                logger.Info("Reinitializing TaskPane");

                // 既存のタスクペインを破棄
                Dispose();

                // 新しいタスクペインを初期化
                Initialize();

                logger.Info("TaskPane reinitialized successfully");
            }
            catch (Exception ex)
            {
                logger.Error(ex, "Failed to reinitialize TaskPane");
                throw;
            }
        }

        /// <summary>
        /// リソースを解放します
        /// </summary>
        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        /// <summary>
        /// リソースを解放します
        /// </summary>
        /// <param name="disposing">マネージドリソースを解放するかどうか</param>
        protected virtual void Dispose(bool disposing)
        {
            if (!isDisposed && disposing)
            {
                try
                {
                    // イベントハンドラを解除（まず最初に実行）
                    if (taskPane != null)
                    {
                        try
                        {
                            taskPane.VisibleChanged -= TaskPane_VisibleChanged;
                            taskPane.DockPositionChanged -= TaskPane_DockPositionChanged;
                            logger.Debug("TaskPane event handlers removed in dispose");
                        }
                        catch (Exception ex)
                        {
                            logger.Warn(ex, "Failed to remove TaskPane event handlers");
                        }
                    }

                    // UserControlを解放
                    if (taskPaneControl != null && !taskPaneControl.IsDisposed)
                    {
                        taskPaneControl.Dispose();
                        taskPaneControl = null;
                        logger.Debug("TaskPaneControl disposed");
                    }

                    // タスクペインを削除
                    if (taskPane != null)
                    {
                        try
                        {
                            // まず非表示にする
                            if (taskPane.Visible)
                            {
                                taskPane.Visible = false;
                                logger.Debug("TaskPane hidden before removal");
                            }

                            // 少し待機してからコレクションから削除
                            System.Threading.Thread.Sleep(100);

                            // CustomTaskPaneCollectionから削除
                            if (Globals.ThisAddIn?.CustomTaskPanes != null)
                            {
                                Globals.ThisAddIn.CustomTaskPanes.Remove(taskPane);
                                logger.Debug("TaskPane removed from collection successfully");
                            }
                        }
                        catch (ObjectDisposedException)
                        {
                            // PowerPoint終了時等でオブジェクトが既に破棄されている場合は正常
                            logger.Debug("TaskPane collection already disposed (normal during shutdown)");
                        }
                        catch (Exception ex)
                        {
                            logger.Warn(ex, "Failed to remove TaskPane from collection");
                        }
                        finally
                        {
                            taskPane = null;
                        }
                    }

                    logger.Debug("TaskPaneManager disposed successfully");
                }
                catch (Exception ex)
                {
                    logger.Error(ex, "Error during TaskPaneManager disposal");
                }

                isDisposed = true;
            }
        }

        /// <summary>
        /// デストラクタ
        /// </summary>
        ~TaskPaneManager()
        {
            Dispose(false);
        }
    }
}