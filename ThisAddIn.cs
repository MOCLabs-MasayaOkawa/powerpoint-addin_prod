using Microsoft.Office.Core;
using Microsoft.Office.Tools;
using NLog;
using PowerPointEfficiencyAddin.Models.Licensing;
using PowerPointEfficiencyAddin.Services;
using PowerPointEfficiencyAddin.Services.Licensing;
using PowerPointEfficiencyAddin.Utils;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PowerPointEfficiencyAddin
{
    /// <summary>
    /// PowerPoint効率化アドインのメインクラス（カスタムペイン対応版）
    /// </summary>
    [System.Runtime.InteropServices.ComVisible(true)]
    public partial class ThisAddIn
    {
        private static readonly Logger logger = LogManager.GetCurrentClassLogger();

        // 既存のフィールドに追加
        private LicenseManager licenseManager;
        private bool isLicenseValid = false;

        // 商用レベルDI管理
        private ApplicationContextManager applicationContextManager;
        private IApplicationProvider applicationProvider;

        // DI対応サービスインスタンス
        private PowerToolService powerToolService;
        private AlignmentService alignmentService;
        private TextFormatService textFormatService;
        private ShapeService shapeService;

        // カスタムタスクペイン管理
        private TaskPaneManager taskPaneManager;
        private readonly Dictionary<string, TaskPaneManager> windowTaskPaneManagers = new Dictionary<string, TaskPaneManager>();

        // キーボードショートカット管理（一旦無効化）
        // private KeyboardHookManager keyboardHookManager;

        /// <summary>
        /// タスクペインマネージャーを取得
        /// </summary>
        public TaskPaneManager TaskPaneManager => taskPaneManager;

        /// <summary>
        /// DI対応PowerToolServiceアクセサー
        /// </summary>
        public PowerToolService PowerToolService => powerToolService;

        /// <summary>
        /// DI対応AlignmentServiceアクセサー
        /// </summary>
        public AlignmentService AlignmentService => alignmentService;

        /// <summary>
        /// DI対応TextFormatServiceアクセサー
        /// </summary>
        public TextFormatService TextFormatService => textFormatService;

        /// <summary>
        /// DI対応ShapeServiceアクセサー
        /// </summary>
        public ShapeService ShapeService => shapeService;

        /// <summary>
        /// ApplicationContextManagerアクセサー（商用デバッグ用）
        /// </summary>
        public ApplicationContextManager ApplicationContextManager => applicationContextManager;


        /// <summary>
        /// アドイン開始時の処理
        /// </summary>
        /// <param name="sender">送信者</param>
        /// <param name="e">イベント引数</param>
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            try
            {
                LoggerConfig.Initialize();
                logger.Info("Commercial PowerPoint Efficiency Addin startup initiated");

                // Step 0: ライセンス認証を最初に実行
                InitializeLicensingAsync();


                // Step 1: 複数インスタンス対応基盤の初期化
                InitializeMultiInstanceSupport();

                // Step 2: DI対応サービス初期化
                InitializeDIServices();

                // Step 3: アプリケーションイベント購読
                SubscribeToApplicationEvents();

                // Step 4: カスタムタスクペイン初期化（DI対応サービス注入）
                InitializeCustomTaskPaneWithDI();

                logger.Info("Commercial multi-instance PowerPoint Efficiency Addin startup completed successfully");
            }
            catch (Exception ex)
            {
                logger.Fatal(ex, "Critical error during commercial addin startup");
                HandleStartupFailure(ex);
            }
        }

        /// <summary>
        /// ライセンス認証の初期化（新規追加）
        /// </summary>
        private async void InitializeLicensingAsync()
        {
            try
            {
                logger.Info("Initializing license management system");

                // LicenseManagerのインスタンス取得
                licenseManager = LicenseManager.Instance;

                // 非同期でライセンス検証（UIをブロックしない）
                var validationResult = await licenseManager.InitializeAsync();

                // 検証結果に基づいて状態を設定
                isLicenseValid = validationResult.IsSuccess ||
                    validationResult.AccessLevel != FeatureAccessLevel.Blocked;

                // ユーザーへの通知
                if (!validationResult.IsSuccess)
                {
                    ShowLicenseNotification(validationResult);
                }

                // タスクペインにライセンス状態を反映
                UpdateTaskPaneLicenseStatus();

                logger.Info($"License initialization completed. Status: {licenseManager.GetStatusMessage()}");
            }
            catch (Exception ex)
            {
                logger.Error(ex, "Failed to initialize licensing");
                isLicenseValid = false;

                // ライセンスエラーでも基本機能は使えるようにする
                ShowLicenseErrorNotification();
            }
        }

        /// <summary>
        /// ライセンス通知を表示（新規追加）
        /// </summary>
        private void ShowLicenseNotification(LicenseValidationResult result)
        {
            try
            {
                string message = result.Message;
                string title = "ライセンス状態";
                System.Windows.Forms.MessageBoxIcon icon;

                switch (result.Type)
                {
                    case ValidationType.OfflineGrace:
                        if (result.AccessLevel == FeatureAccessLevel.Free)
                        {
                            message += "\n\n一部の高度な機能が制限されています。";
                            icon = System.Windows.Forms.MessageBoxIcon.Warning;
                        }
                        else
                        {
                            // 3日以内のオフラインは通知しない（UXを考慮）
                            return;
                        }
                        break;

                    case ValidationType.Expired:
                        message += "\n\nライセンスを更新してください。";
                        icon = System.Windows.Forms.MessageBoxIcon.Error;
                        break;

                    case ValidationType.NoLicense:
                        message = "ライセンスが登録されていません。\n[表示] タブの [効率化ペイン表示] から設定してください。";
                        icon = System.Windows.Forms.MessageBoxIcon.Information;
                        break;

                    default:
                        icon = System.Windows.Forms.MessageBoxIcon.Warning;
                        break;
                }

                // 非同期でメッセージボックス表示（UIスレッドで実行）
                System.Windows.Forms.Application.DoEvents();
                Task.Run(() =>
                {
                    System.Threading.Thread.Sleep(2000); // 少し遅延させて起動を妨げない
                    this.Invoke((Action)(() =>
                    {
                        System.Windows.Forms.MessageBox.Show(
                            message, title,
                            System.Windows.Forms.MessageBoxButtons.OK,
                            icon);
                    }));
                });
            }
            catch (Exception ex)
            {
                logger.Error(ex, "Failed to show license notification");
            }
        }

        /// <summary>
        /// ライセンスエラー通知（新規追加）
        /// </summary>
        private void ShowLicenseErrorNotification()
        {
            try
            {
                var message = "ライセンス認証システムの初期化中にエラーが発生しました。\n" +
                             "基本機能は引き続き利用可能ですが、一部機能が制限される場合があります。";

                Task.Run(() =>
                {
                    System.Threading.Thread.Sleep(3000);
                    this.Invoke((Action)(() =>
                    {
                        System.Windows.Forms.MessageBox.Show(
                            message,
                            "ライセンス初期化エラー",
                            System.Windows.Forms.MessageBoxButtons.OK,
                            System.Windows.Forms.MessageBoxIcon.Warning);
                    }));
                });
            }
            catch (Exception ex)
            {
                logger.Error(ex, "Failed to show license error notification");
            }
        }

        /// <summary>
        /// タスクペインのライセンス状態を更新（新規追加）
        /// </summary>
        private void UpdateTaskPaneLicenseStatus()
        {
            try
            {
                if (taskPaneManager?.TaskPaneControl != null)
                {
                    // タスクペインコントロールにライセンス状態を通知
                    // ※CustomTaskPaneUserControl にメソッド追加が必要
                    var control = taskPaneManager.TaskPaneControl as UI.CustomTaskPaneUserControl;
                    if (control != null)
                    {
                        control.UpdateLicenseStatus(
                            licenseManager.CurrentStatus.AccessLevel,
                            licenseManager.GetStatusMessage()
                        );
                    }
                }
            }
            catch (Exception ex)
            {
                logger.Error(ex, "Failed to update task pane license status");
            }
        }

        /// <summary>
        /// 機能実行前のライセンスチェック
        /// </summary>
        public bool CheckFeatureAccess(string featureName, int objectCount = 0)
        {
            try
            {
                // 開発モードまたはライセンスマネージャー未初期化の場合は許可
                if (licenseManager == null || LicenseManager.DevelopmentMode)
                {
                    return true;
                }

                // 機能の利用可否チェック
                if (!licenseManager.IsFeatureAllowed(featureName))
                {
                    ShowFeatureRestrictedMessage(featureName);
                    return false;
                }

                // オブジェクト数制限チェック（制限モードの場合）
                if (objectCount > 0 && !licenseManager.IsWithinObjectLimit(objectCount))
                {
                    ShowObjectLimitMessage(objectCount);
                    return false;
                }

                return true;
            }
            catch (Exception ex)
            {
                logger.Error(ex, $"Error checking feature access for {featureName}");
                return true; // エラー時は許可（UXを優先）
            }
        }

        /// <summary>
        /// 機能制限メッセージ表示（新規追加）
        /// </summary>
        private void ShowFeatureRestrictedMessage(string featureName)
        {
            var requiredLevel = licenseManager.GetRequiredLevel(featureName);
            var currentLevel = licenseManager.CurrentStatus?.AccessLevel ?? FeatureAccessLevel.Blocked;

            var message = $"「{featureName}」機能は「{requiredLevel.GetDisplayName()}」プラン以上で利用可能です。\n" +
                          $"現在のプラン: {currentLevel.GetDisplayName()}";

            MessageBox.Show(message, "機能制限",
                MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        /// <summary>
        /// オブジェクト数制限メッセージ（新規追加）
        /// </summary>
        private void ShowObjectLimitMessage(int objectCount)
        {
            System.Windows.Forms.MessageBox.Show(
                $"制限モードでは最大10個のオブジェクトまで処理できます。\n選択されているオブジェクト数: {objectCount}",
                "処理数制限",
                System.Windows.Forms.MessageBoxButtons.OK,
                System.Windows.Forms.MessageBoxIcon.Information);
        }

        /// <summary>
        /// 内部ステータス取得（修正版）
        /// </summary>
        public string GetInternalStatus()
        {
            try
            {
                var info = new StringBuilder();
                info.AppendLine("=== PowerPoint Efficiency Addin Status ===");

                // 既存のステータス情報
                info.AppendLine($"TaskPaneManager: {(taskPaneManager != null ? "Initialized" : "Not initialized")}");
                info.AppendLine($"PowerToolService: {(powerToolService != null ? "Initialized" : "Not initialized")}");
                info.AppendLine($"AlignmentService: {(alignmentService != null ? "Initialized" : "Not initialized")}");
                info.AppendLine($"TextFormatService: {(textFormatService != null ? "Initialized" : "Not initialized")}");
                info.AppendLine($"ShapeService: {(shapeService != null ? "Initialized" : "Not initialized")}");
                info.AppendLine($"ApplicationProvider: {applicationProvider?.GetType().Name ?? "Not initialized"}");

                // ★★★ ライセンス情報を追加（新規追加）★★★
                info.AppendLine("=== License Status ===");
                if (licenseManager != null)
                {
                    info.AppendLine($"License Mode: {(LicenseManager.DevelopmentMode ? "Development" : "Production")}");
                    info.AppendLine($"License Valid: {isLicenseValid}");
                    info.AppendLine($"Access Level: {licenseManager.CurrentStatus.AccessLevel}");
                    info.AppendLine($"Plan Type: {licenseManager.CurrentStatus.PlanType ?? "None"}");
                    info.AppendLine($"Status: {licenseManager.GetStatusMessage()}");
                }
                else
                {
                    info.AppendLine("License Manager: Not initialized");
                }

                return info.ToString();
            }
            catch (Exception ex)
            {
                return $"Status retrieval error: {ex.Message}";
            }
        }

        // Invokeヘルパーメソッド（UIスレッドでの実行用）
        private void Invoke(Action action)
        {
            try
            {
                if (System.Windows.Forms.Application.OpenForms.Count > 0)
                {
                    System.Windows.Forms.Application.OpenForms[0].Invoke(action);
                }
                else
                {
                    action();
                }
            }
            catch (Exception ex)
            {
                logger.Error(ex, "Failed to invoke action on UI thread");
            }
        }

        /// <summary>
        /// カスタムタスクペインを初期化します
        /// </summary>
        private void InitializeCustomTaskPane()
        {
            try
            {
                logger.Info("Initializing custom task pane");

                // Step 1: TaskPaneManagerを作成
                logger.Debug("Creating TaskPaneManager");
                taskPaneManager = new TaskPaneManager();
                logger.Debug("TaskPaneManager created successfully");

                // Step 2: カスタムタスクペインを初期化
                logger.Debug("Initializing TaskPane");
                taskPaneManager.Initialize();
                logger.Debug("TaskPane initialization completed");

                // Step 3: 初期状態の復元
                logger.Debug("Restoring TaskPane state");
                taskPaneManager.RestoreTaskPaneState();
                logger.Debug("TaskPane state restoration completed");

                logger.Info("Custom task pane initialized successfully");
            }
            catch (Exception ex)
            {
                logger.Error(ex, "Failed to initialize custom task pane");

                // TaskPaneManagerのクリーンアップ
                if (taskPaneManager != null)
                {
                    try
                    {
                        taskPaneManager.Dispose();
                        taskPaneManager = null;
                    }
                    catch (Exception cleanupEx)
                    {
                        logger.Warn(cleanupEx, "Failed to cleanup TaskPaneManager during error handling");
                    }
                }

                // タスクペイン初期化に失敗してもアドイン全体は動作させる
                ErrorHandler.ExecuteSafely(() =>
                {
                    throw new InvalidOperationException($"効率化ペインの初期化に失敗しました: {ex.Message}。一部機能が利用できない場合があります。");
                }, "カスタムペイン初期化", false);
            }
        }

        /// <summary>
        /// 複数インスタンス対応基盤の初期化
        /// </summary>
        private void InitializeMultiInstanceSupport()
        {
            try
            {
                logger.Info("Initializing multi-instance support infrastructure");

                // ApplicationContextManager初期化
                applicationContextManager = new ApplicationContextManager();

                // DI用ApplicationProvider初期化
                applicationProvider = new MultiInstanceApplicationProvider(applicationContextManager);

                logger.Info("Multi-instance support infrastructure initialized successfully");
            }
            catch (Exception ex)
            {
                logger.Error(ex, "Failed to initialize multi-instance support, using fallback");

                // フォールバック：デフォルト実装
                applicationProvider = new DefaultApplicationProvider();
            }
        }

        /// <summary>
        /// DI対応サービス初期化
        /// </summary>
        private void InitializeDIServices()
        {
            try
            {
                logger.Info("Initializing DI-enabled services");

                // PowerToolService（DI対応）
                powerToolService = new PowerToolService(applicationProvider);

                // AlignmentService（DI対応）
                alignmentService = new AlignmentService(applicationProvider);

                // TextFormatService（DI対応）
                textFormatService = new TextFormatService(applicationProvider);

                // ShapeService（DI対応）
                shapeService = new ShapeService(applicationProvider);

                logger.Info("DI-enabled services initialized successfully");
            }
            catch (Exception ex)
            {
                logger.Error(ex, "Failed to initialize DI services, using default constructors");

                // フォールバック：デフォルトコンストラクタ
                powerToolService = new PowerToolService();
                alignmentService = new AlignmentService();
                textFormatService = new TextFormatService();
                shapeService = new ShapeService();
            }
        }

        /// <summary>
        /// カスタムタスクペイン初期化（DI対応サービス注入）
        /// </summary>
        private void InitializeCustomTaskPaneWithDI()
        {
            try
            {
                logger.Info("Initializing custom task pane with DI services");

                // TaskPaneManager初期化
                taskPaneManager = new TaskPaneManager();

                // DI対応サービス注入
                taskPaneManager.InjectAllServices(powerToolService, alignmentService, textFormatService, shapeService);

                // 初期化実行
                taskPaneManager.Initialize();
                taskPaneManager.RestoreTaskPaneState();

                logger.Info("Custom task pane with DI services initialized successfully");
            }
            catch (Exception ex)
            {
                logger.Error(ex, "Failed to initialize custom task pane with DI");
                CleanupTaskPaneManager();

                ErrorHandler.ExecuteSafely(() =>
                {
                    throw new InvalidOperationException($"効率化ペインの初期化に失敗しました: {ex.Message}。一部機能が利用できない場合があります。");
                }, "カスタムペイン初期化", false);
            }
        }


        /// <summary>
        /// アドイン終了時の処理（商用レベルクリーンアップ）
        /// </summary>
        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            try
            {
                logger.Info("Commercial PowerPoint Efficiency Addin shutdown initiated");

                // アプリケーションイベント購読解除
                UnsubscribeFromApplicationEvents();

                // カスタムタスクペインクリーンアップ
                CleanupTaskPaneManager();

                // DI管理オブジェクトのクリーンアップ
                CleanupDIInfrastructure();

                CleanupMultiWindowTaskPaneManagers();

                // 更新の適用
                try
                {
                    var updateService = UpdateService.Instance;
                    if (updateService.HasPendingUpdate())
                    {
                        logger.Info("Applying pending update on shutdown");
                        updateService.ApplyPendingUpdate();
                    }
                }
                catch (Exception ex)
                {
                    logger.Error(ex, "Failed to apply update on shutdown");
                    // エラーが発生しても終了処理を続行
                }

                // ライセンスマネージャーのクリーンアップ
                if (licenseManager != null)
                {
                    licenseManager.Dispose();
                    licenseManager = null;
                    logger.Debug("License manager disposed");
                }


                // COMオブジェクト強制クリーンアップ
                ComHelper.ForceGarbageCollection();

                logger.Info("Commercial PowerPoint Efficiency Addin shutdown completed successfully");
            }
            catch (Exception ex)
            {
                logger.Error(ex, "Error during commercial addin shutdown");
            }
        }

        /// <summary>
        /// 複数ウィンドウ対応TaskPaneManagerのクリーンアップ
        /// </summary>
        private void CleanupMultiWindowTaskPaneManagers()
        {
            try
            {
                logger.Info("Cleaning up multi-window TaskPaneManagers");

                foreach (var kvp in windowTaskPaneManagers.ToList())
                {
                    try
                    {
                        kvp.Value?.Dispose();
                        logger.Debug($"Disposed TaskPaneManager for window: {kvp.Key}");
                    }
                    catch (Exception ex)
                    {
                        logger.Warn(ex, $"Failed to dispose TaskPaneManager for window: {kvp.Key}");
                    }
                }

                windowTaskPaneManagers.Clear();
                logger.Info("Multi-window TaskPaneManagers cleanup completed");
            }
            catch (Exception ex)
            {
                logger.Error(ex, "Error during multi-window TaskPaneManagers cleanup");
            }
        }

        /// <summary>
        /// DI基盤のクリーンアップ
        /// </summary>
        private void CleanupDIInfrastructure()
        {
            try
            {
                logger.Debug("Cleaning up DI infrastructure");

                // ApplicationContextManager破棄
                if (applicationContextManager != null)
                {
                    applicationContextManager.Dispose();
                    applicationContextManager = null;
                }

                // その他のDI管理オブジェクトクリーンアップ
                applicationProvider = null;
                powerToolService = null;
                alignmentService = null;
                shapeService = null;
                textFormatService = null;

                logger.Debug("DI infrastructure cleanup completed");
            }
            catch (Exception ex)
            {
                logger.Error(ex, "Error cleaning up DI infrastructure");
            }
        }

        /// <summary>
        /// TaskPaneManagerクリーンアップ
        /// </summary>
        private void CleanupTaskPaneManager()
        {
            try
            {
                if (taskPaneManager != null)
                {
                    taskPaneManager.Dispose();
                    taskPaneManager = null;
                }
            }
            catch (Exception ex)
            {
                logger.Error(ex, "Error during custom task pane cleanup");
            }
        }

        /// <summary>
        /// 起動失敗時のハンドリング
        /// </summary>
        private void HandleStartupFailure(Exception ex)
        {
            try
            {
                System.Windows.Forms.MessageBox.Show(
                    "PowerPoint効率化アドインの初期化中にエラーが発生しましたが、基本機能は利用可能です。\n\n" +
                    "複数PowerPointウィンドウ使用時は、各ウィンドウで個別にアドインをご利用ください。",
                    "アドイン初期化警告",
                    System.Windows.Forms.MessageBoxButtons.OK,
                    System.Windows.Forms.MessageBoxIcon.Warning
                );
            }
            catch
            {
                // MessageBox失敗時も継続
            }
        }

        /// <summary>
        /// 商用レベル健全性チェック
        /// </summary>
        public bool PerformCommercialHealthCheck()
        {
            try
            {
                // 基本健全性チェック
                if (Application == null) return false;

                // 複数インスタンス対応の健全性チェック
                if (applicationContextManager != null)
                {
                    var currentApp = applicationContextManager.CurrentApplication;
                    if (currentApp == null) return false;

                    // アクティブアプリケーションの有効性確認
                    try
                    {
                        var _ = currentApp.Version;
                    }
                    catch
                    {
                        return false;
                    }
                }

                // DI対応サービスの健全性チェック
                if (powerToolService == null || alignmentService == null || textFormatService == null || shapeService == null)
                {
                    return false;
                }

                return true;
            }
            catch (Exception ex)
            {
                logger.Error(ex, "Commercial health check failed");
                return false;
            }
        }

        /// <summary>
        /// 複数インスタンス対応状態の取得（商用デバッグ用）
        /// </summary>
        public string GetMultiInstanceStatus()
        {
            try
            {
                var info = new System.Text.StringBuilder();
                info.AppendLine("=== 商用版複数インスタンス対応状態 ===");

                // ApplicationContextManager状態
                if (applicationContextManager != null)
                {
                    info.AppendLine("ApplicationContextManager: Active");
                    info.AppendLine(applicationContextManager.GetDebugInfo());
                }
                else
                {
                    info.AppendLine("ApplicationContextManager: Not initialized");
                }

                // DI状態
                info.AppendLine($"PowerToolService: {(powerToolService != null ? "Initialized" : "Not initialized")}");
                info.AppendLine($"AlignmentService: {(alignmentService != null ? "Initialized" : "Not initialized")}");
                info.AppendLine($"TextFormatService: {(textFormatService != null ? "Initialized" : "Not initialized")}");
                info.AppendLine($"shapeService: {(shapeService != null ? "Initialized" : "Not initialized")}");
                info.AppendLine($"ApplicationProvider: {applicationProvider?.GetType().Name ?? "Not initialized"}");

                return info.ToString();
            }
            catch (Exception ex)
            {
                return $"Status retrieval error: {ex.Message}";
            }
        }

        /// <summary>
        /// カスタムタスクペインの終了処理を行います
        /// </summary>
        private void CleanupCustomTaskPane()
        {
            try
            {
                logger.Info("Cleaning up custom task pane");

                if (taskPaneManager != null)
                {
                    taskPaneManager.Dispose();
                    taskPaneManager = null;
                }

                logger.Debug("Custom task pane cleanup completed");
            }
            catch (Exception ex)
            {
                logger.Error(ex, "Error during custom task pane cleanup");
            }
        }

        /// <summary>
        /// アプリケーションイベントを購読します
        /// </summary>
        private void SubscribeToApplicationEvents()
        {
            try
            {
                // プレゼンテーション関連イベント
                Application.PresentationOpen += Application_PresentationOpen;
                Application.PresentationClose += Application_PresentationClose;
                Application.SlideSelectionChanged += Application_SlideSelectionChanged;

                // ウィンドウ関連イベント（複数ウィンドウ対応）
                Application.WindowActivate += Application_WindowActivate;
                Application.WindowDeactivate += Application_WindowDeactivate;

                logger.Debug("Application events subscribed successfully (with window events)");
            }
            catch (Exception ex)
            {
                logger.Error(ex, "Failed to subscribe to application events");
            }
        }

        /// <summary>
        /// アプリケーションイベントの購読を解除します
        /// </summary>
        private void UnsubscribeFromApplicationEvents()
        {
            try
            {
                if (Application != null)
                {
                    Application.PresentationOpen -= Application_PresentationOpen;
                    Application.PresentationClose -= Application_PresentationClose;
                    Application.SlideSelectionChanged -= Application_SlideSelectionChanged;
                    Application.WindowActivate -= Application_WindowActivate;
                    Application.WindowDeactivate -= Application_WindowDeactivate;
                }

                logger.Debug("Application events unsubscribed successfully (with window events)");
            }
            catch (Exception ex)
            {
                logger.Error(ex, "Failed to unsubscribe from application events");
            }
        }

        /// <summary>
        /// ウィンドウがアクティブになった時の処理
        /// </summary>
        private void Application_WindowActivate(PowerPoint.Presentation presentation, PowerPoint.DocumentWindow window)
        {
            try
            {
                logger.Debug($"Window activated: {presentation.Name}");

                // アクティブウィンドウが変更された場合、必要に応じてTaskPaneの状態を更新
                if (applicationContextManager != null)
                {
                    // ApplicationContextManagerに新しいアクティブウィンドウを通知
                    // （既存のCheckActiveWindowメソッドで処理される）
                }
            }
            catch (Exception ex)
            {
                logger.Debug(ex, "Error in WindowActivate event handler");
            }
        }

        /// <summary>
        /// ウィンドウが非アクティブになった時の処理
        /// </summary>
        private void Application_WindowDeactivate(PowerPoint.Presentation presentation, PowerPoint.DocumentWindow window)
        {
            try
            {
                logger.Debug($"Window deactivated: {presentation.Name}");
            }
            catch (Exception ex)
            {
                logger.Debug(ex, "Error in WindowDeactivate event handler");
            }
        }

        /// <summary>
        /// プレゼンテーション開起時のイベントハンドラ
        /// </summary>
        /// <param name="Pres">開かれたプレゼンテーション</param>
        private void Application_PresentationOpen(PowerPoint.Presentation Pres)
        {
            try
            {
                logger.Info($"Presentation opened: {Pres.Name}");

                // プレゼンテーション開起時にタスクペインの健全性チェック
                if (taskPaneManager != null && !taskPaneManager.PerformHealthCheck())
                {
                    logger.Warn("TaskPane health check failed, attempting to reinitialize");

                    try
                    {
                        taskPaneManager.Reinitialize();
                        logger.Info("TaskPane reinitialized successfully");
                    }
                    catch (Exception reinitEx)
                    {
                        logger.Error(reinitEx, "Failed to reinitialize TaskPane");
                    }
                }
            }
            catch (Exception ex)
            {
                logger.Error(ex, "Error in PresentationOpen event handler");
            }
        }

        /// <summary>
        /// プレゼンテーション終了時のイベントハンドラ
        /// </summary>
        /// <param name="Pres">閉じられるプレゼンテーション</param>
        private void Application_PresentationClose(PowerPoint.Presentation Pres)
        {
            try
            {
                logger.Info($"Presentation closing: {Pres.Name}");

                // プレゼンテーション終了時のクリーンアップ
                ComHelper.ForceGarbageCollection();
            }
            catch (Exception ex)
            {
                logger.Error(ex, "Error in PresentationClose event handler");
            }
        }

        /// <summary>
        /// スライド選択変更時のイベントハンドラ
        /// </summary>
        /// <param name="SldRange">選択されたスライド範囲</param>
        private void Application_SlideSelectionChanged(PowerPoint.SlideRange SldRange)
        {
            try
            {
                logger.Debug($"Slide selection changed: {SldRange.Count} slides selected");
            }
            catch (Exception ex)
            {
                logger.Error(ex, "Error in SlideSelectionChanged event handler");
            }
        }

        /// <summary>
        /// 現在アクティブなプレゼンテーションを取得します
        /// </summary>
        /// <returns>アクティブなプレゼンテーション、またはnull</returns>
        public PowerPoint.Presentation GetActivePresentation()
        {
            try
            {
                // 🔧 複数インスタンス対応
                var application = applicationContextManager?.CurrentApplication ?? this.Application;

                if (application.Presentations.Count > 0)
                {
                    return application.ActivePresentation;
                }
            }
            catch (Exception ex)
            {
                logger.Error(ex, "Failed to get active presentation");
            }
            return null;
        }

        /// <summary>
        /// 現在選択されている図形を取得します
        /// </summary>
        /// <returns>選択された図形の配列、またはnull</returns>
        public PowerPoint.ShapeRange GetSelectedShapes()
        {
            try
            {
                // 🔧 複数インスタンス対応：現在アクティブなアプリケーションから取得
                var application = applicationContextManager?.CurrentApplication ?? this.Application;

                var activeWindow = application.ActiveWindow;
                if (activeWindow == null) return null;

                var selection = activeWindow.Selection;

                switch (selection.Type)
                {
                    case PowerPoint.PpSelectionType.ppSelectionShapes:
                        return selection.ShapeRange;

                    case PowerPoint.PpSelectionType.ppSelectionText:
                        try
                        {
                            var textModeShapeRange = selection.ShapeRange;
                            if (textModeShapeRange?.Count > 0)
                            {
                                return textModeShapeRange;
                            }
                        }
                        catch (Exception ex)
                        {
                            logger.Debug(ex, "Failed to get ShapeRange from text editing mode");
                        }
                        break;
                }
            }
            catch (Exception ex)
            {
                logger.Error(ex, "Failed to get selected shapes");
            }
            return null;
        }

        /// <summary>
        /// アドインのバージョン情報を取得します
        /// </summary>
        /// <returns>バージョン文字列</returns>
        public string GetVersion()
        {
            try
            {
                var assembly = System.Reflection.Assembly.GetExecutingAssembly();
                var version = assembly.GetName().Version;
                return $"{version.Major}.{version.Minor}.{version.Build}";
            }
            catch
            {
                return "1.0.0";
            }
        }

        /// <summary>
        /// アドインの健全性チェックを実行します
        /// </summary>
        /// <returns>健全性チェック結果</returns>
        public bool PerformHealthCheck()
        {
            try
            {
                // PowerPointアプリケーションの確認
                if (Application == null)
                {
                    logger.Error("PowerPoint Application is null");
                    return false;
                }

                // アクティブウィンドウの確認
                if (Application.ActiveWindow == null)
                {
                    logger.Warn("No active window available");
                    return false;
                }

                // タスクペインの健全性チェック
                if (taskPaneManager != null && !taskPaneManager.PerformHealthCheck())
                {
                    logger.Warn("TaskPane health check failed");
                    // タスクペインの問題は全体の健全性を損なわない
                }

                // キーボードフックの健全性チェック（一旦無効化）
                /*
                if (keyboardHookManager != null && !keyboardHookManager.IsHookHealthy())
                {
                    logger.Warn("Keyboard hook health check failed, attempting to reinstall");
                    try
                    {
                        keyboardHookManager.ReinstallHook();
                    }
                    catch (Exception ex)
                    {
                        logger.Error(ex, "Failed to reinstall keyboard hook");
                    }
                }
                */

                logger.Debug("Health check passed");
                return true;
            }
            catch (Exception ex)
            {
                logger.Error(ex, "Health check failed");
                return false;
            }
        }

        /// <summary>
        /// カスタムタスクペインの表示/非表示を切り替えます（外部呼び出し用）
        /// </summary>
        public void ToggleTaskPane()
        {
            try
            {
                logger.Info("*** ToggleTaskPane called (Multi-Window) ***");

                var currentTaskPaneManager = GetOrCreateTaskPaneManagerForCurrentWindow();
                if (currentTaskPaneManager != null)
                {
                    currentTaskPaneManager.ToggleVisibility();
                    logger.Info($"*** TaskPane toggle completed, current visibility: {currentTaskPaneManager.IsVisible} ***");
                }
                else
                {
                    logger.Error("Failed to get TaskPaneManager for current window");
                    throw new InvalidOperationException("効率化ペインの初期化に失敗しました。");
                }
            }
            catch (Exception ex)
            {
                logger.Error(ex, "Failed to toggle task pane externally");
                ErrorHandler.ExecuteSafely(() =>
                {
                    throw new InvalidOperationException("効率化ペインの表示切り替えに失敗しました。");
                }, "ペイン表示切り替え");
            }
        }

        /// <summary>
        /// カスタムタスクペインを強制表示します（VBA用・最もシンプル）
        /// </summary>
        public void ShowPanel()
        {
            try
            {
                logger.Info("*** ShowPanel called (Multi-Window) ***");

                var currentTaskPaneManager = GetOrCreateTaskPaneManagerForCurrentWindow();
                if (currentTaskPaneManager != null)
                {
                    currentTaskPaneManager.Show();
                    logger.Info($"*** Panel shown, visibility: {currentTaskPaneManager.IsVisible} ***");
                }
                else
                {
                    logger.Error("Failed to get TaskPaneManager for current window");
                }
            }
            catch (Exception ex)
            {
                logger.Error(ex, "Error in ShowPanel");
                throw new System.Runtime.InteropServices.COMException($"ShowPanel failed: {ex.Message}");
            }
        }

        /// <summary>
        /// 現在アクティブなウィンドウ用のTaskPaneManagerを取得または作成
        /// </summary>
        /// <returns>TaskPaneManager</returns>
        private TaskPaneManager GetOrCreateTaskPaneManagerForCurrentWindow()
        {
            try
            {
                // 現在アクティブなウィンドウ識別子を取得
                string windowKey = GetCurrentWindowKey();

                if (string.IsNullOrEmpty(windowKey))
                {
                    logger.Warn("Could not get current window key, using fallback");
                    // フォールバック：従来の単一TaskPaneManager
                    if (taskPaneManager == null)
                    {
                        InitializeCustomTaskPane();
                    }
                    return taskPaneManager;
                }

                // 既存のTaskPaneManagerをチェック
                if (windowTaskPaneManagers.ContainsKey(windowKey))
                {
                    var existingManager = windowTaskPaneManagers[windowKey];
                    if (existingManager != null && !existingManager.IsDisposed)
                    {
                        logger.Debug($"Using existing TaskPaneManager for window: {windowKey}");
                        return existingManager;
                    }
                    else
                    {
                        // 無効なManagerを削除
                        windowTaskPaneManagers.Remove(windowKey);
                        logger.Debug($"Removed invalid TaskPaneManager for window: {windowKey}");
                    }
                }

                // 新しいTaskPaneManagerを作成
                logger.Info($"Creating new TaskPaneManager for window: {windowKey}");
                var newManager = CreateTaskPaneManagerForWindow(windowKey);

                if (newManager != null)
                {
                    windowTaskPaneManagers[windowKey] = newManager;

                    // 後方互換性：最初のTaskPaneManagerをtaskPaneManagerにも設定
                    if (taskPaneManager == null)
                    {
                        taskPaneManager = newManager;
                    }

                    logger.Info($"TaskPaneManager created and cached for window: {windowKey}");
                    return newManager;
                }

                return null;
            }
            catch (Exception ex)
            {
                logger.Error(ex, "Failed to get or create TaskPaneManager for current window");

                // フォールバック：従来の単一TaskPaneManager
                if (taskPaneManager == null)
                {
                    InitializeCustomTaskPane();
                }
                return taskPaneManager;
            }
        }

        /// <summary>
        /// 現在アクティブなウィンドウのキーを取得
        /// </summary>
        /// <returns>ウィンドウキー</returns>
        private string GetCurrentWindowKey()
        {
            try
            {
                var currentApp = applicationContextManager?.CurrentApplication ?? Application;
                if (currentApp?.ActiveWindow != null)
                {
                    // ウィンドウキーとして一意な識別子を生成
                    // Caption + ハンドル or プレゼンテーションID等を組み合わせ
                    var windowCaption = currentApp.ActiveWindow.Caption ?? "Unknown";
                    var presentationName = currentApp.ActiveWindow.Presentation?.Name ?? "NoPresentation";
                    var windowKey = $"{windowCaption}_{presentationName}_{currentApp.ActiveWindow.GetHashCode()}";

                    logger.Debug($"Generated window key: {windowKey}");
                    return windowKey;
                }
            }
            catch (Exception ex)
            {
                logger.Debug(ex, "Failed to get current window key");
            }

            return null;
        }

        /// <summary>
        /// 指定ウィンドウ用のTaskPaneManagerを作成
        /// </summary>
        /// <param name="windowKey">ウィンドウキー</param>
        /// <returns>TaskPaneManager</returns>
        private TaskPaneManager CreateTaskPaneManagerForWindow(string windowKey)
        {
            try
            {
                logger.Debug($"Creating TaskPaneManager for window: {windowKey}");

                // 新しいTaskPaneManagerを作成
                var newManager = new TaskPaneManager(applicationProvider);

                // DI対応サービス注入
                newManager.InjectAllServices(powerToolService, alignmentService, textFormatService, shapeService);

                // 初期化実行
                newManager.Initialize();

                logger.Info($"TaskPaneManager successfully created for window: {windowKey}");
                return newManager;
            }
            catch (Exception ex)
            {
                logger.Error(ex, $"Failed to create TaskPaneManager for window: {windowKey}");
                return null;
            }
        }

        /// <summary>
        /// カスタムタスクペインを非表示にします（VBA用）
        /// </summary>
        public void HidePanel()
        {
            try
            {
                logger.Info("*** HidePanel called (Multi-Window) ***");

                var currentTaskPaneManager = GetOrCreateTaskPaneManagerForCurrentWindow();
                if (currentTaskPaneManager != null)
                {
                    currentTaskPaneManager.Hide();
                    logger.Info($"*** Panel hidden, visibility: {currentTaskPaneManager.IsVisible} ***");
                }
                else
                {
                    logger.Warn("TaskPaneManager not found for current window");
                }
            }
            catch (Exception ex)
            {
                logger.Error(ex, "Error in HidePanel");
                throw new System.Runtime.InteropServices.COMException($"HidePanel failed: {ex.Message}");
            }
        }

        /// <summary>
        /// カスタムタスクペインの状態を取得します（VBA用）
        /// </summary>
        public string GetPanelStatus()
        {
            try
            {
                if (taskPaneManager == null)
                {
                    return "TaskPaneManager: Not initialized";
                }

                var status = $"TaskPaneManager: Initialized\n" +
                           $"Panel Visible: {taskPaneManager.IsVisible}\n" +
                           $"Panel Width: {taskPaneManager.Width}\n" +
                           $"Panel Height: {taskPaneManager.Height}";

                logger.Info($"Panel status requested: {status}");
                return status;
            }
            catch (Exception ex)
            {
                logger.Error(ex, "Error getting panel status");
                return $"Error: {ex.Message}";
            }
        }

        #region VSTO generated code

        /// <summary>
        /// デザイナー サポートに必要なメソッドです。このメソッドの内容を
        /// コード エディターで変更しないでください。
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }

        /// <summary>
        /// Ribbon拡張オブジェクトを作成します（最小リボン版）
        /// </summary>
        /// <returns>Ribbon拡張オブジェクト</returns>
        protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            try
            {
                logger.Info("Creating ViewRibbon extensibility object");
                var viewRibbon = new Ribbon.ViewRibbon();
                logger.Info("ViewRibbon extensibility object created successfully");
                return viewRibbon;
            }
            catch (Exception ex)
            {
                logger.Error(ex, "Failed to create ViewRibbon extensibility object");

                // フォールバック：最低限のリボンを作成
                return new Ribbon.ViewRibbon();
            }
        }

        #endregion
    }
}