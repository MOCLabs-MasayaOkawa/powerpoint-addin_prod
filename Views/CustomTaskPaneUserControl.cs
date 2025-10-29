using Microsoft.Office.Core;
using NLog;
using PowerPointEfficiencyAddin.Models;
using PowerPointEfficiencyAddin.Models.Licensing;
using PowerPointEfficiencyAddin.Services.Core.Alignment;
using PowerPointEfficiencyAddin.Services.Core.BuiltIn;
using PowerPointEfficiencyAddin.Services.Core.Image;
using PowerPointEfficiencyAddin.Services.Core.Matrix;
using PowerPointEfficiencyAddin.Services.Core.PowerTool;
using PowerPointEfficiencyAddin.Services.Core.Selection;
using PowerPointEfficiencyAddin.Services.Core.Shape;
using PowerPointEfficiencyAddin.Services.Core.Table;
using PowerPointEfficiencyAddin.Services.Core.Text;
using PowerPointEfficiencyAddin.Services.Infrastructure.Licensing;
using PowerPointEfficiencyAddin.Services.Infrastructure.MultiInstance;
using PowerPointEfficiencyAddin.Services.UI.Dialogs;
using PowerPointEfficiencyAddin.Utils;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PowerPointEfficiencyAddin.UI
{
    /// <summary>
    /// カスタムタスクペインのユーザーコントロール（PDF配置表対応版）
    /// </summary>
    public partial class CustomTaskPaneUserControl : UserControl
    {
        private static readonly Logger logger = LogManager.GetCurrentClassLogger();

        //ライセンス管理
        private Label lblLicenseStatus;
        private Panel licenseStatusPanel;
        private Button btnShowLicenseSettings;

        // UI コンポーネント
        private Panel mainPanel;
        private ToolTip toolTip;

        // サービスクラス
        private ShapeService shapeService;
        private AlignmentService alignmentService;
        private PowerToolService powerToolService;
        private TextFormatService textFormatService;
        private ImageCompressionService imageCompressionService;
        private ShapeSelectionService shapeSelectionService;
        private BuiltInShapeService builtInShapeService;
        private MatrixOperationService matrixOperationService;
        private TableConversionService tableConversionService;

        // 機能定義
        private List<FunctionItem> allFunctions;
        private Dictionary<FunctionCategory, List<FunctionItem>> categorizedFunctions;

        private System.Windows.Forms.Timer statusUpdateTimer;

        public CustomTaskPaneUserControl()
        {
            try
            {
                logger.Info("Initializing CustomTaskPaneUserControl");

                InitializeComponent();
                logger.Debug("InitializeComponent completed");

                InitializeServices();
                logger.Debug("InitializeServices completed");

                InitializeCustomComponents();
                logger.Debug("InitializeCustomComponents completed");

                // 【追加】ライセンスパネルを初期化
                InitializeLicenseStatusPanel();

                InitializeFunctions();
                logger.Debug("InitializeFunctions completed");

                CreateUI();
                logger.Debug("CreateUI completed");

                logger.Info("CustomTaskPaneUserControl initialization completed successfully");
            }
            catch (Exception ex)
            {
                logger.Fatal(ex, "Critical error during CustomTaskPaneUserControl initialization");

                try
                {
                    // 緊急時のフォールバック
                    InitializeComponent();
                    InitializeCustomComponents();
                    CreateMinimalUI();
                    logger.Warn("Fallback minimal UI created");
                }
                catch (Exception fallbackEx)
                {
                    logger.Fatal(fallbackEx, "Failed to create fallback UI");
                }

                ErrorHandler.ExecuteSafely(() =>
                {
                    throw new InvalidOperationException($"カスタムペインの初期化に失敗しました: {ex.Message}");
                }, "カスタムペイン初期化", false);
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

            logger.Info("All DI services injected into CustomTaskPaneUserControl");

            if (allFunctions != null)
            {
                InitializeFunctions(); // サービス参照を更新
                RefreshUI();
            }
        }

        /// <summary>
        /// 🆕 商用デバッグ・ステータス機能
        /// </summary>
        private void AddCommercialDebugFunctions()
        {
            var debugFunctions = new[]
            {
                // 複数インスタンス対応ステータス表示
                new FunctionItem("ShowMultiInstanceStatus", "マルチ状態", "複数PowerPoint対応の状態を表示", "debug_status.png",
                    () => SafeExecuteFunction(() => ShowMultiInstanceStatus(), "複数インスタンス状態表示"),
                    FunctionCategory.PowerTool, 3, 6),
                
                // 商用健全性チェック
                new FunctionItem("PerformHealthCheck", "健全性チェック", "アドインの動作状態をチェック", "health_check.png",
                    () => SafeExecuteFunction(() => PerformCommercialHealthCheck(), "健全性チェック"),
                    FunctionCategory.PowerTool, 3, 7),
                
                // ApplicationContext手動切替（高度な用途）
                new FunctionItem("RefreshApplicationContext", "コンテキスト更新", "アクティブPowerPointコンテキストを更新", "refresh_context.png",
                    () => SafeExecuteFunction(() => RefreshApplicationContext(), "コンテキスト更新"),
                    FunctionCategory.PowerTool, 3, 8),
            };

            allFunctions.AddRange(debugFunctions);
        }

        /// <summary>
        /// 複数インスタンス対応状態表示
        /// </summary>
        private void ShowMultiInstanceStatus()
        {
            try
            {
                var statusInfo = Globals.ThisAddIn.GetMultiInstanceStatus();

                var statusForm = new Form
                {
                    Text = "複数PowerPoint対応状態",
                    Size = new System.Drawing.Size(700, 600),
                    StartPosition = FormStartPosition.CenterParent
                };

                var textBox = new TextBox
                {
                    Text = statusInfo,
                    Multiline = true,
                    ScrollBars = ScrollBars.Both,
                    Dock = DockStyle.Fill,
                    ReadOnly = true,
                    Font = new System.Drawing.Font("Consolas", 9)
                };

                var buttonPanel = new Panel
                {
                    Height = 40,
                    Dock = DockStyle.Bottom
                };

                var refreshButton = new Button
                {
                    Text = "最新状態に更新",
                    Size = new System.Drawing.Size(120, 30),
                    Anchor = AnchorStyles.Right | AnchorStyles.Top
                };
                refreshButton.Location = new System.Drawing.Point(buttonPanel.Width - 130, 5);
                refreshButton.Click += (s, e) =>
                {
                    textBox.Text = Globals.ThisAddIn.GetMultiInstanceStatus();
                };

                buttonPanel.Controls.Add(refreshButton);
                statusForm.Controls.Add(textBox);
                statusForm.Controls.Add(buttonPanel);

                statusForm.ShowDialog();
            }
            catch (Exception ex)
            {
                logger.Error(ex, "Failed to show multi-instance status");
                MessageBox.Show($"ステータス表示エラー: {ex.Message}", "エラー",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// 商用健全性チェック実行
        /// </summary>
        private void PerformCommercialHealthCheck()
        {
            try
            {
                var isHealthy = Globals.ThisAddIn.PerformCommercialHealthCheck();
                var powerPointCount = System.Diagnostics.Process.GetProcessesByName("POWERPNT").Length;

                var message = new System.Text.StringBuilder();
                message.AppendLine($"=== 商用版アドイン健全性チェック ===");
                message.AppendLine($"総合状態: {(isHealthy ? "✅ 正常" : "⚠️ 要注意")}");
                message.AppendLine($"PowerPointプロセス数: {powerPointCount}");
                message.AppendLine($"複数インスタンス対応: {(powerPointCount > 1 ? "有効" : "単一インスタンス")}");

                if (!isHealthy)
                {
                    message.AppendLine();
                    message.AppendLine("推奨対処:");
                    message.AppendLine("• アドインパネルを開き直してください");
                    message.AppendLine("• 使用予定のPowerPointをアクティブにしてください");
                }

                MessageBox.Show(message.ToString(), "健全性チェック結果",
                    MessageBoxButtons.OK,
                    isHealthy ? MessageBoxIcon.Information : MessageBoxIcon.Warning);
            }
            catch (Exception ex)
            {
                logger.Error(ex, "Failed to perform commercial health check");
                MessageBox.Show($"健全性チェックエラー: {ex.Message}", "エラー",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// ApplicationContext手動更新
        /// </summary>
        private void RefreshApplicationContext()
        {
            try
            {
                var contextManager = Globals.ThisAddIn.ApplicationContextManager;
                if (contextManager != null)
                {
                    // 強制的にコンテキスト更新（内部実装詳細に依存）
                    var currentApp = contextManager.CurrentApplication;
                    logger.Info($"Application context refreshed: {currentApp?.Version ?? "Unknown"}");

                    MessageBox.Show(
                        "アプリケーションコンテキストを更新しました。\n" +
                        "現在アクティブなPowerPointウィンドウでアドイン機能が動作します。",
                        "コンテキスト更新完了",
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    MessageBox.Show(
                        "ApplicationContextManagerが初期化されていません。\n" +
                        "デフォルトモードで動作しています。",
                        "コンテキスト更新",
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
            catch (Exception ex)
            {
                logger.Error(ex, "Failed to refresh application context");
                MessageBox.Show($"コンテキスト更新エラー: {ex.Message}", "エラー",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// UI更新（DI対応サービス切り替え後）
        /// </summary>
        private void RefreshUI()
        {
            try
            {
                // 機能ボタンの再バインド（DI対応サービス使用）
                foreach (Control control in this.Controls)
                {
                    if (control is Button button && button.Tag is FunctionItem functionItem)
                    {
                        // ボタンアクションを再設定（新しいサービス参照使用）
                        button.Click -= (s, e) => functionItem.Action();
                        button.Click += (s, e) => functionItem.Action();
                    }
                }

                logger.Debug("UI refreshed with new DI service references");
            }
            catch (Exception ex)
            {
                logger.Error(ex, "Failed to refresh UI");
            }
        }

        /// <summary>
        /// 機能初期化（DI対応版）
        /// </summary>
        private void InitializeFunctions()
        {
            try
            {
                allFunctions = new List<FunctionItem>();

                // 既存機能追加（DI対応サービス使用）
                AddSelectionFunctions();      // 選択
                AddTextFunctions();          // テキスト
                AddShapeFunctions();         // 図形
                AddFormatFunctions();        // 整形
                AddGroupingFunctions();      // グループ化
                AddAlignmentFunctions();     // 整列
                AddShapeOperationFunctions(); // 図形操作プロ
                AddTableOperationFunctions(); // 表操作
                AddSpacingFunctions();       // 間隔
                AddPowerToolFunctions();     // PowerTool

                // 🆕 商用デバッグ機能追加
                AddCommercialDebugFunctions();

                // カテゴリ別整理
                categorizedFunctions = allFunctions
                    .GroupBy(f => f.Category)
                    .ToDictionary(g => g.Key, g => g.ToList());

                logger.Info($"Functions initialized with DI support: {allFunctions.Count} functions");
            }
            catch (Exception ex)
            {
                logger.Error(ex, "Failed to initialize functions with DI");
                allFunctions = CreateMinimalFunctionSet(); // フォールバック
            }
        }

        /// <summary>
        /// サービスクラスを初期化します
        /// </summary>
        private void InitializeServices()
        {
            try
            {
                logger.Debug("Initializing services");

                var applicationProvider = new DefaultApplicationProvider();

                shapeService = new ShapeService();
                logger.Debug("ShapeService created");

                alignmentService = new AlignmentService();
                logger.Debug("AlignmentService created");

                powerToolService = new PowerToolService();
                logger.Debug("PowerToolService created");

                textFormatService = new TextFormatService();
                logger.Debug("TextFormatService created");

                imageCompressionService = new ImageCompressionService(applicationProvider);
                logger.Debug("ImageCompressionService created");

                shapeSelectionService = new ShapeSelectionService(applicationProvider);
                logger.Debug("ShapeSelectionService created");

                // 新しいサービスを初期化
                var helper = new PowerToolServiceHelper(applicationProvider);
                
                tableConversionService = new TableConversionService(applicationProvider, helper);
                logger.Debug("TableConversionService created");

                matrixOperationService = new MatrixOperationService(applicationProvider);
                logger.Debug("MatrixOperationService created");

                builtInShapeService = new BuiltInShapeService(applicationProvider, helper, powerToolService);
                logger.Debug("BuiltInShapeService created");

                logger.Info("All services initialized successfully");
            }
            catch (Exception ex)
            {
                logger.Error(ex, "Failed to initialize services");
                throw new InvalidOperationException("サービスクラスの初期化に失敗しました", ex);
            }
        }