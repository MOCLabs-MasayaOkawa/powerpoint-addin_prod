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

        /// <summary>
        /// カスタムコンポーネントを初期化します
        /// </summary>
        private void InitializeCustomComponents()
        {
            try
            {
                logger.Debug("Initializing custom components");

                // このコントロール自体の設定
                this.BackColor = Color.White;
                this.Size = new Size(280, 800);
                this.AutoScroll = true;

                // ToolTip初期化
                toolTip = new ToolTip()
                {
                    AutoPopDelay = 10000,
                    InitialDelay = 500,
                    ReshowDelay = 100
                };

                // メインパネル初期化
                mainPanel = new Panel()
                {
                    AutoSize = true,
                    Dock = DockStyle.Top,
                    BackColor = Color.White
                };

                this.Controls.Add(mainPanel);

                logger.Info("Custom components initialized");
            }
            catch (Exception ex)
            {
                logger.Error(ex, "Failed to initialize custom components");
                throw new InvalidOperationException("カスタムコンポーネントの初期化に失敗しました", ex);
            }
        }

        #region 機能定義メソッド

        /// <summary>
        /// 選択カテゴリの機能を追加
        /// </summary>
        private void AddSelectionFunctions()
        {
            var functions = new[]
            {
                // 1行目
                new FunctionItem("SelectAll", "すべて選択", "スライド上のすべてのオブジェクトを選択します", "select_all_builtin.png",
                    () => SafeExecuteFunction(() => ExecutePowerPointCommand("SelectAll"), "すべて選択"), FunctionCategory.Selection, 1, 0, true),
                new FunctionItem("SelectSameSize", "サイズ一致選択", "最後に選択した図形と同じサイズの図形を選択します", "select_same_size.png",
                    () => SafeExecuteFunction(() => shapeSelectionService.SelectSameSize(), "サイズ一致選択"), FunctionCategory.Selection, 1, 1),
                new FunctionItem("SelectSameColor", "塗り色一致選択", "最後に選択した図形と同じ塗りつぶし色の図形を選択します", "select_same_color.png",
                    () => SafeExecuteFunction(() => shapeSelectionService.SelectSameColor(), "塗り色一致選択"), FunctionCategory.Selection, 1, 2),
                new FunctionItem("SelectByText", "テキスト一致選択", "特定のテキストを含む図形を選択します", "select_by_text.png",
                    () => SafeExecuteFunction(() => shapeSelectionService.SelectByText(), "テキスト一致選択"), FunctionCategory.Selection, 1, 3)
            };

            allFunctions.AddRange(functions);
        }

        /// <summary>
        /// テキストカテゴリの機能を追加
        /// </summary>
        private void AddTextFunctions()
        {
            var functions = new[]
            {
                // 1行目
                new FunctionItem("AdjustTextSize", "テキスト最適化", "選択図形のテキストサイズを段階的に変更（トグル動作）", "text_size_toggle.png",
                    () => SafeExecuteFunction(() => textFormatService.AdjustTextSize(), "テキスト最適化"), FunctionCategory.Text, 1, 0),
                new FunctionItem("AdjustMarginUp", "余白Up", "選択図形の余白を1.5倍に拡大します", "margin_up.png",
                    () => SafeExecuteFunction(() => textFormatService.AdjustMarginUp(), "余白Up"), FunctionCategory.Text, 1, 1),
                new FunctionItem("AdjustMarginDown", "余白Down", "選択図形の余白を1.5で割って縮小します", "margin_down.png",
                    () => SafeExecuteFunction(() => textFormatService.AdjustMarginDown(), "余白Down"), FunctionCategory.Text, 1, 2),
                new FunctionItem("ShowMarginAdjustDialog", "余白調整", "詳細な余白調整ダイアログを表示します", "margin_adjust.png",
                    () => SafeExecuteFunction(() => textFormatService.ShowMarginAdjustDialog(), "余白調整"), FunctionCategory.Text, 1, 3),
                new FunctionItem("TextBox", "Text Box", "テキストボックスを挿入します", "text_box.png",
                    () => SafeExecuteFunction(() => ExecutePowerPointCommand("TextBox"), "Text"), FunctionCategory.Text, 1, 4, true),
                new FunctionItem("ClearTextsFromSelectedShapes", "テキストクリア", "選択した図形のテキスト内容をすべてクリアします", "clear_text.png",
                    () => SafeExecuteFunction(() => textFormatService.ClearTextsFromSelectedShapes(), "テキストクリア"), FunctionCategory.Text, 1, 5)
            };

            allFunctions.AddRange(functions);
        }

        /// <summary>
        /// 図形カテゴリの機能を追加
        /// </summary>
        private void AddShapeFunctions()
        {
            var functions = new[]
            {
                // 1行目
                new FunctionItem("ShapeRectangle", "四角", "四角形を挿入します", "rectangle.png",
                    () => SafeExecuteFunction(() => ExecutePowerPointCommand("ShapeRectangle"), "四角"), FunctionCategory.Shape, 1, 0, true),
                new FunctionItem("ShapeRoundedRectangle", "角丸", "角丸四角形を挿入します", "rounded_rectangle.png",
                    () => SafeExecuteFunction(() => ExecutePowerPointCommand("ShapeRoundedRectangle"), "角丸"), FunctionCategory.Shape, 1, 1, true),
                new FunctionItem("ShapeOval", "丸", "楕円を挿入します", "oval.png",
                    () => SafeExecuteFunction(() => ExecutePowerPointCommand("ShapeOval"), "丸"), FunctionCategory.Shape, 1, 2, true),
                new FunctionItem("ShapeIsoscelesTriangle", "三角", "三角形を挿入します", "triangle.png",
                    () => SafeExecuteFunction(() => ExecutePowerPointCommand("ShapeIsoscelesTriangle"), "三角"), FunctionCategory.Shape, 1, 3, true),
                new FunctionItem("ShapeRectangularCallout", "吹き出し", "吹き出しを挿入します", "callout.png",
                    () => SafeExecuteFunction(() => ExecutePowerPointCommand("ShapeRectangularCallout"), "吹き出し"), FunctionCategory.Shape, 1, 4, true),
                new FunctionItem("ShapeRightArrow", "矢印（右）", "右向き矢印を挿入します", "arrow_right.png",
                    () => SafeExecuteFunction(() => ExecutePowerPointCommand("ShapeRightArrow"), "矢印（右）"), FunctionCategory.Shape, 1, 5, true),
                new FunctionItem("ShapeDownArrow", "矢印（下）", "下向き矢印を挿入します", "arrow_down.png",
                    () => SafeExecuteFunction(() => ExecutePowerPointCommand("ShapeDownArrow"), "矢印（下）"), FunctionCategory.Shape, 1, 6, true),
                new FunctionItem("ShapeLine", "線", "直線を挿入します", "line.png",
                    () => SafeExecuteFunction(() => ExecutePowerPointCommand("ShapeLine"), "線"), FunctionCategory.Shape, 1, 7, true),
                
                // 2行目
                new FunctionItem("ShapeLineArrow", "矢印線", "矢印付き直線を挿入します", "arrow_line.png",
                    () => SafeExecuteFunction(() => ExecutePowerPointCommand("ShapeLineArrow"), "矢印線"), FunctionCategory.Shape, 2, 0,true),

                new FunctionItem("ShapeElbowConnector", "鍵線", "鍵型コネクタを挿入します", "elbow_connector.png",
                    () => SafeExecuteFunction(() => ExecutePowerPointCommand("ShapeElbowConnector"), "鍵線"), FunctionCategory.Shape, 2, 1, true),
                new FunctionItem("ShapeElbowArrowConnector", "鍵線矢印", "矢印付き鍵型コネクタを挿入します", "elbow_arrow_connector.png",
                    () => SafeExecuteFunction(() => ExecutePowerPointCommand("ShapeElbowArrowConnector"), "鍵線矢印"), FunctionCategory.Shape, 2, 2, true),
                new FunctionItem("ShapeLeftBrace", "中括弧", "中括弧を挿入します", "brace.png",
                    () => SafeExecuteFunction(() => ExecutePowerPointCommand("ShapeLeftBrace"), "中括弧"), FunctionCategory.Shape, 2, 3, true),
                new FunctionItem("ShapePentagon", "五角形", "五角形を作成します", "shape_pentagon.png",
                    () => SafeExecuteFunction(() => builtInShapeService.ExecutePowerPointCommand("ShapePentagon"), "五角形作成"), FunctionCategory.Shape, 2, 4),
                new FunctionItem("ShapeChevron", "シェブロン", "シェブロン（V字型）を作成します", "shape_chevron.png",
                    () => SafeExecuteFunction(() => builtInShapeService.ExecutePowerPointCommand("ShapeChevron"), "シェブロン作成"), FunctionCategory.Shape, 2, 5),
                new FunctionItem("ShapeStyleSettings", "図形スタイル設定", "新規作成図形の塗りつぶし色・枠線色・フォント・フォント色を設定します", "shape_style_settings.png",
                    () => SafeExecuteFunction(() => builtInShapeService.ShowShapeStyleDialog(), "図形スタイル設定"), FunctionCategory.Shape, 2, 6)
            };

            allFunctions.AddRange(functions);
        }

        /// <summary>
        /// 整形カテゴリの機能を追加
        /// </summary>
        private void AddFormatFunctions()
        {
            var functions = new[]
            {
                // 1行目

                new FunctionItem("SizeUpToggle", "サイズUp", "選択した図形のサイズを5%大きくします（累積効果）", "size_up.png",
                    () => SafeExecuteFunction(() => shapeService.SizeUpToggle(), "サイズUp"), FunctionCategory.Format, 1, 0),
                new FunctionItem("SizeDownToggle", "サイズDown", "選択した図形のサイズを5%小さくします（累積効果）", "size_down.png",
                    () => SafeExecuteFunction(() => shapeService.SizeDownToggle(), "サイズDown"), FunctionCategory.Format, 1, 1),
                new FunctionItem("LineWeightUpToggle", "枠線の太さUp", "選択した図形の枠線の太さを0.25pt太くします（累積効果）", "line_weight_up.png",
                    () => SafeExecuteFunction(() => shapeService.LineWeightUpToggle(), "枠線の太さUp"), FunctionCategory.Format, 1, 2),
                new FunctionItem("LineWeightDownToggle", "枠線の太さDown", "選択した図形の枠線の太さを0.25pt細くします（累積効果）", "line_weight_down.png",
                    () => SafeExecuteFunction(() => shapeService.LineWeightDownToggle(), "枠線の太さDown"), FunctionCategory.Format, 1, 3),
                new FunctionItem("DashStyleToggle", "枠線の種類変更トグル", "選択した図形の枠線を実線→点線→破線→鎖線の順に変更します", "dash_style.png",
                    () => SafeExecuteFunction(() => shapeService.DashStyleToggle(), "枠線の種類変更トグル"), FunctionCategory.Format, 1, 4),
                new FunctionItem("TransparencyUpToggle", "透過率Upトグル", "選択した図形の透過率を10%ずつ上げます", "transparency_up.png",
                    () => SafeExecuteFunction(() => shapeSelectionService.TransparencyUpToggle(), "透過率Upトグル"), FunctionCategory.Format, 1, 5),
                new FunctionItem("TransparencyDownToggle", "透過率Downトグル", "選択した図形の透過率を10%ずつ下げます", "transparency_down.png",
                    () => SafeExecuteFunction(() => shapeSelectionService.TransparencyDownToggle(), "透過率Downトグル"), FunctionCategory.Format, 1, 6),

                // 2行目
                new FunctionItem("MatchHeight", "縦幅を揃える", "選択した図形の高さを最後に選択した図形に合わせます", "match_height.png",
                    () => SafeExecuteFunction(() => shapeService.MatchHeight(), "縦幅を揃える"), FunctionCategory.Format, 2, 0),
                new FunctionItem("MatchWidth", "横幅を揃える", "選択した図形の幅を最後に選択した図形に合わせます", "match_width.png",
                    () => SafeExecuteFunction(() => shapeService.MatchWidth(), "横幅を揃える"), FunctionCategory.Format, 2, 1),
                new FunctionItem("MatchSize", "横幅縦幅を揃える", "選択した図形の幅と高さを最後に選択した図形に合わせます", "match_size.png",
                    () => SafeExecuteFunction(() => shapeService.MatchSize(), "横幅縦幅を揃える"), FunctionCategory.Format, 2, 2),
                new FunctionItem("MatchLineWeight", "枠線の太さを揃える", "選択した図形の枠線の太さを最後に選択した図形に合わせます", "match_line_weight.png",
                    () => SafeExecuteFunction(() => shapeService.MatchLineWeight(), "枠線の太さを揃える"), FunctionCategory.Format, 2, 3),
                new FunctionItem("MatchLineColor", "枠線の色を揃える", "選択した図形の枠線の色を最後に選択した図形に合わせます", "match_line_color.png",
                    () => SafeExecuteFunction(() => shapeService.MatchLineColor(), "枠線の色を揃える"), FunctionCategory.Format, 2, 4),
                new FunctionItem("MatchFillColor", "塗りつぶし色を揃える", "選択した図形の塗りつぶし色を最後に選択した図形に合わせます", "match_fill_color.png",
                    () => SafeExecuteFunction(() => shapeService.MatchFillColor(), "塗りつぶし色を揃える"), FunctionCategory.Format, 2, 5),
                new FunctionItem("CompressAllImages", "画像圧縮", "スライド上の全ての画像を一括圧縮します", "compress_images.png",
                    () => SafeExecuteFunction(() => imageCompressionService.CompressAllImages(), "画像圧縮"), FunctionCategory.Format, 2, 6)
            };

            allFunctions.AddRange(functions);
        }

        /// <summary>
        /// グループ化カテゴリの機能を追加
        /// </summary>
        private void AddGroupingFunctions()
        {
            var functions = new[]
            {
                // 1行目
                new FunctionItem("Group", "グループ化", "選択した図形をグループ化します", "group_builtin.png",
                    () => SafeExecuteFunction(() => ExecutePowerPointCommand("ObjectsGroup"), "グループ化"), FunctionCategory.Grouping, 1, 0, true),
                new FunctionItem("Ungroup", "グループ解除", "選択したグループを解除します", "ungroup_builtin.png",
                    () => SafeExecuteFunction(() => ExecutePowerPointCommand("ObjectsUngroup"), "グループ解除"), FunctionCategory.Grouping, 1, 1, true),
                new FunctionItem("GroupByRows", "行でグループ化", "選択したオブジェクトを行別にグループ化します", "group_by_rows.png",
                    () => SafeExecuteFunction(() => alignmentService.GroupByRows(), "行でグループ化"), FunctionCategory.Grouping, 1, 2),
                new FunctionItem("GroupByColumns", "列でグループ化", "選択したオブジェクトを列別にグループ化します", "group_by_columns.png",
                    () => SafeExecuteFunction(() => alignmentService.GroupByColumns(), "列でグループ化"), FunctionCategory.Grouping, 1, 3)
            };

            allFunctions.AddRange(functions);
        }

        /// <summary>
        /// 整列カテゴリの機能を追加
        /// </summary>
        private void AddAlignmentFunctions()
        {
            var functions = new[]
            {
                // 1行目
                new FunctionItem("AlignLeft", "左揃え", "選択した図形を左揃えにします", "align_left_builtin.png",
                    () => SafeExecuteFunction(() => ExecutePowerPointCommand("AlignLeft"), "左揃え"), FunctionCategory.Alignment, 1, 0, true),
                new FunctionItem("AlignCenterHorizontal", "中央揃え", "選択した図形を中央揃えにします", "align_center_builtin.png",
                    () => SafeExecuteFunction(() => ExecutePowerPointCommand("AlignCenterHorizontal"), "中央揃え"), FunctionCategory.Alignment, 1, 1, true),
                new FunctionItem("AlignRight", "右揃え", "選択した図形を右揃えにします", "align_right_builtin.png",
                    () => SafeExecuteFunction(() => ExecutePowerPointCommand("AlignRight"), "右揃え"), FunctionCategory.Alignment, 1, 2, true),
                new FunctionItem("AlignTop", "上揃え", "選択した図形を上揃えにします", "align_top_builtin.png",
                    () => SafeExecuteFunction(() => ExecutePowerPointCommand("AlignTop"), "上揃え"), FunctionCategory.Alignment, 1, 3, true),
                new FunctionItem("AlignCenterVertical", "水平揃え", "選択した図形を水平中央揃えにします", "align_middle_builtin.png",
                    () => SafeExecuteFunction(() => ExecutePowerPointCommand("AlignCenterVertical"), "水平揃え"), FunctionCategory.Alignment, 1, 4, true),
                new FunctionItem("AlignBottom", "下揃え", "選択した図形を下揃えにします", "align_bottom_builtin.png",
                    () => SafeExecuteFunction(() => ExecutePowerPointCommand("AlignBottom"), "下揃え"), FunctionCategory.Alignment, 1, 5, true),

                // 2行目
                new FunctionItem("PlaceLeftToRight", "左端を右端へ", "2つの選択した図形の片方の図形の左端を、もう一方の図形の右端に接着させる", "place_left_to_right.png",
                    () => SafeExecuteFunction(() => alignmentService.PlaceLeftToRight(), "左端を右端へ"), FunctionCategory.Alignment, 2, 0),
                new FunctionItem("PlaceRightToLeft", "右端を左端へ", "2つの選択した図形の片方の図形の右端を、もう一方の図形の左端に接着させる", "place_right_to_left.png",
                    () => SafeExecuteFunction(() => alignmentService.PlaceRightToLeft(), "右端を左端へ"), FunctionCategory.Alignment, 2, 1),
                new FunctionItem("PlaceTopToBottom", "上端を下端へ", "2つの選択した図形の片方の図形の上端を、もう一方の図形の下端に接着させる", "place_top_to_bottom.png",
                    () => SafeExecuteFunction(() => alignmentService.PlaceTopToBottom(), "上端を下端へ"), FunctionCategory.Alignment, 2, 2),
                new FunctionItem("PlaceBottomToTop", "下端を上端へ", "2つの選択した図形の片方の図形の下端を、もう一方の図形の上端に接着させる", "place_bottom_to_top.png",
                    () => SafeExecuteFunction(() => alignmentService.PlaceBottomToTop(), "下端を上端へ"), FunctionCategory.Alignment, 2, 3),
                new FunctionItem("CenterAlign", "水平垂直中央揃え", "選択した図形を水平・垂直中央に配置します", "center_align.png",
                    () => SafeExecuteFunction(() => alignmentService.CenterAlign(), "水平垂直中央揃え"), FunctionCategory.Alignment, 2, 4),

                // 3行目
                new FunctionItem("MakeLineHorizontal", "水平にする", "選択した線の角度を水平（0度）にします", "line_horizontal.png",
                    () => SafeExecuteFunction(() => powerToolService.MakeLineHorizontal(), "水平にする"), FunctionCategory.Alignment, 3, 0),
                new FunctionItem("MakeLineVertical", "垂直にする", "選択した線の角度を垂直（90度）にします", "line_vertical.png",
                    () => SafeExecuteFunction(() => powerToolService.MakeLineVertical(), "垂直にする"), FunctionCategory.Alignment, 3, 1),
                new FunctionItem("MatchRoundCorner", "角丸統一", "選択した図形の角丸具合のある図形の角丸位置を同じにします", "match_round_corner.png",
                    () => SafeExecuteFunction(() => shapeService.MatchRoundCorner(), "角丸統一"), FunctionCategory.Alignment, 3, 2),
                new FunctionItem("MatchEnvironment", "矢羽統一", "選択した図形のハンドル設定のある図形のハンドル位置を同じにします", "match_environment.png",
                    () => SafeExecuteFunction(() => shapeService.MatchEnvironment(), "矢羽統一"), FunctionCategory.Alignment, 3, 3)
            };

            allFunctions.AddRange(functions);
        }

        /// <summary>
        /// 図形操作プロカテゴリの機能を追加
        /// </summary>
        private void AddShapeOperationFunctions()
        {
            var functions = new[]
            {
                // 1行目
                new FunctionItem("SplitShape", "図形分割", "選択した図形を指定したグリッドに分割します", "split_shape.png",
                    () => SafeExecuteFunction(() => shapeService.SplitShape(), "図形分割"), FunctionCategory.ShapeOperation, 1, 0),
                new FunctionItem("DuplicateShape", "図形複製", "選択した図形を指定したグリッドに複製します", "duplicate_shape.png",
                    () => SafeExecuteFunction(() => shapeService.DuplicateShape(), "図形複製"), FunctionCategory.ShapeOperation, 1, 1),
                new FunctionItem("GenerateMatrix", "マトリクス生成", "指定した行数・列数のマトリクスを生成します", "generate_matrix.png",
                    () => SafeExecuteFunction(() => shapeService.GenerateMatrix(), "マトリクス生成"), FunctionCategory.ShapeOperation, 1, 2),
                new FunctionItem("AddSequentialNumbers", "連番付与", "選択図形に左上基準で1からの連番を付与します", "sequential_numbers.png",
                    () => SafeExecuteFunction(() => powerToolService.AddSequentialNumbers(), "連番付与"), FunctionCategory.ShapeOperation, 1, 3),
                new FunctionItem("MergeText", "テキスト図形統合", "選択した図形のテキストを改行区切りで合成し、新しいテキストボックスを作成します", "merge_text.png",
                    () => SafeExecuteFunction(() => powerToolService.MergeText(), "テキスト図形統合"), FunctionCategory.ShapeOperation, 1, 4),
                new FunctionItem("SwapPositions", "図形位置の交換", "2つの選択した図形の位置を交換します", "swap_positions.png",
                    () => SafeExecuteFunction(() => powerToolService.SwapPositions(), "図形位置の交換"), FunctionCategory.ShapeOperation, 1, 5)
            };

            allFunctions.AddRange(functions);
        }

        /// <summary>
        /// 表操作カテゴリの機能を追加
        /// </summary>
        private void AddTableOperationFunctions()
        {
            var functions = new[]
            {
                // 1行目
                new FunctionItem("ConvertTableToTextBoxes", "表→オブジェクト", "選択した表をセル毎のテキストボックスに変換します", "table_to_textbox.png",
                    () => SafeExecuteFunction(() => tableConversionService.ConvertTableToTextBoxes(), "表→オブジェクト"), FunctionCategory.TableOperation, 1, 0),
                new FunctionItem("ConvertTextBoxesToTable", "オブジェクト→表", "グリッド配置されたテキストボックスを表に変換します", "textbox_to_table.png",
                    () => SafeExecuteFunction(() => tableConversionService.ConvertTextBoxesToTable(), "オブジェクト→表"), FunctionCategory.TableOperation, 1, 1),
                new FunctionItem("OptimizeMatrixRowHeights", "行高さ最適化", "選択したマトリクスの行高さをテキスト量に応じて最適化します", "optimize_row_heights.png",
                    () => SafeExecuteFunction(() => matrixOperationService.OptimizeMatrixRowHeights(), "行高さ最適化"), FunctionCategory.TableOperation, 1, 2),
                new FunctionItem("OptimizeTableComplete", "表最適化", "選択した表の列幅と行高を同時最適化し、最もコンパクトな表を作成します", "optimize_table_complete.png",
                    () => SafeExecuteFunction(() => matrixOperationService.OptimizeTableComplete(), "表最適化"), FunctionCategory.TableOperation, 1, 3),
                new FunctionItem("EqualizeRowHeights", "行高統一", "選択した表またはオブジェクトマトリクスに行高を統一の高さにします", "equalize_row_heights.png",
                    () => SafeExecuteFunction(() => matrixOperationService.EqualizeRowHeights(), "行高統一"), FunctionCategory.TableOperation, 1, 4),
                new FunctionItem("EqualizeColumnWidths", "列幅統一", "選択した表またはオブジェクトマトリクスに列幅を等幅にします", "equalize_column_widths.png",
                    () => SafeExecuteFunction(() => matrixOperationService.EqualizeColumnWidths(), "列幅統一"), FunctionCategory.TableOperation, 1, 5),
                new FunctionItem("ExcelToPptx", "ExcelToPPT", "クリップボードのExcelデータをPowerPointに貼り付けます", "excel_to_pptx.png",
                    () => SafeExecuteFunction(() => tableConversionService.ExcelToPptx(), "ExcelToPPT"), FunctionCategory.TableOperation, 1, 6),

                // 2行目
                new FunctionItem("AddMatrixRowSeparators", "行間区切り線", "選択したオブジェクトマトリクスの行間に区切り線を追加します", "add_row_separators.png",
                    () => SafeExecuteFunction(() => matrixOperationService.AddMatrixRowSeparators(), "行間区切り線"), FunctionCategory.TableOperation, 2, 0),
                new FunctionItem("AlignShapesToCells", "図形セル整列", "マトリクス上の図形をセル中央に整列します", "align_shapes_to_cells.png",
                    () => SafeExecuteFunction(() => matrixOperationService.AlignShapesToCells(), "図形セル整列"), FunctionCategory.TableOperation, 2, 1),
                new FunctionItem("AddHeaderRowToMatrix", "見出し行付与", "表またはグリッドレイアウトに見出し行を付与します", "add_header_row.png",
                    () => SafeExecuteFunction(() => matrixOperationService.AddHeaderRowToMatrix(), "見出し行付与"), FunctionCategory.TableOperation, 2, 2),
                new FunctionItem("SetCellMargins", "セルマージン設定", "選択した表のセルまたはテキストボックスのマージンを設定します", "cell_margin.png",
                    () => SafeExecuteFunction(() => tableConversionService.SetCellMargins(), "セルマージン設定"), FunctionCategory.TableOperation, 2, 3),
                new FunctionItem("AddMatrixRow", "行追加", "選択した表またはオブジェクトマトリクスに行を追加します", "add_matrix_row.png",
                    () => SafeExecuteFunction(() => matrixOperationService.AddMatrixRow(), "行追加"), FunctionCategory.TableOperation, 2, 4),
                new FunctionItem("AddMatrixColumn", "列追加", "選択した表またはオブジェクトマトリクスに列を追加します", "add_matrix_column.png",
                    () => SafeExecuteFunction(() => matrixOperationService.AddMatrixColumn(), "列追加"), FunctionCategory.TableOperation, 2, 5),

                // 3行目 - Matrix Tuner を追加
                new FunctionItem("MatrixTuner", "Matrix Tuner", "マトリクス配置の高度な調整（サイズ・間隔・ロック）", "matrix_tuner.png",
                    () => SafeExecuteFunction(() => matrixOperationService.MatrixTuner(), "Matrix Tuner"), FunctionCategory.TableOperation, 3, 0)

            };

            allFunctions.AddRange(functions);
        }

        /// <summary>
        /// 間隔カテゴリの機能を追加
        /// </summary>
        private void AddSpacingFunctions()
        {
            var functions = new[]
            {
                // 1行目
                new FunctionItem("DistributeHorizontally", "水平に整列", "選択した図形を水平方向に等間隔で配置します", "distribute_horizontal_builtin.png",
                    () => SafeExecuteFunction(() => ExecutePowerPointCommand("DistributeHorizontally"), "水平に整列"), FunctionCategory.Spacing, 1, 0, true),
                new FunctionItem("DistributeVertically", "垂直に整列", "選択した図形を垂直方向に等間隔で配置します", "distribute_vertical_builtin.png",
                    () => SafeExecuteFunction(() => ExecutePowerPointCommand("DistributeVertically"), "垂直に整列"), FunctionCategory.Spacing, 1, 1, true),
                new FunctionItem("AdjustEqualSpacing", "等間隔調整", "選択した図形を表形式に整頓し、指定間隔で配置します", "equal_spacing.png",
                    () => SafeExecuteFunction(() => shapeService.AdjustEqualSpacing(), "等間隔調整"), FunctionCategory.Spacing, 1, 2)
            };

            allFunctions.AddRange(functions);
        }

        /// <summary>
        /// PowerToolカテゴリの機能を追加
        /// </summary>
        private void AddPowerToolFunctions()
        {
            var functions = new[]
            {
                // 1行目
                new FunctionItem("ShowFontColorChangeDialog", "文字色変更", "選択図形のテキストの色を一括変更します", "font_color_change.png",
                    () => SafeExecuteFunction(() => powerToolService.ShowFontColorChangeDialog(), "文字色変更"), FunctionCategory.PowerTool, 1, 0),
                new FunctionItem("ShowFillColorChangeDialog", "塗り色変更", "選択図形の塗りつぶし色を一括変更します", "fill_color_change.png",
                    () => SafeExecuteFunction(() => powerToolService.ShowFillColorChangeDialog(), "塗り色変更"), FunctionCategory.PowerTool, 1, 1),
                new FunctionItem("ShowLineColorChangeDialog", "枠線色変更", "選択図形の枠線色を一括変更します", "line_color_change.png",
                    () => SafeExecuteFunction(() => powerToolService.ShowLineColorChangeDialog(), "枠線色変更"), FunctionCategory.PowerTool, 1, 2),
                new FunctionItem("ShowFontChangeDialog", "フォント変更", "選択図形のフォントを一括変更します", "font_change.png",
                    () => SafeExecuteFunction(() => powerToolService.ShowFontChangeDialog(), "フォント変更"), FunctionCategory.PowerTool, 1, 3),
                new FunctionItem("ShowCreateLegendDialog", "凡例作成", "選択図形から凡例を作成します", "create_legend.png",
                    () => SafeExecuteFunction(() => powerToolService.ShowCreateLegendDialog(), "凡例作成"), FunctionCategory.PowerTool, 1, 4),
                new FunctionItem("CreateTable", "表作成", "指定した行数・列数の表を作成します", "create_table.png",
                    () => SafeExecuteFunction(() => powerToolService.CreateTable(), "表作成"), FunctionCategory.PowerTool, 1, 5),

                // 2行目
                new FunctionItem("ConvertToSmartArt", "SmartArt変換", "選択図形をSmartArtに変換します", "convert_to_smartart.png",
                    () => SafeExecuteFunction(() => powerToolService.ConvertToSmartArt(), "SmartArt変換"), FunctionCategory.PowerTool, 2, 0),
                new FunctionItem("ShowPDFPlacementDialog", "PDF配置", "PDFを指定配置で貼り付けます", "pdf_placement.png",
                    () => SafeExecuteFunction(() => powerToolService.ShowPDFPlacementDialog(), "PDF配置"), FunctionCategory.PowerTool, 2, 1),
                new FunctionItem("ApplyTemplateColor", "テンプレート色適用", "選択図形にテンプレートの配色を適用します", "apply_template_color.png",
                    () => SafeExecuteFunction(() => powerToolService.ApplyTemplateColor(), "テンプレート色適用"), FunctionCategory.PowerTool, 2, 2),
                new FunctionItem("CreateShapeFromText", "テキストから図形", "選択テキストから図形を作成します", "create_shape_from_text.png",
                    () => SafeExecuteFunction(() => powerToolService.CreateShapeFromText(), "テキストから図形"), FunctionCategory.PowerTool, 2, 3),
                new FunctionItem("ShowBulkReplaceDialog", "一括置換", "スライド内のテキストを一括置換します", "bulk_replace.png",
                    () => SafeExecuteFunction(() => powerToolService.ShowBulkReplaceDialog(), "一括置換"), FunctionCategory.PowerTool, 2, 4),
                new FunctionItem("ExportSlideAsImage", "スライド画像出力", "現在のスライドを画像として出力します", "export_slide_image.png",
                    () => SafeExecuteFunction(() => powerToolService.ExportSlideAsImage(), "スライド画像出力"), FunctionCategory.PowerTool, 2, 5),

                // 3行目
                new FunctionItem("ShowSlideNavigator", "スライド移動", "スライド間を素早く移動します", "slide_navigator.png",
                    () => SafeExecuteFunction(() => powerToolService.ShowSlideNavigator(), "スライド移動"), FunctionCategory.PowerTool, 3, 0),
                new FunctionItem("ShowAnimationTiming", "アニメ調整", "アニメーションタイミングを調整します", "animation_timing.png",
                    () => SafeExecuteFunction(() => powerToolService.ShowAnimationTiming(), "アニメ調整"), FunctionCategory.PowerTool, 3, 1),
                new FunctionItem("ShowGridSettings", "グリッド設定", "グリッドとガイドの設定を変更します", "grid_settings.png",
                    () => SafeExecuteFunction(() => powerToolService.ShowGridSettings(), "グリッド設定"), FunctionCategory.PowerTool, 3, 2),
                new FunctionItem("ShowSlideSize", "スライドサイズ", "スライドサイズを変更します", "slide_size.png",
                    () => SafeExecuteFunction(() => powerToolService.ShowSlideSize(), "スライドサイズ"), FunctionCategory.PowerTool, 3, 3),
                new FunctionItem("ShowPresenterView", "発表者画面", "発表者画面を表示します", "presenter_view.png",
                    () => SafeExecuteFunction(() => powerToolService.ShowPresenterView(), "発表者画面"), FunctionCategory.PowerTool, 3, 4),
                new FunctionItem("ShowSectionManager", "セクション管理", "セクションを管理します", "section_manager.png",
                    () => SafeExecuteFunction(() => powerToolService.ShowSectionManager(), "セクション管理"), FunctionCategory.PowerTool, 3, 5)
            };

            allFunctions.AddRange(functions);
        }

        #endregion

        #region UI作成メソッド

        /// <summary>
        /// UIを作成します
        /// </summary>
        private void CreateUI()
        {
            try
            {
                logger.Debug("Creating UI");

                mainPanel.Controls.Clear();

                int currentY = 10;

                // カテゴリ順にセクション作成
                var orderedCategories = new[]
                {
                    FunctionCategory.Selection,
                    FunctionCategory.Text,
                    FunctionCategory.Shape,
                    FunctionCategory.Format,
                    FunctionCategory.Grouping,
                    FunctionCategory.Alignment,
                    FunctionCategory.ShapeOperation,
                    FunctionCategory.TableOperation,
                    FunctionCategory.Spacing,
                    FunctionCategory.PowerTool
                };

                foreach (var category in orderedCategories)
                {
                    if (categorizedFunctions.ContainsKey(category))
                    {
                        var section = CreateFunctionSection(category, categorizedFunctions[category]);
                        section.Location = new Point(0, currentY);
                        mainPanel.Controls.Add(section);
                        currentY += section.Height + 5;
                    }
                }

                // ライセンス状態バーを作成
                CreateLicenseStatusBarWithButton();

                // ライセンスステータス定期更新タイマー開始
                StartLicenseStatusUpdateTimer();

                logger.Info("UI created successfully");
            }
            catch (Exception ex)
            {
                logger.Error(ex, "Failed to create UI");
                throw new InvalidOperationException("UIの作成に失敗しました", ex);
            }
        }

        /// <summary>
        /// 機能セクションを作成します
        /// </summary>
        private Panel CreateFunctionSection(FunctionCategory category, List<FunctionItem> functions)
        {
            logger.Debug($"Creating section for category: {category}");

            var panel = new Panel
            {
                AutoSize = false,
                Width = 270,
                BorderStyle = BorderStyle.None,
                BackColor = Color.White
            };

            // カテゴリヘッダー
            var headerLabel = new Label
            {
                Text = GetCategoryName(category),
                Location = new Point(5, 5),
                Size = new Size(260, 25),
                BackColor = GetCategoryColor(category),
                ForeColor = Color.White,
                Font = new Font("Yu Gothic UI", 10, FontStyle.Bold),
                TextAlign = ContentAlignment.MiddleLeft,
                Padding = new Padding(10, 0, 0, 0)
            };

            panel.Controls.Add(headerLabel);

            // グリッドレイアウトで機能ボタンを配置
            var maxRow = functions.Max(f => f.Row);
            var maxCol = functions.Max(f => f.Column);

            int buttonSize = 34;
            int spacing = 2;
            int startX = 5;
            int startY = 35;

            for (int row = 1; row <= maxRow; row++)
            {
                for (int col = 0; col <= maxCol; col++)
                {
                    var function = functions.FirstOrDefault(f => f.Row == row && f.Column == col);
                    if (function != null)
                    {
                        var button = CreateFunctionButton(function);
                        button.Location = new Point(
                            startX + col * (buttonSize + spacing),
                            startY + (row - 1) * (buttonSize + spacing)
                        );
                        button.Size = new Size(buttonSize, buttonSize);

                        panel.Controls.Add(button);

                        // ToolTip設定
                        toolTip.SetToolTip(button, $"{function.DisplayName}\n{function.Description}");
                    }
                }
            }

            // パネル高さを自動計算
            var totalHeight = startY + (maxRow * (buttonSize + spacing)) + 10;
            panel.Height = totalHeight;

            logger.Debug($"Section created for {category}: {functions.Count} functions, height={totalHeight}");

            return panel;
        }

        /// <summary>
        /// カテゴリ名を取得します
        /// </summary>
        private string GetCategoryName(FunctionCategory category)
        {
            return category switch
            {
                FunctionCategory.Selection => "選択",
                FunctionCategory.Text => "テキスト",
                FunctionCategory.Shape => "図形",
                FunctionCategory.Format => "整形",
                FunctionCategory.Grouping => "グループ化",
                FunctionCategory.Alignment => "整列",
                FunctionCategory.ShapeOperation => "図形操作プロ",
                FunctionCategory.TableOperation => "表操作",
                FunctionCategory.Spacing => "間隔",
                FunctionCategory.PowerTool => "PowerTool",
                _ => category.ToString()
            };
        }

        /// <summary>
        /// カテゴリ色を取得します
        /// </summary>
        private Color GetCategoryColor(FunctionCategory category)
        {
            return category switch
            {
                FunctionCategory.Selection => Color.FromArgb(52, 152, 219),
                FunctionCategory.Text => Color.FromArgb(46, 204, 113),
                FunctionCategory.Shape => Color.FromArgb(155, 89, 182),
                FunctionCategory.Format => Color.FromArgb(241, 196, 15),
                FunctionCategory.Grouping => Color.FromArgb(230, 126, 34),
                FunctionCategory.Alignment => Color.FromArgb(231, 76, 60),
                FunctionCategory.ShapeOperation => Color.FromArgb(26, 188, 156),
                FunctionCategory.TableOperation => Color.FromArgb(52, 73, 94),
                FunctionCategory.Spacing => Color.FromArgb(149, 165, 166),
                FunctionCategory.PowerTool => Color.FromArgb(192, 57, 43),
                _ => Color.Gray
            };
        }

        /// <summary>
        /// 機能ボタンを作成します
        /// </summary>
        private Button CreateFunctionButton(FunctionItem function)
        {
            var button = new Button
            {
                BackColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                UseVisualStyleBackColor = false,
                Cursor = Cursors.Hand,
                Tag = function,
                Text = "",
                Font = new Font("Yu Gothic UI", 7, FontStyle.Bold),
                ForeColor = Color.Black
            };

            // Built-in機能の場合は背景色を薄いグレーに
            if (function.IsBuiltIn)
            {
                button.BackColor = Color.FromArgb(248, 248, 248);
            }

            button.FlatAppearance.BorderSize = 0;
            button.FlatAppearance.BorderColor = Color.FromArgb(150, 150, 150);
            button.FlatAppearance.MouseOverBackColor = Color.FromArgb(230, 230, 230);

            // アイコンまたはテキスト設定
            try
            {
                var icon = function.GetIcon();
                if (icon != null)
                {
                    // アイコンを24×24にリサイズ
                    var resizedIcon = new Bitmap(icon, new Size(26, 26));
                    button.Image = resizedIcon;
                    button.ImageAlign = ContentAlignment.MiddleCenter;
                    logger.Debug($"Icon set successfully for {function.Name}");
                }
                else
                {
                    throw new Exception("Icon is null");
                }
            }
            catch (Exception ex)
            {
                logger.Warn(ex, $"Failed to set icon for function {function.Name}");
                // アイコン失敗時はテキスト表示
                button.Text = function.GetShortName();
                button.ForeColor = Color.FromArgb(50, 50, 50);
                button.TextAlign = ContentAlignment.MiddleCenter;
                logger.Debug($"Using text '{button.Text}' for {function.Name}");
            }

            button.Click += (sender, e) =>
            {
                try
                {
                    logger.Debug($"Button clicked: {function.Name}");
                    function.Action?.Invoke();
                }
                catch (Exception ex)
                {
                    logger.Error(ex, $"Error executing function: {function.Name}");
                    SafeExecuteFunction(() => throw ex, function.Name);
                }
            };

            return button;
        }

        private void CreateLicenseStatusBarWithButton()
        {
            licenseStatusPanel = new Panel
            {
                Dock = DockStyle.Bottom,
                Height = 35,
                BackColor = Color.FromArgb(240, 240, 240),
                BorderStyle = BorderStyle.FixedSingle,
                Padding = new Padding(5)
            };

            // 状態表示ラベル
            lblLicenseStatus = new Label
            {
                Location = new Point(5, 8),
                Size = new Size(190, 20),
                TextAlign = ContentAlignment.MiddleLeft,
                Font = new Font("Segoe UI", 8.25F, FontStyle.Regular),
                ForeColor = Color.FromArgb(64, 64, 64),
                Text = "ライセンス確認中..."
            };

            // 設定ボタン（小さく配置）
            var btnSettings = new Button
            {
                Text = "設定",
                Location = new Point(200, 5),
                Size = new Size(50, 23),
                Font = new Font("Segoe UI", 8F),
                FlatStyle = FlatStyle.System
            };
            btnSettings.Click += (s, e) => ShowLicenseSettingsDialog();

            licenseStatusPanel.Controls.Add(lblLicenseStatus);
            licenseStatusPanel.Controls.Add(btnSettings);
            this.Controls.Add(licenseStatusPanel);
        }

        #endregion

        #region ライセンス関連

        /// <summary>
        /// ライセンスステータスパネル初期化
        /// </summary>
        private void InitializeLicenseStatusPanel()
        {
            // ライセンス状態を初回取得して反映
            UpdateLicenseStatus();
        }

        /// <summary>
        /// ライセンスステータス更新タイマーを開始
        /// </summary>
        private void StartLicenseStatusUpdateTimer()
        {
            statusUpdateTimer = new System.Windows.Forms.Timer
            {
                Interval = 60000 // 60秒ごと
            };
            statusUpdateTimer.Tick += (s, e) => UpdateLicenseStatus();
            statusUpdateTimer.Start();
            logger.Debug("License status update timer started (60s interval)");
        }

        /// <summary>
        /// ライセンスステータスを更新
        /// </summary>
        private void UpdateLicenseStatus()
        {
            try
            {
                var licenseManager = LicenseManagerService.Instance;
                var info = licenseManager.GetCurrentLicenseInfo();

                if (info.IsValid)
                {
                    lblLicenseStatus.Text = $"{info.PlanName} | 有効期限: {info.ExpiryDate:yyyy/MM/dd}";
                    lblLicenseStatus.ForeColor = Color.FromArgb(39, 174, 96);
                }
                else
                {
                    lblLicenseStatus.Text = "ライセンス未認証";
                    lblLicenseStatus.ForeColor = Color.FromArgb(231, 76, 60);
                }
            }
            catch (Exception ex)
            {
                logger.Error(ex, "Failed to update license status");
                lblLicenseStatus.Text = "ライセンス状態取得エラー";
                lblLicenseStatus.ForeColor = Color.FromArgb(192, 57, 43);
            }
        }

        /// <summary>
        /// ライセンス設定ダイアログを表示
        /// </summary>
        private void ShowLicenseSettingsDialog()
        {
            try
            {
                var dialog = new LicenseSettingsDialog();
                var result = dialog.ShowDialog();

                if (result == DialogResult.OK)
                {
                    // ライセンス状態を即座に更新
                    UpdateLicenseStatus();
                    logger.Info("License settings dialog closed with OK - status updated");
                }
            }
            catch (Exception ex)
            {
                logger.Error(ex, "Failed to show license settings dialog");
                MessageBox.Show(
                    $"ライセンス設定の表示に失敗しました。\n\n{ex.Message}",
                    "エラー",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error
                );
            }
        }

        #endregion

        #region ヘルパーメソッド

        /// <summary>
        /// PowerPoint標準コマンドを実行します
        /// </summary>
        private void ExecutePowerPointCommand(string commandName)
        {
            try
            {
                builtInShapeService.ExecutePowerPointCommand(commandName);
                logger.Debug($"Executed PowerPoint command via BuiltInShapeService: {commandName}");
            }
            catch (Exception ex)
            {
                logger.Error(ex, $"Failed to execute PowerPoint command: {commandName}");
                throw new InvalidOperationException($"PowerPoint標準コマンド '{commandName}' の実行に失敗しました: {ex.Message}");
            }
        }

        /// <summary>
        /// 機能実行の安全ラッパー
        /// </summary>
        private void SafeExecuteFunction(Action action, string functionName)
        {
            try
            {
                action();
            }
            catch (Exception ex)
            {
                logger.Error(ex, $"Error executing function: {functionName}");
                MessageBox.Show(
                    $"機能「{functionName}」の実行中にエラーが発生しました。\n\n{ex.Message}",
                    "機能実行エラー",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning
                );
            }
        }

        /// <summary>
        /// 最小限の機能セットを作成（エラー時のフォールバック）
        /// </summary>
        private List<FunctionItem> CreateMinimalFunctionSet()
        {
            return new List<FunctionItem>
            {
                new FunctionItem("TestFunction", "テスト機能", "テスト用の機能です", "test.png",
                    () => MessageBox.Show("テスト機能が実行されました", "テスト",
                        MessageBoxButtons.OK, MessageBoxIcon.Information), FunctionCategory.PowerTool)
            };
        }

        /// <summary>
        /// 最小限のUIを作成（エラー時のフォールバック）
        /// </summary>
        private void CreateMinimalUI()
        {
            try
            {
                logger.Debug("Creating minimal UI");

                mainPanel.Controls.Clear();

                var errorLabel = new Label
                {
                    Text = "カスタムペインの初期化に失敗しました。\n一部機能が利用できない可能性があります。",
                    Location = new Point(10, 10),
                    Size = new Size(260, 60),
                    TextAlign = ContentAlignment.MiddleCenter,
                    BackColor = Color.White,
                    ForeColor = Color.Red,
                    Font = new Font("Yu Gothic UI", 9),
                    BorderStyle = BorderStyle.FixedSingle
                };

                var testButton = new Button
                {
                    Text = "テスト",
                    Location = new Point(10, 80),
                    Size = new Size(100, 30),
                    BackColor = Color.LightBlue,
                    ForeColor = Color.Black,
                    UseVisualStyleBackColor = false
                };

                testButton.Click += (sender, e) =>
                {
                    MessageBox.Show("テストボタンが動作しています", "テスト", MessageBoxButtons.OK, MessageBoxIcon.Information);
                };

                mainPanel.Controls.Add(errorLabel);
                mainPanel.Controls.Add(testButton);

                logger.Debug("Minimal UI created successfully");
            }
            catch (Exception ex)
            {
                logger.Fatal(ex, "Failed to create minimal UI");
            }
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                statusUpdateTimer?.Stop();
                statusUpdateTimer?.Dispose();
                // 他のリソースも解放
            }
            base.Dispose(disposing);
        }

        #endregion
    }
}