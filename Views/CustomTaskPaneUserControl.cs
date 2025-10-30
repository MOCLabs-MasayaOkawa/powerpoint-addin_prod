using Microsoft.Office.Core;
using NLog;
using PowerPointEfficiencyAddin.Models;
using PowerPointEfficiencyAddin.Models.Licensing;
using PowerPointEfficiencyAddin.Services.Core.Alignment;
using PowerPointEfficiencyAddin.Services.Core.Image;
using PowerPointEfficiencyAddin.Services.Core.PowerTool;
using PowerPointEfficiencyAddin.Services.Core.Selection;
using PowerPointEfficiencyAddin.Services.Core.Shape;
using PowerPointEfficiencyAddin.Services.Core.Text;
using PowerPointEfficiencyAddin.Services.Core.BuiltIn;
using PowerPointEfficiencyAddin.Services.Core.Table;
using PowerPointEfficiencyAddin.Services.Core.Matrix;
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
        private TableConversionService tableConversionService;
        private MatrixOperationService matrixOperationService;
        private BuiltInShapeService builtInShapeService;

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

                tableConversionService = new TableConversionService(applicationProvider, new PowerToolServiceHelper(applicationProvider));
                logger.Debug("TableConversionService created");

                matrixOperationService = new MatrixOperationService(applicationProvider);
                logger.Debug("MatrixOperationService created");

                builtInShapeService = new BuiltInShapeService(applicationProvider, new PowerToolServiceHelper(applicationProvider), powerToolService);
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
                    AutoPopDelay = 5000,
                    InitialDelay = 500,
                    ReshowDelay = 100,
                    ShowAlways = true
                };

                // メインパネル
                mainPanel = new Panel()
                {
                    Name = "mainPanel",
                    Location = new Point(0, 0),
                    Size = new Size(270, 800), // 初期サイズ、後で調整
                    BackColor = Color.White,
                    Padding = new Padding(5),
                    AutoScroll = false
                };

                // ライセンス状態表示パネルの初期化
                InitializeLicenseStatusPanel();

                this.Controls.Clear();
                this.Controls.Add(mainPanel);

                logger.Debug("Custom components initialized successfully");
            }
            catch (Exception ex)
            {
                logger.Error(ex, "Failed to initialize custom components");
                throw;
            }
        }


        #region 機能定義（PDF配置表対応）

        /// <summary>
        /// 選択カテゴリの機能を追加
        /// </summary>
        private void AddSelectionFunctions()
        {
            var functions = new[]
            {
                // 1行目
                new FunctionItem("SelectSameColorShapes", "色で選択", "同スライド内の同じ色の図形を一括選択します", "select_same_color.png",
                    () => SafeExecuteFunction(() => shapeSelectionService.SelectSameColorShapes(), "色で選択"), FunctionCategory.Selection, 1, 0),
                new FunctionItem("SelectSameSizeShapes", "サイズで選択", "同スライド内の同じサイズの図形を一括選択します", "select_same_size.png",
                    () => SafeExecuteFunction(() => shapeSelectionService.SelectSameSizeShapes(), "サイズで選択"), FunctionCategory.Selection, 1, 1),
                new FunctionItem("SelectSimilarShapes", "同種図形で選択", "選択した図形と同じ種類の図形をスライド内で一括選択します", "select_similar.png",
                    () => SafeExecuteFunction(() => powerToolService.SelectSimilarShapes(), "同種図形で選択"), FunctionCategory.Selection, 1, 2)
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
                new FunctionItem("ToggleTextWrap", "折り返しトグル", "図形内でテキストを折り返す設定を切り替えます", "text_wrap.png",
                    () => SafeExecuteFunction(() => textFormatService.ToggleTextWrap(), "折り返しトグル"), FunctionCategory.Text, 1, 0),
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
                    () => SafeExecuteFunction(() => powerToolService.ShowShapeStyleDialog(), "図形スタイル設定"), FunctionCategory.Shape, 2, 6)
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
                new FunctionItem("MatchFormat", "書式を揃える", "選択した図形の書式を最後に選択した図形に合わせます", "match_format.png",
                    () => SafeExecuteFunction(() => shapeService.MatchFormat(), "書式を揃える"), FunctionCategory.Format, 2, 3),

                // 3行目
                new FunctionItem("AlignSizeLeft", "左端を揃える", "選択した図形の左端を基準図形の左端まで伸ばし、左端を揃える", "align_left.png",
                    () => SafeExecuteFunction(() => alignmentService.AlignSizeLeft(), "左端を揃える"), FunctionCategory.Format, 3, 0),
                new FunctionItem("AlignSizeRight", "右端を揃える", "選択した図形の右端を基準図形の右端まで伸ばし、右端を揃える", "align_right.png",
                    () => SafeExecuteFunction(() => alignmentService.AlignSizeRight(), "右端を揃える"), FunctionCategory.Format, 3, 1),
                new FunctionItem("AlignSizeTop", "上端を揃える", "選択した図形の上端を基準図形の上端まで伸ばし、上端を揃える", "align_top.png",
                    () => SafeExecuteFunction(() => alignmentService.AlignSizeTop(), "上端を揃える"), FunctionCategory.Format, 3, 2),
                new FunctionItem("AlignSizeBottom", "下端を揃える", "選択した図形の下端を基準図形の下端まで伸ばし、下端を揃える", "align_bottom.png",
                    () => SafeExecuteFunction(() => alignmentService.AlignSizeBottom(), "下端を揃える"), FunctionCategory.Format, 3, 3),
                new FunctionItem("AlignLineLength", "線の長さを揃える", "選択した線の中で最も長いものを基準に長さを揃えます", "align_line_length.png",
                    () => SafeExecuteFunction(() => powerToolService.AlignLineLength(), "線の長さを揃える"), FunctionCategory.Format, 3, 4)
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
                new FunctionItem("GroupShapes", "グループ化", "選択した図形をグループ化します", "group.png",
                    () => SafeExecuteFunction(() => ExecutePowerPointCommand("GroupObjects"), "グループ化"), FunctionCategory.Grouping, 1, 0, true),
                new FunctionItem("UngroupShapes", "グループ解除", "選択したグループを解除します", "ungroup.png",
                    () => SafeExecuteFunction(() => ExecutePowerPointCommand("UngroupObjects"), "グループ解除"), FunctionCategory.Grouping, 1, 1, true),
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
                    () => SafeExecuteFunction(() => matrixOperationService.ExcelToPptx(), "ExcelToPPT"), FunctionCategory.TableOperation, 1, 6),

                // 2行目
                new FunctionItem("AddMatrixRowSeparators", "行間区切り線", "選択したオブジェクトマトリクスの行間に区切り線を追加します", "add_row_separators.png",
                    () => SafeExecuteFunction(() => matrixOperationService.AddMatrixRowSeparators(), "行間区切り線"), FunctionCategory.TableOperation, 2, 0),
                new FunctionItem("AlignShapesToCells", "図形セル整列", "マトリクス上の図形をセル中央に整列します", "align_shapes_to_cells.png",
                    () => SafeExecuteFunction(() => matrixOperationService.AlignShapesToCells(), "図形セル整列"), FunctionCategory.TableOperation, 2, 1),
                new FunctionItem("AddHeaderRowToMatrix", "見出し行付与", "表またはグリッドレイアウトに見出し行を付与します", "add_header_row.png",
                    () => SafeExecuteFunction(() => matrixOperationService.AddHeaderRowToMatrix(), "見出し行付与"), FunctionCategory.TableOperation, 2, 2),
                new FunctionItem("SetCellMargins", "セルマージン設定", "選択した表のセルまたはテキストボックスのマージンを設定します", "cell_margin.png",
                    () => SafeExecuteFunction(() => matrixOperationService.SetCellMargins(), "セルマージン設定"), FunctionCategory.TableOperation, 2, 3),
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
                new FunctionItem("RemoveSpacing", "間隔をなくす", "図形間の隙間を削除して隣接させます", "remove_spacing.png",
                    () => SafeExecuteFunction(() => alignmentService.RemoveSpacing(), "間隔をなくす"), FunctionCategory.Spacing, 1, 0),
                new FunctionItem("AdjustHorizontalSpacing", "水平間隔調整", "水平方向の図形間隔を詳細調整します", "adjust_horizontal_spacing.png",
                    () => SafeExecuteFunction(() => alignmentService.AdjustHorizontalSpacing(), "水平間隔調整"), FunctionCategory.Spacing, 1, 1),
                new FunctionItem("AdjustVerticalSpacing", "垂直間隔調整", "垂直方向の図形間隔を詳細調整します", "adjust_vertical_spacing.png",
                    () => SafeExecuteFunction(() => alignmentService.AdjustVerticalSpacing(), "垂直間隔調整"), FunctionCategory.Spacing, 1, 2),
                new FunctionItem("AdjustEqualSpacing", "間隔調整", "選択図形を表形式に整頓し、指定間隔で配置します", "adjust_equal_spacing.png",
                    () => SafeExecuteFunction(() => shapeService.AdjustEqualSpacing(), "間隔調整"), FunctionCategory.Spacing, 1, 3),
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
                new FunctionItem("UnifyFont", "テキスト一括置換", "全ページのすべてのテキストを指定フォントに統一します", "unify_font.png",
                    () => SafeExecuteFunction(() => powerToolService.UnifyFont(), "テキスト一括置換"), FunctionCategory.PowerTool, 1, 0),
                new FunctionItem("CompressImages", "画像圧縮", "選択した画像のファイルサイズを高品質圧縮で削減します", "compress_images.png",
                    () => SafeExecuteFunction(() => imageCompressionService.CompressImages(), "画像圧縮"), FunctionCategory.PowerTool, 1, 1)
            };

            allFunctions.AddRange(functions);
        }

        #endregion

        #region ライセンス管理

        /// <summary>
        /// ライセンス状態表示パネルを初期化（新規追加）
        /// </summary>
        private void InitializeLicenseStatusPanel()
        {
            try
            {
                logger.Info("InitializeLicenseStatusPanel started");

                // 既存のパネルがあれば削除
                if (licenseStatusPanel != null)
                {
                    this.Controls.Remove(licenseStatusPanel);
                    licenseStatusPanel.Dispose();
                }

                // クラスフィールドに代入（var を使わない）
                licenseStatusPanel = new Panel
                {
                    Name = "licenseStatusPanel",
                    Dock = DockStyle.Top,
                    Height = 35,
                    BackColor = Color.FromArgb(240, 240, 240),
                    BorderStyle = BorderStyle.FixedSingle
                };

                // クラスフィールドに代入
                lblLicenseStatus = new Label
                {
                    Name = "lblLicenseStatus",
                    Location = new Point(5, 8),
                    Size = new Size(150, 20),
                    Font = new Font("Segoe UI", 9F),
                    ForeColor = Color.DarkGray,
                    Text = "ライセンス: 確認中...",
                    AutoSize = false
                };

                var btnSettings = new Button
                {
                    Name = "btnSettings",
                    Text = "設定",
                    Location = new Point(160, 6),
                    Size = new Size(45, 23),
                    Font = new Font("Segoe UI", 8F),
                    UseVisualStyleBackColor = true
                };
                btnSettings.Click += (s, e) =>
                {
                    logger.Info("Settings button clicked");
                    ShowLicenseSettingsDialog();
                };

                var btnUpdate = new Button
                {
                    Name = "btnUpdate",
                    Text = "更新",
                    Location = new Point(210, 6),
                    Size = new Size(60, 23),
                    Font = new Font("Segoe UI", 8F, FontStyle.Bold),
                    Visible = true,  // テスト用に表示
                    BackColor = Color.Orange,
                    ForeColor = Color.White
                };
                btnUpdate.Click += async (s, e) =>
                {
                    logger.Info("Update button clicked");
                    await HandleUpdateClick();
                };

                // コントロールを追加
                licenseStatusPanel.Controls.Add(lblLicenseStatus);
                licenseStatusPanel.Controls.Add(btnSettings);
                licenseStatusPanel.Controls.Add(btnUpdate);

                // UserControlに追加
                this.Controls.Add(licenseStatusPanel);
                licenseStatusPanel.BringToFront();

                logger.Info($"License panel added. Controls count: {licenseStatusPanel.Controls.Count}");
                logger.Info($"Panel visible: {licenseStatusPanel.Visible}, Size: {licenseStatusPanel.Size}");

                // 初期表示（テスト用）
                lblLicenseStatus.Text = "ライセンス: テスト中";
                lblLicenseStatus.ForeColor = Color.Green;
                btnUpdate.Text = "v2.0.0";

                logger.Info("InitializeLicenseStatusPanel completed successfully");
            }
            catch (Exception ex)
            {
                logger.Error(ex, "Failed to initialize license status panel");
                MessageBox.Show($"ライセンスパネル初期化エラー: {ex.Message}", "エラー");
            }
        }

        /// <summary>
        /// ライセンス状態を更新（新規追加）
        /// </summary>
        /// <param name="accessLevel">アクセスレベル</param>
        /// <param name="statusMessage">ステータスメッセージ</param>
        public void UpdateLicenseStatus(FeatureAccessLevel accessLevel, string statusMessage)
        {
            try
            {
                if (this.InvokeRequired)
                {
                    this.Invoke(new Action(() => UpdateLicenseStatus(accessLevel, statusMessage)));
                    return;
                }

                if (lblLicenseStatus == null)
                {
                    InitializeLicenseStatusPanel();
                }

                // ステータステキストを設定
                lblLicenseStatus.Text = statusMessage ?? "ライセンス状態不明";

                // アクセスレベルに応じて色を変更
                switch (accessLevel)
                {
                    case FeatureAccessLevel.Pro:
                        lblLicenseStatus.ForeColor = Color.Green;
                        licenseStatusPanel.BackColor = Color.FromArgb(240, 255, 240);
                        break;

                    case FeatureAccessLevel.Free:
                        lblLicenseStatus.ForeColor = Color.Orange;
                        licenseStatusPanel.BackColor = Color.FromArgb(255, 250, 230);
                        break;

                    case FeatureAccessLevel.Blocked:
                        lblLicenseStatus.ForeColor = Color.Red;
                        licenseStatusPanel.BackColor = Color.FromArgb(255, 240, 240);
                        break;

                    default:
                        lblLicenseStatus.ForeColor = Color.Gray;
                        licenseStatusPanel.BackColor = Color.FromArgb(240, 240, 240);
                        break;
                }

                // 機能ボタンの有効/無効を切り替え
                UpdateFunctionButtonsState(accessLevel);

                logger.Debug($"License status updated: {accessLevel} - {statusMessage}");
            }
            catch (Exception ex)
            {
                logger.Error(ex, "Failed to update license status");
            }
        }

        /// <summary>
        /// 機能ボタンの有効/無効状態を更新（新規追加）
        /// </summary>
        private void UpdateFunctionButtonsState(FeatureAccessLevel accessLevel)
        {
            try
            {
                if (accessLevel == FeatureAccessLevel.Blocked)
                {
                    // すべての機能ボタンを無効化
                    foreach (Control control in this.Controls)
                    {
                        if (control is Button button && button != btnShowLicenseSettings)
                        {
                            button.Enabled = false;
                        }
                    }
                }
                else if (accessLevel == FeatureAccessLevel.Free)
                {
                    // 高度な機能のみ無効化
                    // TODO: 機能ごとの制限を実装
                    foreach (Control control in this.Controls)
                    {
                        if (control is Button button)
                        {
                            // 基本機能のタグを持つボタンのみ有効
                            button.Enabled = IsBasicFunction(button.Tag?.ToString());
                        }
                    }
                }
                else
                {
                    // すべての機能ボタンを有効化
                    foreach (Control control in this.Controls)
                    {
                        if (control is Button button)
                        {
                            button.Enabled = true;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                logger.Error(ex, "Failed to update function buttons state");
            }
        }

        /// <summary>
        /// 基本機能かどうかを判定（新規追加）
        /// </summary>
        private bool IsBasicFunction(string functionTag)
        {
            if (string.IsNullOrEmpty(functionTag))
                return false;

            string[] basicFunctions =
            {
                "AlignLeft", "AlignRight", "AlignTop", "AlignBottom",
                "AlignCenter", "AlignMiddle", "DistributeHorizontally",
                "DistributeVertically", "MatchSize"
            };

            return Array.Exists(basicFunctions, f =>
                functionTag.IndexOf(f, StringComparison.OrdinalIgnoreCase) >= 0);
        }

        /// <summary>
        /// ライセンス設定ダイアログを表示（新規追加）
        /// </summary>
        private void ShowLicenseSettingsDialog()
        {
            try
            {
                using (var dialog = new LicenseSettingsDialog())
                {
                    if (dialog.ShowDialog() == DialogResult.OK)
                    {
                        // ライセンスキーが更新された場合の処理
                        logger.Info("License key updated via settings dialog");
                    }
                }
            }
            catch (Exception ex)
            {
                logger.Error(ex, "Failed to show license settings dialog");
                MessageBox.Show(
                    "ライセンス設定画面の表示に失敗しました。",
                    "エラー",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }

        // 新規メソッド：ライセンスと更新ステータスの更新
        private void UpdateLicenseAndUpdateStatus(Label lblLicenseStatus, Button btnUpdate)
        {
            try
            {
                var licenseManager = LicenseManager.Instance;
                var status = licenseManager.CurrentStatus;

                if (status != null && status.IsValid)
                {
                    lblLicenseStatus.Text = $"ライセンス: {status.PlanType ?? "有効"}";
                    lblLicenseStatus.ForeColor = Color.Green;

                    // 【追加】更新チェック
                    if (licenseManager.HasPendingUpdate())
                    {
                        var update = licenseManager.GetPendingUpdate();
                        if (update != null)
                        {
                            btnUpdate.Visible = true;

                            // 重要更新の場合は赤色
                            if (update.IsCritical)
                            {
                                btnUpdate.BackColor = Color.FromArgb(220, 53, 69);
                                btnUpdate.ForeColor = Color.White;
                                btnUpdate.Text = "重要更新";

                                // ツールチップで詳細表示
                                var toolTip = new ToolTip();
                                toolTip.SetToolTip(btnUpdate,
                                    $"重要な更新があります: v{update.Version}\n" +
                                    "セキュリティ修正が含まれています。");
                            }
                            else
                            {
                                btnUpdate.BackColor = Color.FromArgb(255, 193, 7);
                                btnUpdate.ForeColor = Color.Black;
                                btnUpdate.Text = "更新可";

                                var toolTip = new ToolTip();
                                toolTip.SetToolTip(btnUpdate,
                                    $"新しいバージョンがあります: v{update.Version}");
                            }
                        }
                    }
                    else
                    {
                        btnUpdate.Visible = false;
                    }
                }
                else
                {
                    lblLicenseStatus.Text = "ライセンス: 未登録";
                    lblLicenseStatus.ForeColor = Color.Orange;
                    btnUpdate.Visible = false;
                }
            }
            catch (Exception ex)
            {
                logger.Error(ex, "Failed to update license and update status");
            }
        }

        // 新規メソッド：更新ボタンクリック処理
        private async Task HandleUpdateClick()
        {
            try
            {
                var licenseManager = LicenseManager.Instance;
                var update = licenseManager.GetPendingUpdate();

                if (update == null)
                {
                    MessageBox.Show("更新情報が見つかりません。", "更新",
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                // 更新ダイアログ表示
                var message = $"新しいバージョン {update.Version} が利用可能です。\n\n" +
                             $"リリース日: {update.ReleaseDate:yyyy/MM/dd}\n";

                if (!string.IsNullOrEmpty(update.ReleaseNotes))
                {
                    message += $"\n更新内容:\n{update.ReleaseNotes}\n";
                }

                message += "\n今すぐダウンロードしますか？\n" +
                          "（ダウンロード後、PowerPoint終了時に自動的にインストールされます）";

                var result = MessageBox.Show(message, "アップデート確認",
                    MessageBoxButtons.YesNo, MessageBoxIcon.Information);

                if (result == DialogResult.Yes)
                {
                    // プログレスダイアログ表示
                    var progressForm = new Form
                    {
                        Text = "更新をダウンロード中...",
                        Size = new Size(400, 120),
                        StartPosition = FormStartPosition.CenterScreen,
                        FormBorderStyle = FormBorderStyle.FixedDialog,
                        MaximizeBox = false,
                        MinimizeBox = false
                    };

                    var progressBar = new ProgressBar
                    {
                        Style = ProgressBarStyle.Marquee,
                        Location = new Point(20, 20),
                        Size = new Size(340, 30)
                    };

                    var lblStatus = new Label
                    {
                        Text = "更新ファイルをダウンロードしています...",
                        Location = new Point(20, 60),
                        Size = new Size(340, 20)
                    };

                    progressForm.Controls.Add(progressBar);
                    progressForm.Controls.Add(lblStatus);

                    progressForm.Show();

                    // ダウンロード実行
                    var success = await licenseManager.DownloadUpdateAsync();

                    progressForm.Close();

                    if (success)
                    {
                        MessageBox.Show(
                            "更新のダウンロードが完了しました。\n" +
                            "PowerPoint終了時に自動的にインストールされます。",
                            "ダウンロード完了",
                            MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    else
                    {
                        MessageBox.Show(
                            "更新のダウンロードに失敗しました。\n" +
                            "後で再試行してください。",
                            "ダウンロード失敗",
                            MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }
                }
            }
            catch (Exception ex)
            {
                logger.Error(ex, "Failed to handle update click");
                MessageBox.Show(
                    "更新処理中にエラーが発生しました。",
                    "エラー",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        #endregion

        #region UI作成

        /// <summary>
        /// UIを作成します（PDF配置表対応）
        /// </summary>
        private void CreateUI()
        {
            try
            {
                logger.Debug("Starting CreateUI");

                // ライセンスパネルが未初期化の場合は初期化
                if (licenseStatusPanel == null)
                {
                    InitializeLicenseStatusPanel();
                }

                // 既存のコントロールをクリア
                mainPanel.Controls.Clear();

                var yPosition = 10;
                const int buttonSize = 26; // 24×24に変更
                const int buttonsPerRow = 8; // 仕様書通り8個/行
                const int buttonMargin = 4;
                const int headerHeight = 20;

                // カテゴリ順にUI作成
                var categories = CategoryInfo.GetAllCategories();
                logger.Debug($"Processing {categories.Length} categories");

                //CreateLicenseStatusBarWithButton();

                foreach (var categoryInfo in categories)
                {
                    try
                    {
                        var categoryFunctions = allFunctions
                            .Where(f => f.Category == categoryInfo.Category)
                            .OrderBy(f => f.RowPosition)     // まず行位置でソート
                            .ThenBy(f => f.Order)            // 同じ行内ではOrder値でソート（名前ソートを削除）
                            .ToList();

                        if (categoryFunctions.Count == 0)
                        {
                            logger.Debug($"No functions for category {categoryInfo.Category}");
                            continue;
                        }

                        logger.Debug($"Creating UI for category {categoryInfo.DisplayName} with {categoryFunctions.Count} functions");

                        // カテゴリヘッダー
                        var headerLabel = CreateCategoryHeader(categoryInfo.DisplayName, categoryInfo.HeaderColor);
                        headerLabel.Location = new Point(10, yPosition);
                        headerLabel.Size = new Size(250, headerHeight);
                        mainPanel.Controls.Add(headerLabel);
                        yPosition += headerHeight + 5;

                        // 行別にボタンを配置
                        var rowGroups = categoryFunctions.GroupBy(f => f.RowPosition).OrderBy(g => g.Key);

                        foreach (var rowGroup in rowGroups)
                        {
                            var rowFunctions = rowGroup.OrderBy(f => f.Order).ToList(); // Order順でソート
                            logger.Debug($"Creating row {rowGroup.Key} with {rowFunctions.Count} functions in order: {string.Join(", ", rowFunctions.Select(f => $"{f.Name}(Order:{f.Order})"))}");

                            for (int i = 0; i < rowFunctions.Count; i++)
                            {
                                if (i > 0 && i % buttonsPerRow == 0)
                                {
                                    // 次の行へ
                                    yPosition += buttonSize + buttonMargin;
                                }

                                var function = rowFunctions[i];
                                var button = CreateFunctionButton(function, buttonSize);

                                var col = i % buttonsPerRow;
                                var xPosition = 10 + col * (buttonSize + buttonMargin);

                                button.Location = new Point(xPosition, yPosition);
                                mainPanel.Controls.Add(button);

                                toolTip.SetToolTip(button, $"{function.Name}\n{function.Description}");

                                logger.Debug($"Created button for {function.Name} (Order:{function.Order}) at ({xPosition}, {yPosition})");
                            }

                            yPosition += buttonSize + buttonMargin + 2; // 行間隔
                        }

                        yPosition += 10; // カテゴリ間隔
                    }
                    catch (Exception categoryEx)
                    {
                        logger.Error(categoryEx, $"Error creating UI for category {categoryInfo.Category}");
                    }
                }

                // パネルサイズを動的に調整
                var finalHeight = Math.Max(yPosition + 50, 800);
                mainPanel.Size = new Size(270, finalHeight);

                // 強制再描画
                mainPanel.Invalidate();
                this.Invalidate();

                logger.Info($"Created UI with {allFunctions.Count} function buttons across {categories.Length} categories. Panel height: {finalHeight}");
            }
            catch (Exception ex)
            {
                logger.Error(ex, "Critical error in CreateUI");
                CreateMinimalUI();
            }
        }

        /// <summary>
        /// カテゴリヘッダーを作成します
        /// </summary>
        private Label CreateCategoryHeader(string categoryName, Color headerColor)
        {
            return new Label
            {
                Text = categoryName,
                Font = new Font("Yu Gothic UI", 9, FontStyle.Bold),
                BackColor = headerColor,
                ForeColor = Color.Black,
                TextAlign = ContentAlignment.MiddleLeft,
                Padding = new Padding(5, 0, 0, 0),
                AutoSize = false
            };
        }

        /// <summary>
        /// 機能ボタンを作成します
        /// </summary>
        private Button CreateFunctionButton(FunctionItem function, int buttonSize)
        {
            var button = new Button
            {
                Size = new Size(buttonSize, buttonSize),
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.White,
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