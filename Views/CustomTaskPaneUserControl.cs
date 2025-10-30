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
    /// カスタムタスクペインのユーザーコントロール(PDF配置表対応版)
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
                this.AutoScroll = true;
                this.BackColor = Color.White;

                // ToolTip
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


        #region 機能定義(PDF配置表対応)

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
                new FunctionItem("AdjustMarginUp", "余白Up", "図形内の余白サイズを1.5倍に拡大します", "margin_up.png",
                    () => SafeExecuteFunction(() => textFormatService.AdjustMarginUp(), "余白Up"), FunctionCategory.Text, 1, 1),
                new FunctionItem("AdjustMarginDown", "余白Down", "図形内の余白サイズを1.5で割って縮小します", "margin_down.png",
                    () => SafeExecuteFunction(() => textFormatService.AdjustMarginDown(), "余白Down"), FunctionCategory.Text, 1, 2),
                new FunctionItem("ShowMarginAdjustDialog", "余白調整", "図形内余白の詳細調整ダイアログを開きます", "margin_adjust.png",
                    () => SafeExecuteFunction(() => textFormatService.ShowMarginAdjustDialog(), "余白調整"), FunctionCategory.Text, 1, 3)
            };

            allFunctions.AddRange(functions);
        }

        /// <summary>
        /// 図形カテゴリの機能を追加(PDF配置表対応)
        /// </summary>
        private void AddShapeFunctions()
        {
            var functions = new[]
            {
                // 1行目
                new FunctionItem("SelectAllShapes", "すべて選択", "スライド内のすべての図形を選択します", "select_all.png",
                    () => SafeExecuteFunction(() => shapeService.SelectAllShapes(), "すべて選択"), FunctionCategory.Shape, 1, 0),
                new FunctionItem("CopyShapeFormatting", "書式コピー", "選択図形の書式を他の図形にコピーします", "copy_formatting.png",
                    () => SafeExecuteFunction(() => shapeService.CopyShapeFormatting(), "書式コピー"), FunctionCategory.Shape, 1, 1),
                new FunctionItem("ResetShapeSize", "サイズリセット", "図形のサイズを初期状態にリセットします", "reset_size.png",
                    () => SafeExecuteFunction(() => shapeService.ResetShapeSize(), "サイズリセット"), FunctionCategory.Shape, 1, 2)
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
                new FunctionItem("MatchHeight", "縦幅を揃える", "選択図形の高さを揃えます", "match_height.png",
                    () => SafeExecuteFunction(() => shapeService.MatchHeight(), "縦幅を揃える"), FunctionCategory.Format, 1, 0),
                new FunctionItem("MatchWidth", "横幅を揃える", "選択図形の幅を揃えます", "match_width.png",
                    () => SafeExecuteFunction(() => shapeService.MatchWidth(), "横幅を揃える"), FunctionCategory.Format, 1, 1),
                new FunctionItem("MatchSize", "サイズを揃える", "選択図形のサイズを揃えます", "match_size.png",
                    () => SafeExecuteFunction(() => shapeService.MatchSize(), "サイズを揃える"), FunctionCategory.Format, 1, 2)
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
                new FunctionItem("GroupShapes", "グループ化", "選択図形をグループ化します", "group.png",
                    () => SafeExecuteFunction(() => shapeService.GroupShapes(), "グループ化"), FunctionCategory.Grouping, 1, 0),
                new FunctionItem("UngroupShapes", "グループ解除", "グループを解除します", "ungroup.png",
                    () => SafeExecuteFunction(() => shapeService.UngroupShapes(), "グループ解除"), FunctionCategory.Grouping, 1, 1),
                new FunctionItem("GroupByRows", "行でグループ化", "図形を行ごとにグループ化します", "group_by_rows.png",
                    () => SafeExecuteFunction(() => alignmentService.GroupByRows(), "行でグループ化"), FunctionCategory.Grouping, 1, 2),
                new FunctionItem("GroupByColumns", "列でグループ化", "図形を列ごとにグループ化します", "group_by_columns.png",
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
                new FunctionItem("AlignLeft", "左揃え", "選択図形を左揃えにします", "align_left.png",
                    () => SafeExecuteFunction(() => alignmentService.AlignLeft(), "左揃え"), FunctionCategory.Alignment, 1, 0),
                new FunctionItem("AlignCenter", "中央揃え", "選択図形を中央揃えにします", "align_center.png",
                    () => SafeExecuteFunction(() => alignmentService.AlignCenter(), "中央揃え"), FunctionCategory.Alignment, 1, 1),
                new FunctionItem("AlignRight", "右揃え", "選択図形を右揃えにします", "align_right.png",
                    () => SafeExecuteFunction(() => alignmentService.AlignRight(), "右揃え"), FunctionCategory.Alignment, 1, 2),
                new FunctionItem("AlignTop", "上揃え", "選択図形を上揃えにします", "align_top.png",
                    () => SafeExecuteFunction(() => alignmentService.AlignTop(), "上揃え"), FunctionCategory.Alignment, 1, 3),
                new FunctionItem("AlignMiddle", "水平中央揃え", "選択図形を水平中央に揃えます", "align_middle.png",
                    () => SafeExecuteFunction(() => alignmentService.AlignMiddle(), "水平中央揃え"), FunctionCategory.Alignment, 1, 4),
                new FunctionItem("AlignBottom", "下揃え", "選択図形を下揃えにします", "align_bottom.png",
                    () => SafeExecuteFunction(() => alignmentService.AlignBottom(), "下揃え"), FunctionCategory.Alignment, 1, 5),

                // 2行目
                new FunctionItem("DistributeHorizontally", "水平に整列", "選択図形を水平方向に等間隔で配置します", "distribute_horizontal.png",
                    () => SafeExecuteFunction(() => alignmentService.DistributeHorizontally(), "水平に整列"), FunctionCategory.Alignment, 2, 0),
                new FunctionItem("DistributeVertically", "垂直に整列", "選択図形を垂直方向に等間隔で配置します", "distribute_vertical.png",
                    () => SafeExecuteFunction(() => alignmentService.DistributeVertically(), "垂直に整列"), FunctionCategory.Alignment, 2, 1)
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
                new FunctionItem("SplitShape", "図形分割", "選択図形を指定したグリッドに分割します", "split_shape.png",
                    () => SafeExecuteFunction(() => shapeService.SplitShape(), "図形分割"), FunctionCategory.ShapeOperation, 1, 0),
                new FunctionItem("DuplicateShape", "図形複製", "選択図形を指定したグリッドに複製します", "duplicate_shape.png",
                    () => SafeExecuteFunction(() => shapeService.DuplicateShape(), "図形複製"), FunctionCategory.ShapeOperation, 1, 1),
                new FunctionItem("GenerateMatrix", "マトリクス生成", "指定した行列のマトリクスを生成します", "generate_matrix.png",
                    () => SafeExecuteFunction(() => shapeService.GenerateMatrix(), "マトリクス生成"), FunctionCategory.ShapeOperation, 1, 2)
            };

            allFunctions.AddRange(functions);
        }

        /// <summary>
        /// 表操作カテゴリの機能を追加(PDF配置表対応)
        /// </summary>
        private void AddTableOperationFunctions()
        {
            var functions = new[]
            {
                // 1行目
                new FunctionItem("ConvertTableToTextBox", "表→テキストボックス", "表をテキストボックスに変換します", "table_to_textbox.png",
                    () => SafeExecuteFunction(() => powerToolService.ConvertTableToTextBoxes(), "表→テキストボックス"), FunctionCategory.TableOperation, 1, 0),
                new FunctionItem("ConvertTextBoxToTable", "テキストボックス→表", "テキストボックスを表に変換します", "textbox_to_table.png",
                    () => SafeExecuteFunction(() => powerToolService.ConvertTextBoxesToTable(), "テキストボックス→表"), FunctionCategory.TableOperation, 1, 1),
                new FunctionItem("OptimizeMatrix", "マトリクス最適化", "マトリクスの行高を最適化します", "optimize_matrix.png",
                    () => SafeExecuteFunction(() => powerToolService.OptimizeMatrixRowHeights(), "マトリクス最適化"), FunctionCategory.TableOperation, 1, 2),
                new FunctionItem("OptimizeTable", "表最適化", "表の幅と高さを最適化します", "optimize_table.png",
                    () => SafeExecuteFunction(() => powerToolService.OptimizeTableComplete(), "表最適化"), FunctionCategory.TableOperation, 1, 3)
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
        /// ライセンス状態表示パネルを初期化(新規追加)
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

                // クラスフィールドに代入(var を使わない)
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

                logger.Info($"License panel added. Controls count: {this.Controls.Count}");

                // 初回ライセンスステータス更新を呼び出し
                UpdateLicenseAndUpdateStatus(lblLicenseStatus, btnUpdate);

                // ライセンスステータス定期更新タイマーを開始
                StartLicenseStatusTimer();
            }
            catch (Exception ex)
            {
                logger.Error(ex, "InitializeLicenseStatusPanel error");
            }
        }

        /// <summary>
        /// ライセンスステータス定期更新タイマーを開始
        /// </summary>
        private void StartLicenseStatusTimer()
        {
            try
            {
                if (statusUpdateTimer != null)
                {
                    statusUpdateTimer.Stop();
                    statusUpdateTimer.Dispose();
                }

                statusUpdateTimer = new System.Windows.Forms.Timer
                {
                    Interval = 30000 // 30秒ごとに更新
                };

                statusUpdateTimer.Tick += (s, e) =>
                {
                    logger.Debug("License status timer tick");
                    UpdateLicenseAndUpdateStatus(lblLicenseStatus, 
                        licenseStatusPanel?.Controls.OfType<Button>().FirstOrDefault(b => b.Name == "btnUpdate"));
                };

                statusUpdateTimer.Start();
                logger.Info("License status timer started");
            }
            catch (Exception ex)
            {
                logger.Error(ex, "Failed to start license status timer");
            }
        }

        /// <summary>
        /// 機能初期化
        /// </summary>
        private void InitializeFunctions()
        {
            try
            {
                logger.Debug("Initializing functions");

                allFunctions = new List<FunctionItem>();

                // 各カテゴリの機能を追加
                AddSelectionFunctions();
                AddTextFunctions();
                AddShapeFunctions();
                AddFormatFunctions();
                AddGroupingFunctions();
                AddAlignmentFunctions();
                AddShapeOperationFunctions();
                AddTableOperationFunctions();
                AddSpacingFunctions();
                AddPowerToolFunctions();

                // カテゴリ別に整理
                categorizedFunctions = allFunctions
                    .GroupBy(f => f.Category)
                    .ToDictionary(g => g.Key, g => g.ToList());

                logger.Info($"Functions initialized: {allFunctions.Count} functions across {categorizedFunctions.Count} categories");
            }
            catch (Exception ex)
            {
                logger.Error(ex, "Failed to initialize functions");
                allFunctions = CreateMinimalFunctionSet();
            }
        }

        #endregion

        #region UI作成

        /// <summary>
        /// UIを作成します
        /// </summary>
        private void CreateUI()
        {
            try
            {
                logger.Debug("Creating UI");

                var categories = new[]
                {
                    new { Category = FunctionCategory.Selection, Name = "選択", Color = Color.FromArgb(52, 152, 219) },
                    new { Category = FunctionCategory.Text, Name = "テキスト", Color = Color.FromArgb(46, 204, 113) },
                    new { Category = FunctionCategory.Shape, Name = "図形", Color = Color.FromArgb(155, 89, 182) },
                    new { Category = FunctionCategory.Format, Name = "整形", Color = Color.FromArgb(241, 196, 15) },
                    new { Category = FunctionCategory.Grouping, Name = "グループ化", Color = Color.FromArgb(230, 126, 34) },
                    new { Category = FunctionCategory.Alignment, Name = "整列", Color = Color.FromArgb(231, 76, 60) },
                    new { Category = FunctionCategory.ShapeOperation, Name = "図形操作プロ", Color = Color.FromArgb(26, 188, 156) },
                    new { Category = FunctionCategory.TableOperation, Name = "表操作", Color = Color.FromArgb(52, 73, 94) },
                    new { Category = FunctionCategory.Spacing, Name = "間隔", Color = Color.FromArgb(149, 165, 166) },
                    new { Category = FunctionCategory.PowerTool, Name = "PowerTool", Color = Color.FromArgb(192, 57, 43) }
                };

                int yPosition = 10;
                int buttonSize = 34;
                int buttonMargin = 2;
                int buttonsPerRow = 8;

                foreach (var categoryInfo in categories)
                {
                    try
                    {
                        if (!categorizedFunctions.ContainsKey(categoryInfo.Category))
                        {
                            logger.Debug($"No functions for category {categoryInfo.Category}, skipping");
                            continue;
                        }

                        var categoryFunctions = categorizedFunctions[categoryInfo.Category];

                        // カテゴリヘッダー
                        var header = new Label
                        {
                            Text = categoryInfo.Name,
                            Location = new Point(10, yPosition),
                            Size = new Size(250, 25),
                            BackColor = categoryInfo.Color,
                            ForeColor = Color.White,
                            Font = new Font("Yu Gothic UI", 10, FontStyle.Bold),
                            TextAlign = ContentAlignment.MiddleLeft,
                            Padding = new Padding(10, 0, 0, 0)
                        };

                        mainPanel.Controls.Add(header);
                        yPosition += 30;

                        // 行ごとにグループ化してボタンを配置
                        var rowGroups = categoryFunctions
                            .GroupBy(f => f.Row)
                            .OrderBy(g => g.Key);

                        foreach (var rowGroup in rowGroups)
                        {
                            var rowFunctions = rowGroup.OrderBy(f => f.Order).ToList();

                            logger.Debug($"Creating row {rowGroup.Key} with {rowFunctions.Count} functions");

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

                logger.Info($"Created UI with {allFunctions.Count} function buttons. Panel height: {finalHeight}");
            }
            catch (Exception ex)
            {
                logger.Error(ex, "Critical error in CreateUI");
                CreateMinimalUI();
            }
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

            button.FlatAppearance.BorderSize = 0;
            button.FlatAppearance.BorderColor = Color.FromArgb(150, 150, 150);
            button.FlatAppearance.MouseOverBackColor = Color.FromArgb(230, 230, 230);

            // アイコンまたはテキスト設定
            try
            {
                var icon = function.GetIcon();
                if (icon != null)
                {
                    var resizedIcon = new Bitmap(icon, new Size(26, 26));
                    button.Image = resizedIcon;
                    button.ImageAlign = ContentAlignment.MiddleCenter;
                }
                else
                {
                    throw new Exception("Icon not found");
                }
            }
            catch (Exception ex)
            {
                logger.Warn(ex, $"Failed to load icon for {function.Name}, using text");
                button.Text = function.DisplayName;
                button.TextAlign = ContentAlignment.MiddleCenter;
            }

            button.Click += (sender, e) =>
            {
                try
                {
                    logger.Info($"Function button clicked: {function.Name}");
                    function.Action?.Invoke();
                }
                catch (Exception ex)
                {
                    logger.Error(ex, $"Error executing function {function.Name}");
                    MessageBox.Show(
                        $"機能「{function.DisplayName}」の実行中にエラーが発生しました。\n\n{ex.Message}",
                        "エラー",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Error);
                }
            };

            return button;
        }

        #endregion

        #region ヘルパーメソッド

        /// <summary>
        /// 機能実行の安全ラッパー
        /// </summary>
        private void SafeExecuteFunction(Action action, string functionName)
        {
            ErrorHandler.ExecuteSafely(action, functionName);
        }

        /// <summary>
        /// ライセンス設定ダイアログを表示(新規追加)
        /// </summary>
        private void ShowLicenseSettingsDialog()
        {
            try
            {
                using (var dialog = new LicenseSettingsDialog())
                {
                    if (dialog.ShowDialog() == DialogResult.OK)
                    {
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

        // 新規メソッド:ライセンスと更新ステータスの更新
        private void UpdateLicenseAndUpdateStatus(Label lblLicenseStatus, Button btnUpdate)
        {
            try
            {
                var licenseManager = LicenseManager.Instance;
                var status = licenseManager.CurrentStatus;

                if (status != null && status.IsValid)
                {
                    lblLicenseStatus.Text = $"ライセンス: {status.PlanType ?? "Free"}";
                    lblLicenseStatus.ForeColor = Color.Green;
                }
                else
                {
                    lblLicenseStatus.Text = "ライセンス: 未認証";
                    lblLicenseStatus.ForeColor = Color.Red;
                }

                // 更新チェック
                var updateManager = UpdateManager.Instance;
                if (updateManager.IsUpdateAvailable())
                {
                    if (btnUpdate != null)
                    {
                        btnUpdate.Visible = true;
                        btnUpdate.BackColor = Color.Orange;
                        logger.Debug("Update button shown - update available");
                    }
                }
                else
                {
                    if (btnUpdate != null)
                    {
                        btnUpdate.Visible = false;
                        logger.Debug("Update button hidden - no update");
                    }
                }
            }
            catch (Exception ex)
            {
                logger.Error(ex, "Failed to update license and update status");
            }
        }

        private async Task HandleUpdateClick()
        {
            try
            {
                var updateManager = UpdateManager.Instance;
                var result = MessageBox.Show(
                    "新しいバージョンが利用可能です。今すぐ更新しますか?\n\nPowerPointは自動的に再起動されます。",
                    "更新の確認",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Question);

                if (result == DialogResult.Yes)
                {
                    await updateManager.ApplyUpdate();
                }
            }
            catch (Exception ex)
            {
                logger.Error(ex, "Failed to handle update click");
                MessageBox.Show($"更新処理中にエラーが発生しました。\n{ex.Message}",
                    "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// 最小限の機能セットを作成(フォールバック用)
        /// </summary>
        private List<FunctionItem> CreateMinimalFunctionSet()
        {
            return new List<FunctionItem>
            {
                new FunctionItem("TestFunction", "テスト", "テスト機能です", "test.png",
                    () => MessageBox.Show("テスト機能が実行されました", "テスト",
                        MessageBoxButtons.OK, MessageBoxIcon.Information), FunctionCategory.PowerTool)
            };
        }

        /// <summary>
        /// 最小限のUIを作成(エラー時のフォールバック)
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
            }
            base.Dispose(disposing);
        }

        #endregion
    }
}
