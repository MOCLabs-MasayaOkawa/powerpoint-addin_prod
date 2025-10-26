using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using PowerPointEfficiencyAddin.Models;
using NLog;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PowerPointEfficiencyAddin.UI
{
    /// <summary>
    /// Matrix Tuner ダイアログ（新仕様版）
    /// 矩形オブジェクトのマトリックス配置を簡潔に調整
    /// </summary>
    public partial class MatrixTunerDialog : Form
    {
        private static readonly Logger logger = LogManager.GetCurrentClassLogger();

        #region Fields
        private readonly List<ShapeInfo> shapes;
        private readonly PowerPointEfficiencyAddin.Services.PowerToolService.GridInfo gridInfo;
        private readonly List<ShapePosition> originalPositions;

        // 基準値（初期状態）
        private List<ShapePosition> baselinePositions;
        private List<ShapePosition> previewPositions; // プレビュー用の一時的な位置
        private const float DEFAULT_SPACING_CM = 0.2f; // デフォルト間隔 0.2cm
        private const float ADJUSTMENT_UNIT_CM = 0.1f; // 調整単位 0.1cm
        private const float CM_TO_POINTS = 28.3465f; // 1cm = 28.3465 points

        // UI Controls
        private GroupBox grpTarget;
        private RadioButton radioRow;
        private RadioButton radioColumn;
        private NumericUpDown numSizeAdjust;
        private Label lblSizeUnit;
        private Button btnSizeOK; // サイズ調整用OKボタン

        private MatrixPreviewPanel previewPanel;

        private GroupBox grpSpacing;
        private NumericUpDown numSpacing;
        private Label lblSpacingUnit;

        private Button btnSelectAll;
        private Button btnDeselectAll;
        private Button btnSelectOdd;
        private Button btnSelectEven;
        private Button btnSelectEdge;

        private Button btnOK;
        private Button btnApply;
        private Button btnCancel;
        private Button btnReset;

        // State
        private bool[,] selectedCells;
        private Timer updateTimer;
        private bool isUpdating = false;
        private bool isPreview = true; // プレビューモードフラグ
        #endregion

        public MatrixTunerDialog(List<ShapeInfo> shapes, PowerPointEfficiencyAddin.Services.PowerToolService.GridInfo gridInfo)
        {
            this.shapes = shapes;
            this.gridInfo = gridInfo;
            this.originalPositions = SaveCurrentPositions();
            this.baselinePositions = new List<ShapePosition>();
            this.previewPositions = new List<ShapePosition>();
            this.selectedCells = new bool[gridInfo.Rows, gridInfo.Columns];

            // タイマー初期化
            updateTimer = new Timer();
            updateTimer.Interval = 200;
            updateTimer.Tick += UpdateTimer_Tick;

            InitializeComponent();
            InitializeValues();

            // 初期配置（0.2cm間隔）後の状態を基準値として更新
            UpdateBaselinePositions();
        }

        #region Initialization
        private void InitializeComponent()
        {
            // Form設定
            this.Text = "Matrix Tuner";
            this.Size = new Size(500, 550);
            this.MinimumSize = new Size(500, 550);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;

            // サイズ変更対象選択グループ
            grpTarget = new GroupBox
            {
                Text = "サイズ変更対象選択",
                Location = new Point(12, 12),
                Size = new Size(460, 320)
            };

            var lblTarget = new Label
            {
                Text = "対象:",
                Location = new Point(10, 25),
                Size = new Size(40, 20)
            };

            radioRow = new RadioButton
            {
                Text = "行",
                Location = new Point(55, 25),
                Size = new Size(50, 20),
                Checked = true
            };

            radioColumn = new RadioButton
            {
                Text = "列",
                Location = new Point(115, 25),
                Size = new Size(50, 20)
            };

            var lblSize = new Label
            {
                Text = "サイズ調整:",
                Location = new Point(180, 25),
                Size = new Size(80, 20)
            };

            numSizeAdjust = new NumericUpDown
            {
                Location = new Point(265, 22),
                Size = new Size(70, 20),
                Minimum = -5,  // 最小値は-5cm
                Maximum = 5,   // 最大値は5cm
                Value = 0,
                DecimalPlaces = 1,
                Increment = (decimal)ADJUSTMENT_UNIT_CM  // 0.1cm単位
            };

            lblSizeUnit = new Label
            {
                Text = "cm",
                Location = new Point(340, 25),
                Size = new Size(25, 20)
            };

            btnSizeOK = new Button
            {
                Text = "OK",
                Location = new Point(375, 20),
                Size = new Size(50, 25)
            };

            // プレビューパネル
            previewPanel = new MatrixPreviewPanel(gridInfo.Rows, gridInfo.Columns)
            {
                Location = new Point(10, 55),
                Size = new Size(440, 200),
                BorderStyle = BorderStyle.FixedSingle
            };

            // 選択ボタン群
            var buttonY = 265;
            btnSelectAll = new Button
            {
                Text = "全選択",
                Location = new Point(10, buttonY),
                Size = new Size(80, 25)
            };

            btnDeselectAll = new Button
            {
                Text = "全解除",
                Location = new Point(95, buttonY),
                Size = new Size(80, 25)
            };

            btnSelectOdd = new Button
            {
                Text = "奇数",
                Location = new Point(180, buttonY),
                Size = new Size(80, 25)
            };

            btnSelectEven = new Button
            {
                Text = "偶数",
                Location = new Point(265, buttonY),
                Size = new Size(80, 25)
            };

            btnSelectEdge = new Button
            {
                Text = "外枠",
                Location = new Point(350, buttonY),
                Size = new Size(80, 25)
            };

            grpTarget.Controls.AddRange(new Control[] {
                lblTarget, radioRow, radioColumn,
                lblSize, numSizeAdjust, lblSizeUnit, btnSizeOK,
                previewPanel,
                btnSelectAll, btnDeselectAll, btnSelectOdd, btnSelectEven, btnSelectEdge
            });

            // 間隔グループ（最下部）
            grpSpacing = new GroupBox
            {
                Text = "間隔設定（全セル共通）",
                Location = new Point(12, 340),
                Size = new Size(460, 55)
            };

            var lblSpacing = new Label
            {
                Text = "間隔:",
                Location = new Point(10, 23),
                Size = new Size(40, 20)
            };

            numSpacing = new NumericUpDown
            {
                Location = new Point(55, 20),
                Size = new Size(70, 20),
                Minimum = 0,  // 最小値は0cm
                Maximum = 2,  // 最大値は2cm
                Value = (decimal)DEFAULT_SPACING_CM,
                DecimalPlaces = 1,
                Increment = (decimal)ADJUSTMENT_UNIT_CM  // 0.1cm単位
            };

            lblSpacingUnit = new Label
            {
                Text = "cm",
                Location = new Point(130, 23),
                Size = new Size(25, 20)
            };

            grpSpacing.Controls.AddRange(new Control[] { lblSpacing, numSpacing, lblSpacingUnit });

            // ボタン群
            btnOK = new Button
            {
                Text = "OK",
                Location = new Point(220, 410),
                Size = new Size(80, 30),
                DialogResult = DialogResult.OK
            };

            btnApply = new Button
            {
                Text = "適用",
                Location = new Point(310, 410),
                Size = new Size(80, 30)
            };

            btnCancel = new Button
            {
                Text = "キャンセル",
                Location = new Point(400, 410),
                Size = new Size(80, 30),
                DialogResult = DialogResult.Cancel
            };

            btnReset = new Button
            {
                Text = "リセット",
                Location = new Point(12, 410),
                Size = new Size(80, 30)
            };

            // イベントハンドラ設定
            radioRow.CheckedChanged += RadioTarget_CheckedChanged;
            radioColumn.CheckedChanged += RadioTarget_CheckedChanged;
            numSizeAdjust.ValueChanged += NumericUpDown_ValueChanged;
            numSpacing.ValueChanged += NumericUpDown_ValueChanged;

            previewPanel.CellClicked += PreviewPanel_CellClicked;
            previewPanel.RowClicked += PreviewPanel_RowClicked;
            previewPanel.ColumnClicked += PreviewPanel_ColumnClicked;

            btnSelectAll.Click += BtnSelectAll_Click;
            btnDeselectAll.Click += BtnDeselectAll_Click;
            btnSelectOdd.Click += BtnSelectOdd_Click;
            btnSelectEven.Click += BtnSelectEven_Click;
            btnSelectEdge.Click += BtnSelectEdge_Click;

            btnSizeOK.Click += BtnSizeOK_Click;
            btnApply.Click += BtnApply_Click;
            btnReset.Click += BtnReset_Click;
            btnOK.Click += BtnOK_Click;

            // フォームにコントロールを追加
            this.Controls.AddRange(new Control[] { grpTarget, grpSpacing, btnOK, btnApply, btnCancel, btnReset });
        }

        private void InitializeValues()
        {
            // プレビューパネルの初期設定
            previewPanel.SetLockMode(LockMode.Row);
            previewPanel.UpdateLockedCells(selectedCells);

            // 初期化時に0.2cm間隔で再配置
            isUpdating = true;
            try
            {
                var initialSpacingCm = DEFAULT_SPACING_CM;
                var initialSpacingPt = CmToPoints(initialSpacingCm);

                logger.Info($"Applying initial spacing: {initialSpacingCm}cm = {initialSpacingPt:F2}pt");

                // 確実に0.2cm間隔で再配置
                ApplySpacingAdjustmentInternal(initialSpacingPt, false); // プレビューではなく実際に適用

                // ShapeInfoのプロパティも更新（GridInfoとの整合性を保つ）
                foreach (var shape in shapes)
                {
                    shape.Left = shape.Shape.Left;
                    shape.Top = shape.Shape.Top;
                    shape.Width = shape.Shape.Width;
                    shape.Height = shape.Shape.Height;
                }

                logger.Info($"Successfully initialized matrix with {initialSpacingCm}cm spacing");
            }
            catch (Exception ex)
            {
                logger.Error(ex, "Failed to initialize spacing");
            }
            finally
            {
                isUpdating = false;
            }
        }
        #endregion

        #region Helper Methods
        private List<ShapePosition> SaveCurrentPositions()
        {
            var positions = new List<ShapePosition>();
            foreach (var shape in shapes)
            {
                positions.Add(new ShapePosition
                {
                    Shape = shape.Shape,
                    Left = shape.Shape.Left,
                    Top = shape.Shape.Top,
                    Width = shape.Shape.Width,
                    Height = shape.Shape.Height
                });
            }
            return positions;
        }

        private void RestorePositions(List<ShapePosition> positions)
        {
            foreach (var pos in positions)
            {
                pos.Shape.Left = pos.Left;
                pos.Shape.Top = pos.Top;
                pos.Shape.Width = pos.Width;
                pos.Shape.Height = pos.Height;
            }
        }

        /// <summary>
        /// cm単位をポイント単位に変換
        /// </summary>
        private float CmToPoints(float cm)
        {
            // 1cm = 28.3465 points (正確な変換)
            return cm * CM_TO_POINTS;
        }

        /// <summary>
        /// 現在の状態を基準値として保存
        /// </summary>
        private void UpdateBaselinePositions()
        {
            baselinePositions.Clear();
            foreach (var shape in shapes)
            {
                baselinePositions.Add(new ShapePosition
                {
                    Shape = shape.Shape,
                    Left = shape.Shape.Left,
                    Top = shape.Shape.Top,
                    Width = shape.Shape.Width,
                    Height = shape.Shape.Height
                });
            }

            // プレビュー用位置も更新
            previewPositions = new List<ShapePosition>(baselinePositions);
            logger.Debug($"Updated baseline positions for {baselinePositions.Count} shapes");
        }
        #endregion

        #region Event Handlers
        private void RadioTarget_CheckedChanged(object sender, EventArgs e)
        {
            if (isUpdating) return;

            // 選択対象切り替え時に選択状態をリセット
            for (int r = 0; r < gridInfo.Rows; r++)
                for (int c = 0; c < gridInfo.Columns; c++)
                    selectedCells[r, c] = false;

            previewPanel.SetLockMode(radioRow.Checked ? LockMode.Row : LockMode.Column);
            previewPanel.UpdateLockedCells(selectedCells);
        }

        private void NumericUpDown_ValueChanged(object sender, EventArgs e)
        {
            if (isUpdating) return;

            updateTimer.Stop();
            updateTimer.Start();
        }

        private void UpdateTimer_Tick(object sender, EventArgs e)
        {
            updateTimer.Stop();
            ApplyPreviewAdjustments();
        }

        private void PreviewPanel_CellClicked(int row, int col)
        {
            if (radioRow.Checked)
            {
                // 行モード：行全体を切り替え
                bool newState = !selectedCells[row, 0];
                for (int c = 0; c < gridInfo.Columns; c++)
                {
                    selectedCells[row, c] = newState;
                }
            }
            else
            {
                // 列モード：列全体を切り替え
                bool newState = !selectedCells[0, col];
                for (int r = 0; r < gridInfo.Rows; r++)
                {
                    selectedCells[r, col] = newState;
                }
            }

            previewPanel.UpdateLockedCells(selectedCells);
        }

        private void PreviewPanel_RowClicked(int row)
        {
            if (!radioRow.Checked) return;

            bool newState = !selectedCells[row, 0];
            for (int c = 0; c < gridInfo.Columns; c++)
            {
                selectedCells[row, c] = newState;
            }

            previewPanel.UpdateLockedCells(selectedCells);
        }

        private void PreviewPanel_ColumnClicked(int col)
        {
            if (!radioColumn.Checked) return;

            bool newState = !selectedCells[0, col];
            for (int r = 0; r < gridInfo.Rows; r++)
            {
                selectedCells[r, col] = newState;
            }

            previewPanel.UpdateLockedCells(selectedCells);
        }

        private void BtnSelectAll_Click(object sender, EventArgs e)
        {
            for (int r = 0; r < gridInfo.Rows; r++)
                for (int c = 0; c < gridInfo.Columns; c++)
                    selectedCells[r, c] = true;

            previewPanel.UpdateLockedCells(selectedCells);
        }

        private void BtnDeselectAll_Click(object sender, EventArgs e)
        {
            for (int r = 0; r < gridInfo.Rows; r++)
                for (int c = 0; c < gridInfo.Columns; c++)
                    selectedCells[r, c] = false;

            previewPanel.UpdateLockedCells(selectedCells);
        }

        private void BtnSelectOdd_Click(object sender, EventArgs e)
        {
            if (radioRow.Checked)
            {
                for (int r = 0; r < gridInfo.Rows; r++)
                {
                    bool select = (r % 2 == 0); // 0-indexed, so even indices are odd rows
                    for (int c = 0; c < gridInfo.Columns; c++)
                        selectedCells[r, c] = select;
                }
            }
            else
            {
                for (int c = 0; c < gridInfo.Columns; c++)
                {
                    bool select = (c % 2 == 0);
                    for (int r = 0; r < gridInfo.Rows; r++)
                        selectedCells[r, c] = select;
                }
            }

            previewPanel.UpdateLockedCells(selectedCells);
        }

        private void BtnSelectEven_Click(object sender, EventArgs e)
        {
            if (radioRow.Checked)
            {
                for (int r = 0; r < gridInfo.Rows; r++)
                {
                    bool select = (r % 2 == 1); // 0-indexed, so odd indices are even rows
                    for (int c = 0; c < gridInfo.Columns; c++)
                        selectedCells[r, c] = select;
                }
            }
            else
            {
                for (int c = 0; c < gridInfo.Columns; c++)
                {
                    bool select = (c % 2 == 1);
                    for (int r = 0; r < gridInfo.Rows; r++)
                        selectedCells[r, c] = select;
                }
            }

            previewPanel.UpdateLockedCells(selectedCells);
        }

        private void BtnSelectEdge_Click(object sender, EventArgs e)
        {
            // 全解除してから外枠のみ選択
            for (int r = 0; r < gridInfo.Rows; r++)
                for (int c = 0; c < gridInfo.Columns; c++)
                    selectedCells[r, c] = false;

            if (radioRow.Checked)
            {
                // 最初と最後の行
                for (int c = 0; c < gridInfo.Columns; c++)
                {
                    selectedCells[0, c] = true;
                    selectedCells[gridInfo.Rows - 1, c] = true;
                }
            }
            else
            {
                // 最初と最後の列
                for (int r = 0; r < gridInfo.Rows; r++)
                {
                    selectedCells[r, 0] = true;
                    selectedCells[r, gridInfo.Columns - 1] = true;
                }
            }

            previewPanel.UpdateLockedCells(selectedCells);
        }

        private void BtnSizeOK_Click(object sender, EventArgs e)
        {
            // サイズ変更を確定
            ApplyFinalAdjustments();

            // 確定後の状態を基準値として更新
            UpdateBaselinePositions();

            // コントロールをリセット
            isUpdating = true;
            try
            {
                numSizeAdjust.Value = 0;

                // 選択状態もクリア
                for (int r = 0; r < gridInfo.Rows; r++)
                    for (int c = 0; c < gridInfo.Columns; c++)
                        selectedCells[r, c] = false;

                previewPanel.UpdateLockedCells(selectedCells);
            }
            finally
            {
                isUpdating = false;
            }

            logger.Info("Size adjustment applied and reset via Size OK button");
        }

        private void BtnApply_Click(object sender, EventArgs e)
        {
            ApplyFinalAdjustments();
            UpdateBaselinePositions();
            logger.Info("All adjustments applied via Apply button");
        }

        private void BtnOK_Click(object sender, EventArgs e)
        {
            // OK押下時に最終適用
            ApplyFinalAdjustments();
            logger.Info("All adjustments applied via OK button");
        }

        private void BtnReset_Click(object sender, EventArgs e)
        {
            isUpdating = true;
            try
            {
                // 元のサイズに戻す
                RestorePositions(originalPositions);

                // コントロール値をリセット
                numSizeAdjust.Value = 0;
                numSpacing.Value = (decimal)DEFAULT_SPACING_CM;

                // 選択状態もリセット
                for (int r = 0; r < gridInfo.Rows; r++)
                    for (int c = 0; c < gridInfo.Columns; c++)
                        selectedCells[r, c] = false;

                previewPanel.UpdateLockedCells(selectedCells);

                // 0.2cm間隔で再配置
                var initialSpacingPt = CmToPoints(DEFAULT_SPACING_CM);
                ApplySpacingAdjustmentInternal(initialSpacingPt, false);

                // ShapeInfoのプロパティも更新
                foreach (var shape in shapes)
                {
                    shape.Left = shape.Shape.Left;
                    shape.Top = shape.Shape.Top;
                    shape.Width = shape.Shape.Width;
                    shape.Height = shape.Shape.Height;
                }

                // リセット後の状態を基準値として更新
                UpdateBaselinePositions();
            }
            finally
            {
                isUpdating = false;
            }
        }
        #endregion

        #region Adjustment Methods
        /// <summary>
        /// プレビュー用の調整（リアルタイム表示）
        /// </summary>
        private void ApplyPreviewAdjustments()
        {
            if (isUpdating) return;

            isUpdating = true;
            try
            {
                // 1. まずサイズ調整を適用
                var sizeAdjustCm = (float)numSizeAdjust.Value;
                if (sizeAdjustCm != 0)
                {
                    var adjustmentPt = CmToPoints(sizeAdjustCm);
                    ApplySizeAdjustment(adjustmentPt, true); // プレビューモード
                    logger.Debug($"Preview size adjustment: {sizeAdjustCm}cm = {adjustmentPt:F2}pt");
                }
                else
                {
                    // サイズ調整が0の場合は基準サイズに戻す
                    foreach (var pos in baselinePositions)
                    {
                        pos.Shape.Width = pos.Width;
                        pos.Shape.Height = pos.Height;
                    }
                }

                // 2. サイズ調整後、必ず間隔を再適用（間隔を一定に保つ）
                var spacingCm = (float)numSpacing.Value;
                var spacingPt = CmToPoints(spacingCm);
                ApplySpacingAdjustmentInternal(spacingPt, true); // プレビューモード

                logger.Debug($"Preview adjustments: Size={sizeAdjustCm}cm, Spacing={spacingCm}cm");
            }
            catch (Exception ex)
            {
                logger.Error(ex, "Failed to apply preview adjustments");
            }
            finally
            {
                isUpdating = false;
            }
        }

        /// <summary>
        /// 最終的な調整の適用
        /// </summary>
        private void ApplyFinalAdjustments()
        {
            isUpdating = true;
            try
            {
                // 1. サイズ調整を適用
                var sizeAdjustCm = (float)numSizeAdjust.Value;
                if (sizeAdjustCm != 0)
                {
                    var adjustmentPt = CmToPoints(sizeAdjustCm);
                    ApplySizeAdjustment(adjustmentPt, false); // 最終適用
                }

                // 2. 間隔を適用
                var spacingCm = (float)numSpacing.Value;
                var spacingPt = CmToPoints(spacingCm);
                ApplySpacingAdjustmentInternal(spacingPt, false); // 最終適用

                // ShapeInfoのプロパティも更新
                foreach (var shape in shapes)
                {
                    shape.Left = shape.Shape.Left;
                    shape.Top = shape.Shape.Top;
                    shape.Width = shape.Shape.Width;
                    shape.Height = shape.Shape.Height;
                }

                logger.Info($"Final adjustments applied: Size={sizeAdjustCm}cm, Spacing={spacingCm}cm");
            }
            catch (Exception ex)
            {
                logger.Error(ex, "Failed to apply final adjustments");
            }
            finally
            {
                isUpdating = false;
            }
        }

        private void ApplySizeAdjustment(float adjustmentPt, bool isPreview)
        {
            if (radioRow.Checked)
            {
                // 行モード：選択された行の高さを調整
                for (int row = 0; row < gridInfo.Rows; row++)
                {
                    bool isSelected = selectedCells[row, 0];
                    if (isSelected)
                    {
                        foreach (var shape in gridInfo.ShapeGrid[row])
                        {
                            var baseline = baselinePositions.FirstOrDefault(p => p.Shape == shape.Shape);
                            if (baseline != null)
                            {
                                shape.Shape.Height = Math.Max(1, baseline.Height + adjustmentPt);
                            }
                        }
                    }
                }
            }
            else
            {
                // 列モード：選択された列の幅を調整
                for (int col = 0; col < gridInfo.Columns; col++)
                {
                    bool isSelected = selectedCells[0, col];
                    if (isSelected)
                    {
                        for (int row = 0; row < gridInfo.Rows; row++)
                        {
                            if (col < gridInfo.ShapeGrid[row].Count)
                            {
                                var shape = gridInfo.ShapeGrid[row][col];
                                var baseline = baselinePositions.FirstOrDefault(p => p.Shape == shape.Shape);
                                if (baseline != null)
                                {
                                    shape.Shape.Width = Math.Max(1, baseline.Width + adjustmentPt);
                                }
                            }
                        }
                    }
                }
            }
        }

        /// <summary>
        /// 間隔調整の内部実装
        /// </summary>
        private void ApplySpacingAdjustmentInternal(float spacingPt, bool isPreview)
        {
            // 現在のサイズを保持したまま、間隔のみ調整

            // 各行の現在の高さと各列の現在の幅を取得
            var rowHeights = new float[gridInfo.Rows];
            var columnWidths = new float[gridInfo.Columns];

            for (int row = 0; row < gridInfo.Rows; row++)
            {
                float maxHeight = 0;
                foreach (var shape in gridInfo.ShapeGrid[row])
                {
                    maxHeight = Math.Max(maxHeight, shape.Shape.Height);
                }
                rowHeights[row] = maxHeight;
            }

            for (int col = 0; col < gridInfo.Columns; col++)
            {
                float maxWidth = 0;
                for (int row = 0; row < gridInfo.Rows; row++)
                {
                    if (col < gridInfo.ShapeGrid[row].Count)
                    {
                        maxWidth = Math.Max(maxWidth, gridInfo.ShapeGrid[row][col].Shape.Width);
                    }
                }
                columnWidths[col] = maxWidth;
            }

            // グリッドの左上基準点を元の位置から取得
            float baseLeft = float.MaxValue;
            float baseTop = float.MaxValue;
            foreach (var pos in originalPositions)
            {
                baseLeft = Math.Min(baseLeft, pos.Left);
                baseTop = Math.Min(baseTop, pos.Top);
            }

            // 位置を再配置（現在のサイズを維持しつつ、指定された間隔で配置）
            float currentTop = baseTop;
            for (int row = 0; row < gridInfo.Rows; row++)
            {
                float currentLeft = baseLeft;
                for (int col = 0; col < gridInfo.Columns; col++)
                {
                    if (col < gridInfo.ShapeGrid[row].Count)
                    {
                        var shape = gridInfo.ShapeGrid[row][col];

                        // 位置のみ更新（サイズは変更しない）
                        shape.Shape.Left = currentLeft;
                        shape.Shape.Top = currentTop;
                    }

                    currentLeft += columnWidths[col] + spacingPt;
                }
                currentTop += rowHeights[row] + spacingPt;
            }
        }
        #endregion

        #region Inner Classes
        private class ShapePosition
        {
            public PowerPoint.Shape Shape { get; set; }
            public float Left { get; set; }
            public float Top { get; set; }
            public float Width { get; set; }
            public float Height { get; set; }
        }
        #endregion
    }

    #region MatrixPreviewPanel
    /// <summary>
    /// マトリックスプレビューパネル（既存クラスを流用）
    /// </summary>
    public class MatrixPreviewPanel : Panel
    {
        private readonly int rows;
        private readonly int columns;
        private bool[,] lockedCells;
        private LockMode lockMode = LockMode.Row;

        public delegate void CellClickedHandler(int row, int col);
        public event CellClickedHandler CellClicked;

        public delegate void RowClickedHandler(int row);
        public event RowClickedHandler RowClicked;

        public delegate void ColumnClickedHandler(int col);
        public event ColumnClickedHandler ColumnClicked;

        public MatrixPreviewPanel(int rows, int columns)
        {
            this.rows = rows;
            this.columns = columns;
            this.lockedCells = new bool[rows, columns];
            this.DoubleBuffered = true;
            this.SetStyle(ControlStyles.AllPaintingInWmPaint | ControlStyles.UserPaint | ControlStyles.ResizeRedraw, true);
        }

        public void SetLockMode(LockMode mode)
        {
            this.lockMode = mode;
            Invalidate();
        }

        public void UpdateLockedCells(bool[,] lockedCells)
        {
            this.lockedCells = lockedCells;
            Invalidate();
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            base.OnPaint(e);

            var g = e.Graphics;
            g.Clear(Color.White);

            var cellWidth = (float)(Width - 30) / columns;
            var cellHeight = (float)(Height - 30) / rows;

            // ヘッダー領域の背景
            using (var headerBrush = new SolidBrush(Color.FromArgb(240, 240, 240)))
            {
                g.FillRectangle(headerBrush, 0, 0, 30, Height);
                g.FillRectangle(headerBrush, 0, 0, Width, 30);
            }

            // セル描画
            for (int r = 0; r < rows; r++)
            {
                for (int c = 0; c < columns; c++)
                {
                    var x = 30 + c * cellWidth;
                    var y = 30 + r * cellHeight;

                    // セル背景（選択状態により色を変更）
                    var isLocked = lockedCells[r, c];
                    using (var brush = new SolidBrush(isLocked ? Color.FromArgb(255, 200, 200) : Color.White))
                    {
                        g.FillRectangle(brush, x, y, cellWidth, cellHeight);
                    }

                    // セル枠線
                    using (var pen = new Pen(Color.Gray, 1))
                    {
                        g.DrawRectangle(pen, x, y, cellWidth, cellHeight);
                    }
                }
            }

            // 行ヘッダー
            using (var font = new Font("Arial", 8))
            using (var brush = new SolidBrush(Color.Black))
            {
                for (int r = 0; r < rows; r++)
                {
                    var y = 30 + r * cellHeight;
                    var text = (r + 1).ToString();
                    var textSize = g.MeasureString(text, font);
                    g.DrawString(text, font, brush, 15 - textSize.Width / 2, y + cellHeight / 2 - textSize.Height / 2);
                }
            }

            // 列ヘッダー
            using (var font = new Font("Arial", 8))
            using (var brush = new SolidBrush(Color.Black))
            {
                for (int c = 0; c < columns; c++)
                {
                    var x = 30 + c * cellWidth;
                    var text = ((char)('A' + c)).ToString();
                    var textSize = g.MeasureString(text, font);
                    g.DrawString(text, font, brush, x + cellWidth / 2 - textSize.Width / 2, 15 - textSize.Height / 2);
                }
            }
        }

        protected override void OnMouseClick(MouseEventArgs e)
        {
            base.OnMouseClick(e);

            if (e.X < 30 && e.Y >= 30)
            {
                // 行ヘッダークリック
                var row = (int)((e.Y - 30) / ((Height - 30) / (float)rows));
                if (row >= 0 && row < rows)
                {
                    RowClicked?.Invoke(row);
                }
            }
            else if (e.Y < 30 && e.X >= 30)
            {
                // 列ヘッダークリック
                var col = (int)((e.X - 30) / ((Width - 30) / (float)columns));
                if (col >= 0 && col < columns)
                {
                    ColumnClicked?.Invoke(col);
                }
            }
            else if (e.X >= 30 && e.Y >= 30)
            {
                // セルクリック
                var col = (int)((e.X - 30) / ((Width - 30) / (float)columns));
                var row = (int)((e.Y - 30) / ((Height - 30) / (float)rows));
                if (row >= 0 && row < rows && col >= 0 && col < columns)
                {
                    CellClicked?.Invoke(row, col);
                }
            }
        }
    }

    public enum LockMode
    {
        Row,
        Column
    }
    #endregion
}