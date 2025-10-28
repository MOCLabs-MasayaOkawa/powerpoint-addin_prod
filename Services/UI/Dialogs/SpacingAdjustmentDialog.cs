using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using PowerPointEfficiencyAddin.Models;
using NLog;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PowerPointEfficiencyAddin.Services.UI.Dialogs
{
    /// <summary>
    /// 間隔調整ダイアログ（リアルタイム更新対応）
    /// </summary>
    public partial class SpacingAdjustmentDialog : Form
    {
        private static readonly Logger logger = LogManager.GetCurrentClassLogger();

        // コントロール
        private NumericUpDown numSpacing;
        private RadioButton radioResize;
        private RadioButton radioMove;
        private Button btnOK;
        private Button btnCancel;
        private Timer updateTimer;

        // データ
        private readonly List<ShapeInfo> shapes;
        private readonly bool isVertical;
        private readonly float originalSpacing;
        private readonly List<ShapePosition> originalPositions;

        // 設定
        private SpacingSettings currentSettings;
        private bool isUpdating = false;

        public SpacingAdjustmentDialog(string title, float currentSpacing, List<ShapeInfo> shapes, bool isVertical)
        {
            this.shapes = shapes;
            this.isVertical = isVertical;
            originalSpacing = currentSpacing;
            originalPositions = SaveOriginalPositions();

            // タイマー初期化を最初に実行（InitializeComponentより前）
            updateTimer = new Timer();
            updateTimer.Interval = 200; // 200ms間隔で更新
            updateTimer.Tick += UpdateTimer_Tick;

            InitializeComponent();
            Text = title;

            // デフォルト設定
            currentSettings = new SpacingSettings
            {
                Spacing = currentSpacing,
                AdjustmentMethod = SpacingAdjustmentMethod.MoveObjects
            };

            // 初期値設定
            numSpacing.Value = (decimal)Math.Max(0, currentSpacing);
            radioMove.Checked = true;

            logger.Info($"SpacingAdjustmentDialog initialized: {title}, spacing: {currentSpacing:F2}cm");
        }

        private void InitializeComponent()
        {
            SuspendLayout();

            // フォーム設定
            Size = new Size(350, 200);
            StartPosition = FormStartPosition.CenterScreen;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;

            // 間隔調整ラベル
            var labelSpacing = new Label()
            {
                Text = "間隔(cm):",
                Location = new Point(20, 20),
                Size = new Size(80, 20)
            };

            // 間隔調整数値入力
            numSpacing = new NumericUpDown()
            {
                Location = new Point(110, 18),
                Size = new Size(120, 20),
                Minimum = 0.0M,
                Maximum = 50.0M,
                DecimalPlaces = 2,
                Increment = 0.1M
            };
            numSpacing.ValueChanged += NumSpacing_ValueChanged;

            // 調整方法グループボックス
            var groupMethod = new GroupBox()
            {
                Text = "調整方法",
                Location = new Point(20, 60),
                Size = new Size(280, 70)
            };

            // ラジオボタン：サイズ変更
            radioResize = new RadioButton()
            {
                Text = "オブジェクトのサイズを変更して間隔を調整",
                Location = new Point(10, 20),
                Size = new Size(260, 20)
            };
            radioResize.CheckedChanged += RadioButton_CheckedChanged;

            // ラジオボタン：移動
            radioMove = new RadioButton()
            {
                Text = "オブジェクトを移動して間隔を調整",
                Location = new Point(10, 45),
                Size = new Size(260, 20),
                Checked = true
            };
            radioMove.CheckedChanged += RadioButton_CheckedChanged;

            // ボタン
            btnOK = new Button()
            {
                Text = "OK",
                Location = new Point(180, 150),
                Size = new Size(75, 30),
                DialogResult = DialogResult.OK
            };

            btnCancel = new Button()
            {
                Text = "キャンセル",
                Location = new Point(265, 150),
                Size = new Size(75, 30),
                DialogResult = DialogResult.Cancel
            };
            btnCancel.Click += BtnCancel_Click;

            // コントロール追加
            groupMethod.Controls.AddRange(new Control[] { radioResize, radioMove });
            Controls.AddRange(new Control[]
            {
                labelSpacing, numSpacing, groupMethod, btnOK, btnCancel
            });

            AcceptButton = btnOK;
            CancelButton = btnCancel;

            ResumeLayout(false);
        }

        /// <summary>
        /// 元の位置情報を保存
        /// </summary>
        private List<ShapePosition> SaveOriginalPositions()
        {
            if (shapes == null) return new List<ShapePosition>();

            return shapes.Where(s => s?.Shape != null).Select(s => new ShapePosition
            {
                Shape = s.Shape,
                Left = s.Left,
                Top = s.Top,
                Width = s.Width,
                Height = s.Height
            }).ToList();
        }

        /// <summary>
        /// 元の位置に復元
        /// </summary>
        private void RestoreOriginalPositions()
        {
            try
            {
                if (originalPositions == null) return;

                foreach (var pos in originalPositions.Where(p => p?.Shape != null))
                {
                    pos.Shape.Left = pos.Left;
                    pos.Shape.Top = pos.Top;
                    pos.Shape.Width = pos.Width;
                    pos.Shape.Height = pos.Height;
                }
                logger.Debug("Restored original positions");
            }
            catch (Exception ex)
            {
                logger.Error(ex, "Failed to restore original positions");
            }
        }

        /// <summary>
        /// 数値変更イベント
        /// </summary>
        private void NumSpacing_ValueChanged(object sender, EventArgs e)
        {
            if (isUpdating) return;

            currentSettings.Spacing = (float)numSpacing.Value;

            // タイマーを再起動（パフォーマンス考慮）
            if (updateTimer != null)
            {
                updateTimer.Stop();
                updateTimer.Start();
            }
        }

        /// <summary>
        /// ラジオボタン変更イベント
        /// </summary>
        private void RadioButton_CheckedChanged(object sender, EventArgs e)
        {
            if (isUpdating) return;

            if (radioResize.Checked)
            {
                currentSettings.AdjustmentMethod = SpacingAdjustmentMethod.ResizeObjects;
            }
            else
            {
                currentSettings.AdjustmentMethod = SpacingAdjustmentMethod.MoveObjects;
            }

            // 即座に更新
            ApplyCurrentSettings();
        }

        /// <summary>
        /// タイマー更新イベント（パフォーマンス考慮）
        /// </summary>
        private void UpdateTimer_Tick(object sender, EventArgs e)
        {
            if (updateTimer != null)
            {
                updateTimer.Stop();
            }
            ApplyCurrentSettings();
        }

        /// <summary>
        /// キャンセルボタンクリック
        /// </summary>
        private void BtnCancel_Click(object sender, EventArgs e)
        {
            RestoreOriginalPositions();
        }

        /// <summary>
        /// 現在の設定を適用
        /// </summary>
        private void ApplyCurrentSettings()
        {
            try
            {
                if (shapes == null || originalPositions == null) return;

                isUpdating = true;

                var spacingPoints = currentSettings.Spacing * 28.35f; // cmをポイントに変換
                var sortedShapes = isVertical
                    ? shapes.OrderBy(s => s.Top).ToList()
                    : shapes.OrderBy(s => s.Left).ToList();

                if (isVertical)
                {
                    ApplyVerticalSpacingPreview(sortedShapes, spacingPoints);
                }
                else
                {
                    ApplyHorizontalSpacingPreview(sortedShapes, spacingPoints);
                }

                logger.Debug($"Applied spacing preview: {currentSettings.Spacing:F2}cm, method: {currentSettings.AdjustmentMethod}");
            }
            catch (Exception ex)
            {
                logger.Error(ex, "Failed to apply spacing preview");
            }
            finally
            {
                isUpdating = false;
            }
        }

        /// <summary>
        /// 垂直間隔プレビューを適用
        /// </summary>
        private void ApplyVerticalSpacingPreview(List<ShapeInfo> sortedShapes, float spacingPoints)
        {
            var topBound = originalPositions.Min(p => p.Top);
            var bottomBound = originalPositions.Max(p => p.Top + p.Height);

            // 垂直揃え（左端揃え）
            var leftmostX = originalPositions.Min(p => p.Left);

            if (currentSettings.AdjustmentMethod == SpacingAdjustmentMethod.MoveObjects)
            {
                // オブジェクトを移動
                var currentTop = topBound;
                for (int i = 0; i < sortedShapes.Count; i++)
                {
                    var shape = sortedShapes[i];
                    var originalPos = originalPositions.First(p => p.Shape == shape.Shape);

                    shape.Shape.Left = leftmostX; // 左端揃えを維持
                    shape.Shape.Top = currentTop;
                    shape.Shape.Width = originalPos.Width;
                    shape.Shape.Height = originalPos.Height;

                    currentTop += originalPos.Height + spacingPoints;
                }
            }
            else
            {
                // オブジェクトサイズを変更
                var totalSpacing = spacingPoints * (sortedShapes.Count - 1);
                var availableSpace = bottomBound - topBound - totalSpacing;
                var newHeight = Math.Max(5f, availableSpace / sortedShapes.Count); // 最小5pt

                var currentTop = topBound;
                for (int i = 0; i < sortedShapes.Count; i++)
                {
                    var shape = sortedShapes[i];
                    var originalPos = originalPositions.First(p => p.Shape == shape.Shape);

                    shape.Shape.Left = leftmostX; // 左端揃えを維持
                    shape.Shape.Top = currentTop;
                    shape.Shape.Width = originalPos.Width;
                    shape.Shape.Height = newHeight;

                    currentTop += newHeight + spacingPoints;
                }
            }
        }

        /// <summary>
        /// 水平間隔プレビューを適用
        /// </summary>
        private void ApplyHorizontalSpacingPreview(List<ShapeInfo> sortedShapes, float spacingPoints)
        {
            var leftBound = originalPositions.Min(p => p.Left);
            var rightBound = originalPositions.Max(p => p.Left + p.Width);

            // 水平揃え（上端揃え）
            var topmostY = originalPositions.Min(p => p.Top);

            if (currentSettings.AdjustmentMethod == SpacingAdjustmentMethod.MoveObjects)
            {
                // オブジェクトを移動
                var currentLeft = leftBound;
                for (int i = 0; i < sortedShapes.Count; i++)
                {
                    var shape = sortedShapes[i];
                    var originalPos = originalPositions.First(p => p.Shape == shape.Shape);

                    shape.Shape.Top = topmostY; // 上端揃えを維持
                    shape.Shape.Left = currentLeft;
                    shape.Shape.Width = originalPos.Width;
                    shape.Shape.Height = originalPos.Height;

                    currentLeft += originalPos.Width + spacingPoints;
                }
            }
            else
            {
                // オブジェクトサイズを変更
                var totalSpacing = spacingPoints * (sortedShapes.Count - 1);
                var availableSpace = rightBound - leftBound - totalSpacing;
                var newWidth = Math.Max(5f, availableSpace / sortedShapes.Count); // 最小5pt

                var currentLeft = leftBound;
                for (int i = 0; i < sortedShapes.Count; i++)
                {
                    var shape = sortedShapes[i];
                    var originalPos = originalPositions.First(p => p.Shape == shape.Shape);

                    shape.Shape.Top = topmostY; // 上端揃えを維持
                    shape.Shape.Left = currentLeft;
                    shape.Shape.Width = newWidth;
                    shape.Shape.Height = originalPos.Height;

                    currentLeft += newWidth + spacingPoints;
                }
            }
        }

        /// <summary>
        /// 設定を取得
        /// </summary>
        public SpacingSettings GetSettings()
        {
            return currentSettings;
        }

        /// <summary>
        /// リソース解放
        /// </summary>
        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                updateTimer?.Dispose();
            }
            base.Dispose(disposing);
        }

        /// <summary>
        /// 図形位置情報
        /// </summary>
        private class ShapePosition
        {
            public PowerPoint.Shape Shape { get; set; }
            public float Left { get; set; }
            public float Top { get; set; }
            public float Width { get; set; }
            public float Height { get; set; }
        }
    }
}