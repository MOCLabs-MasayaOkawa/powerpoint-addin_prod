using System;
using System.Drawing;
using System.Windows.Forms;
using NLog;

namespace PowerPointEfficiencyAddin.Services.UI.Dialogs
{
    /// <summary>
    /// セルマージン設定ダイアログ
    /// </summary>
    public partial class MarginAdjustmentDialog : Form
    {
        private static readonly Logger logger = LogManager.GetCurrentClassLogger();

        #region UI Controls
        // プリセット選択
        private RadioButton radioNone;
        private RadioButton radioNormal;
        private RadioButton radioNarrow;
        private RadioButton radioWide;
        private RadioButton radioCustomize;

        // カスタム数値入力
        private NumericUpDown numTop;
        private NumericUpDown numBottom;
        private NumericUpDown numLeft;
        private NumericUpDown numRight;

        // ボタン
        private Button btnOK;
        private Button btnCancel;
        #endregion

        #region Data
        private float currentTop;
        private float currentBottom;
        private float currentLeft;
        private float currentRight;
        private bool isUpdating = false;
        #endregion

        public MarginAdjustmentDialog(string title = "セルマージン設定")
        {
            InitializeComponent();
            Text = title;

            // デフォルト値設定（Normalプリセット）
            SetPresetValues(0.13f, 0.13f, 0.25f, 0.25f);
            radioNormal.Checked = true;

            logger.Info($"MarginAdjustmentDialog initialized: {title}");
        }

        private void InitializeComponent()
        {
            SuspendLayout();

            // フォーム設定
            Size = new Size(380, 320);
            StartPosition = FormStartPosition.CenterScreen;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;

            // プリセットグループボックス
            var groupPresets = new GroupBox()
            {
                Text = "プリセット",
                Location = new Point(20, 20),
                Size = new Size(320, 140)
            };

            // ラジオボタン：None
            radioNone = new RadioButton()
            {
                Text = "None (全て 0cm)",
                Location = new Point(15, 25),
                Size = new Size(150, 20)
            };
            radioNone.CheckedChanged += RadioPreset_CheckedChanged;

            // ラジオボタン：Normal
            radioNormal = new RadioButton()
            {
                Text = "Normal (上下:0.13cm, 左右:0.25cm)",
                Location = new Point(15, 50),
                Size = new Size(280, 20)
            };
            radioNormal.CheckedChanged += RadioPreset_CheckedChanged;

            // ラジオボタン：Narrow
            radioNarrow = new RadioButton()
            {
                Text = "Narrow (上下:0.13cm, 左右:0.13cm)",
                Location = new Point(15, 75),
                Size = new Size(280, 20)
            };
            radioNarrow.CheckedChanged += RadioPreset_CheckedChanged;

            // ラジオボタン：Wide
            radioWide = new RadioButton()
            {
                Text = "Wide (全て 0.38cm)",
                Location = new Point(15, 100),
                Size = new Size(200, 20)
            };
            radioWide.CheckedChanged += RadioPreset_CheckedChanged;

            // ラジオボタン：Customize
            radioCustomize = new RadioButton()
            {
                Text = "Customize",
                Location = new Point(15, 115),
                Size = new Size(100, 20)
            };
            radioCustomize.CheckedChanged += RadioPreset_CheckedChanged;

            // カスタム設定グループボックス
            var groupCustom = new GroupBox()
            {
                Text = "詳細設定 (cm)",
                Location = new Point(20, 170),
                Size = new Size(320, 80)
            };

            // Top
            var labelTop = new Label()
            {
                Text = "上:",
                Location = new Point(15, 25),
                Size = new Size(30, 20)
            };
            numTop = new NumericUpDown()
            {
                Location = new Point(50, 23),
                Size = new Size(60, 20),
                Minimum = 0.0M,
                Maximum = 5.0M,
                DecimalPlaces = 2,
                Increment = 0.01M,
                Value = 0.13M
            };
            numTop.ValueChanged += NumCustom_ValueChanged;

            // Bottom
            var labelBottom = new Label()
            {
                Text = "下:",
                Location = new Point(125, 25),
                Size = new Size(30, 20)
            };
            numBottom = new NumericUpDown()
            {
                Location = new Point(160, 23),
                Size = new Size(60, 20),
                Minimum = 0.0M,
                Maximum = 5.0M,
                DecimalPlaces = 2,
                Increment = 0.01M,
                Value = 0.13M
            };
            numBottom.ValueChanged += NumCustom_ValueChanged;

            // Left
            var labelLeft = new Label()
            {
                Text = "左:",
                Location = new Point(15, 50),
                Size = new Size(30, 20)
            };
            numLeft = new NumericUpDown()
            {
                Location = new Point(50, 48),
                Size = new Size(60, 20),
                Minimum = 0.0M,
                Maximum = 5.0M,
                DecimalPlaces = 2,
                Increment = 0.01M,
                Value = 0.25M
            };
            numLeft.ValueChanged += NumCustom_ValueChanged;

            // Right
            var labelRight = new Label()
            {
                Text = "右:",
                Location = new Point(125, 50),
                Size = new Size(30, 20)
            };
            numRight = new NumericUpDown()
            {
                Location = new Point(160, 48),
                Size = new Size(60, 20),
                Minimum = 0.0M,
                Maximum = 5.0M,
                DecimalPlaces = 2,
                Increment = 0.01M,
                Value = 0.25M
            };
            numRight.ValueChanged += NumCustom_ValueChanged;

            // ボタン
            btnOK = new Button()
            {
                Text = "OK",
                Location = new Point(210, 270),
                Size = new Size(75, 30),
                DialogResult = DialogResult.OK
            };

            btnCancel = new Button()
            {
                Text = "キャンセル",
                Location = new Point(295, 270),
                Size = new Size(75, 30),
                DialogResult = DialogResult.Cancel
            };

            // コントロール追加
            groupPresets.Controls.AddRange(new Control[] { radioNone, radioNormal, radioNarrow, radioWide, radioCustomize });
            groupCustom.Controls.AddRange(new Control[] { labelTop, numTop, labelBottom, numBottom, labelLeft, numLeft, labelRight, numRight });
            Controls.AddRange(new Control[] { groupPresets, groupCustom, btnOK, btnCancel });

            AcceptButton = btnOK;
            CancelButton = btnCancel;

            ResumeLayout(false);
        }

        /// <summary>
        /// プリセットラジオボタン選択イベント
        /// </summary>
        private void RadioPreset_CheckedChanged(object sender, EventArgs e)
        {
            if (isUpdating) return;

            var radio = sender as RadioButton;
            if (radio == null || !radio.Checked) return;

            isUpdating = true;

            try
            {
                if (radio == radioNone)
                {
                    SetPresetValues(0.0f, 0.0f, 0.0f, 0.0f);
                }
                else if (radio == radioNormal)
                {
                    SetPresetValues(0.13f, 0.13f, 0.25f, 0.25f);
                }
                else if (radio == radioNarrow)
                {
                    SetPresetValues(0.13f, 0.13f, 0.13f, 0.13f);
                }
                else if (radio == radioWide)
                {
                    SetPresetValues(0.38f, 0.38f, 0.38f, 0.38f);
                }
                // radioCustomize の場合は何もしない（現在の数値を維持）
            }
            finally
            {
                isUpdating = false;
            }
        }

        /// <summary>
        /// カスタム数値変更イベント
        /// </summary>
        private void NumCustom_ValueChanged(object sender, EventArgs e)
        {
            if (isUpdating) return;

            // カスタム数値が変更されたらCustomizeラジオボタンを選択
            isUpdating = true;
            radioCustomize.Checked = true;
            isUpdating = false;

            UpdateCurrentValues();
        }

        /// <summary>
        /// プリセット値を設定
        /// </summary>
        private void SetPresetValues(float top, float bottom, float left, float right)
        {
            numTop.Value = (decimal)top;
            numBottom.Value = (decimal)bottom;
            numLeft.Value = (decimal)left;
            numRight.Value = (decimal)right;

            UpdateCurrentValues();
        }

        /// <summary>
        /// 現在の値を更新
        /// </summary>
        private void UpdateCurrentValues()
        {
            currentTop = (float)numTop.Value;
            currentBottom = (float)numBottom.Value;
            currentLeft = (float)numLeft.Value;
            currentRight = (float)numRight.Value;
        }

        /// <summary>
        /// 設定値を取得
        /// </summary>
        public (float top, float bottom, float left, float right) GetMarginValues()
        {
            return (currentTop, currentBottom, currentLeft, currentRight);
        }
    }
}