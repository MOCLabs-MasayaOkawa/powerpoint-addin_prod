using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Windows.Forms;
using Microsoft.Office.Core;

namespace PowerPointEfficiencyAddin.Services.UI.Dialogs
{
    /// <summary>
    /// 行間区切り線設定ダイアログ（PowerPoint標準準拠版）
    /// </summary>
    public partial class LineSeparatorDialog : Form
    {
        #region プロパティ

        /// <summary>
        /// 選択された線の種類
        /// </summary>
        public MsoLineDashStyle LineStyle { get; private set; }

        /// <summary>
        /// 選択された線の太さ（ポイント）
        /// </summary>
        public float LineWeight { get; private set; }

        /// <summary>
        /// 選択された線の色（RGB）
        /// </summary>
        public Color LineColor { get; private set; }

        #endregion

        #region コンストラクタ

        /// <summary>
        /// 線設定ダイアログを初期化します
        /// </summary>
        public LineSeparatorDialog()
        {
            InitializeComponent();
            InitializeSettings();
        }

        #endregion

        #region 初期化

        /// <summary>
        /// 設定値を初期化します
        /// </summary>
        private void InitializeSettings()
        {
            // 線の種類を設定（OwnerDrawで線サンプル表示）
            comboLineStyle.DrawMode = DrawMode.OwnerDrawFixed;
            comboLineStyle.ItemHeight = 20; // サンプル表示用の高さ
            comboLineStyle.DropDownStyle = ComboBoxStyle.DropDownList;
            comboLineStyle.Items.Clear();
            comboLineStyle.Items.Add(new LineStyleItem("実線", MsoLineDashStyle.msoLineSolid));
            comboLineStyle.Items.Add(new LineStyleItem("点線（角）", MsoLineDashStyle.msoLineSquareDot));
            comboLineStyle.Items.Add(new LineStyleItem("点線（丸）", MsoLineDashStyle.msoLineRoundDot));
            comboLineStyle.Items.Add(new LineStyleItem("破線", MsoLineDashStyle.msoLineDash));
            comboLineStyle.Items.Add(new LineStyleItem("一点鎖線", MsoLineDashStyle.msoLineDashDot));
            comboLineStyle.Items.Add(new LineStyleItem("長破線", MsoLineDashStyle.msoLineLongDash));
            comboLineStyle.Items.Add(new LineStyleItem("長鎖線", MsoLineDashStyle.msoLineLongDashDot));
            comboLineStyle.SelectedIndex = 0; // 実線をデフォルト

            // 線の太さを設定（PowerPoint標準：0.25pt単位）
            comboLineWeight.Items.Clear();
            comboLineWeight.Items.Add("0.25pt");
            comboLineWeight.Items.Add("0.5pt");
            comboLineWeight.Items.Add("0.75pt");
            comboLineWeight.Items.Add("1.0pt");
            comboLineWeight.Items.Add("1.25pt");
            comboLineWeight.Items.Add("1.5pt");
            comboLineWeight.Items.Add("2.0pt");
            comboLineWeight.Items.Add("3.0pt");
            comboLineWeight.Items.Add("4.5pt");
            comboLineWeight.Items.Add("6.0pt");
            comboLineWeight.SelectedIndex = 3; // 1.0ptをデフォルト

            // デフォルト値を設定
            LineStyle = MsoLineDashStyle.msoLineSolid;
            LineWeight = 1.0f;
            LineColor = Color.Black;

            // イベントハンドラを設定
            comboLineStyle.SelectedIndexChanged += ComboLineStyle_SelectedIndexChanged;
            comboLineStyle.DrawItem += ComboLineStyle_DrawItem;
            comboLineWeight.SelectedIndexChanged += ComboLineWeight_SelectedIndexChanged;
            buttonSelectColor.Click += ButtonSelectColor_Click;

            // 色表示ラベルの初期設定
            UpdateColorDisplay();
        }

        #endregion

        #region イベントハンドラ

        /// <summary>
        /// 線の種類が変更された時の処理
        /// </summary>
        private void ComboLineStyle_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboLineStyle.SelectedItem is LineStyleItem item)
            {
                LineStyle = item.Style;
            }
        }

        /// <summary>
        /// 線の種類コンボボックスの描画処理（カスタム描画）
        /// </summary>
        private void ComboLineStyle_DrawItem(object sender, DrawItemEventArgs e)
        {
            if (e.Index < 0) return;

            var combo = (ComboBox)sender;
            var item = (LineStyleItem)combo.Items[e.Index];

            e.DrawBackground();

            // 線のサンプルを描画
            using (var pen = CreatePenFromStyle(item.Style, 2.0f, Color.Black))
            {
                var lineY = e.Bounds.Y + e.Bounds.Height / 2;
                var lineStartX = e.Bounds.X + 5;
                var lineEndX = e.Bounds.X + 120;

                e.Graphics.DrawLine(pen, lineStartX, lineY, lineEndX, lineY);
            }

            // テキストを描画
            using (var brush = new SolidBrush(e.ForeColor))
            {
                var textRect = new Rectangle(e.Bounds.X + 130, e.Bounds.Y,
                    e.Bounds.Width - 135, e.Bounds.Height);
                e.Graphics.DrawString(item.Name, e.Font, brush, textRect,
                    new StringFormat
                    {
                        Alignment = StringAlignment.Near,
                        LineAlignment = StringAlignment.Center
                    });
            }

            e.DrawFocusRectangle();
        }

        /// <summary>
        /// 線の太さが変更された時の処理
        /// </summary>
        private void ComboLineWeight_SelectedIndexChanged(object sender, EventArgs e)
        {
            var weightText = comboLineWeight.SelectedItem?.ToString() ?? "1.0pt";
            var weightValue = weightText.Replace("pt", "");

            if (float.TryParse(weightValue, out float weight))
            {
                LineWeight = weight;
            }
        }

        /// <summary>
        /// 色選択ボタンがクリックされた時の処理
        /// </summary>
        private void ButtonSelectColor_Click(object sender, EventArgs e)
        {
            using (var colorDialog = new ColorDialog())
            {
                colorDialog.Color = LineColor;
                colorDialog.FullOpen = true;

                // PowerPoint標準色を追加
                colorDialog.CustomColors = new int[]
                {
                    ColorTranslator.ToOle(Color.Black),
                    ColorTranslator.ToOle(Color.White),
                    ColorTranslator.ToOle(Color.Red),
                    ColorTranslator.ToOle(Color.Green),
                    ColorTranslator.ToOle(Color.Blue),
                    ColorTranslator.ToOle(Color.Yellow),
                    ColorTranslator.ToOle(Color.Magenta),
                    ColorTranslator.ToOle(Color.Cyan),
                    ColorTranslator.ToOle(Color.Orange),
                    ColorTranslator.ToOle(Color.Purple),
                    ColorTranslator.ToOle(Color.Brown),
                    ColorTranslator.ToOle(Color.Gray),
                    ColorTranslator.ToOle(Color.LightGray),
                    ColorTranslator.ToOle(Color.DarkGray),
                    ColorTranslator.ToOle(Color.Navy),
                    ColorTranslator.ToOle(Color.DarkRed)
                };

                if (colorDialog.ShowDialog() == DialogResult.OK)
                {
                    LineColor = colorDialog.Color;
                    UpdateColorDisplay();
                }
            }
        }

        /// <summary>
        /// 色表示を更新します
        /// </summary>
        private void UpdateColorDisplay()
        {
            labelColorDisplay.BackColor = LineColor;
            labelColorDisplay.Text = $"RGB({LineColor.R}, {LineColor.G}, {LineColor.B})";
            labelColorDisplay.ForeColor = GetContrastColor(LineColor);
        }

        /// <summary>
        /// 背景色に対してコントラストの高い前景色を取得します
        /// </summary>
        private Color GetContrastColor(Color backgroundColor)
        {
            var luminance = (0.299 * backgroundColor.R + 0.587 * backgroundColor.G + 0.114 * backgroundColor.B) / 255;
            return luminance > 0.5 ? Color.Black : Color.White;
        }

        /// <summary>
        /// OKボタンがクリックされた時の処理
        /// </summary>
        private void ButtonOK_Click(object sender, EventArgs e)
        {
            DialogResult = DialogResult.OK;
            Close();
        }

        /// <summary>
        /// キャンセルボタンがクリックされた時の処理
        /// </summary>
        private void ButtonCancel_Click(object sender, EventArgs e)
        {
            DialogResult = DialogResult.Cancel;
            Close();
        }

        #endregion

        #region ヘルパーメソッド

        /// <summary>
        /// 線スタイルからPenを作成します
        /// </summary>
        private Pen CreatePenFromStyle(MsoLineDashStyle style, float width, Color color)
        {
            var pen = new Pen(color, width);

            switch (style)
            {
                case MsoLineDashStyle.msoLineSolid:
                    pen.DashStyle = DashStyle.Solid;
                    break;
                case MsoLineDashStyle.msoLineSquareDot:
                case MsoLineDashStyle.msoLineRoundDot:
                    pen.DashStyle = DashStyle.Dot;
                    break;
                case MsoLineDashStyle.msoLineDash:
                    pen.DashStyle = DashStyle.Dash;
                    break;
                case MsoLineDashStyle.msoLineDashDot:
                    pen.DashStyle = DashStyle.DashDot;
                    break;
                case MsoLineDashStyle.msoLineLongDash:
                    pen.DashPattern = new float[] { 8.0f, 3.0f };
                    break;
                case MsoLineDashStyle.msoLineLongDashDot:
                    pen.DashPattern = new float[] { 8.0f, 3.0f, 2.0f, 3.0f };
                    break;
                default:
                    pen.DashStyle = DashStyle.Solid;
                    break;
            }

            return pen;
        }

        #endregion

        #region 内部クラス

        /// <summary>
        /// 線スタイル項目クラス
        /// </summary>
        private class LineStyleItem
        {
            public string Name { get; }
            public MsoLineDashStyle Style { get; }

            public LineStyleItem(string name, MsoLineDashStyle style)
            {
                Name = name;
                Style = style;
            }

            public override string ToString()
            {
                return Name;
            }
        }

        #endregion

        #region デザイナー生成コード

        private System.ComponentModel.IContainer components = null;
        private ComboBox comboLineStyle;
        private ComboBox comboLineWeight;
        private Button buttonSelectColor;
        private Label labelColorDisplay;
        private Button buttonOK;
        private Button buttonCancel;
        private Label labelLineStyle;
        private Label labelLineWeight;
        private Label labelColor;

        /// <summary>
        /// 必要なデザイナー変数です。
        /// </summary>
        protected override void Dispose(bool disposing)
        {
            if (disposing && components != null)
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        /// <summary>
        /// デザイナー サポートに必要なメソッドです。このメソッドの内容を
        /// コード エディターで変更しないでください。
        /// </summary>
        private void InitializeComponent()
        {
            comboLineStyle = new ComboBox();
            comboLineWeight = new ComboBox();
            buttonSelectColor = new Button();
            labelColorDisplay = new Label();
            buttonOK = new Button();
            buttonCancel = new Button();
            labelLineStyle = new Label();
            labelLineWeight = new Label();
            labelColor = new Label();

            SuspendLayout();

            // 
            // labelLineStyle
            // 
            labelLineStyle.Location = new Point(15, 15);
            labelLineStyle.Size = new Size(80, 23);
            labelLineStyle.Text = "線の種類：";
            labelLineStyle.TextAlign = ContentAlignment.MiddleLeft;

            // 
            // comboLineStyle
            // 
            comboLineStyle.DropDownStyle = ComboBoxStyle.DropDownList;
            comboLineStyle.Location = new Point(100, 15);
            comboLineStyle.Size = new Size(280, 23);

            // 
            // labelLineWeight
            // 
            labelLineWeight.Location = new Point(15, 50);
            labelLineWeight.Size = new Size(80, 23);
            labelLineWeight.Text = "線の太さ：";
            labelLineWeight.TextAlign = ContentAlignment.MiddleLeft;

            // 
            // comboLineWeight
            // 
            comboLineWeight.DropDownStyle = ComboBoxStyle.DropDownList;
            comboLineWeight.Location = new Point(100, 50);
            comboLineWeight.Size = new Size(100, 23);

            // 
            // labelColor
            // 
            labelColor.Location = new Point(15, 85);
            labelColor.Size = new Size(80, 23);
            labelColor.Text = "線の色：";
            labelColor.TextAlign = ContentAlignment.MiddleLeft;

            // 
            // buttonSelectColor
            // 
            buttonSelectColor.Location = new Point(100, 85);
            buttonSelectColor.Size = new Size(100, 25);
            buttonSelectColor.Text = "色を選択...";
            buttonSelectColor.UseVisualStyleBackColor = true;

            // 
            // labelColorDisplay
            // 
            labelColorDisplay.Location = new Point(210, 85);
            labelColorDisplay.Size = new Size(150, 25);
            labelColorDisplay.BorderStyle = BorderStyle.FixedSingle;
            labelColorDisplay.TextAlign = ContentAlignment.MiddleCenter;
            labelColorDisplay.BackColor = Color.Black;
            labelColorDisplay.Text = "RGB(0, 0, 0)";
            labelColorDisplay.ForeColor = Color.White;

            // 
            // buttonOK
            // 
            buttonOK.Location = new Point(200, 130);
            buttonOK.Size = new Size(80, 30);
            buttonOK.Text = "OK";
            buttonOK.UseVisualStyleBackColor = true;
            buttonOK.Click += ButtonOK_Click;

            // 
            // buttonCancel
            // 
            buttonCancel.Location = new Point(290, 130);
            buttonCancel.Size = new Size(80, 30);
            buttonCancel.Text = "キャンセル";
            buttonCancel.UseVisualStyleBackColor = true;
            buttonCancel.Click += ButtonCancel_Click;

            // 
            // LineSeparatorDialog
            // 
            ClientSize = new Size(400, 180);
            Controls.Add(labelLineStyle);
            Controls.Add(comboLineStyle);
            Controls.Add(labelLineWeight);
            Controls.Add(comboLineWeight);
            Controls.Add(labelColor);
            Controls.Add(buttonSelectColor);
            Controls.Add(labelColorDisplay);
            Controls.Add(buttonOK);
            Controls.Add(buttonCancel);
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            StartPosition = FormStartPosition.CenterParent;
            Text = "行間区切り線の設定";

            ResumeLayout(false);
        }

        #endregion
    }
}