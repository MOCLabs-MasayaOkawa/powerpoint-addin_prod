using System;
using System.Drawing;
using System.Windows.Forms;
using PowerPointEfficiencyAddin.Models;
using PowerPointEfficiencyAddin.Services;
using PowerPointEfficiencyAddin.Utils;
using NLog;

namespace PowerPointEfficiencyAddin.UI
{
    /// <summary>
    /// 図形スタイル設定ダイアログ
    /// 塗りつぶし色、枠線色、フォント、フォント色の設定UI
    /// </summary>
    public partial class ShapeStyleDialog : Form
    {
        #region フィールド

        private static readonly Logger logger = LogManager.GetCurrentClassLogger();
        private ShapeStyleSettings _settings;
        private ShapeStyleSettings _originalSettings;

        // UIコントロール
        private CheckBox chkEnableStyling;
        private Button btnFillColor;
        private Button btnLineColor;
        private Button btnFontColor;
        private Button btnOK;
        private Button btnCancel;
        private Button btnReset;
        private GroupBox grpSettings;

        #endregion

        #region プロパティ

        /// <summary>
        /// 設定された図形スタイル設定を取得
        /// </summary>
        public ShapeStyleSettings Settings => _settings?.Clone();

        #endregion

        #region コンストラクタ

        /// <summary>
        /// 現在の設定を基に初期化
        /// </summary>
        public ShapeStyleDialog()
        {
            try
            {
                logger.Info("Initializing ShapeStyleDialog");

                // 現在の設定を読み込み
                _originalSettings = SettingsService.Instance.LoadShapeStyleSettings();
                _settings = _originalSettings.Clone();

                InitializeComponent();
                InitializeValues();

                logger.Info("ShapeStyleDialog initialized successfully");
            }
            catch (Exception ex)
            {
                logger.Error(ex, "Failed to initialize ShapeStyleDialog");
                throw;
            }
        }

        /// <summary>
        /// 指定設定で初期化
        /// </summary>
        /// <param name="initialSettings">初期設定</param>
        public ShapeStyleDialog(ShapeStyleSettings initialSettings) : this()
        {
            if (initialSettings != null)
            {
                _settings = initialSettings.Clone();
                _originalSettings = initialSettings.Clone();
                InitializeValues();
            }
        }

        #endregion

        #region フォーム設計

        /// <summary>
        /// UIコンポーネントを初期化
        /// </summary>
        private void InitializeComponent()
        {
            SuspendLayout();

            // フォーム設定
            Text = "図形スタイル設定";
            Size = new Size(420, 300);
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            StartPosition = FormStartPosition.CenterParent;
            ShowInTaskbar = false;

            // スタイル有効化チェックボックス
            chkEnableStyling = new CheckBox
            {
                Text = "図形スタイリングを有効にする",
                Location = new Point(20, 20),
                Size = new Size(350, 25),
                Font = new Font("Segoe UI", 9F)
            };
            chkEnableStyling.CheckedChanged += ChkEnableStyling_CheckedChanged;

            // 設定グループボックス
            grpSettings = new GroupBox
            {
                Text = "スタイル設定",
                Location = new Point(15, 55),
                Size = new Size(380, 120),
                Font = new Font("Segoe UI", 9F)
            };

            // 塗りつぶし色ボタン
            var lblFillColor = new Label
            {
                Text = "塗りつぶし色:",
                Location = new Point(20, 30),
                Size = new Size(100, 20),
                Font = new Font("Segoe UI", 9F)
            };

            btnFillColor = new Button
            {
                Location = new Point(130, 25),
                Size = new Size(60, 30),
                FlatStyle = FlatStyle.Flat,
                Text = ""
            };
            btnFillColor.Click += BtnFillColor_Click;

            // 枠線色ボタン
            var lblLineColor = new Label
            {
                Text = "枠線色:",
                Location = new Point(220, 30),
                Size = new Size(60, 20),
                Font = new Font("Segoe UI", 9F)
            };

            btnLineColor = new Button
            {
                Location = new Point(290, 25),
                Size = new Size(60, 30),
                FlatStyle = FlatStyle.Flat,
                Text = ""
            };
            btnLineColor.Click += BtnLineColor_Click;

            // フォント色ボタン
            var lblFontColor = new Label
            {
                Text = "フォント色:",
                Location = new Point(20, 80),
                Size = new Size(80, 20),
                Font = new Font("Segoe UI", 9F)
            };

            btnFontColor = new Button
            {
                Location = new Point(130, 75),
                Size = new Size(60, 30),
                FlatStyle = FlatStyle.Flat,
                Text = ""
            };
            btnFontColor.Click += BtnFontColor_Click;

            // ボタン群
            btnOK = new Button
            {
                Text = "OK",
                Location = new Point(155, 190),
                Size = new Size(75, 30),
                DialogResult = DialogResult.OK,
                Font = new Font("Segoe UI", 9F)
            };
            btnOK.Click += BtnOK_Click;

            btnCancel = new Button
            {
                Text = "キャンセル",
                Location = new Point(240, 190),
                Size = new Size(75, 30),
                DialogResult = DialogResult.Cancel,
                Font = new Font("Segoe UI", 9F)
            };

            btnReset = new Button
            {
                Text = "リセット",
                Location = new Point(325, 190),
                Size = new Size(65, 30),
                Font = new Font("Segoe UI", 9F)
            };
            btnReset.Click += BtnReset_Click;

            // コントロールをフォームに追加
            Controls.Add(chkEnableStyling);
            Controls.Add(grpSettings);
            Controls.Add(btnOK);
            Controls.Add(btnCancel);
            Controls.Add(btnReset);

            // グループボックスにコントロールを追加
            grpSettings.Controls.Add(lblFillColor);
            grpSettings.Controls.Add(btnFillColor);
            grpSettings.Controls.Add(lblLineColor);
            grpSettings.Controls.Add(btnLineColor);
            grpSettings.Controls.Add(lblFontColor);
            grpSettings.Controls.Add(btnFontColor);

            AcceptButton = btnOK;
            CancelButton = btnCancel;

            ResumeLayout(false);
        }

        #endregion

        #region 初期化・更新

        /// <summary>
        /// コントロールに現在の設定値を反映
        /// </summary>
        private void InitializeValues()
        {
            try
            {
                chkEnableStyling.Checked = _settings.EnableStyling;

                UpdateColorButton(btnFillColor, _settings.FillColor);
                UpdateColorButton(btnLineColor, _settings.LineColor);
                UpdateColorButton(btnFontColor, _settings.FontColor);

                UpdateControlsEnabled();
            }
            catch (Exception ex)
            {
                logger.Error(ex, "Failed to initialize dialog values");
            }
        }

        /// <summary>
        /// コントロールの有効/無効状態を更新
        /// </summary>
        private void UpdateControlsEnabled()
        {
            grpSettings.Enabled = chkEnableStyling.Checked;
        }

        /// <summary>
        /// 色ボタンの外観を更新
        /// </summary>
        /// <param name="button">色ボタン</param>
        /// <param name="color">表示する色</param>
        private void UpdateColorButton(Button button, Color color)
        {
            button.BackColor = color;

            // 色が暗い場合は白い境界線、明るい場合は黒い境界線
            var brightness = (color.R + color.G + color.B) / 3;
            button.ForeColor = brightness > 128 ? Color.Black : Color.White;
            button.FlatAppearance.BorderColor = brightness > 128 ? Color.Black : Color.White;
        }

        #endregion

        #region イベントハンドラ

        /// <summary>
        /// スタイリング有効化チェックボックス変更時
        /// </summary>
        private void ChkEnableStyling_CheckedChanged(object sender, EventArgs e)
        {
            try
            {
                _settings.EnableStyling = chkEnableStyling.Checked;
                UpdateControlsEnabled();
                logger.Debug($"Styling enabled changed to: {_settings.EnableStyling}");
            }
            catch (Exception ex)
            {
                logger.Error(ex, "Error in ChkEnableStyling_CheckedChanged");
            }
        }

        /// <summary>
        /// 塗りつぶし色ボタンクリック時
        /// </summary>
        private void BtnFillColor_Click(object sender, EventArgs e)
        {
            ShowColorDialog(_settings.FillColor, color =>
            {
                _settings.FillColor = color;
                UpdateColorButton(btnFillColor, color);
                logger.Debug($"Fill color changed to: {color}");
            });
        }

        /// <summary>
        /// 枠線色ボタンクリック時
        /// </summary>
        private void BtnLineColor_Click(object sender, EventArgs e)
        {
            ShowColorDialog(_settings.LineColor, color =>
            {
                _settings.LineColor = color;
                UpdateColorButton(btnLineColor, color);
                logger.Debug($"Line color changed to: {color}");
            });
        }

        /// <summary>
        /// フォント色ボタンクリック時
        /// </summary>
        private void BtnFontColor_Click(object sender, EventArgs e)
        {
            ShowColorDialog(_settings.FontColor, color =>
            {
                _settings.FontColor = color;
                UpdateColorButton(btnFontColor, color);
                logger.Debug($"Font color changed to: {color}");
            });
        }


        /// <summary>
        /// OKボタンクリック時
        /// </summary>
        private void BtnOK_Click(object sender, EventArgs e)
        {
            try
            {
                // 設定を保存
                if (SettingsService.Instance.SaveShapeStyleSettings(_settings))
                {
                    logger.Info("Shape style settings saved successfully");
                    DialogResult = DialogResult.OK;
                    Close();
                }
                else
                {
                    MessageBox.Show(this, "設定の保存に失敗しました。", "エラー",
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            catch (Exception ex)
            {
                logger.Error(ex, "Error saving settings");
                MessageBox.Show(this, "設定の保存中にエラーが発生しました。", "エラー",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// リセットボタンクリック時
        /// </summary>
        private void BtnReset_Click(object sender, EventArgs e)
        {
            try
            {
                var result = MessageBox.Show(this,
                    "すべての設定をデフォルト値にリセットしますか？",
                    "設定リセット",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Question);

                if (result == DialogResult.Yes)
                {
                    _settings.SetDefaults();
                    InitializeValues();
                    logger.Info("Settings reset to defaults");
                }
            }
            catch (Exception ex)
            {
                logger.Error(ex, "Error resetting settings");
            }
        }

        #endregion

        #region ヘルパーメソッド

        /// <summary>
        /// カラーダイアログを表示
        /// </summary>
        /// <param name="currentColor">現在の色</param>
        /// <param name="onColorSelected">色選択時のコールバック</param>
        private void ShowColorDialog(Color currentColor, Action<Color> onColorSelected)
        {
            try
            {
                using (var colorDialog = new ColorDialog())
                {
                    colorDialog.Color = currentColor;
                    colorDialog.FullOpen = true;
                    colorDialog.AnyColor = true;

                    if (colorDialog.ShowDialog(this) == DialogResult.OK)
                    {
                        onColorSelected?.Invoke(colorDialog.Color);
                    }
                }
            }
            catch (Exception ex)
            {
                logger.Error(ex, "Error in color selection");
                MessageBox.Show(this, "色選択でエラーが発生しました。", "エラー",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        #endregion

        #region リソース解放

        /// <summary>
        /// リソースを解放
        /// </summary>
        /// <param name="disposing">マネージドリソースも解放するか</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                // 明示的に作成したコントロールの解放
                chkEnableStyling?.Dispose();
                btnFillColor?.Dispose();
                btnLineColor?.Dispose();
                btnFontColor?.Dispose();
                btnOK?.Dispose();
                btnCancel?.Dispose();
                btnReset?.Dispose();
                grpSettings?.Dispose();
            }

            base.Dispose(disposing);
            logger.Debug("ShapeStyleDialog disposed");
        }

        #endregion
    }
}