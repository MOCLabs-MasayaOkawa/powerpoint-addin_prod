using System;
using System.Drawing;
using System.Threading.Tasks;
using System.Windows.Forms;
using NLog;
using PowerPointEfficiencyAddin.Models.Licensing;
using PowerPointEfficiencyAddin.Services.Infrastructure.Licensing;

namespace PowerPointEfficiencyAddin.Services.UI.Dialogs
{
    /// <summary>
    /// ライセンス設定ダイアログ
    /// </summary>
    public class LicenseSettingsDialog : Form
    {
        private static readonly Logger logger = LogManager.GetCurrentClassLogger();

        // UIコンポーネント
        private Label lblTitle;
        private Label lblLicenseKey;
        private TextBox txtLicenseKey;
        private Label lblStatus;
        private Label lblCurrentStatus;
        private Button btnValidate;
        private Button btnOK;
        private Button btnCancel;
        private ProgressBar progressBar;
        private LinkLabel lnkPurchase;

        private LicenseManager licenseManager;
        private bool isValidating = false;

        public LicenseSettingsDialog()
        {
            InitializeComponent();
            LoadCurrentLicenseInfo();
        }

        /// <summary>
        /// コンポーネントを初期化
        /// </summary>
        private void InitializeComponent()
        {
            Text = "ライセンス設定";
            Size = new Size(500, 350);
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            StartPosition = FormStartPosition.CenterParent;
            Icon = SystemIcons.Shield;

            // タイトル
            lblTitle = new Label
            {
                Text = "PowerPoint効率化アドイン - ライセンス設定",
                Location = new Point(20, 20),
                Size = new Size(440, 30),
                Font = new Font("Segoe UI", 12F, FontStyle.Bold),
                ForeColor = Color.FromArgb(0, 51, 102)
            };

            // 現在の状態表示
            lblCurrentStatus = new Label
            {
                Text = "現在の状態:",
                Location = new Point(20, 60),
                Size = new Size(80, 20),
                Font = new Font("Segoe UI", 9F)
            };

            lblStatus = new Label
            {
                Text = "確認中...",
                Location = new Point(100, 60),
                Size = new Size(360, 40),
                Font = new Font("Segoe UI", 9F),
                ForeColor = Color.Gray,
                AutoSize = false
            };

            // ライセンスキー入力
            lblLicenseKey = new Label
            {
                Text = "ライセンスキー:",
                Location = new Point(20, 110),
                Size = new Size(100, 20),
                Font = new Font("Segoe UI", 9F)
            };

            txtLicenseKey = new TextBox
            {
                Location = new Point(20, 135),
                Size = new Size(350, 25),
                Font = new Font("Consolas", 10F),
                MaxLength = 100,
                CharacterCasing = CharacterCasing.Upper
            };
            txtLicenseKey.TextChanged += TxtLicenseKey_TextChanged;

            // 検証ボタン
            btnValidate = new Button
            {
                Text = "検証",
                Location = new Point(380, 133),
                Size = new Size(80, 27),
                Font = new Font("Segoe UI", 9F),
                Enabled = false
            };
            btnValidate.Click += BtnValidate_Click;

            // プログレスバー
            progressBar = new ProgressBar
            {
                Location = new Point(20, 170),
                Size = new Size(440, 10),
                Style = ProgressBarStyle.Marquee,
                Visible = false
            };

            // 購入リンク
            lnkPurchase = new LinkLabel
            {
                Text = "ライセンスをお持ちでない方はこちらから購入",
                Location = new Point(20, 190),
                Size = new Size(300, 20),
                Font = new Font("Segoe UI", 9F)
            };
            lnkPurchase.LinkClicked += LnkPurchase_LinkClicked;

            // 区切り線
            var separator = new Label
            {
                BorderStyle = BorderStyle.Fixed3D,
                Location = new Point(20, 220),
                Size = new Size(440, 2)
            };

            // OKボタン
            btnOK = new Button
            {
                Text = "OK",
                Location = new Point(285, 240),
                Size = new Size(85, 30),
                Font = new Font("Segoe UI", 9F),
                DialogResult = DialogResult.OK,
                Enabled = false
            };

            // キャンセルボタン
            btnCancel = new Button
            {
                Text = "キャンセル",
                Location = new Point(375, 240),
                Size = new Size(85, 30),
                Font = new Font("Segoe UI", 9F),
                DialogResult = DialogResult.Cancel
            };

            // 注意事項
            var lblNote = new Label
            {
                Text = "※ ライセンスキーは大文字・小文字を区別しません\n※ インターネット接続が必要です",
                Location = new Point(20, 280),
                Size = new Size(440, 35),
                Font = new Font("Segoe UI", 8F),
                ForeColor = Color.Gray
            };

            // コントロールを追加
            Controls.AddRange(new Control[] {
                lblTitle, lblCurrentStatus, lblStatus,
                lblLicenseKey, txtLicenseKey, btnValidate,
                progressBar, lnkPurchase, separator,
                btnOK, btnCancel, lblNote
            });
        }

        /// <summary>
        /// 現在のライセンス情報を読み込み
        /// </summary>
        private void LoadCurrentLicenseInfo()
        {
            try
            {
                licenseManager = LicenseManager.Instance;
                var status = licenseManager.CurrentStatus;

                if (status != null)
                {
                    UpdateStatusDisplay(status);

                    // 既存のライセンスキーがある場合は表示（マスク）
                    var cache = new LicenseCache(
                        new Infrastructure.Security.RegistryManager());
                    var cachedKey = cache.GetCachedLicenseKey();

                    if (!string.IsNullOrEmpty(cachedKey))
                    {
                        // 最初と最後の数文字のみ表示
                        if (cachedKey.Length > 10)
                        {
                            txtLicenseKey.Text = cachedKey.Substring(0, 5) +
                                new string('*', cachedKey.Length - 8) +
                                cachedKey.Substring(cachedKey.Length - 3);
                        }
                    }
                }
                else
                {
                    lblStatus.Text = "ライセンスが登録されていません";
                    lblStatus.ForeColor = Color.Orange;
                }
            }
            catch (Exception ex)
            {
                logger.Error(ex, "Failed to load current license info");
                lblStatus.Text = "ライセンス情報の読み込みに失敗しました";
                lblStatus.ForeColor = Color.Red;
            }
        }

        /// <summary>
        /// ライセンス状態表示を更新
        /// </summary>
        private void UpdateStatusDisplay(LicenseStatus status)
        {
            string statusText = "";
            Color statusColor = Color.Gray;

            if (status.IsValid)
            {
                statusText = $"有効 - {status.PlanType ?? "Standard"}プラン";

                if (status.ExpiryDate.HasValue)
                {
                    var daysRemaining = (status.ExpiryDate.Value - DateTime.Now).TotalDays;
                    if (daysRemaining > 0)
                    {
                        statusText += $"\n有効期限: {status.ExpiryDate.Value:yyyy/MM/dd} (残り{(int)daysRemaining}日)";
                        statusColor = Color.Green;
                    }
                    else
                    {
                        statusText += "\n期限切れ";
                        statusColor = Color.Red;
                    }
                }
                else
                {
                    statusColor = Color.Green;
                }

                if (status.IsOfflineMode)
                {
                    statusText += $"\nオフラインモード (残り{status.GetOfflineGraceDaysRemaining()}日)";
                    statusColor = Color.Orange;
                }
            }
            else
            {
                switch (status.AccessLevel)
                {
                    case FeatureAccessLevel.Free:
                        statusText = "制限モード - 一部機能のみ利用可能";
                        statusColor = Color.Orange;
                        break;
                    case FeatureAccessLevel.Blocked:
                        statusText = "無効 - ライセンスの更新が必要です";
                        statusColor = Color.Red;
                        break;
                    default:
                        statusText = status.Message ?? "ライセンス未登録";
                        statusColor = Color.Gray;
                        break;
                }
            }

            lblStatus.Text = statusText;
            lblStatus.ForeColor = statusColor;
        }

        /// <summary>
        /// ライセンスキー入力変更時
        /// </summary>
        private void TxtLicenseKey_TextChanged(object sender, EventArgs e)
        {
            // マスクされた文字列でない新規入力の場合のみ検証ボタンを有効化
            bool isNewInput = !txtLicenseKey.Text.Contains("*") &&
                             txtLicenseKey.Text.Length >= 10;
            btnValidate.Enabled = isNewInput && !isValidating;
        }

        /// <summary>
        /// 検証ボタンクリック時
        /// </summary>
        private async void BtnValidate_Click(object sender, EventArgs e)
        {
            if (isValidating || string.IsNullOrWhiteSpace(txtLicenseKey.Text))
                return;

            isValidating = true;
            btnValidate.Enabled = false;
            btnOK.Enabled = false;
            progressBar.Visible = true;

            try
            {
                logger.Info("Starting license validation");

                // 非同期でライセンス検証
                var result = await licenseManager.SetLicenseKeyAsync(txtLicenseKey.Text.Trim());

                progressBar.Visible = false;

                if (result.IsSuccess)
                {
                    MessageBox.Show(
                        "ライセンスが正常に認証されました。",
                        "認証成功",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information);

                    // 状態表示を更新
                    UpdateStatusDisplay(licenseManager.CurrentStatus);
                    btnOK.Enabled = true;
                }
                else
                {
                    MessageBox.Show(
                        $"ライセンス認証に失敗しました。\n\n{result.Message}",
                        "認証失敗",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning);

                    lblStatus.Text = result.Message;
                    lblStatus.ForeColor = Color.Red;
                }
            }
            catch (Exception ex)
            {
                logger.Error(ex, "License validation error");
                progressBar.Visible = false;

                MessageBox.Show(
                    "ライセンス検証中にエラーが発生しました。\nインターネット接続を確認してください。",
                    "エラー",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
            finally
            {
                isValidating = false;
                btnValidate.Enabled = !string.IsNullOrWhiteSpace(txtLicenseKey.Text) &&
                                     !txtLicenseKey.Text.Contains("*");
            }
        }

        /// <summary>
        /// 購入リンククリック時
        /// </summary>
        private void LnkPurchase_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            try
            {
                // TODO: 実際の購入ページURLに変更
                System.Diagnostics.Process.Start("https://www.yourcompany.com/purchase");
            }
            catch (Exception ex)
            {
                logger.Error(ex, "Failed to open purchase link");
                MessageBox.Show(
                    "Webページを開けませんでした。",
                    "エラー",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }
    }
}