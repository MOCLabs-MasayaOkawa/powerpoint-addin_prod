using System;
using System.ComponentModel;
using System.Drawing;
using System.IO;
using System.Threading.Tasks;
using System.Windows.Forms;
using ImageMagick;
using NLog;

namespace PowerPointEfficiencyAddin.Services
{
    /// <summary>
    /// 高機能画像圧縮設定ダイアログ
    /// </summary>
    public partial class ImageCompressionDialog : Form
    {
        private static readonly Logger logger = LogManager.GetCurrentClassLogger();

        #region UI Controls
        // モード選択
        private RadioButton radioAutoMode;
        private RadioButton radioManualMode;

        // 出力形式
        private RadioButton radioJPEG;
        private RadioButton radioPngReduced;
        private RadioButton radioPngLossless;

        // JPEG設定
        private TrackBar trackBarQuality;
        private Label labelQuality;
        private Panel panelJpegSettings;

        // PNG減色設定
        private RadioButton radio256Colors;
        private RadioButton radio192Colors;
        private RadioButton radio128Colors;
        private Panel panelPngSettings;

        // 長辺リサイズ
        private RadioButton radioResizeOff;
        private RadioButton radioResize640;
        private RadioButton radioResize800;
        private RadioButton radioResize1024;
        private RadioButton radioResize1280;
        private RadioButton radioResize1920;
        private RadioButton radioResize2560;

        // 透過処理
        private RadioButton radioTransparencyAuto;
        private RadioButton radioTransparencyRemove;

        // メタデータ
        private RadioButton radioMetadataRemove;
        private RadioButton radioMetadataKeep;

        // ファイルサイズ情報
        private Label labelOriginalSize;
        private Label labelCompressedSize;
        private Label labelReduction;

        // ボタン
        private Button btnExecute;
        private Button btnCancel;

        // プログレス
        private ProgressBar progressBar;
        private Label labelProgress;
        #endregion

        #region Data
        private readonly byte[] originalImageData;
        private readonly MagickImage originalImage;
        private readonly ImageAnalysisResult analysisResult;
        private AdvancedCompressionSettings settings;
        private BackgroundWorker compressionWorker;
        private bool isCompressing = false;
        private bool isInitializing = true;
        #endregion

        public AdvancedCompressionSettings CompressionSettings => settings;

        public ImageCompressionDialog(byte[] imageData)
        {
            this.originalImageData = imageData;
            this.originalImage = new MagickImage(imageData);

            // 画像解析を実行
            this.analysisResult = AnalyzeImage(originalImage);

            // 自動設定を生成
            this.settings = GenerateAutoSettings(analysisResult);

            InitializeComponent();
            InitializeWorker();
            SetupInitialValues();

            isInitializing = false;
            UpdateUI();
            StartCompressionCalculation();
        }

        private void InitializeComponent()
        {
            this.SuspendLayout();

            // フォーム設定（幅と高さを適切に設定）
            this.Text = "画像圧縮";
            this.Size = new Size(520, 670); // 高さを650に拡大
            this.StartPosition = FormStartPosition.CenterScreen;
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;

            int yPos = 20;

            // モード選択グループ
            var groupMode = CreateGroupBox("モード", 20, yPos, 460, 50);
            radioAutoMode = CreateRadioButton("自動（推奨）", 15, 25, 120, 20, true);
            radioManualMode = CreateRadioButton("手動設定", 150, 25, 100, 20, false);

            groupMode.Controls.AddRange(new Control[] { radioAutoMode, radioManualMode });
            yPos += 70;

            // 出力形式グループ（高さを詰めて行間を調整）
            var groupFormat = CreateGroupBox("出力形式", 20, yPos, 460, 100); // 高さを100に調整
            radioJPEG = CreateRadioButton("JPEG", 15, 22, 80, 20, false); // Y位置を22に調整
            radioPngReduced = CreateRadioButton("PNG（減色）", 15, 44, 120, 20, false); // Y位置を44に調整（1pt間隔）
            radioPngLossless = CreateRadioButton("PNG（無減色）", 15, 66, 140, 20, false); // Y位置を66に調整（1pt間隔）

            // JPEG設定パネル
            panelJpegSettings = new Panel()
            {
                Location = new Point(170, 18), // Y位置を18に調整
                Size = new Size(270, 30),
                Visible = false
            };
            var labelJpegQuality = new Label()
            {
                Text = "品質(%):",
                Location = new Point(0, 8),
                Size = new Size(60, 20)
            };
            trackBarQuality = new TrackBar()
            {
                Location = new Point(65, 0),
                Size = new Size(160, 30),
                Minimum = 10,
                Maximum = 100,
                Value = 85,
                TickFrequency = 10
            };
            labelQuality = new Label()
            {
                Text = "85",
                Location = new Point(230, 8),
                Size = new Size(30, 20)
            };
            panelJpegSettings.Controls.AddRange(new Control[] { labelJpegQuality, trackBarQuality, labelQuality });

            // PNG減色設定パネル
            panelPngSettings = new Panel()
            {
                Location = new Point(170, 40), // Y位置を40に調整
                Size = new Size(270, 30),
                Visible = false
            };
            var labelPngColors = new Label()
            {
                Text = "色数:",
                Location = new Point(0, 5),
                Size = new Size(40, 20)
            };
            radio256Colors = CreateRadioButton("256色", 45, 3, 60, 20, true);
            radio192Colors = CreateRadioButton("192色", 110, 3, 60, 20, false);
            radio128Colors = CreateRadioButton("128色", 175, 3, 60, 20, false);
            panelPngSettings.Controls.AddRange(new Control[] { labelPngColors, radio256Colors, radio192Colors, radio128Colors });

            groupFormat.Controls.AddRange(new Control[] { radioJPEG, radioPngReduced, radioPngLossless, panelJpegSettings, panelPngSettings });
            yPos += 120; // 出力形式グループの高さ分だけ進める

            // 長辺リサイズグループ（コントロールの配置を2行に分ける）
            var groupResize = CreateGroupBox("長辺リサイズ", 20, yPos, 460, 100); // 高さを100に拡大
            radioResizeOff = CreateRadioButton("オフ", 15, 25, 50, 20, true);
            radioResize640 = CreateRadioButton("640", 70, 25, 50, 20, false);
            radioResize800 = CreateRadioButton("800", 125, 25, 50, 20, false);
            radioResize1024 = CreateRadioButton("1024", 180, 25, 60, 20, false); // 幅を60に拡大
                                                                                 // 2行目に配置
            radioResize1280 = CreateRadioButton("1280", 15, 50, 60, 20, false); // 2行目、幅を60に拡大
            radioResize1920 = CreateRadioButton("1920", 80, 50, 60, 20, false); // 2行目、幅を60に拡大
            radioResize2560 = CreateRadioButton("2560", 145, 50, 60, 20, false); // 2行目、幅を60に拡大

            var labelResizeNote = new Label()
            {
                Text = "※元画像の長辺より大きい値を選んだ場合はリサイズ処理なし",
                Location = new Point(15, 75),
                Size = new Size(430, 20),
                Font = new Font(this.Font.FontFamily, 7.5f),
                ForeColor = Color.Gray
            };

            groupResize.Controls.AddRange(new Control[] {
                 radioResizeOff, radioResize640, radioResize800, radioResize1024,
                radioResize1280, radioResize1920, radioResize2560, labelResizeNote
            });
            yPos += 120; // 高さ分だけ進める

            // 透過処理グループ
            var groupTransparency = CreateGroupBox("透過処理", 20, yPos, 460, 50);
            radioTransparencyAuto = CreateRadioButton("自動（透過保持）", 15, 25, 150, 20, true); // 幅を150に拡大
            radioTransparencyRemove = CreateRadioButton("透過削除", 180, 25, 100, 20, false); // 位置調整、幅を100に拡大
            groupTransparency.Controls.AddRange(new Control[] { radioTransparencyAuto, radioTransparencyRemove });
            yPos += 70;

            // メタデータグループ
            var groupMetadata = CreateGroupBox("メタデータ", 20, yPos, 460, 50);
            radioMetadataRemove = CreateRadioButton("削除する", 15, 25, 100, 20, true); // 幅を100に拡大
            radioMetadataKeep = CreateRadioButton("削除しない", 130, 25, 100, 20, false); // 位置調整、幅を100に拡大
            groupMetadata.Controls.AddRange(new Control[] { radioMetadataRemove, radioMetadataKeep });
            yPos += 70;

            // ファイルサイズ情報グループ
            var groupFileSize = CreateGroupBox("ファイルサイズ見込み", 20, yPos, 460, 80);
            labelOriginalSize = new Label()
            {
                Text = $"圧縮前: {FormatFileSize(originalImageData.Length)}",
                Location = new Point(15, 25),
                Size = new Size(140, 20)
            };
            labelCompressedSize = new Label()
            {
                Text = "圧縮後: 計算中...",
                Location = new Point(170, 25),
                Size = new Size(140, 20)
            };
            labelReduction = new Label()
            {
                Text = "",
                Location = new Point(320, 25),
                Size = new Size(120, 20),
                ForeColor = Color.Green,
                Font = new Font(this.Font, FontStyle.Bold)
            };
            var labelSizeNote = new Label()
            {
                Text = "※圧縮後サイズが元以上の場合は処理を行わない",
                Location = new Point(15, 50),
                Size = new Size(430, 20),
                Font = new Font(this.Font.FontFamily, 7.5f),
                ForeColor = Color.Gray
            };
            groupFileSize.Controls.AddRange(new Control[] { labelOriginalSize, labelCompressedSize, labelReduction, labelSizeNote });
            yPos += 100;

            // プログレスバー
            progressBar = new ProgressBar()
            {
                Location = new Point(20, yPos),
                Size = new Size(250, 20), // 幅を350に調整（ボタンのスペースを確保）
                Style = ProgressBarStyle.Continuous,
                Visible = false
            };
            labelProgress = new Label()
            {
                Text = "",
                Location = new Point(20, yPos + 25),
                Size = new Size(350, 20),
                Visible = false
            };

            // ボタン（右端に配置）
            btnExecute = new Button()
            {
                Text = "実行",
                Location = new Point(310, yPos), // プログレスバーと同じ高さに配置
                Size = new Size(70, 30),
                DialogResult = DialogResult.OK,
                Enabled = false
            };
            btnCancel = new Button()
            {
                Text = "キャンセル",
                Location = new Point(385, yPos), // 
                Size = new Size(100, 30),
                DialogResult = DialogResult.Cancel
            };

            // イベントハンドラ設定
            radioAutoMode.CheckedChanged += Mode_CheckedChanged;
            radioManualMode.CheckedChanged += Mode_CheckedChanged;
            radioJPEG.CheckedChanged += Format_CheckedChanged;
            radioPngReduced.CheckedChanged += Format_CheckedChanged;
            radioPngLossless.CheckedChanged += Format_CheckedChanged;
            trackBarQuality.ValueChanged += Settings_Changed;
            radio256Colors.CheckedChanged += Settings_Changed;
            radio192Colors.CheckedChanged += Settings_Changed;
            radio128Colors.CheckedChanged += Settings_Changed;

            // リサイズ関連
            radioResizeOff.CheckedChanged += Settings_Changed;
            radioResize640.CheckedChanged += Settings_Changed;
            radioResize800.CheckedChanged += Settings_Changed;
            radioResize1024.CheckedChanged += Settings_Changed;
            radioResize1280.CheckedChanged += Settings_Changed;
            radioResize1920.CheckedChanged += Settings_Changed;
            radioResize2560.CheckedChanged += Settings_Changed;

            radioTransparencyAuto.CheckedChanged += Settings_Changed;
            radioTransparencyRemove.CheckedChanged += Settings_Changed;
            radioMetadataRemove.CheckedChanged += Settings_Changed;
            radioMetadataKeep.CheckedChanged += Settings_Changed;

            // コントロール追加
            this.Controls.AddRange(new Control[]
            {
                groupMode, groupFormat, groupResize, groupTransparency, groupMetadata,
                 groupFileSize, progressBar, labelProgress, btnExecute, btnCancel
            });

            this.AcceptButton = btnExecute;
            this.CancelButton = btnCancel;
            this.ResumeLayout(false);
        }

        #region UI Helper Methods
        private GroupBox CreateGroupBox(string text, int x, int y, int width, int height)
        {
            return new GroupBox()
            {
                Text = text,
                Location = new Point(x, y),
                Size = new Size(width, height)
            };
        }

        private RadioButton CreateRadioButton(string text, int x, int y, int width, int height, bool isChecked)
        {
            return new RadioButton()
            {
                Text = text,
                Location = new Point(x, y),
                Size = new Size(width, height),
                Checked = isChecked
            };
        }
        #endregion

        #region Event Handlers
        private void Mode_CheckedChanged(object sender, EventArgs e)
        {
            if (isInitializing) return;
            UpdateUI();
            StartCompressionCalculation();
        }

        private void Format_CheckedChanged(object sender, EventArgs e)
        {
            if (isInitializing) return;
            UpdateFormatPanels();
            UpdateSettingsFromUI();
            StartCompressionCalculation();
        }

        private void Settings_Changed(object sender, EventArgs e)
        {
            if (isInitializing) return;

            if (sender == trackBarQuality)
            {
                labelQuality.Text = trackBarQuality.Value.ToString();
            }

            UpdateSettingsFromUI();
            StartCompressionCalculation();
        }
        #endregion

        #region UI Update Methods
        private void UpdateUI()
        {
            bool isManualMode = radioManualMode.Checked;

            // 手動モード時のみ各項目を活性化
            EnableControlsRecursively(this, isManualMode, new Control[] {
                radioAutoMode, radioManualMode, btnExecute, btnCancel,
                labelOriginalSize, labelCompressedSize, labelReduction,
                progressBar, labelProgress
            });

            UpdateFormatPanels();
        }

        private void UpdateFormatPanels()
        {
            panelJpegSettings.Visible = radioJPEG.Checked;
            panelPngSettings.Visible = radioPngReduced.Checked;
        }

        private void EnableControlsRecursively(Control parent, bool enabled, Control[] exceptions)
        {
            foreach (Control control in parent.Controls)
            {
                if (Array.IndexOf(exceptions, control) == -1)
                {
                    if (control.HasChildren)
                    {
                        EnableControlsRecursively(control, enabled, exceptions);
                    }
                    else
                    {
                        control.Enabled = enabled;
                    }
                }
            }
        }

        private void SetupInitialValues()
        {
            // 自動判定結果をUIに反映
            switch (settings.OutputFormat)
            {
                case OutputFormat.JPEG:
                    radioJPEG.Checked = true;
                    break;
                case OutputFormat.PngReduced:
                    radioPngReduced.Checked = true;
                    break;
                case OutputFormat.PngLossless:
                    radioPngLossless.Checked = true;
                    break;
            }

            trackBarQuality.Value = settings.JpegQuality;
            labelQuality.Text = settings.JpegQuality.ToString();

            switch (settings.PngColors)
            {
                case 256:
                    radio256Colors.Checked = true;
                    break;
                case 192:
                    radio192Colors.Checked = true;
                    break;
                case 128:
                    radio128Colors.Checked = true;
                    break;
            }

            // リサイズ設定
            if (settings.MaxDimension == 0)
                radioResizeOff.Checked = true;
            else if (settings.MaxDimension <= 640)
                radioResize640.Checked = true;
            else if (settings.MaxDimension <= 800)
                radioResize800.Checked = true;
            else if (settings.MaxDimension <= 1024)
                radioResize1024.Checked = true;
            else if (settings.MaxDimension <= 1280)
                radioResize1280.Checked = true;
            else if (settings.MaxDimension <= 1920)
                radioResize1920.Checked = true;
            else
                radioResize2560.Checked = true;

            radioTransparencyAuto.Checked = settings.PreserveTransparency;
            radioTransparencyRemove.Checked = !settings.PreserveTransparency;
            radioMetadataRemove.Checked = settings.RemoveMetadata;
            radioMetadataKeep.Checked = !settings.RemoveMetadata;
        }

        private void UpdateSettingsFromUI()
        {
            if (radioAutoMode.Checked) return; // 自動モードでは設定変更なし

            // 出力形式
            if (radioJPEG.Checked)
                settings.OutputFormat = OutputFormat.JPEG;
            else if (radioPngReduced.Checked)
                settings.OutputFormat = OutputFormat.PngReduced;
            else if (radioPngLossless.Checked)
                settings.OutputFormat = OutputFormat.PngLossless;

            // JPEG品質
            settings.JpegQuality = trackBarQuality.Value;

            // PNG色数
            if (radio256Colors.Checked)
                settings.PngColors = 256;
            else if (radio192Colors.Checked)
                settings.PngColors = 192;
            else if (radio128Colors.Checked)
                settings.PngColors = 128;

            // リサイズ
            if (radioResizeOff.Checked)
                settings.MaxDimension = 0;
            else if (radioResize640.Checked)
                settings.MaxDimension = 640;
            else if (radioResize800.Checked)
                settings.MaxDimension = 800;
            else if (radioResize1024.Checked)
                settings.MaxDimension = 1024;
            else if (radioResize1280.Checked)
                settings.MaxDimension = 1280;
            else if (radioResize1920.Checked)
                settings.MaxDimension = 1920;
            else if (radioResize2560.Checked)
                settings.MaxDimension = 2560;

            // 透過・メタデータ
            settings.PreserveTransparency = radioTransparencyAuto.Checked;
            settings.RemoveMetadata = radioMetadataRemove.Checked;
        }
        #endregion

        #region Image Analysis and Auto Settings
        private ImageAnalysisResult AnalyzeImage(MagickImage image)
        {
            var result = new ImageAnalysisResult();

            try
            {
                result.Width = image.Width;
                result.Height = image.Height;
                result.HasTransparency = image.HasAlpha;
                result.ColorCount = (int)image.TotalColors;
                result.Format = image.Format.ToString().ToLower();

                // 長辺計算
                result.LongerSide = Math.Max(image.Width, image.Height);

                // 画像種類判定（簡易版）
                if (result.HasTransparency)
                {
                    result.ImageType = ImageType.Graphic; // 透明ありは図/アイコン扱い
                }
                else if (result.ColorCount > 50000 || image.Format == MagickFormat.Jpeg)
                {
                    result.ImageType = ImageType.Photo; // 多色またはJPEGは写真扱い
                }
                else
                {
                    result.ImageType = ImageType.Graphic; // その他は図/スクショ扱い
                }

                logger.Debug($"Image analysis: {result.Width}x{result.Height}, " +
                           $"colors: {result.ColorCount}, type: {result.ImageType}, " +
                           $"transparency: {result.HasTransparency}");
            }
            catch (Exception ex)
            {
                logger.Error(ex, "Failed to analyze image");
            }

            return result;
        }

        private AdvancedCompressionSettings GenerateAutoSettings(ImageAnalysisResult analysis)
        {
            var settings = new AdvancedCompressionSettings
            {
                OriginalSize = originalImageData.Length,
                RemoveMetadata = true,
                PreserveTransparency = true
            };

            // 長辺リサイズ判定（800px基準）
            if (analysis.LongerSide > 800)
            {
                settings.MaxDimension = 800;
            }
            else
            {
                settings.MaxDimension = 0; // リサイズなし
            }

            // 出力形式判定
            if (analysis.HasTransparency)
            {
                // 透明あり → PNG
                if (analysis.ImageType == ImageType.Photo)
                {
                    settings.OutputFormat = OutputFormat.PngLossless; // 写真寄り
                }
                else
                {
                    settings.OutputFormat = OutputFormat.PngReduced; // 図/スクショ
                    settings.PngColors = analysis.ColorCount > 256 ? 256 :
                                       analysis.ColorCount > 192 ? 192 : 128;
                }
            }
            else
            {
                // 透明なし
                if (analysis.ImageType == ImageType.Photo)
                {
                    settings.OutputFormat = OutputFormat.JPEG;
                    settings.JpegQuality = 85;
                }
                else
                {
                    settings.OutputFormat = OutputFormat.PngReduced;
                    settings.PngColors = 256;
                }
            }

            logger.Debug($"Auto settings generated: format={settings.OutputFormat}, " +
                       $"quality={settings.JpegQuality}, colors={settings.PngColors}, " +
                       $"resize={settings.MaxDimension}");

            return settings;
        }
        #endregion

        #region Compression Worker
        private void InitializeWorker()
        {
            compressionWorker = new BackgroundWorker()
            {
                WorkerReportsProgress = true,
                WorkerSupportsCancellation = true
            };

            compressionWorker.DoWork += CompressionWorker_DoWork;
            compressionWorker.ProgressChanged += CompressionWorker_ProgressChanged;
            compressionWorker.RunWorkerCompleted += CompressionWorker_RunWorkerCompleted;
        }

        private void StartCompressionCalculation()
        {
            if (isCompressing) return;

            isCompressing = true;
            btnExecute.Enabled = false;
            progressBar.Visible = true;
            labelProgress.Visible = true;
            labelProgress.Text = "圧縮サイズを計算中...";

            labelCompressedSize.Text = "圧縮後: 計算中...";
            labelReduction.Text = "";

            compressionWorker.RunWorkerAsync(settings.Clone());
        }

        private void CompressionWorker_DoWork(object sender, DoWorkEventArgs e)
        {
            try
            {
                var testSettings = e.Argument as AdvancedCompressionSettings;
                compressionWorker.ReportProgress(10, "画像処理中...");

                using (var testImage = new MagickImage(originalImageData))
                {
                    // 圧縮パイプライン実行
                    var compressedData = ProcessImageCompression(testImage, testSettings, compressionWorker);

                    e.Result = new CompressionResult
                    {
                        CompressedSize = compressedData?.Length ?? 0,
                        Success = compressedData != null,
                        Data = compressedData
                    };
                }
            }
            catch (Exception ex)
            {
                logger.Error(ex, "Compression calculation failed");
                e.Result = new CompressionResult
                {
                    Success = false,
                    ErrorMessage = ex.Message
                };
            }
        }

        private void CompressionWorker_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            progressBar.Value = e.ProgressPercentage;
            if (e.UserState is string message)
            {
                labelProgress.Text = message;
            }
        }

        private void CompressionWorker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            isCompressing = false;
            progressBar.Visible = false;
            labelProgress.Visible = false;

            if (e.Result is CompressionResult result)
            {
                if (result.Success && result.CompressedSize > 0)
                {
                    settings.CompressedSize = result.CompressedSize;

                    labelCompressedSize.Text = $"圧縮後: {FormatFileSize(result.CompressedSize)}";

                    if (result.CompressedSize >= settings.OriginalSize)
                    {
                        labelReduction.Text = "（効果なし）";
                        labelReduction.ForeColor = Color.Orange;
                        btnExecute.Enabled = false; // 効果なしなら実行無効
                    }
                    else
                    {
                        var reduction = (1.0 - (double)result.CompressedSize / settings.OriginalSize) * 100;
                        labelReduction.Text = $"（約{reduction:F0}%削減）";
                        labelReduction.ForeColor = Color.Green;
                        btnExecute.Enabled = true;
                    }
                }
                else
                {
                    labelCompressedSize.Text = "圧縮後: 計算エラー";
                    labelReduction.Text = $"エラー: {result.ErrorMessage}";
                    labelReduction.ForeColor = Color.Red;
                    btnExecute.Enabled = false;
                }
            }
        }
        #endregion

        #region Compression Pipeline
        private byte[] ProcessImageCompression(MagickImage image, AdvancedCompressionSettings settings, BackgroundWorker worker)
        {
            try
            {
                worker?.ReportProgress(20, "リサイズ処理中...");

                // 1. 長辺リサイズ（縮小のみ）
                if (settings.MaxDimension > 0)
                {
                    var longerSide = Math.Max(image.Width, image.Height);
                    if (longerSide > settings.MaxDimension)
                    {
                        var scale = (double)settings.MaxDimension / longerSide;
                        var newWidth = (int)(image.Width * scale);
                        var newHeight = (int)(image.Height * scale);
                        image.Resize(newWidth, newHeight);
                        logger.Debug($"Resized to {newWidth}x{newHeight}");
                    }
                }

                worker?.ReportProgress(40, "処理中...");

                // 2. 出力形式に応じた圧縮
                switch (settings.OutputFormat)
                {
                    case OutputFormat.JPEG:
                        ApplyJpegCompression(image, settings);
                        break;
                    case OutputFormat.PngReduced:
                        ApplyPngReducedCompression(image, settings);
                        break;
                    case OutputFormat.PngLossless:
                        ApplyPngLosslessCompression(image, settings);
                        break;
                }

                worker?.ReportProgress(60, "透過処理中...");

                // 3. 透過処理
                if (!settings.PreserveTransparency && image.HasAlpha)
                {
                    image.Alpha(AlphaOption.Remove);
                    image.BackgroundColor = MagickColors.White;
                    image.Alpha(AlphaOption.Background);
                }

                worker?.ReportProgress(80, "メタデータ処理中...");

                // 4. メタデータ処理
                if (settings.RemoveMetadata)
                {
                    image.Strip();
                }

                worker?.ReportProgress(100, "完了");

                return image.ToByteArray();
            }
            catch (Exception ex)
            {
                logger.Error(ex, "Image compression processing failed");
                return null;
            }
        }

        private void ApplyJpegCompression(MagickImage image, AdvancedCompressionSettings settings)
        {
            image.Format = MagickFormat.Jpeg;
            image.Quality = settings.JpegQuality;
            image.Interlace = Interlace.Jpeg;
            image.ColorSpace = ColorSpace.sRGB;
            image.Settings.SetDefine(MagickFormat.Jpeg, "optimize-coding", "true");
        }

        private void ApplyPngReducedCompression(MagickImage image, AdvancedCompressionSettings settings)
        {
            image.Format = MagickFormat.Png;

            // 減色処理
            var quantizeSettings = new QuantizeSettings()
            {
                Colors = settings.PngColors,
                DitherMethod = DitherMethod.FloydSteinberg, // Dither ON
                ColorSpace = ColorSpace.sRGB
            };
            image.Quantize(quantizeSettings);

            // PNG最適化
            image.Depth = 8;
            image.Settings.SetDefine(MagickFormat.Png, "compression-level", "9");
        }

        private void ApplyPngLosslessCompression(MagickImage image, AdvancedCompressionSettings settings)
        {
            image.Format = MagickFormat.Png;
            image.Settings.SetDefine(MagickFormat.Png, "compression-level", "9");
            image.Settings.SetDefine(MagickFormat.Png, "compression-strategy", "1");
        }
        #endregion

        #region Utility Methods
        private string FormatFileSize(long bytes)
        {
            if (bytes < 1024)
                return $"{bytes} B";
            else if (bytes < 1024 * 1024)
                return $"{bytes / 1024.0:F0} KB";
            else
                return $"{bytes / (1024.0 * 1024.0):F1} MB";
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                originalImage?.Dispose();
                compressionWorker?.Dispose();
            }
            base.Dispose(disposing);
        }
        #endregion

        #region Data Classes
        public class AdvancedCompressionSettings
        {
            public OutputFormat OutputFormat { get; set; } = OutputFormat.JPEG;
            public int JpegQuality { get; set; } = 85;
            public int PngColors { get; set; } = 256;
            public int MaxDimension { get; set; } = 0; // 0 = no resize
            public bool PreserveTransparency { get; set; } = true;
            public bool RemoveMetadata { get; set; } = true;
            public long OriginalSize { get; set; }
            public long CompressedSize { get; set; }

            public AdvancedCompressionSettings Clone()
            {
                return new AdvancedCompressionSettings
                {
                    OutputFormat = this.OutputFormat,
                    JpegQuality = this.JpegQuality,
                    PngColors = this.PngColors,
                    MaxDimension = this.MaxDimension,
                    PreserveTransparency = this.PreserveTransparency,
                    RemoveMetadata = this.RemoveMetadata,
                    OriginalSize = this.OriginalSize,
                    CompressedSize = this.CompressedSize
                };
            }
        }

        public enum OutputFormat
        {
            JPEG,
            PngReduced,
            PngLossless
        }

        private class ImageAnalysisResult
        {
            public int Width { get; set; }
            public int Height { get; set; }
            public int LongerSide { get; set; }
            public bool HasTransparency { get; set; }
            public int ColorCount { get; set; }
            public string Format { get; set; }
            public ImageType ImageType { get; set; }
        }

        private enum ImageType
        {
            Photo,    // 写真
            Graphic   // 図/スクショ
        }

        private class CompressionResult
        {
            public bool Success { get; set; }
            public long CompressedSize { get; set; }
            public byte[] Data { get; set; }
            public string ErrorMessage { get; set; }
        }
        #endregion
    }
}