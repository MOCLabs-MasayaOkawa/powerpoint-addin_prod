using ImageMagick;
using Microsoft.Office.Core;
using NLog;
using PowerPointEfficiencyAddin.Services.Infrastructure.MultiInstance;
using PowerPointEfficiencyAddin.Services.UI.Dialogs;
using PowerPointEfficiencyAddin.Utils;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PowerPointEfficiencyAddin.Services.Core.Image
{
    /// <summary>
    /// 画像圧縮機能を提供するサービスクラス
    /// </summary>
    public class ImageCompressionService
    {
        private static readonly Logger logger = LogManager.GetCurrentClassLogger();
        private readonly IApplicationProvider applicationProvider;

        /// <summary>
        /// DI対応コンストラクタ
        /// </summary>
        /// <param name="applicationProvider">アプリケーションプロバイダー</param>
        public ImageCompressionService(IApplicationProvider applicationProvider)
        {
            this.applicationProvider = applicationProvider ?? throw new ArgumentNullException(nameof(applicationProvider));
            logger.Debug("ImageCompressionService initialized with DI application provider");
        }

        #region Public Methods

        /// <summary>
        /// 選択画像を圧縮（高機能版）
        /// 選択した画像を高品質圧縮で最適化し、品質とファイルサイズのバランスを調整
        /// </summary>
        public void CompressImages()
        {
            if (!Globals.ThisAddIn.CheckFeatureAccess("CompressImages")) return;

            logger.Info("CompressImages operation started");

            var selectedShapes = GetSelectedShapeInfos();
            if (!ValidateSelection(selectedShapes, 1, 0, "画像圧縮")) return;

            // 画像図形のみを抽出
            var imageShapes = selectedShapes.Where(s => IsImageShape(s.Shape)).ToList();
            if (imageShapes.Count == 0)
            {
                ErrorHandler.ExecuteSafely(() =>
                {
                    throw new InvalidOperationException("画像が選択されていません。挿入画像を選択してください。");
                }, "画像圧縮");
                return;
            }

            if (imageShapes.Count > 1)
            {
                ErrorHandler.ExecuteSafely(() =>
                {
                    throw new InvalidOperationException("一度に圧縮できる画像は1つだけです。");
                }, "画像圧縮");
                return;
            }

            var targetImageShape = imageShapes.First();

            ComHelper.ExecuteWithComCleanup(() =>
            {
                try
                {
                    // トリミング削除による画像データ抽出（最善方法）
                    var imageData = ExtractVisibleImageData(targetImageShape.Shape);
                    if (imageData == null)
                    {
                        ErrorHandler.ExecuteSafely(() =>
                        {
                            throw new InvalidOperationException("画像データの抽出に失敗しました。");
                        }, "画像圧縮");
                        return;
                    }

                    // 高機能圧縮ダイアログを表示
                    using (var dialog = new ImageCompressionDialog(imageData))
                    {
                        if (dialog.ShowDialog() == DialogResult.OK)
                        {
                            var settings = dialog.CompressionSettings;

                            // サイズ増チェック（ダイアログ内でも行うが二重チェック）
                            if (settings.CompressedSize >= settings.OriginalSize)
                            {
                                MessageBox.Show(
                                    "圧縮効果が見込めないため、処理を中止しました。\n" +
                                    "画像の差し替えは行われていません。",
                                    "画像圧縮 - 処理中止",
                                    MessageBoxButtons.OK,
                                    MessageBoxIcon.Information
                                );
                                logger.Info("Compression cancelled due to no size reduction benefit");
                                return;
                            }

                            // 最終圧縮処理
                            var compressedData = ExecuteFinalCompressionInternal(imageData, settings);
                            if (compressedData != null)
                            {
                                // 画像を置き換え
                                ReplaceImageInShape(targetImageShape.Shape, compressedData, settings.OutputFormat);

                                var originalSizeMB = imageData.Length / (1024.0 * 1024.0);
                                var compressedSizeMB = compressedData.Length / (1024.0 * 1024.0);
                                var reduction = (1.0 - (double)compressedData.Length / imageData.Length) * 100;

                                logger.Info($"Image compression completed: {originalSizeMB:F2}MB → {compressedSizeMB:F2}MB ({reduction:F1}% reduction)");

                                MessageBox.Show(
                                    $"画像の圧縮が完了しました。\n\n" +
                                    $"圧縮前: {FormatFileSize(imageData.Length)}\n" +
                                    $"圧縮後: {FormatFileSize(compressedData.Length)}\n" +
                                    $"削減率: {reduction:F1}%\n\n" +
                                    $"※元に戻すには Ctrl+Z を押してください。",
                                    "画像圧縮完了",
                                    MessageBoxButtons.OK,
                                    MessageBoxIcon.Information
                                );
                            }
                            else
                            {
                                ErrorHandler.ExecuteSafely(() =>
                                {
                                    throw new InvalidOperationException("最終圧縮処理に失敗しました。");
                                }, "画像圧縮");
                            }
                        }
                        else
                        {
                            logger.Info("Image compression cancelled by user");
                        }
                    }
                }
                catch (Exception ex)
                {
                    logger.Error(ex, "Failed to compress image");
                    ErrorHandler.ExecuteSafely(() =>
                    {
                        throw new InvalidOperationException($"画像圧縮中にエラーが発生しました: {ex.Message}");
                    }, "画像圧縮");
                }
            }, selectedShapes.Select(s => s.Shape).ToArray());

            logger.Info("CompressImages completed");
        }

        #endregion

        #region Private Helper Methods

        /// <summary>
        /// 選択されている図形の情報を取得します
        /// </summary>
        private List<Models.ShapeInfo> GetSelectedShapeInfos()
        {
            var shapeInfos = new List<Models.ShapeInfo>();

            try
            {
                var application = applicationProvider.GetCurrentApplication();
                var activeWindow = application.ActiveWindow;

                if (activeWindow?.Selection == null)
                {
                    logger.Debug("No active window or selection");
                    return shapeInfos;
                }

                var selection = activeWindow.Selection;
                logger.Debug($"Selection type: {selection.Type}");

                switch (selection.Type)
                {
                    case PowerPoint.PpSelectionType.ppSelectionShapes:
                        var normalShapeRange = selection.ShapeRange;
                        if (normalShapeRange != null)
                        {
                            for (int i = 1; i <= normalShapeRange.Count; i++)
                            {
                                var shape = normalShapeRange[i];
                                shapeInfos.Add(new Models.ShapeInfo(shape, i - 1));
                            }
                        }
                        break;
                }

                logger.Debug($"Retrieved {shapeInfos.Count} shape(s)");
            }
            catch (Exception ex)
            {
                logger.Error(ex, "Failed to get selected shape infos");
            }

            return shapeInfos;
        }

        /// <summary>
        /// 選択状態を検証します
        /// </summary>
        private bool ValidateSelection(List<Models.ShapeInfo> shapeInfos, int minRequired, int maxAllowed, string operationName)
        {
            return ErrorHandler.ValidateSelection(shapeInfos.Count, minRequired, maxAllowed, operationName);
        }

        /// <summary>
        /// 現在のスライドを取得します
        /// </summary>
        private PowerPoint.Slide GetCurrentSlide()
        {
            try
            {
                var application = applicationProvider.GetCurrentApplication();
                var activeWindow = application.ActiveWindow;

                if (activeWindow.ViewType == PowerPoint.PpViewType.ppViewSlide ||
                    activeWindow.ViewType == PowerPoint.PpViewType.ppViewNormal)
                {
                    return activeWindow.View.Slide;
                }
            }
            catch (Exception ex)
            {
                logger.Error(ex, "Failed to get current slide");
            }

            return null;
        }

        /// <summary>
        /// 図形が画像かどうかを判定します
        /// </summary>
        /// <param name="shape">図形</param>
        /// <returns>画像の場合true</returns>
        private bool IsImageShape(PowerPoint.Shape shape)
        {
            try
            {
                return shape.Type == MsoShapeType.msoPicture;
            }
            catch (Exception ex)
            {
                logger.Error(ex, $"Failed to determine if shape {shape.Name} is an image");
                return false;
            }
        }

        /// <summary>
        /// トリミング削除による画像データ抽出（最善方法）
        /// Shape.Exportを使用して「見えている状態」を直接抽出
        /// </summary>
        /// <param name="shape">画像図形</param>
        /// <returns>抽出された画像データ</returns>
        private byte[] ExtractVisibleImageData(PowerPoint.Shape shape)
        {
            try
            {
                if (!IsImageShape(shape))
                {
                    logger.Error($"Shape {shape.Name} is not an image");
                    return null;
                }

                // 一時ディレクトリ作成
                var tempDirectory = Path.Combine(Path.GetTempPath(), "PowerPointEfficiencyAddin", "ImageCompression");
                if (!Directory.Exists(tempDirectory))
                {
                    Directory.CreateDirectory(tempDirectory);
                }

                var tempFileName = $"visible_image_{Guid.NewGuid():N}";
                var tempFilePath = Path.Combine(tempDirectory, tempFileName);

                try
                {
                    // PNG形式で見えている状態をエクスポート（最高品質）
                    shape.Export(tempFilePath, PowerPoint.PpShapeFormat.ppShapeFormatPNG);

                    // PNG形式で保存されるため拡張子を追加
                    var pngFilePath = tempFilePath + ".png";
                    if (File.Exists(pngFilePath))
                    {
                        tempFilePath = pngFilePath;
                    }

                    if (!File.Exists(tempFilePath))
                    {
                        logger.Error($"Exported image file not found: {tempFilePath}");
                        return null;
                    }

                    // ファイルから画像データを読み取り
                    var imageBytes = File.ReadAllBytes(tempFilePath);

                    logger.Debug($"Extracted visible image data: {imageBytes.Length} bytes from shape export");

                    return imageBytes;
                }
                finally
                {
                    // 一時ファイルをクリーンアップ
                    CleanupTempFilesInternal(tempFilePath);
                }
            }
            catch (Exception ex)
            {
                logger.Error(ex, $"Failed to extract visible image data from shape {shape.Name}");
                return null;
            }
        }

        /// <summary>
        /// 最終圧縮処理を実行します
        /// </summary>
        /// <param name="originalData">元画像データ</param>
        /// <param name="settings">圧縮設定</param>
        /// <returns>圧縮後画像データ</returns>
        private byte[] ExecuteFinalCompressionInternal(byte[] originalData, ImageCompressionDialog.AdvancedCompressionSettings settings)
        {
            try
            {
                using (var magickImage = new MagickImage(originalData))
                {
                    logger.Debug($"Original image: {magickImage.Width}x{magickImage.Height}, " +
                                $"format: {magickImage.Format}, size: {originalData.Length} bytes");

                    // 長辺リサイズ（縮小のみ）
                    if (settings.MaxDimension > 0)
                    {
                        var longerSide = Math.Max(magickImage.Width, magickImage.Height);
                        if (longerSide > settings.MaxDimension)
                        {
                            var scale = (double)settings.MaxDimension / longerSide;
                            var newWidth = (int)(magickImage.Width * scale);
                            var newHeight = (int)(magickImage.Height * scale);
                            magickImage.Resize(newWidth, newHeight);
                            logger.Debug($"Resized to {newWidth}x{newHeight}");
                        }
                    }

                    // 出力形式に応じた圧縮
                    switch (settings.OutputFormat)
                    {
                        case ImageCompressionDialog.OutputFormat.JPEG:
                            ApplyJpegCompression(magickImage, settings);
                            break;
                        case ImageCompressionDialog.OutputFormat.PngReduced:
                            ApplyPngReducedCompression(magickImage, settings);
                            break;
                        case ImageCompressionDialog.OutputFormat.PngLossless:
                            ApplyPngLosslessCompression(magickImage, settings);
                            break;
                    }

                    // 透過処理
                    if (!settings.PreserveTransparency && magickImage.HasAlpha)
                    {
                        magickImage.Alpha(AlphaOption.Remove);
                        magickImage.BackgroundColor = MagickColors.White;
                        magickImage.Alpha(AlphaOption.Background);
                    }

                    // メタデータ処理
                    if (settings.RemoveMetadata)
                    {
                        magickImage.Strip();
                    }

                    var compressedData = magickImage.ToByteArray();

                    logger.Debug($"Compressed image: format: {magickImage.Format}, size: {compressedData.Length} bytes");

                    return compressedData;
                }
            }
            catch (Exception ex)
            {
                logger.Error(ex, "Failed to execute final compression");
                return null;
            }
        }

        /// <summary>
        /// JPEG圧縮を適用します
        /// </summary>
        private void ApplyJpegCompression(MagickImage image, ImageCompressionDialog.AdvancedCompressionSettings settings)
        {
            image.Format = MagickFormat.Jpeg;
            image.Quality = settings.JpegQuality;
            image.Interlace = Interlace.Jpeg;
            image.ColorSpace = ColorSpace.sRGB;
            image.Settings.SetDefine(MagickFormat.Jpeg, "optimize-coding", "true");
        }

        /// <summary>
        /// PNG減色圧縮を適用します
        /// </summary>
        private void ApplyPngReducedCompression(MagickImage image, ImageCompressionDialog.AdvancedCompressionSettings settings)
        {
            image.Format = MagickFormat.Png;

            // 減色処理
            var quantizeSettings = new QuantizeSettings()
            {
                Colors = settings.PngColors,
                DitherMethod = DitherMethod.FloydSteinberg,
                ColorSpace = ColorSpace.sRGB
            };
            image.Quantize(quantizeSettings);

            // PNG最適化
            image.Depth = 8;
            image.Settings.SetDefine(MagickFormat.Png, "compression-level", "9");
        }

        /// <summary>
        /// PNG無減色圧縮を適用します
        /// </summary>
        private void ApplyPngLosslessCompression(MagickImage image, ImageCompressionDialog.AdvancedCompressionSettings settings)
        {
            image.Format = MagickFormat.Png;
            image.Settings.SetDefine(MagickFormat.Png, "compression-level", "9");
            image.Settings.SetDefine(MagickFormat.Png, "compression-strategy", "1");
        }

        /// <summary>
        /// 図形の画像を置き換えます
        /// </summary>
        /// <param name="shape">対象図形</param>
        /// <param name="newImageData">新しい画像データ</param>
        /// <param name="format">画像形式</param>
        private void ReplaceImageInShape(PowerPoint.Shape shape, byte[] newImageData,
            ImageCompressionDialog.OutputFormat format)
        {
            try
            {
                var slide = GetCurrentSlide();
                if (slide == null)
                {
                    throw new InvalidOperationException("アクティブなスライドが見つかりません。");
                }

                // 現在の図形の位置とサイズを保存
                var left = shape.Left;
                var top = shape.Top;
                var width = shape.Width;
                var height = shape.Height;
                var name = shape.Name;

                // 一時ファイルに新しい画像を保存
                var tempDirectory = Path.Combine(Path.GetTempPath(), "PowerPointEfficiencyAddin");
                if (!Directory.Exists(tempDirectory))
                {
                    Directory.CreateDirectory(tempDirectory);
                }

                var extension = format == ImageCompressionDialog.OutputFormat.JPEG ? ".jpg" : ".png";
                var tempFilePath = Path.Combine(tempDirectory, $"compressed_image_{Guid.NewGuid():N}{extension}");

                try
                {
                    File.WriteAllBytes(tempFilePath, newImageData);

                    // 元の図形を削除
                    shape.Delete();

                    // 新しい画像を挿入
                    var newShape = slide.Shapes.AddPicture(
                        tempFilePath,
                        MsoTriState.msoFalse,  // LinkToFile
                        MsoTriState.msoTrue,   // SaveWithDocument
                        left, top, width, height
                    );

                    // 名前を復元
                    try
                    {
                        newShape.Name = name;
                    }
                    catch
                    {
                        // 名前設定失敗は無視
                    }

                    // 新しい図形を選択状態にする
                    newShape.Select();

                    logger.Debug($"Replaced image in shape: {name}");
                }
                finally
                {
                    // 一時ファイルを削除
                    CleanupTempFilesInternal(tempFilePath);
                }
            }
            catch (Exception ex)
            {
                logger.Error(ex, "Failed to replace image in shape");
                throw;
            }
        }

        /// <summary>
        /// 一時ファイルをクリーンアップします
        /// </summary>
        /// <param name="filePath">削除するファイルパス</param>
        private void CleanupTempFilesInternal(string filePath)
        {
            try
            {
                if (string.IsNullOrEmpty(filePath)) return;

                // 基本ファイル削除
                if (File.Exists(filePath))
                {
                    File.Delete(filePath);
                    logger.Debug($"Deleted temp file: {filePath}");
                }

                // 拡張子付きファイルも削除
                var extensions = new[] { ".png", ".jpg", ".jpeg" };
                foreach (var ext in extensions)
                {
                    var fileWithExt = filePath + ext;
                    if (File.Exists(fileWithExt))
                    {
                        File.Delete(fileWithExt);
                        logger.Debug($"Deleted temp file: {fileWithExt}");
                    }
                }
            }
            catch (Exception ex)
            {
                logger.Warn(ex, $"Failed to cleanup temp file: {filePath}");
            }
        }

        /// <summary>
        /// ファイルサイズを人間が読みやすい形式でフォーマットします
        /// </summary>
        /// <param name="bytes">バイト数</param>
        /// <returns>フォーマットされた文字列</returns>
        private string FormatFileSize(long bytes)
        {
            if (bytes < 1024)
                return $"{bytes} B";
            else if (bytes < 1024 * 1024)
                return $"{bytes / 1024.0:F0} KB";
            else
                return $"{bytes / (1024.0 * 1024.0):F1} MB";
        }

        #endregion
    }
}
