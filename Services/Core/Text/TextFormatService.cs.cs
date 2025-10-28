using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using PowerPointEfficiencyAddin.Models;
using PowerPointEfficiencyAddin.Utils;
using NLog;
using PowerPointEfficiencyAddin.Services.Infrastructure.MultiInstance;

namespace PowerPointEfficiencyAddin.Services.Core.Text
{
    /// <summary>
    /// テキスト書式・余白調整機能を提供するサービスクラス
    /// </summary>
    public class TextFormatService
    {
        private static readonly Logger logger = LogManager.GetCurrentClassLogger();
        private readonly IApplicationProvider applicationProvider;

        //ライセンスチェック共通メソッド
        private bool CheckFeatureAccess(string featureName)
        {
            return Globals.ThisAddIn?.CheckFeatureAccess(featureName, 0) ?? true;
        }

        // DI対応コンストラクタ
        public TextFormatService(IApplicationProvider applicationProvider)
        {
            this.applicationProvider = applicationProvider ?? throw new ArgumentNullException(nameof(applicationProvider));
            logger.Debug("TextFormatService initialized with DI application provider");
        }

        // 既存コンストラクタ（後方互換性維持）
        public TextFormatService() : this(new DefaultApplicationProvider())
        {
            logger.Debug("TextFormatService initialized with default application provider");
        }

        // ポイントからcmへの変換係数（1cm = 28.35ポイント）
        private const float PointsToCm = 28.35f;

        #region テキスト折り返し機能

        /// <summary>
        /// テキスト折り返し設定トグル
        /// 選択した図形の「図形内でテキストを折り返す」設定を切り替える
        /// </summary>
        public void ToggleTextWrap()
        {
            if (!Globals.ThisAddIn.CheckFeatureAccess("ToggleTextWrap")) return;

            logger.Info("ToggleTextWrap operation started (text editing mode supported)");

            var selectedShapes = GetSelectedShapeInfos();
            if (!ValidateSelection(selectedShapes, 1, 0, "テキスト折り返し")) return;

            ComHelper.ExecuteWithComCleanup(() =>
            {
                var processedCount = 0;
                var toggledToWrap = false;
                var toggledToNoWrap = false;

                foreach (var shapeInfo in selectedShapes)
                {
                    try
                    {
                        if (shapeInfo.HasTextFrame)
                        {
                            var textFrame = shapeInfo.Shape.TextFrame;
                            var currentWrapSetting = textFrame.WordWrap;

                            // 折り返し設定をトグル
                            var newWrapSetting = currentWrapSetting == MsoTriState.msoTrue
                                ? MsoTriState.msoFalse
                                : MsoTriState.msoTrue;

                            textFrame.WordWrap = newWrapSetting;

                            var currentState = currentWrapSetting == MsoTriState.msoTrue ? "ON" : "OFF";
                            var newState = newWrapSetting == MsoTriState.msoTrue ? "ON" : "OFF";

                            logger.Debug($"Text wrap toggled for {shapeInfo.Name}: {currentState} → {newState}");

                            if (newWrapSetting == MsoTriState.msoTrue)
                                toggledToWrap = true;
                            else
                                toggledToNoWrap = true;

                            processedCount++;
                        }
                        else
                        {
                            logger.Debug($"Shape {shapeInfo.Name} has no text frame, skipping");
                        }
                    }
                    catch (Exception ex)
                    {
                        logger.Error(ex, $"Failed to toggle text wrap for shape {shapeInfo.Name}");
                    }
                }

                // 結果をユーザーに通知
                if (processedCount > 0)
                {
                    var message = processedCount == 1
                        ? toggledToWrap ? "テキストの折り返しを有効にしました" : "テキストの折り返しを無効にしました"
                        : $"{processedCount}個の図形のテキスト折り返し設定を変更しました";

                    logger.Info($"ToggleTextWrap completed: {message}");
                }
                else
                {
                    logger.Warn("No shapes with text frame were found");
                }

            }, selectedShapes.Select(s => s.Shape).ToArray());

            logger.Info("ToggleTextWrap completed");
        }

        /// <summary>
        /// テキスト編集モードかどうかを判定します
        /// </summary>
        /// <returns>テキスト編集中の場合true</returns>
        private bool IsInTextEditingMode()
        {
            try
            {
                var application = applicationProvider.GetCurrentApplication();
                var activeWindow = application.ActiveWindow;

                if (activeWindow?.Selection != null)
                {
                    return activeWindow.Selection.Type == PowerPoint.PpSelectionType.ppSelectionText;
                }
            }
            catch (Exception ex)
            {
                logger.Debug(ex, "Failed to check text editing mode");
            }

            return false;
        }


        /// <summary>
        /// テキスト編集中の図形を取得します
        /// </summary>
        /// <returns>テキスト編集中の図形、取得できない場合はnull</returns>
        private PowerPoint.Shape GetTextEditingShape()
        {
            try
            {
                var application = applicationProvider.GetCurrentApplication();
                var activeWindow = application.ActiveWindow;
                var selection = activeWindow?.Selection;

                if (selection?.Type == PowerPoint.PpSelectionType.ppSelectionText)
                {
                    // Method 1: TextRangeから取得
                    try
                    {
                        var textRange = selection.TextRange;
                        if (textRange?.Parent?.Parent is PowerPoint.Shape shape)
                        {
                            logger.Debug($"Got text editing shape via TextRange: {shape.Name}");
                            return shape;
                        }
                    }
                    catch (Exception ex)
                    {
                        logger.Debug(ex, "Method 1 (TextRange) failed");
                    }

                    // Method 2: ShapeRangeから取得
                    try
                    {
                        var shapeRange = selection.ShapeRange;
                        if (shapeRange?.Count > 0)
                        {
                            var shape = shapeRange[1];
                            logger.Debug($"Got text editing shape via ShapeRange: {shape.Name}");
                            return shape;
                        }
                    }
                    catch (Exception ex)
                    {
                        logger.Debug(ex, "Method 2 (ShapeRange) failed");
                    }
                }
            }
            catch (Exception ex)
            {
                logger.Error(ex, "Failed to get text editing shape");
            }

            return null;
        }

        #endregion

        #region 余白調整機能

        /// <summary>
        /// 余白Up（×1.5）
        /// 選択した図形の上下左右余白を1.5倍に調整
        /// </summary>
        public void AdjustMarginUp()
        {
            if (!Globals.ThisAddIn.CheckFeatureAccess("AdjustMarginUp")) return;

            logger.Info("AdjustMarginUp operation started");

            var selectedShapes = GetSelectedShapeInfos();
            if (!ValidateSelection(selectedShapes, 1, 0, "余白Up")) return;

            ComHelper.ExecuteWithComCleanup(() =>
            {
                foreach (var shapeInfo in selectedShapes)
                {
                    try
                    {
                        if (shapeInfo.HasTextFrame)
                        {
                            AdjustShapeMargins(shapeInfo.Shape, 1.5f);
                            logger.Debug($"Increased margins for {shapeInfo.Name} by 1.5x");
                        }
                        else
                        {
                            logger.Warn($"Shape {shapeInfo.Name} has no text frame");
                        }
                    }
                    catch (Exception ex)
                    {
                        logger.Error(ex, $"Failed to adjust margins up for shape {shapeInfo.Name}");
                    }
                }
            }, selectedShapes.Select(s => s.Shape).ToArray());

            logger.Info($"AdjustMarginUp completed for {selectedShapes.Count} shapes");
        }

        /// <summary>
        /// 余白Down（÷1.5）
        /// 選択した図形の上下左右余白を1.5で割る（縮小）
        /// </summary>
        public void AdjustMarginDown()
        {
            if (!Globals.ThisAddIn.CheckFeatureAccess("AdjustMarginDown")) return;

            logger.Info("AdjustMarginDown operation started");

            var selectedShapes = GetSelectedShapeInfos();
            if (!ValidateSelection(selectedShapes, 1, 0, "余白Down")) return;

            ComHelper.ExecuteWithComCleanup(() =>
            {
                foreach (var shapeInfo in selectedShapes)
                {
                    try
                    {
                        if (shapeInfo.HasTextFrame)
                        {
                            AdjustShapeMargins(shapeInfo.Shape, 1.0f / 1.5f);
                            logger.Debug($"Decreased margins for {shapeInfo.Name} by 1/1.5");
                        }
                        else
                        {
                            logger.Warn($"Shape {shapeInfo.Name} has no text frame");
                        }
                    }
                    catch (Exception ex)
                    {
                        logger.Error(ex, $"Failed to adjust margins down for shape {shapeInfo.Name}");
                    }
                }
            }, selectedShapes.Select(s => s.Shape).ToArray());

            logger.Info($"AdjustMarginDown completed for {selectedShapes.Count} shapes");
        }

        /// <summary>
        /// 余白調整ダイアログ
        /// 詳細な余白設定ダイアログを表示
        /// </summary>
        public void ShowMarginAdjustDialog()
        {
            if (!Globals.ThisAddIn.CheckFeatureAccess("ShowMarginAdjustDialog")) return;

            logger.Info("ShowMarginAdjustDialog operation started");

            var selectedShapes = GetSelectedShapeInfos();
            if (!ValidateSelection(selectedShapes, 1, 0, "余白調整")) return;

            // 基準となる図形の現在の余白を取得
            var referenceShape = selectedShapes.First();
            var currentMargins = GetShapeMargins(referenceShape.Shape);

            if (currentMargins == null)
            {
                ErrorHandler.ExecuteSafely(() =>
                {
                    throw new InvalidOperationException("選択した図形にテキストフレームがありません。");
                }, "余白調整");
                return;
            }

            // ダイアログを表示
            var newMargins = ShowMarginDialog(currentMargins.Value);
            if (newMargins == null)
            {
                logger.Info("Margin adjustment cancelled by user");
                return;
            }

            ComHelper.ExecuteWithComCleanup(() =>
            {
                foreach (var shapeInfo in selectedShapes)
                {
                    try
                    {
                        if (shapeInfo.HasTextFrame)
                        {
                            SetShapeMargins(shapeInfo.Shape, newMargins.Value);
                            logger.Debug($"Set custom margins for {shapeInfo.Name}");
                        }
                        else
                        {
                            logger.Warn($"Shape {shapeInfo.Name} has no text frame");
                        }
                    }
                    catch (Exception ex)
                    {
                        logger.Error(ex, $"Failed to set margins for shape {shapeInfo.Name}");
                    }
                }
            }, selectedShapes.Select(s => s.Shape).ToArray());

            logger.Info($"ShowMarginAdjustDialog completed for {selectedShapes.Count} shapes");
        }

        #endregion

        #region Private Helper Methods

        /// <summary>
        /// 現在選択されている図形の情報を取得します
        /// </summary>
        private List<ShapeInfo> GetSelectedShapeInfos()
        {
            var shapeInfos = new List<ShapeInfo>();

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

                // 選択状態の種類を確認
                logger.Debug($"Selection type: {selection.Type}");

                switch (selection.Type)
                {
                    case PowerPoint.PpSelectionType.ppSelectionShapes:
                        // 通常の図形選択状態
                        logger.Debug("Normal shape selection mode");
                        var normalShapeRange = selection.ShapeRange;
                        if (normalShapeRange != null)
                        {
                            for (int i = 1; i <= normalShapeRange.Count; i++)
                            {
                                var shape = normalShapeRange[i];
                                shapeInfos.Add(new ShapeInfo(shape, i - 1));
                            }
                        }
                        break;

                    case PowerPoint.PpSelectionType.ppSelectionText:
                        // テキスト編集モード
                        logger.Debug("Text editing mode detected");
                        try
                        {
                            // テキスト編集中の図形を取得
                            var textRange = selection.TextRange;
                            if (textRange?.Parent != null)
                            {
                                var parentShape = textRange.Parent.Parent; // TextFrame -> Shape
                                if (parentShape is PowerPoint.Shape shape)
                                {
                                    shapeInfos.Add(new ShapeInfo(shape, 0));
                                    logger.Debug($"Added text editing shape: {shape.Name}");
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            logger.Debug(ex, "Failed to get shape from text editing mode, trying alternative method");

                            // 代替方法：ShapeRangeを試行
                            try
                            {
                                var textModeShapeRange = selection.ShapeRange;
                                if (textModeShapeRange?.Count > 0)
                                {
                                    var shape = textModeShapeRange[1];
                                    shapeInfos.Add(new ShapeInfo(shape, 0));
                                    logger.Debug($"Added shape via alternative method: {shape.Name}");
                                }
                            }
                            catch (Exception ex2)
                            {
                                logger.Warn(ex2, "Alternative method also failed to get text editing shape");
                            }
                        }
                        break;

                    case PowerPoint.PpSelectionType.ppSelectionNone:
                        logger.Debug("No selection");
                        break;

                    default:
                        logger.Debug($"Other selection type: {selection.Type}");
                        break;
                }

                logger.Debug($"Retrieved {shapeInfos.Count} shape(s) including text editing mode");
            }
            catch (Exception ex)
            {
                logger.Error(ex, "Failed to get selected shape infos including text editing mode");
            }

            return shapeInfos;
        }


        /// <summary>
        /// 選択状態を検証します
        /// </summary>
        private bool ValidateSelection(List<ShapeInfo> shapeInfos, int minRequired, int maxAllowed, string operationName)
        {
            return ErrorHandler.ValidateSelection(shapeInfos.Count, minRequired, maxAllowed, operationName);
        }

        /// <summary>
        /// 図形の余白を指定倍率で調整します
        /// </summary>
        /// <param name="shape">対象図形</param>
        /// <param name="multiplier">倍率</param>
        private void AdjustShapeMargins(PowerPoint.Shape shape, float multiplier)
        {
            try
            {
                var textFrame = shape.TextFrame;

                // 現在の余白を取得（ポイント単位）
                var currentLeft = textFrame.MarginLeft;
                var currentRight = textFrame.MarginRight;
                var currentTop = textFrame.MarginTop;
                var currentBottom = textFrame.MarginBottom;

                // 倍率を適用
                textFrame.MarginLeft = currentLeft * multiplier;
                textFrame.MarginRight = currentRight * multiplier;
                textFrame.MarginTop = currentTop * multiplier;
                textFrame.MarginBottom = currentBottom * multiplier;

                logger.Debug($"Adjusted margins by {multiplier}x: " +
                    $"L:{currentLeft * multiplier:F1}, R:{currentRight * multiplier:F1}, " +
                    $"T:{currentTop * multiplier:F1}, B:{currentBottom * multiplier:F1}");
            }
            catch (Exception ex)
            {
                logger.Error(ex, $"Failed to adjust margins for shape {shape.Name}");
                throw;
            }
        }

        /// <summary>
        /// 図形の現在の余白を取得します（cm単位）
        /// </summary>
        /// <param name="shape">対象図形</param>
        /// <returns>余白設定、取得できない場合はnull</returns>
        private MarginSettings? GetShapeMargins(PowerPoint.Shape shape)
        {
            try
            {
                if (shape.HasTextFrame == MsoTriState.msoTrue)
                {
                    var textFrame = shape.TextFrame;
                    return new MarginSettings
                    {
                        Left = textFrame.MarginLeft / PointsToCm,
                        Right = textFrame.MarginRight / PointsToCm,
                        Top = textFrame.MarginTop / PointsToCm,
                        Bottom = textFrame.MarginBottom / PointsToCm
                    };
                }
            }
            catch (Exception ex)
            {
                logger.Error(ex, $"Failed to get margins for shape {shape.Name}");
            }

            return null;
        }

        /// <summary>
        /// 図形に余白を設定します（cm単位）
        /// </summary>
        /// <param name="shape">対象図形</param>
        /// <param name="margins">余白設定</param>
        private void SetShapeMargins(PowerPoint.Shape shape, MarginSettings margins)
        {
            try
            {
                var textFrame = shape.TextFrame;

                // cm単位をポイント単位に変換して設定
                textFrame.MarginLeft = margins.Left * PointsToCm;
                textFrame.MarginRight = margins.Right * PointsToCm;
                textFrame.MarginTop = margins.Top * PointsToCm;
                textFrame.MarginBottom = margins.Bottom * PointsToCm;

                logger.Debug($"Set margins (cm): L:{margins.Left:F2}, R:{margins.Right:F2}, " +
                    $"T:{margins.Top:F2}, B:{margins.Bottom:F2}");
            }
            catch (Exception ex)
            {
                logger.Error(ex, $"Failed to set margins for shape {shape.Name}");
                throw;
            }
        }

        /// <summary>
        /// 余白調整ダイアログを表示します
        /// </summary>
        /// <param name="currentMargins">現在の余白設定</param>
        /// <returns>新しい余白設定、キャンセル時はnull</returns>
        private MarginSettings? ShowMarginDialog(MarginSettings currentMargins)
        {
            using (var form = new Form())
            {
                form.Text = "余白調整";
                form.Size = new System.Drawing.Size(320, 250);
                form.StartPosition = FormStartPosition.CenterScreen;
                form.FormBorderStyle = FormBorderStyle.FixedDialog;
                form.MaximizeBox = false;
                form.MinimizeBox = false;

                // 左余白（テキスト簡潔化）
                var labelLeft = new Label()
                {
                    Text = "左余白:",
                    Location = new System.Drawing.Point(20, 20),
                    Size = new System.Drawing.Size(80, 20)
                };
                var numLeft = new NumericUpDown()
                {
                    Location = new System.Drawing.Point(110, 18),
                    Size = new System.Drawing.Size(120, 20),
                    Minimum = 0.0M,
                    Maximum = 10.0M,
                    Value = (decimal)currentMargins.Left,
                    DecimalPlaces = 2,
                    Increment = 0.1M
                };

                // 右余白（テキスト簡潔化）
                var labelRight = new Label()
                {
                    Text = "右余白:",
                    Location = new System.Drawing.Point(20, 50),
                    Size = new System.Drawing.Size(80, 20)
                };
                var numRight = new NumericUpDown()
                {
                    Location = new System.Drawing.Point(110, 48),
                    Size = new System.Drawing.Size(120, 20),
                    Minimum = 0.0M,
                    Maximum = 10.0M,
                    Value = (decimal)currentMargins.Right,
                    DecimalPlaces = 2,
                    Increment = 0.1M
                };

                // 上余白（テキスト簡潔化）
                var labelTop = new Label()
                {
                    Text = "上余白:",
                    Location = new System.Drawing.Point(20, 80),
                    Size = new System.Drawing.Size(80, 20)
                };
                var numTop = new NumericUpDown()
                {
                    Location = new System.Drawing.Point(110, 78),
                    Size = new System.Drawing.Size(120, 20),
                    Minimum = 0.0M,
                    Maximum = 10.0M,
                    Value = (decimal)currentMargins.Top,
                    DecimalPlaces = 2,
                    Increment = 0.1M
                };

                // 下余白（テキスト簡潔化）
                var labelBottom = new Label()
                {
                    Text = "下余白:",
                    Location = new System.Drawing.Point(20, 110),
                    Size = new System.Drawing.Size(80, 20)
                };
                var numBottom = new NumericUpDown()
                {
                    Location = new System.Drawing.Point(110, 108),
                    Size = new System.Drawing.Size(120, 20),
                    Minimum = 0.0M,
                    Maximum = 10.0M,
                    Value = (decimal)currentMargins.Bottom,
                    DecimalPlaces = 2,
                    Increment = 0.1M
                };

                // OKボタン
                var btnOK = new Button()
                {
                    Text = "OK",
                    Location = new System.Drawing.Point(110, 160),
                    Size = new System.Drawing.Size(75, 25),
                    DialogResult = DialogResult.OK
                };

                // キャンセルボタン
                var btnCancel = new Button()
                {
                    Text = "キャンセル",
                    Location = new System.Drawing.Point(200, 160),
                    Size = new System.Drawing.Size(75, 25),
                    DialogResult = DialogResult.Cancel
                };

                // コントロール追加
                form.Controls.AddRange(new Control[]
                {
            labelLeft, numLeft,
            labelRight, numRight,
            labelTop, numTop,
            labelBottom, numBottom,
            btnOK, btnCancel
                });

                form.AcceptButton = btnOK;
                form.CancelButton = btnCancel;

                if (form.ShowDialog() == DialogResult.OK)
                {
                    return new MarginSettings
                    {
                        Left = (float)numLeft.Value,
                        Right = (float)numRight.Value,
                        Top = (float)numTop.Value,
                        Bottom = (float)numBottom.Value
                    };
                }

                return null;
            }
        }

        #endregion

        /// <summary>
        /// 余白設定を表すクラス
        /// </summary>
        public struct MarginSettings
        {
            public float Left { get; set; }
            public float Right { get; set; }
            public float Top { get; set; }
            public float Bottom { get; set; }
        }

        #region テキストクリア機能

        /// <summary>
        /// 選択図形からテキストをクリアします
        /// 複数の図形を選択した状態で実行すると、全ての図形のテキスト内容を空にします
        /// </summary>
        public void ClearTextsFromSelectedShapes()
        {
            if (!Globals.ThisAddIn.CheckFeatureAccess("ClearTextsFromSelectedShapes")) return;

            logger.Info("ClearTextsFromSelectedShapes operation started");

            var selectedShapes = GetSelectedShapeInfos(); // ★既存メソッド流用
            if (!ValidateSelection(selectedShapes, 1, 0, "テキストクリア")) return; // ★既存メソッド流用

            ComHelper.ExecuteWithComCleanup(() => // ★既存パターン流用
            {
                int clearedCount = 0;
                int skippedCount = 0;

                foreach (var shapeInfo in selectedShapes)
                {
                    try
                    {
                        // テキストフレームの存在確認（既存パターン流用）
                        if (shapeInfo.HasTextFrame)
                        {
                            // テキストの存在確認（効率化のため）
                            if (shapeInfo.Shape.TextFrame.HasText == MsoTriState.msoTrue)
                            {
                                var originalText = shapeInfo.Shape.TextFrame.TextRange.Text;
                                shapeInfo.Shape.TextFrame.TextRange.Text = ""; // ★テキストクリア実行
                                clearedCount++;

                                logger.Debug($"Cleared text from shape: {shapeInfo.Name} (original: '{TruncateText(originalText)}')");
                            }
                            else
                            {
                                logger.Debug($"Shape {shapeInfo.Name} already has no text, skipping");
                                skippedCount++;
                            }
                        }
                        else
                        {
                            // テキストフレームがない図形はスキップ
                            skippedCount++;
                            logger.Debug($"Skipped shape without text frame: {shapeInfo.Name}");
                        }
                    }
                    catch (System.Runtime.InteropServices.COMException comEx)
                    {
                        // COMオブジェクト操作時のエラーハンドリング（既存パターン流用）
                        logger.Error(comEx, $"COM error clearing text from shape {shapeInfo.Name}: {comEx.Message}");
                        skippedCount++;
                    }
                    catch (UnauthorizedAccessException authEx)
                    {
                        // 読み取り専用ファイル等のアクセス拒否
                        logger.Error(authEx, $"Access denied for shape {shapeInfo.Name}");
                        skippedCount++;
                    }
                    catch (Exception ex)
                    {
                        // 一般的な例外処理（既存パターン流用）
                        logger.Error(ex, $"Failed to clear text from shape {shapeInfo.Name}: {ex.Message}");
                        skippedCount++;
                    }
                }

                // 結果の報告（既存パターン流用）
                if (clearedCount > 0)
                {
                    var message = $"{clearedCount}個の図形のテキストをクリアしました";
                    if (skippedCount > 0)
                    {
                        message += $"（{skippedCount}個の図形をスキップ）";
                    }

                    logger.Info($"ClearTextsFromSelectedShapes completed: {message}");
                }
                else
                {
                    // エラーハンドリング（既存パターン流用）
                    ErrorHandler.ExecuteSafely(() =>
                    {
                        throw new InvalidOperationException(
                            "テキストをクリアできる図形が見つかりませんでした。\n" +
                            "テキストを含む図形を選択してください。");
                    }, "テキストクリア");
                }

            }, selectedShapes.Select(s => s.Shape).ToArray()); // ★既存パターン流用

            logger.Info("ClearTextsFromSelectedShapes completed");
        }

        /// <summary>
        /// テキストを表示用に省略します（ログ出力用）
        /// </summary>
        /// <param name="text">元テキスト</param>
        /// <returns>省略されたテキスト</returns>
        private string TruncateText(string text)
        {
            if (string.IsNullOrEmpty(text)) return "(empty)";

            const int maxLength = 30;
            if (text.Length <= maxLength) return $"'{text}'";

            return $"'{text.Substring(0, maxLength)}...'";
        }

        #endregion

    }
}