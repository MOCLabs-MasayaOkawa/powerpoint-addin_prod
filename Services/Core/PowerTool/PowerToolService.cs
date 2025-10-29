﻿using ImageMagick;
using Microsoft.Office.Core;
using NLog;
using PowerPointEfficiencyAddin.Models;
using PowerPointEfficiencyAddin.Services.Infrastructure.MultiInstance;
using PowerPointEfficiencyAddin.Services.Infrastructure.Settings;
using PowerPointEfficiencyAddin.Services.UI.Dialogs;
using PowerPointEfficiencyAddin.UI;
using PowerPointEfficiencyAddin.Utils;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PowerPointEfficiencyAddin.Services.Core.PowerTool
{
    /// <summary>
    /// パワーツール・特殊機能を提供するサービスクラス
    /// </summary>
    public class PowerToolService
    {
        private static readonly Logger logger = LogManager.GetCurrentClassLogger();
        private readonly IApplicationProvider applicationProvider;
        private readonly PowerToolServiceHelper helper;

        // DI対応コンストラクタ（商用レベル）
        public PowerToolService(IApplicationProvider applicationProvider)
        {
            this.applicationProvider = applicationProvider ?? throw new ArgumentNullException(nameof(applicationProvider));
            logger.Debug("PowerToolService initialized with DI application provider");
            helper = new PowerToolServiceHelper(applicationProvider);
        }

        // 既存コンストラクタ（後方互換性維持）
        public PowerToolService() : this(new DefaultApplicationProvider())
        {
            logger.Debug("PowerToolService initialized with default application provider");
        }

        #region パワーツールグループ (16-23)

        /// <summary>
        /// テキスト合成（16番機能）
        /// 選択した図形のテキストを改行区切りで合成し、基準図形に設定。他の図形を削除
        /// </summary>
        public void MergeText()
        {

            if (!Globals.ThisAddIn.CheckFeatureAccess("MergeText")) return;

            logger.Info("MergeText operation started");

            var selectedShapes = helper.GetSelectedShapeInfos();
            if (!helper.ValidateSelection(selectedShapes, 2, 0, "テキスト合成")) return;

            ComHelper.ExecuteWithComCleanup(() =>
            {
                var textParts = new List<string>();

                // 選択順にテキストを収集
                foreach (var shapeInfo in selectedShapes.OrderBy(s => s.SelectionOrder))
                {
                    if (shapeInfo.HasTextFrame && !string.IsNullOrEmpty(shapeInfo.Text))
                    {
                        textParts.Add(shapeInfo.Text.Trim());
                    }
                }

                if (textParts.Count == 0)
                {
                    ErrorHandler.ExecuteSafely(() =>
                    {
                        throw new InvalidOperationException("選択した図形にテキストが含まれていません。");
                    }, "テキスト合成");
                    return;
                }

                // 改行区切りでテキストを合成
                var mergedText = string.Join(Environment.NewLine, textParts);

                // 最初の図形（基準図形）にテキストを設定
                var referenceShape = selectedShapes.OrderBy(s => s.SelectionOrder).First();
                var targetShapes = selectedShapes.Skip(1).ToList(); // 基準図形以外

                try
                {
                    // 基準図形にテキストを設定
                    if (referenceShape.HasTextFrame)
                    {
                        referenceShape.Shape.TextFrame.TextRange.Text = mergedText;
                    }
                    else
                    {
                        // テキストフレームがない場合は、テキストボックスに変換
                        referenceShape.Shape.TextFrame.TextRange.Text = mergedText;
                    }

                    // サイズを調整（必要に応じて高さを拡張）
                    var lineCount = textParts.Count;
                    var currentHeight = referenceShape.Height;
                    var estimatedHeight = currentHeight * lineCount * 0.8f; // 概算
                    if (estimatedHeight > currentHeight)
                    {
                        referenceShape.Shape.Height = estimatedHeight;
                    }

                    logger.Debug($"Merged text set to reference shape: {referenceShape.Name}");

                    // 基準図形以外を削除
                    foreach (var shapeInfo in targetShapes)
                    {
                        try
                        {
                            shapeInfo.Shape.Delete();
                            logger.Debug($"Deleted shape: {shapeInfo.Name}");
                        }
                        catch (Exception ex)
                        {
                            logger.Error(ex, $"Failed to delete shape: {shapeInfo.Name}");
                        }
                    }

                    logger.Info($"MergeText completed: merged {textParts.Count} texts, deleted {targetShapes.Count} shapes");
                }
                catch (Exception ex)
                {
                    logger.Error(ex, "Failed to set merged text to reference shape");
                    ErrorHandler.ExecuteSafely(() =>
                    {
                        throw new InvalidOperationException("テキスト合成に失敗しました。");
                    }, "テキスト合成");
                }
            }, selectedShapes.Select(s => s.Shape).ToArray());

            logger.Info("MergeText completed");
        }

        /// <summary>
        /// 線を水平にする（18番機能）
        /// 選択した線の角度を水平（0度）にし、線の長さを保持する
        /// </summary>
        public void MakeLineHorizontal()
        {
            if (!Globals.ThisAddIn.CheckFeatureAccess("MakeLineHorizontal")) return;

            logger.Info("MakeLineHorizontal operation started");

            var selectedShapes = helper.GetSelectedShapeInfos();
            if (!helper.ValidateSelection(selectedShapes, 1, 0, "線を水平にする")) return;

            var lineShapes = selectedShapes.Where(s => helper.IsLineShape(s.Shape)).ToList();
            if (lineShapes.Count == 0)
            {
                ErrorHandler.ExecuteSafely(() =>
                {
                    throw new InvalidOperationException("線図形を選択してください。");
                }, "線を水平にする");
                return;
            }

            ComHelper.ExecuteWithComCleanup(() =>
            {
                foreach (var shapeInfo in lineShapes)
                {
                    try
                    {
                        MakeLineHorizontalInternal(shapeInfo.Shape);
                        logger.Debug($"Made line horizontal: {shapeInfo.Name}");
                    }
                    catch (Exception ex)
                    {
                        logger.Error(ex, $"Failed to make line horizontal: {shapeInfo.Name}");
                    }
                }
            }, lineShapes.Select(s => s.Shape).ToArray());

            logger.Info($"MakeLineHorizontal completed for {lineShapes.Count} lines");
        }

        /// <summary>
        /// 線を垂直にする（19番機能）
        /// 選択した線の角度を垂直（90度）にし、線の長さを保持する
        /// </summary>
        public void MakeLineVertical()
        {
            if (!Globals.ThisAddIn.CheckFeatureAccess("MakeLineVertical")) return;

            logger.Info("MakeLineVertical operation started");

            var selectedShapes = helper.GetSelectedShapeInfos();
            if (!helper.ValidateSelection(selectedShapes, 1, 0, "線を垂直にする")) return;

            var lineShapes = selectedShapes.Where(s => helper.IsLineShape(s.Shape)).ToList();
            if (lineShapes.Count == 0)
            {
                ErrorHandler.ExecuteSafely(() =>
                {
                    throw new InvalidOperationException("線図形を選択してください。");
                }, "線を垂直にする");
                return;
            }

            ComHelper.ExecuteWithComCleanup(() =>
            {
                foreach (var shapeInfo in lineShapes)
                {
                    try
                    {
                        MakeLineVerticalInternal(shapeInfo.Shape);
                        logger.Debug($"Made line vertical: {shapeInfo.Name}");
                    }
                    catch (Exception ex)
                    {
                        logger.Error(ex, $"Failed to make line vertical: {shapeInfo.Name}");
                    }
                }
            }, lineShapes.Select(s => s.Shape).ToArray());

            logger.Info($"MakeLineVertical completed for {lineShapes.Count} lines");
        }

        /// <summary>
        /// 図形位置の交換（20番機能）
        /// 2つの選択した図形の位置を交換する
        /// </summary>
        public void SwapPositions()
        {
            if (!Globals.ThisAddIn.CheckFeatureAccess("SwapPositions")) return;

            logger.Info("SwapPositions operation started");

            var selectedShapes = helper.GetSelectedShapeInfos();
            if (!helper.ValidateSelection(selectedShapes, 2, 2, "図形位置の交換")) return;

            var shape1 = selectedShapes[0];
            var shape2 = selectedShapes[1];

            ComHelper.ExecuteWithComCleanup(() =>
            {
                try
                {
                    // 位置を保存
                    var temp1Left = shape1.Left;
                    var temp1Top = shape1.Top;

                    // 位置を交換
                    shape1.Shape.Left = shape2.Left;
                    shape1.Shape.Top = shape2.Top;
                    shape2.Shape.Left = temp1Left;
                    shape2.Shape.Top = temp1Top;

                    logger.Debug($"Swapped positions of {shape1.Name} and {shape2.Name}");
                }
                catch (Exception ex)
                {
                    logger.Error(ex, "Failed to swap positions");
                }
            }, selectedShapes.Select(s => s.Shape).ToArray());

            logger.Info("SwapPositions completed");
        }

        /// <summary>
        /// 同種図形に一括選択（21番機能）
        /// 選択した図形と同種の図形をスライド内で一括選択
        /// </summary>
        public void SelectSimilarShapes()
        {
            if (!Globals.ThisAddIn.CheckFeatureAccess("SelectSimilarShapes")) return;

            logger.Info("SelectSimilarShapes operation started");

            var selectedShapes = helper.GetSelectedShapeInfos();
            if (!helper.ValidateSelection(selectedShapes, 1, 1, "同種図形に一括選択")) return;

            var referenceShape = selectedShapes.First();

            ComHelper.ExecuteWithComCleanup(() =>
            {
                var slide = helper.GetCurrentSlide();
                if (slide == null) return;

                var similarShapes = new List<PowerPoint.Shape>();

                for (int i = 1; i <= slide.Shapes.Count; i++)
                {
                    var shape = slide.Shapes[i];
                    if (helper.IsSimilarShape(referenceShape.Shape, shape))
                    {
                        similarShapes.Add(shape);
                    }
                }

                if (similarShapes.Count > 1)
                {
                    // 新しい選択を作成
                    helper.SelectShapes(similarShapes);
                    logger.Info($"Selected {similarShapes.Count} similar shapes");
                }
                else
                {
                    ErrorHandler.ExecuteSafely(() =>
                    {
                        throw new InvalidOperationException("同種の図形が見つかりませんでした。");
                    }, "同種図形に一括選択");
                }
            }, selectedShapes.Select(s => s.Shape).ToArray());

            logger.Info("SelectSimilarShapes completed");
        }

        private void MakeLineHorizontalInternal(PowerPoint.Shape lineShape)
        {
            try
            {
                if (lineShape.Type == MsoShapeType.msoLine)
                {
                    // 線の現在の長さと中心点を計算
                    var currentLength = GetLineLength(lineShape);
                    var centerX = lineShape.Left + lineShape.Width / 2;
                    var centerY = lineShape.Top + lineShape.Height / 2;

                    // 水平線として設定（長さを保持）
                    lineShape.Left = centerX - currentLength / 2;
                    lineShape.Top = centerY;
                    lineShape.Width = currentLength;
                    lineShape.Height = 0;

                    logger.Debug($"Converted line to horizontal with length {currentLength}");
                }
                else if (lineShape.Type == MsoShapeType.msoFreeform && lineShape.Connector == MsoTriState.msoTrue)
                {
                    // コネクタの場合の処理
                    var currentLength = GetLineLength(lineShape);
                    var centerX = lineShape.Left + lineShape.Width / 2;
                    var centerY = lineShape.Top + lineShape.Height / 2;

                    // コネクタを水平線として再配置
                    lineShape.Left = centerX - currentLength / 2;
                    lineShape.Top = centerY;
                    lineShape.Width = currentLength;
                    lineShape.Height = 0;

                    logger.Debug($"Converted connector to horizontal with length {currentLength}");
                }
            }
            catch (Exception ex)
            {
                logger.Error(ex, $"Failed to make line horizontal: {lineShape.Name}");
                throw;
            }
        }

        /// <summary>
        /// 線を垂直にします（長さを保持）
        /// </summary>
        private void MakeLineVerticalInternal(PowerPoint.Shape lineShape)
        {
            try
            {
                if (lineShape.Type == MsoShapeType.msoLine)
                {
                    // 線の現在の長さと中心点を計算
                    var currentLength = GetLineLength(lineShape);
                    var centerX = lineShape.Left + lineShape.Width / 2;
                    var centerY = lineShape.Top + lineShape.Height / 2;

                    // 垂直線として設定（長さを保持）
                    lineShape.Left = centerX;
                    lineShape.Top = centerY - currentLength / 2;
                    lineShape.Width = 0;
                    lineShape.Height = currentLength;

                    logger.Debug($"Converted line to vertical with length {currentLength}");
                }
                else if (lineShape.Type == MsoShapeType.msoFreeform && lineShape.Connector == MsoTriState.msoTrue)
                {
                    // コネクタの場合の処理
                    var currentLength = GetLineLength(lineShape);
                    var centerX = lineShape.Left + lineShape.Width / 2;
                    var centerY = lineShape.Top + lineShape.Height / 2;

                    // コネクタを垂直線として再配置
                    lineShape.Left = centerX;
                    lineShape.Top = centerY - currentLength / 2;
                    lineShape.Width = 0;
                    lineShape.Height = currentLength;

                    logger.Debug($"Converted connector to vertical with length {currentLength}");
                }
            }
            catch (Exception ex)
            {
                logger.Error(ex, $"Failed to make line vertical: {lineShape.Name}");
                throw;
            }
        }

        #endregion

        #region New Feature Helper Methods

        /// <summary>
        /// フォント選択ダイアログを表示します
        /// </summary>
        /// <returns>選択されたフォント名、キャンセル時は空文字</returns>
        private string ShowFontSelectionDialog()
        {
            string selectedFont = "";

            try
            {
                using (var form = new Form())
                {
                    form.Text = "フォント一括統一";
                    form.Size = new Size(380, 250);
                    form.StartPosition = FormStartPosition.CenterScreen;
                    form.FormBorderStyle = FormBorderStyle.FixedDialog;
                    form.MaximizeBox = false;
                    form.MinimizeBox = false;

                    var label = new Label()
                    {
                        Text = "プレゼンテーション全体に適用するフォントを選択してください:",
                        Location = new Point(20, 20),
                        Size = new Size(340, 40),
                        AutoSize = false
                    };

                    var comboBox = new ComboBox()
                    {
                        Location = new Point(20, 70),
                        Size = new Size(320, 25),
                        DropDownStyle = ComboBoxStyle.DropDownList,
                        Sorted = true
                    };

                    // 推奨フォントを最初に追加
                    var recommendedFonts = new[]
                    {
                        "Meiryo UI",
                        "Yu Gothic UI",
                        "MS Gothic",
                        "MS PGothic",
                        "Arial",
                        "Calibri",
                        "Times New Roman"
                    };

                    foreach (var font in recommendedFonts)
                    {
                        comboBox.Items.Add(font);
                    }

                    // システムの全フォントを取得
                    try
                    {
                        var fontFamilies = FontFamily.Families;
                        foreach (var fontFamily in fontFamilies)
                        {
                            if (!comboBox.Items.Contains(fontFamily.Name))
                            {
                                comboBox.Items.Add(fontFamily.Name);
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        logger.Warn(ex, "Failed to get system fonts, using recommended fonts only");
                    }

                    // デフォルトでMeiryo UIを選択
                    if (comboBox.Items.Contains("Meiryo UI"))
                    {
                        comboBox.SelectedItem = "Meiryo UI";
                    }
                    else if (comboBox.Items.Count > 0)
                    {
                        comboBox.SelectedIndex = 0;
                    }

                    var warningLabel = new Label()
                    {
                        Text = "注意: この操作は元に戻せません。\n必要に応じて事前にファイルを保存してください。",
                        Location = new Point(20, 110),
                        Size = new Size(320, 40),
                        ForeColor = Color.DarkRed,
                        AutoSize = false
                    };

                    var okButton = new Button()
                    {
                        Text = "実行",
                        Location = new Point(180, 170),
                        Size = new Size(75, 25),
                        DialogResult = DialogResult.OK
                    };

                    var cancelButton = new Button()
                    {
                        Text = "キャンセル",
                        Location = new Point(265, 170),
                        Size = new Size(75, 25),
                        DialogResult = DialogResult.Cancel
                    };

                    form.Controls.AddRange(new Control[]
                    {
                        label, comboBox, warningLabel, okButton, cancelButton
                    });

                    form.AcceptButton = okButton;
                    form.CancelButton = cancelButton;

                    if (form.ShowDialog() == DialogResult.OK)
                    {
                        selectedFont = comboBox.SelectedItem?.ToString() ?? "";
                        logger.Info($"User selected font: '{selectedFont}'");
                    }
                    else
                    {
                        logger.Info("Font selection cancelled by user");
                    }
                }
            }
            catch (Exception ex)
            {
                logger.Error(ex, "Failed to show font selection dialog");
                MessageBox.Show(
                    "フォント選択ダイアログの表示に失敗しました。",
                    "エラー",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error
                );
            }

            return selectedFont;
        }

        /// <summary>
        /// 線の長さを取得します
        /// </summary>
        /// <param name="lineShape">線図形</param>
        /// <returns>線の長さ</returns>
        private float GetLineLength(PowerPoint.Shape lineShape)
        {
            try
            {
                if (lineShape.Type == MsoShapeType.msoLine)
                {
                    // 直線の場合、幅と高さから斜辺を計算
                    var width = Math.Abs(lineShape.Width);
                    var height = Math.Abs(lineShape.Height);
                    return (float)Math.Sqrt(width * width + height * height);
                }
                else if (lineShape.Type == MsoShapeType.msoFreeform)
                {
                    // フリーフォーム（コネクタ等）の場合は幅と高さの最大値を使用
                    return Math.Max(Math.Abs(lineShape.Width), Math.Abs(lineShape.Height));
                }
            }
            catch (Exception ex)
            {
                logger.Error(ex, $"Failed to get line length for {lineShape.Name}");
            }

            return 0f;
        }

        /// <summary>
        /// 線の長さを調整します（中心点固定）
        /// </summary>
        /// <param name="lineShape">線図形</param>
        /// <param name="targetLength">目標の長さ</param>
        private void AdjustLineLength(PowerPoint.Shape lineShape, float targetLength)
        {
            try
            {
                if (lineShape.Type == MsoShapeType.msoLine)
                {
                    var currentLength = GetLineLength(lineShape);
                    if (currentLength <= 0) return;

                    // 現在の中心点を保存
                    var centerX = lineShape.Left + lineShape.Width / 2;
                    var centerY = lineShape.Top + lineShape.Height / 2;

                    var ratio = targetLength / currentLength;

                    // 新しいサイズを計算
                    var newWidth = lineShape.Width * ratio;
                    var newHeight = lineShape.Height * ratio;

                    // 中心点を維持して新しいサイズを設定
                    lineShape.Left = centerX - newWidth / 2;
                    lineShape.Top = centerY - newHeight / 2;
                    lineShape.Width = newWidth;
                    lineShape.Height = newHeight;

                    logger.Debug($"Adjusted line {lineShape.Name}: length {currentLength:F1} → {targetLength:F1}, center maintained at ({centerX:F1}, {centerY:F1})");
                }
                else if (lineShape.Type == MsoShapeType.msoFreeform)
                {
                    // フリーフォーム（コネクタ等）の場合も中心点を維持
                    var centerX = lineShape.Left + lineShape.Width / 2;
                    var centerY = lineShape.Top + lineShape.Height / 2;

                    if (Math.Abs(lineShape.Width) > Math.Abs(lineShape.Height))
                    {
                        // 水平方向が主体の場合
                        var newWidth = lineShape.Width > 0 ? targetLength : -targetLength;
                        lineShape.Left = centerX - newWidth / 2;
                        lineShape.Width = newWidth;
                    }
                    else
                    {
                        // 垂直方向が主体の場合
                        var newHeight = lineShape.Height > 0 ? targetLength : -targetLength;
                        lineShape.Top = centerY - newHeight / 2;
                        lineShape.Height = newHeight;
                    }

                    logger.Debug($"Adjusted connector {lineShape.Name}: target length {targetLength:F1}, center maintained at ({centerX:F1}, {centerY:F1})");
                }
            }
            catch (Exception ex)
            {
                logger.Error(ex, $"Failed to adjust line length for {lineShape.Name}");
            }
        }

        #endregion

        #region 特殊機能グループ (24-27)



        /// <summary>
        /// フォント一括統一（修正版）
        /// 全ページのすべてのテキストを指定フォントに完全統一
        /// </summary>
        public void UnifyFont()
        {
            logger.Info("UnifyFont operation started (improved version)");

            // フォント選択ダイアログを表示
            var selectedFont = ShowFontSelectionDialog();
            if (string.IsNullOrEmpty(selectedFont))
            {
                logger.Info("Font unification cancelled by user");
                return;
            }

            ComHelper.ExecuteWithComCleanup(() =>
            {
                var addin = Globals.ThisAddIn;
                var activePresentation = addin.GetActivePresentation();

                if (activePresentation == null)
                {
                    ErrorHandler.ExecuteSafely(() =>
                    {
                        throw new InvalidOperationException("アクティブなプレゼンテーションが見つかりません。");
                    }, "フォント一括統一");
                    return;
                }

                int changedCount = 0;
                int errorCount = 0;

                logger.Info($"Processing {activePresentation.Slides.Count} slides for font unification to '{selectedFont}'");

                // すべてのスライドを処理
                for (int i = 1; i <= activePresentation.Slides.Count; i++)
                {
                    var slide = activePresentation.Slides[i];
                    var slideChangedCount = 0;

                    try
                    {
                        logger.Debug($"Processing slide {i}");

                        // 1. スライド内のすべての図形を処理（通常の図形）
                        for (int j = 1; j <= slide.Shapes.Count; j++)
                        {
                            var shape = slide.Shapes[j];
                            var shapeChangedCount = ProcessShapeFont(shape, selectedFont);
                            slideChangedCount += shapeChangedCount;

                            if (shapeChangedCount > 0)
                            {
                                logger.Debug($"Slide {i}, Shape {j} ({shape.Name}): Changed {shapeChangedCount} text ranges");
                            }
                        }

                        // 2. スライドのプレースホルダーを処理
                        try
                        {
                            for (int k = 1; k <= slide.Shapes.Placeholders.Count; k++)
                            {
                                var placeholder = slide.Shapes.Placeholders[k];
                                var placeholderChangedCount = ProcessPlaceholderFont(placeholder, selectedFont);
                                slideChangedCount += placeholderChangedCount;

                                if (placeholderChangedCount > 0)
                                {
                                    logger.Debug($"Slide {i}, Placeholder {k}: Changed {placeholderChangedCount} text ranges");
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            logger.Warn(ex, $"Failed to process placeholders on slide {i}");
                        }

                        changedCount += slideChangedCount;
                        logger.Debug($"Slide {i} completed: {slideChangedCount} text ranges changed");
                    }
                    catch (Exception ex)
                    {
                        logger.Error(ex, $"Failed to process slide {i}");
                        errorCount++;
                    }
                }

                logger.Info($"UnifyFont completed: changed {changedCount} text ranges to '{selectedFont}' (errors: {errorCount})");

                // 結果表示
                var message = errorCount > 0
                    ? $"フォントを「{selectedFont}」に統一しました。\n変更されたテキスト数: {changedCount}\n処理エラー: {errorCount}件"
                    : $"フォントを「{selectedFont}」に統一しました。\n変更されたテキスト数: {changedCount}";

                MessageBox.Show(
                    message,
                    "フォント一括統一 完了",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information
                );
            });

            logger.Info("UnifyFont completed (improved version)");
        }

        /// <summary>
        /// 図形のフォントを処理します
        /// </summary>
        /// <param name="shape">処理対象の図形</param>
        /// <param name="targetFont">設定するフォント名</param>
        /// <returns>変更されたテキスト範囲数</returns>
        private int ProcessShapeFont(PowerPoint.Shape shape, string targetFont)
        {
            int changedCount = 0;

            try
            {
                // 1. 通常のテキストフレーム処理
                if (shape.HasTextFrame == MsoTriState.msoTrue)
                {
                    changedCount += ProcessTextFrameFont(shape.TextFrame, targetFont, shape.Name);
                }

                // 2. 表の処理
                if (shape.HasTable == MsoTriState.msoTrue)
                {
                    changedCount += ProcessTableFont(shape.Table, targetFont, shape.Name);
                }

                // 3. グループ図形の処理
                if (shape.Type == MsoShapeType.msoGroup)
                {
                    for (int i = 1; i <= shape.GroupItems.Count; i++)
                    {
                        var groupItem = shape.GroupItems[i];
                        changedCount += ProcessShapeFont(groupItem, targetFont);
                    }
                }

                // 4. SmartArt、グラフなどの特殊図形の処理
                if (shape.Type == MsoShapeType.msoChart ||
                    shape.Type == MsoShapeType.msoSmartArt ||
                    shape.Type == MsoShapeType.msoDiagram)
                {
                    // 基本的なテキストフレームのみ処理（詳細なSmartArt処理は複雑すぎるため省略）
                    if (shape.HasTextFrame == MsoTriState.msoTrue)
                    {
                        changedCount += ProcessTextFrameFont(shape.TextFrame, targetFont, shape.Name);
                    }
                }
            }
            catch (Exception ex)
            {
                logger.Error(ex, $"Failed to process font for shape {shape.Name}");
            }

            return changedCount;
        }

        /// <summary>
        /// プレースホルダーのフォントを処理します
        /// </summary>
        /// <param name="placeholder">プレースホルダー</param>
        /// <param name="targetFont">設定するフォント名</param>
        /// <returns>変更されたテキスト範囲数</returns>
        private int ProcessPlaceholderFont(PowerPoint.Shape placeholder, string targetFont)
        {
            int changedCount = 0;

            try
            {
                if (placeholder.HasTextFrame == MsoTriState.msoTrue)
                {
                    changedCount += ProcessTextFrameFont(placeholder.TextFrame, targetFont, $"Placeholder_{placeholder.Name}");
                }
            }
            catch (Exception ex)
            {
                logger.Error(ex, $"Failed to process font for placeholder {placeholder.Name}");
            }

            return changedCount;
        }

        /// <summary>
        /// テキストフレームのフォントを詳細処理します
        /// </summary>
        /// <param name="textFrame">テキストフレーム</param>
        /// <param name="targetFont">設定するフォント名</param>
        /// <param name="shapeName">図形名（ログ用）</param>
        /// <returns>変更されたテキスト範囲数</returns>
        private int ProcessTextFrameFont(PowerPoint.TextFrame textFrame, string targetFont, string shapeName)
        {
            int changedCount = 0;

            try
            {
                if (textFrame.HasText != MsoTriState.msoTrue)
                {
                    return 0;
                }

                var textRange = textFrame.TextRange;
                if (textRange == null || string.IsNullOrEmpty(textRange.Text))
                {
                    return 0;
                }

                // 方法1: 全体のフォントを一括変更
                try
                {
                    textRange.Font.Name = targetFont;
                    changedCount++;
                    logger.Debug($"Changed font for entire text range in {shapeName}");
                }
                catch (Exception ex)
                {
                    logger.Warn(ex, $"Failed to change font for entire text range in {shapeName}, trying character-by-character");

                    // 方法2: 文字単位での変更（フォールバック）
                    try
                    {
                        for (int i = 1; i <= textRange.Length; i++)
                        {
                            var charRange = textRange.Characters(i, 1);
                            try
                            {
                                charRange.Font.Name = targetFont;
                            }
                            catch (Exception charEx)
                            {
                                logger.Debug(charEx, $"Failed to change font for character {i} in {shapeName}");
                            }
                        }
                        changedCount++;
                        logger.Debug($"Changed font character-by-character for text range in {shapeName}");
                    }
                    catch (Exception charEx)
                    {
                        logger.Warn(charEx, $"Failed to change font character-by-character for {shapeName}");
                    }
                }

                // 方法3: 段落単位での変更（追加の保険）
                try
                {
                    for (int i = 1; i <= textRange.Paragraphs().Count; i++)
                    {
                        var paragraph = textRange.Paragraphs(i);
                        try
                        {
                            paragraph.Font.Name = targetFont;
                        }
                        catch (Exception paragraphEx)
                        {
                            logger.Debug(paragraphEx, $"Failed to change font for paragraph {i} in {shapeName}");
                        }
                    }
                }
                catch (Exception ex)
                {
                    logger.Debug(ex, $"Failed to process paragraphs for {shapeName}");
                }
            }
            catch (Exception ex)
            {
                logger.Error(ex, $"Failed to process text frame font for {shapeName}");
            }

            return changedCount;
        }

        /// <summary>
        /// 表のフォントを処理します
        /// </summary>
        /// <param name="table">表</param>
        /// <param name="targetFont">設定するフォント名</param>
        /// <param name="shapeName">図形名（ログ用）</param>
        /// <returns>変更されたテキスト範囲数</returns>
        private int ProcessTableFont(PowerPoint.Table table, string targetFont, string shapeName)
        {
            int changedCount = 0;

            try
            {
                for (int row = 1; row <= table.Rows.Count; row++)
                {
                    for (int col = 1; col <= table.Columns.Count; col++)
                    {
                        try
                        {
                            var cell = table.Cell(row, col);
                            if (cell.Shape.HasTextFrame == MsoTriState.msoTrue)
                            {
                                changedCount += ProcessTextFrameFont(cell.Shape.TextFrame, targetFont, $"{shapeName}_Cell[{row},{col}]");
                            }
                        }
                        catch (Exception ex)
                        {
                            logger.Debug(ex, $"Failed to process table cell [{row},{col}] in {shapeName}");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                logger.Error(ex, $"Failed to process table font for {shapeName}");
            }

            return changedCount;
        }

        /// <summary>
        /// 線の長さを揃える（新機能C）
        /// 選択した線の中で最初に選択したものを基準に長さを揃え、上端を揃える
        /// </summary>
        public void AlignLineLength()
        {
            if (!Globals.ThisAddIn.CheckFeatureAccess("AlignLineLength")) return;

            logger.Info("AlignLineLength operation started");

            var selectedShapes = helper.GetSelectedShapeInfos();
            if (!helper.ValidateSelection(selectedShapes, 2, 0, "線の長さを揃える")) return;

            var lineShapes = selectedShapes.Where(s => helper.IsLineShape(s.Shape)).ToList();
            if (lineShapes.Count < 2)
            {
                ErrorHandler.ExecuteSafely(() =>
                {
                    throw new InvalidOperationException("最低2つの線図形を選択してください。");
                }, "線の長さを揃える");
                return;
            }

            ComHelper.ExecuteWithComCleanup(() =>
            {
                // 最初に選択した線を基準として取得
                var referenceLine = lineShapes.First();
                var referenceLength = GetLineLength(referenceLine.Shape);

                logger.Debug($"Reference line: {referenceLine.Name}, Length: {referenceLength} (最初に選択)");

                // 他の線を調整（位置は移動せず長さのみ調整）
                foreach (var lineInfo in lineShapes.Skip(1))
                {
                    try
                    {
                        AdjustLineLength(lineInfo.Shape, referenceLength);

                        logger.Debug($"Adjusted line {lineInfo.Name} to length {referenceLength} (基準: {referenceLine.Name}, 位置維持)");
                    }
                    catch (Exception ex)
                    {
                        logger.Error(ex, $"Failed to adjust line {lineInfo.Name}");
                    }
                }

                logger.Info($"AlignLineLength completed for {lineShapes.Count} lines (基準: 最初選択, 位置移動なし)");
            }, lineShapes.Select(l => l.Shape).ToArray());

            logger.Info("AlignLineLength completed (基準: 最初選択, 位置移動なし)");
        }

        /// <summary>
        /// 図形に連番付与（新機能D）
        /// 選択図形に左上基準で1からの連番を既存テキストの後に追加
        /// </summary>
        public void AddSequentialNumbers()
        {
            if (!Globals.ThisAddIn.CheckFeatureAccess("AddSequentialNumbers")) return;

            logger.Info("AddSequentialNumbers operation started");

            var selectedShapes = helper.GetSelectedShapeInfos();
            if (!helper.ValidateSelection(selectedShapes, 1, 0, "図形に連番付与")) return;

            ComHelper.ExecuteWithComCleanup(() =>
            {
                // 左上基準でソート（上から下、左から右）
                var sortedShapes = selectedShapes.OrderBy(s => s.Top).ThenBy(s => s.Left).ToList();

                for (int i = 0; i < sortedShapes.Count; i++)
                {
                    var shapeInfo = sortedShapes[i];
                    var sequenceNumber = (i + 1).ToString();

                    try
                    {
                        // テキストフレームがない場合は作成
                        if (shapeInfo.Shape.HasTextFrame != MsoTriState.msoTrue)
                        {
                            // 図形にテキストフレームを追加する方法を試行
                            try
                            {
                                shapeInfo.Shape.TextFrame.TextRange.Text = sequenceNumber;
                            }
                            catch
                            {
                                logger.Warn($"Cannot add text frame to shape {shapeInfo.Name}");
                                continue;
                            }
                        }
                        else
                        {
                            // 既存テキストの後に番号を追加
                            var currentText = "";
                            if (shapeInfo.Shape.TextFrame.HasText == MsoTriState.msoTrue)
                            {
                                currentText = shapeInfo.Shape.TextFrame.TextRange.Text;
                            }

                            var newText = string.IsNullOrEmpty(currentText)
                                ? sequenceNumber
                                : currentText + sequenceNumber;

                            shapeInfo.Shape.TextFrame.TextRange.Text = newText;
                        }

                        logger.Debug($"Added sequence number {sequenceNumber} to shape {shapeInfo.Name}");
                    }
                    catch (Exception ex)
                    {
                        logger.Error(ex, $"Failed to add sequence number to shape {shapeInfo.Name}");
                    }
                }

                logger.Info($"AddSequentialNumbers completed for {selectedShapes.Count} shapes");
            }, selectedShapes.Select(s => s.Shape).ToArray());

            logger.Info("AddSequentialNumbers completed");
        }

        #endregion



        #region Built-in機能ヘルパーメソッド

        /// <summary>
        /// PowerPoint標準コマンドを実行します
        /// </summary>
        /// <param name="commandName">コマンド名</param>
        public void ExecutePowerPointCommand(string commandName)
        {

            if (!Globals.ThisAddIn.CheckFeatureAccess(commandName)) return;

            logger.Info($"ExecutePowerPointCommand operation started: {commandName}");

            try
            {
                // コマンドタイプに応じて処理を分岐
                if (IsShapeCreationCommand(commandName))
                {
                    logger.Debug($"Executing as shape creation command: {commandName}");
                    ExecuteShapeCreationCommand(commandName);
                }
                else
                {
                    logger.Debug($"Executing as standard PowerPoint command: {commandName}");
                    ExecuteStandardCommand(commandName);
                }

                logger.Info($"ExecutePowerPointCommand completed: {commandName}");
            }
            catch (Exception ex)
            {
                logger.Error(ex, $"Error in ExecutePowerPointCommand: {commandName}");
                ErrorHandler.ExecuteSafely(() => throw ex, $"コマンド実行: {GetCommandDisplayName(commandName)}");
            }
        }

        /// <summary>
        /// 図形作成系コマンドかどうかを判定します
        /// </summary>
        /// <param name="commandName">コマンド名</param>
        /// <returns>図形作成系の場合true</returns>
        private bool IsShapeCreationCommand(string commandName)
        {
            // 図形作成系コマンドのリスト
            var shapeCreationCommands = new HashSet<string>
    {
        // Shape関連
        "ShapeRectangle",
        "ShapeRoundedRectangle",
        "ShapeOval",
        "ShapeIsoscelesTriangle",
        "ShapeRectangularCallout",
        "ShapeRightArrow",
        "ShapeDownArrow",
        "ShapeLeftArrow",
        "ShapeUpArrow",
        "ShapeLine",
        "ShapeLineArrow",
        "ShapeElbowConnector",
        "ShapeElbowArrowConnector",
        "ShapeLeftBrace",
        "ShapePentagon",
        "ShapeChevron",
        
        // Text関連
        "TextBox"
    };

            return shapeCreationCommands.Contains(commandName);
        }

        /// <summary>
        /// 図形作成系コマンドを実行します
        /// </summary>
        /// <param name="commandName">コマンド名</param>
        private void ExecuteShapeCreationCommand(string commandName)
        {
            ComHelper.ExecuteWithComCleanup(() =>
            {
                var slide = helper.GetCurrentSlide();
                if (slide == null)
                {
                    throw new InvalidOperationException("アクティブなスライドが見つかりません。");
                }

                // デバッグ：配置前の状況をログ
                LogShapePlacementStatus(slide, "Before placement");

                // スライド左上基準で図形を配置
                var shape = CreateShapeAtOptimalPosition(slide, commandName);
                if (shape != null)
                {
                    shape.Select();
                    logger.Info($"Successfully placed {GetCommandDisplayName(commandName)} at position ({shape.Left:F1}, {shape.Top:F1})");

                    // デバッグ：配置後の状況をログ
                    LogShapePlacementStatus(slide, "After placement");
                }
                else
                {
                    throw new InvalidOperationException($"図形の作成に失敗しました: {commandName}");
                }
            });
        }

        /// <summary>
        /// PowerPoint標準機能コマンドを実行します
        /// </summary>
        /// <param name="commandName">コマンド名</param>
        private void ExecuteStandardCommand(string commandName)
        {
            ComHelper.ExecuteWithComCleanup(() =>
            {
                var application = applicationProvider.GetCurrentApplication();

                try
                {
                    switch (commandName)
                    {
                        // 整列系コマンド
                        case "AlignLeft":
                            ExecuteAlignmentCommand(MsoAlignCmd.msoAlignLefts);
                            break;

                        case "AlignCenterHorizontal":
                            ExecuteAlignmentCommand(MsoAlignCmd.msoAlignCenters);
                            break;

                        case "AlignRight":
                            ExecuteAlignmentCommand(MsoAlignCmd.msoAlignRights);
                            break;

                        case "AlignTop":
                            ExecuteAlignmentCommand(MsoAlignCmd.msoAlignTops);
                            break;

                        case "AlignCenterVertical":
                            ExecuteAlignmentCommand(MsoAlignCmd.msoAlignMiddles);
                            break;

                        case "AlignBottom":
                            ExecuteAlignmentCommand(MsoAlignCmd.msoAlignBottoms);
                            break;

                        // グループ化系コマンド
                        case "GroupObjects":
                            ExecuteGroupCommand();
                            break;

                        case "UngroupObjects":
                            ExecuteUngroupCommand();
                            break;

                        default:
                            logger.Warn($"Unknown standard command: {commandName}");
                            // フォールバック: ExecuteMsoを試行
                            try
                            {
                                application.CommandBars.ExecuteMso(commandName);
                                logger.Info($"Executed via ExecuteMso: {commandName}");
                            }
                            catch (Exception msoEx)
                            {
                                logger.Error(msoEx, $"Failed to execute via ExecuteMso: {commandName}");
                                throw new InvalidOperationException($"サポートされていないコマンドです: {commandName}");
                            }
                            break;
                    }

                    logger.Info($"Standard command executed successfully: {commandName}");
                }
                catch (COMException comEx)
                {
                    logger.Error(comEx, $"COM error executing standard command: {commandName}");
                    throw new InvalidOperationException($"PowerPoint標準機能の実行に失敗しました: {commandName}", comEx);
                }
            });
        }

        /// <summary>
        /// 整列コマンドを実行します
        /// </summary>
        /// <param name="alignCommand">整列タイプ</param>
        private void ExecuteAlignmentCommand(MsoAlignCmd alignCommand)
        {
            var application = applicationProvider.GetCurrentApplication();

            try
            {
                // 選択されている図形を取得
                var selection = application.ActiveWindow.Selection;

                if (selection.Type != PowerPoint.PpSelectionType.ppSelectionShapes)
                {
                    throw new InvalidOperationException("図形が選択されていません。整列を行うには図形を選択してください。");
                }

                if (selection.ShapeRange.Count < 2)
                {
                    throw new InvalidOperationException("整列を行うには2つ以上の図形を選択してください。");
                }

                // 整列実行
                selection.ShapeRange.Align(alignCommand, MsoTriState.msoFalse);
                logger.Debug($"Alignment executed: {alignCommand} on {selection.ShapeRange.Count} shapes");
            }
            catch (COMException comEx)
            {
                logger.Error(comEx, $"Failed to execute alignment command: {alignCommand}");
                throw new InvalidOperationException($"整列コマンドの実行に失敗しました: {alignCommand}", comEx);
            }
        }

        /// <summary>
        /// グループ化コマンドを実行します
        /// </summary>
        private void ExecuteGroupCommand()
        {
            var application = applicationProvider.GetCurrentApplication();

            try
            {
                var selection = application.ActiveWindow.Selection;

                if (selection.Type != PowerPoint.PpSelectionType.ppSelectionShapes)
                {
                    throw new InvalidOperationException("図形が選択されていません。グループ化を行うには図形を選択してください。");
                }

                if (selection.ShapeRange.Count < 2)
                {
                    throw new InvalidOperationException("グループ化を行うには2つ以上の図形を選択してください。");
                }

                // グループ化実行
                var groupedShape = selection.ShapeRange.Group();
                groupedShape.Select();
                logger.Debug($"Grouped {selection.ShapeRange.Count} shapes");
            }
            catch (COMException comEx)
            {
                logger.Error(comEx, "Failed to execute group command");
                throw new InvalidOperationException("グループ化コマンドの実行に失敗しました", comEx);
            }
        }

        /// <summary>
        /// グループ解除コマンドを実行します
        /// </summary>
        private void ExecuteUngroupCommand()
        {
            var application = applicationProvider.GetCurrentApplication();

            try
            {
                var selection = application.ActiveWindow.Selection;

                if (selection.Type != PowerPoint.PpSelectionType.ppSelectionShapes)
                {
                    throw new InvalidOperationException("図形が選択されていません。グループ解除を行うには図形を選択してください。");
                }

                if (selection.ShapeRange.Count != 1)
                {
                    throw new InvalidOperationException("グループ解除を行うには1つのグループを選択してください。");
                }

                var shape = selection.ShapeRange[1];
                if (shape.Type != MsoShapeType.msoGroup)
                {
                    throw new InvalidOperationException("選択されている図形はグループではありません。");
                }

                // グループ解除実行
                var ungroupedShapes = shape.Ungroup();
                ungroupedShapes.Select();
                logger.Debug($"Ungrouped shape into {ungroupedShapes.Count} shapes");
            }
            catch (COMException comEx)
            {
                logger.Error(comEx, "Failed to execute ungroup command");
                throw new InvalidOperationException("グループ解除コマンドの実行に失敗しました", comEx);
            }
        }

        /// <summary>
        /// 最適な位置に図形を作成します（2列斜め配置システム + スタイル適用）
        /// </summary>
        /// <param name="slide">対象スライド</param>
        /// <param name="commandName">コマンド名</param>
        /// <returns>作成された図形</returns>
        private PowerPoint.Shape CreateShapeAtOptimalPosition(PowerPoint.Slide slide, string commandName)
        {
            // 図形の標準サイズを取得
            var shapeSize = GetStandardShapeSize(commandName);

            // 2列斜め配置の次の位置を計算
            var position = CalculateDiagonalColumnPosition(slide, shapeSize.Width, shapeSize.Height);

            logger.Info($"Creating {commandName} at diagonal position: ({position.Left:F1}, {position.Top:F1})");

            // 図形を作成
            var createdShape = CreateShapeAtPosition(slide, commandName, position.Left, position.Top, shapeSize.Width, shapeSize.Height);

            if (createdShape != null)
            {
                // 【新機能】作成した図形にスタイルを適用
                ApplyShapeStyle(createdShape, commandName);
            }

            return createdShape;
        }


        /// <summary>
        /// 2列斜め配置の次の位置を計算します
        /// </summary>
        /// <param name="slide">対象スライド</param>
        /// <param name="shapeWidth">図形幅</param>
        /// <param name="shapeHeight">図形高さ</param>
        /// <returns>次の配置位置</returns>
        private (float Left, float Top) CalculateDiagonalColumnPosition(PowerPoint.Slide slide, float shapeWidth, float shapeHeight)
        {
            // 基準設定
            var column1BaseLeft = 30f;   // 1列目の基準左位置
            var column2BaseLeft = 180f;  // 2列目の基準左位置（1列目から150pt右）
            var baseTop = 50f;           // 基準上位置
            var diagonalOffsetStep = 15f; // 斜めオフセットのステップ（右下に15ptずつ）
            var maxShapesPerColumn = 10;  // 1列あたりの最大図形数

            // 現在のスライドの図形数を取得
            int currentShapeCount = slide.Shapes.Count;

            logger.Debug($"Current shape count: {currentShapeCount}");

            // 図形番号（0から始まる）
            int shapeIndex = currentShapeCount;

            // どちらの列かを判定
            bool isColumn2 = shapeIndex >= maxShapesPerColumn;
            int columnIndex = isColumn2 ? shapeIndex - maxShapesPerColumn : shapeIndex;

            // 2列目でも上限を超えた場合は、さらに右に新しい列を作成
            if (shapeIndex >= maxShapesPerColumn * 2)
            {
                var additionalColumns = (shapeIndex - maxShapesPerColumn * 2) / maxShapesPerColumn;
                var columnBaseLeft = column1BaseLeft + (additionalColumns + 2) * 150f; // 150ptずつ右にずらす
                columnIndex = (shapeIndex - maxShapesPerColumn * 2) % maxShapesPerColumn;

                var left = columnBaseLeft + columnIndex * diagonalOffsetStep;
                var top = baseTop + columnIndex * diagonalOffsetStep;

                logger.Info($"Column {additionalColumns + 3}: shape {shapeIndex} at ({left:F1}, {top:F1})");
                return (left, top);
            }

            // 1列目または2列目の位置計算
            var baseLeft = isColumn2 ? column2BaseLeft : column1BaseLeft;
            var calculatedLeft = baseLeft + columnIndex * diagonalOffsetStep;
            var calculatedTop = baseTop + columnIndex * diagonalOffsetStep;

            // スライド境界チェック
            var slideWidth = slide.Parent.PageSetup.SlideWidth;
            var slideHeight = slide.Parent.PageSetup.SlideHeight;

            if (calculatedLeft + shapeWidth > slideWidth - 20)
            {
                calculatedLeft = slideWidth - shapeWidth - 20;
                logger.Debug($"Adjusted for slide width boundary: left = {calculatedLeft:F1}");
            }

            if (calculatedTop + shapeHeight > slideHeight - 20)
            {
                calculatedTop = slideHeight - shapeHeight - 20;
                logger.Debug($"Adjusted for slide height boundary: top = {calculatedTop:F1}");
            }

            var columnName = isColumn2 ? "Column 2" : "Column 1";
            logger.Info($"{columnName}: shape {shapeIndex} (column index {columnIndex}) at ({calculatedLeft:F1}, {calculatedTop:F1})");

            return (calculatedLeft, calculatedTop);
        }



        /// <summary>
        /// 図形の標準サイズを取得します
        /// </summary>
        /// <param name="commandName">コマンド名</param>
        /// <returns>幅と高さ</returns>
        private (float Width, float Height) GetStandardShapeSize(string commandName)
        {
            switch (commandName)
            {
                case "ShapeRectangle":
                case "ShapeRoundedRectangle":
                    return (100f, 60f);

                case "ShapeOval":
                    return (80f, 80f);

                case "ShapeRightArrow":
                case "ShapeLeftArrow":
                    return (100f, 60f); // 横向き矢印は横長

                case "ShapeDownArrow":
                case "ShapeUpArrow":
                    return (60f, 100f); // 縦向き矢印は縦長

                case "ShapeRectangularCallout":
                    return (120f, 70f);

                case "ShapeIsoscelesTriangle":
                    return (80f, 70f);

                case "ShapeLine":
                case "ShapeLineArrow":
                    return (100f, 0f); // 線は高さ0

                case "ShapeElbowConnector":
                case "ShapeElbowArrowConnector":
                    return (80f, 50f);

                case "TextBox":
                    return (150f, 40f);

                case "ShapeLeftBrace":
                    return (30f, 100f);

                case "ShapePentagon":
                    return (80f, 80f); // 五角形は正方形ベース

                case "ShapeChevron":
                    return (100f, 60f); // シェブロンは横長

                default:
                    return (80f, 60f); // デフォルトサイズ
            }
        }



        /// <summary>
        /// 図形位置情報クラス
        /// </summary>
        private class ShapePositionInfo
        {
            public string Name { get; set; }
            public float Left { get; set; }
            public float Top { get; set; }
            public float Width { get; set; }
            public float Height { get; set; }
            public float Right { get; set; }
            public float Bottom { get; set; }
        }


        /// <summary>
        /// 指定位置に図形を作成します
        /// </summary>
        /// <param name="slide">対象スライド</param>
        /// <param name="commandName">コマンド名</param>
        /// <param name="left">左位置</param>
        /// <param name="top">上位置</param>
        /// <param name="width">幅</param>
        /// <param name="height">高さ</param>
        /// <returns>作成された図形</returns>
        private PowerPoint.Shape CreateShapeAtPosition(PowerPoint.Slide slide, string commandName,
            float left, float top, float width, float height)
        {
            try
            {
                PowerPoint.Shape shape = null;

                switch (commandName)
                {
                    case "ShapeRectangle":
                        shape = slide.Shapes.AddShape(MsoAutoShapeType.msoShapeRectangle, left, top, width, height);
                        break;

                    case "ShapeRoundedRectangle":
                        shape = slide.Shapes.AddShape(MsoAutoShapeType.msoShapeRoundedRectangle, left, top, width, height);
                        break;

                    case "ShapeOval":
                        shape = slide.Shapes.AddShape(MsoAutoShapeType.msoShapeOval, left, top, width, height);
                        break;

                    case "ShapeIsoscelesTriangle":
                        shape = slide.Shapes.AddShape(MsoAutoShapeType.msoShapeIsoscelesTriangle, left, top, width, height);
                        break;

                    case "ShapeRightArrow":
                        shape = slide.Shapes.AddShape(MsoAutoShapeType.msoShapeRightArrow, left, top, width, height);
                        logger.Debug($"Created right arrow shape at ({left:F1}, {top:F1})");
                        break;

                    case "ShapeDownArrow":
                        shape = slide.Shapes.AddShape(MsoAutoShapeType.msoShapeDownArrow, left, top, width, height);
                        logger.Debug($"Created down arrow shape at ({left:F1}, {top:F1})");
                        break;

                    case "ShapeLeftArrow":
                        shape = slide.Shapes.AddShape(MsoAutoShapeType.msoShapeLeftArrow, left, top, width, height);
                        logger.Debug($"Created left arrow shape at ({left:F1}, {top:F1})");
                        break;

                    case "ShapeUpArrow":
                        shape = slide.Shapes.AddShape(MsoAutoShapeType.msoShapeUpArrow, left, top, width, height);
                        logger.Debug($"Created up arrow shape at ({left:F1}, {top:F1})");
                        break;

                    case "ShapeRectangularCallout":
                        shape = slide.Shapes.AddShape(MsoAutoShapeType.msoShapeRectangularCallout, left, top, width, height);
                        shape.TextFrame.TextRange.Text = "テキストを入力";
                        break;

                    case "ShapeLine":
                        shape = slide.Shapes.AddLine(left, top + height / 2, left + width, top + height / 2);
                        break;

                    case "ShapeLineArrow":
                        shape = slide.Shapes.AddLine(left, top + height / 2, left + width, top + height / 2);
                        shape.Line.EndArrowheadStyle = MsoArrowheadStyle.msoArrowheadTriangle;
                        shape.Line.EndArrowheadWidth = MsoArrowheadWidth.msoArrowheadWidthMedium;
                        shape.Line.EndArrowheadLength = MsoArrowheadLength.msoArrowheadLengthMedium;
                        break;

                    case "ShapeElbowConnector":
                        shape = slide.Shapes.AddConnector(MsoConnectorType.msoConnectorElbow,
                            left, top, left + width, top + height);
                        break;

                    case "ShapeElbowArrowConnector":
                        shape = slide.Shapes.AddConnector(MsoConnectorType.msoConnectorElbow,
                            left, top, left + width, top + height);
                        shape.Line.EndArrowheadStyle = MsoArrowheadStyle.msoArrowheadTriangle;
                        shape.Line.EndArrowheadWidth = MsoArrowheadWidth.msoArrowheadWidthMedium;
                        shape.Line.EndArrowheadLength = MsoArrowheadLength.msoArrowheadLengthMedium;
                        break;

                    case "TextBox":
                        shape = slide.Shapes.AddTextbox(MsoTextOrientation.msoTextOrientationHorizontal,
                            left, top, width, height);
                        shape.TextFrame.TextRange.Text = "テキストを入力";
                        // 自動サイズ調整を無効にする
                        shape.TextFrame.AutoSize = PowerPoint.PpAutoSize.ppAutoSizeNone;
                        logger.Debug($"Created text box with fixed size (no auto-resize)");
                        break;

                    case "ShapeLeftBrace":
                        shape = slide.Shapes.AddShape(MsoAutoShapeType.msoShapeLeftBrace, left, top, width, height);
                        break;

                    case "ShapePentagon":
                        shape = slide.Shapes.AddShape(MsoAutoShapeType.msoShapePentagon, left, top, width, height);
                        logger.Debug($"Created pentagon shape at ({left:F1}, {top:F1})");
                        break;

                    case "ShapeChevron":
                        shape = slide.Shapes.AddShape(MsoAutoShapeType.msoShapeChevron, left, top, width, height);
                        logger.Debug($"Created chevron shape at ({left:F1}, {top:F1})");
                        break;

                    default:
                        logger.Warn($"Unknown command name: {commandName}, creating rectangle as fallback");
                        shape = slide.Shapes.AddShape(MsoAutoShapeType.msoShapeRectangle, left, top, width, height);
                        break;
                }

                if (shape != null)
                {
                    logger.Debug($"Created {commandName} at ({left:F1}, {top:F1}) size {width:F1}x{height:F1}");
                }

                return shape;
            }
            catch (Exception ex)
            {
                logger.Error(ex, $"Failed to create shape: {commandName}");
                return null;
            }
        }

        /// <summary>
        /// コマンドの表示名を取得
        /// </summary>
        /// <param name="commandName">コマンド名</param>
        /// <returns>表示名</returns>
        private string GetCommandDisplayName(string commandName)
        {
            switch (commandName)
            {
                case "ShapeRectangle":
                    return "四角形";
                case "ShapeRoundedRectangle":
                    return "角丸四角形";
                case "ShapeOval":
                    return "楕円";
                case "ShapeIsoscelesTriangle":
                    return "三角形";
                case "ShapeRectangularCallout":
                    return "吹き出し";
                case "ShapeLine":
                    return "直線";
                case "ShapeLineArrow":
                    return "矢印線";
                case "ShapeElbowConnector":
                    return "鍵型コネクタ";
                case "ShapeElbowArrowConnector":
                    return "鍵型矢印コネクタ";
                case "TextBox":
                    return "テキストボックス";
                case "ShapeLeftBrace":
                    return "中括弧";
                case "ShapePentagon":
                    return "五角形";
                case "ShapeChevron":
                    return "シェブロン";
                default:
                    return commandName;
            }
        }


        #region 図形スタイル設定機能（新機能）

        /// <summary>
        /// 図形スタイル設定ダイアログを表示
        /// </summary>
        public void ShowShapeStyleDialog()
        {
            if (!Globals.ThisAddIn.CheckFeatureAccess("ShapeStyleSettings")) return;

            logger.Info("ShowShapeStyleDialog operation started");

            try
            {
                using (var dialog = new ShapeStyleDialog())
                {
                    var result = dialog.ShowDialog();

                    if (result == DialogResult.OK)
                    {
                        logger.Info("Shape style settings updated successfully");
                        MessageBox.Show(
                            "図形スタイル設定が保存されました。\n新しく作成する図形に設定が適用されます。",
                            "図形スタイル設定",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Information
                        );
                    }
                    else
                    {
                        logger.Info("Shape style settings dialog cancelled");
                    }
                }
            }
            catch (Exception ex)
            {
                logger.Error(ex, "Error showing shape style dialog");
                ErrorHandler.ExecuteSafely(() => throw ex, "図形スタイル設定");
            }
        }

        /// <summary>
        /// 作成された図形にスタイルを適用
        /// </summary>
        /// <param name="shape">対象図形</param>
        /// <param name="commandName">コマンド名（図形種別の判定用）</param>
        /// <summary>
        /// 作成された図形にスタイルを適用
        /// </summary>
        /// <param name="shape">対象図形</param>
        /// <param name="commandName">コマンド名（図形種別の判定用）</param>
        /// <summary>
        /// 作成された図形にスタイルを適用
        /// </summary>
        /// <param name="shape">対象図形</param>
        /// <param name="commandName">コマンド名（図形種別の判定用）</param>
        private void ApplyShapeStyle(PowerPoint.Shape shape, string commandName)
        {
            if (shape == null) return;

            // 【修正】テキストボックスには一括スタイル設定を適用しない
            if (IsTextBoxCommand(commandName))
            {
                logger.Debug($"Skipping style application for TextBox: {shape.Name}");
                return;
            }

            try
            {
                // 現在の設定を読み込み
                var settings = SettingsService.Instance.LoadShapeStyleSettings();

                if (!settings.IsApplicable())
                {
                    logger.Debug("Shape styling is disabled or invalid settings");
                    return;
                }

                logger.Debug($"Applying shape style to {shape.Name} (command: {commandName})");

                // 色設定を適用
                ApplyFillStyle(shape, settings, commandName);
                ApplyLineStyle(shape, settings, commandName);

                // フォント色のみを適用
                ApplyFontColorStyle(shape, settings);

                logger.Info($"Shape style applied successfully to {shape.Name}");
            }
            catch (Exception ex)
            {
                logger.Error(ex, $"Failed to apply shape style to {shape.Name}");
                // スタイル適用の失敗は図形作成を妨げない
            }
        }

        /// <summary>
        /// フォント色のみを適用
        /// </summary>
        /// <param name="shape">対象図形</param>
        /// <param name="settings">スタイル設定</param>
        private void ApplyFontColorStyle(PowerPoint.Shape shape, ShapeStyleSettings settings)
        {
            if (shape == null) return;

            try
            {
                // テキストを持つ図形のフォント色を変更
                if (shape.HasTextFrame == MsoTriState.msoTrue && shape.TextFrame?.TextRange?.Font != null)
                {
                    shape.TextFrame.TextRange.Font.Color.RGB = ColorTranslator.ToOle(settings.FontColor);
                    logger.Debug($"Font color applied to {shape.Name}: {settings.FontColor.Name}");
                }
            }
            catch (Exception ex)
            {
                logger.Warn(ex, $"Failed to apply font color to {shape.Name}");
            }
        }

        /// <summary>
        /// 塗りつぶしスタイルを適用
        /// </summary>
        /// <param name="shape">対象図形</param>
        /// <param name="settings">スタイル設定</param>
        /// <param name="commandName">コマンド名</param>
        private void ApplyFillStyle(PowerPoint.Shape shape, ShapeStyleSettings settings, string commandName)
        {
            try
            {
                // 線系図形は塗りつぶし不要
                if (IsLineCommand(commandName)) return;

                if (shape.Fill != null)
                {
                    shape.Fill.Visible = MsoTriState.msoTrue;
                    shape.Fill.ForeColor.RGB = ColorTranslator.ToOle(settings.FillColor);

                    logger.Debug($"Fill color applied: {settings.FillColor.Name}");
                }
            }
            catch (Exception ex)
            {
                logger.Warn(ex, "Failed to apply fill style");
            }
        }

        /// <summary>
        /// 線スタイルを適用
        /// </summary>
        /// <param name="shape">対象図形</param>
        /// <param name="settings">スタイル設定</param>
        /// <param name="commandName">コマンド名</param>
        private void ApplyLineStyle(PowerPoint.Shape shape, ShapeStyleSettings settings, string commandName)
        {
            try
            {
                if (shape.Line != null)
                {
                    shape.Line.Visible = MsoTriState.msoTrue;
                    shape.Line.ForeColor.RGB = ColorTranslator.ToOle(settings.LineColor);

                    // 線系図形の場合は線幅を適切に設定
                    if (IsLineCommand(commandName))
                    {
                        shape.Line.Weight = 2.0f; // デフォルト線幅
                    }

                    logger.Debug($"Line color applied: {settings.LineColor.Name}");
                }
            }
            catch (Exception ex)
            {
                logger.Warn(ex, "Failed to apply line style");
            }
        }

        /// <summary>
        /// テキストボックス系コマンドかどうかを判定します
        /// </summary>
        /// <param name="commandName">コマンド名</param>
        /// <returns>テキストボックス系の場合true</returns>
        private bool IsTextBoxCommand(string commandName)
        {
            // テキストボックス系コマンドのリスト
            var textBoxCommands = new HashSet<string>
    {
        "TextBox"
        // 将来的に他のテキスト系図形が追加される場合はここに追加
    };

            return textBoxCommands.Contains(commandName);
        }

        /// <summary>
        /// 線系コマンドかどうかを判定
        /// </summary>
        /// <param name="commandName">コマンド名</param>
        /// <returns>線系コマンドの場合true</returns>
        private bool IsLineCommand(string commandName)
        {
            var lineCommands = new[]
            {
        "ShapeLine",
        "ShapeLineArrow",
        "ShapeElbowConnector",
        "ShapeElbowArrowConnector"
    };

            return lineCommands.Contains(commandName);
        }

        #endregion


        /// <summary>
        /// スライド上の図形配置状況をログ出力（デバッグ用）
        /// </summary>
        /// <param name="slide">対象スライド</param>
        private void LogShapePlacementStatus(PowerPoint.Slide slide, string phase)
        {
            try
            {
                logger.Debug($"=== {phase}: Diagonal Column Layout ===");
                logger.Debug($"Slide dimensions: {slide.Parent.PageSetup.SlideWidth:F1} x {slide.Parent.PageSetup.SlideHeight:F1}");
                logger.Debug($"Total shapes: {slide.Shapes.Count}");

                if (slide.Shapes.Count == 0)
                {
                    logger.Debug("No shapes on slide");
                    return;
                }

                // 列ごとに分類してログ出力
                var column1Shapes = new List<string>();
                var column2Shapes = new List<string>();
                var otherShapes = new List<string>();

                for (int i = 1; i <= slide.Shapes.Count; i++)
                {
                    var shape = slide.Shapes[i];
                    var shapeInfo = $"{i:D2}:({shape.Left:F1},{shape.Top:F1})";

                    if (shape.Left < 150f)
                    {
                        column1Shapes.Add(shapeInfo);
                    }
                    else if (shape.Left < 300f)
                    {
                        column2Shapes.Add(shapeInfo);
                    }
                    else
                    {
                        otherShapes.Add(shapeInfo);
                    }
                }

                logger.Debug($"Column 1 ({column1Shapes.Count} shapes): {string.Join(", ", column1Shapes)}");
                logger.Debug($"Column 2 ({column2Shapes.Count} shapes): {string.Join(", ", column2Shapes)}");
                if (otherShapes.Count > 0)
                {
                    logger.Debug($"Other columns ({otherShapes.Count} shapes): {string.Join(", ", otherShapes)}");
                }

                logger.Debug($"=== End {phase} ===");
            }
            catch (Exception ex)
            {
                logger.Warn(ex, $"Failed to log diagonal column placement status for {phase}");
            }
        }

        /// <summary>
        /// 図形からコマンド名を推定
        /// </summary>
        /// <param name="shape">図形</param>
        /// <returns>推定されたコマンド名</returns>
        private string EstimateCommandNameFromShape(PowerPoint.Shape shape)
        {
            try
            {
                switch (shape.Type)
                {
                    case MsoShapeType.msoAutoShape:
                        switch (shape.AutoShapeType)
                        {
                            case MsoAutoShapeType.msoShapeRectangle:
                                return "ShapeRectangle";
                            case MsoAutoShapeType.msoShapeRoundedRectangle:
                                return "ShapeRoundedRectangle";
                            case MsoAutoShapeType.msoShapeOval:
                                return "ShapeOval";
                            case MsoAutoShapeType.msoShapeRightArrow:
                                return "ShapeRightArrow";
                            case MsoAutoShapeType.msoShapeLeftArrow:
                                return "ShapeLeftArrow";
                            case MsoAutoShapeType.msoShapeDownArrow:
                                return "ShapeDownArrow";
                            case MsoAutoShapeType.msoShapeUpArrow:
                                return "ShapeUpArrow";
                            case MsoAutoShapeType.msoShapeIsoscelesTriangle:
                                return "ShapeIsoscelesTriangle";
                            default:
                                return "ShapeRectangle"; // デフォルト
                        }
                    case MsoShapeType.msoTextBox:
                        return "TextBox";
                    case MsoShapeType.msoLine:
                        return "ShapeLine";
                    default:
                        return "ShapeRectangle"; // デフォルト
                }
            }
            catch (Exception ex)
            {
                logger.Warn(ex, $"Failed to estimate command name for shape {shape.Name}");
                return "ShapeRectangle"; // フォールバック
            }
        }


        /// <summary>
        /// テキストを表示用に省略します（ログ出力用）
        /// ★TextFormatService.csのTruncateText()と同じパターンを流用
        /// </summary>
        /// <param name="text">元テキスト</param>
        /// <returns>省略されたテキスト</returns>
        private string TruncateText(string text)
        {
            if (string.IsNullOrEmpty(text)) return "(empty)";

            const int maxLength = 20;
            if (text.Length <= maxLength) return text;

            return text.Substring(0, maxLength) + "...";
        }

        #endregion




    }
}