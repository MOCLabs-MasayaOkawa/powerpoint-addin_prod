using ImageMagick;
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

        /// <summary>
        /// Excel to PPTX（22番機能）
        /// Excelの表データを解析し、セル数と同じ数の図形を作成して文字列を配置
        /// </summary>
        public void ExcelToPptx()
        {
            if (!Globals.ThisAddIn.CheckFeatureAccess("ExcelToPptx")) return;

            logger.Info("ExcelToPptx operation started (paste to existing matrix)");

            var selectedShapes = helper.GetSelectedShapeInfos(); // ★既存メソッド流用
            if (!helper.ValidateSelection(selectedShapes, 1, 0, "Excel貼り付け")) return; // ★既存メソッド流用

            ComHelper.ExecuteWithComCleanup(() => // ★既存パターン流用
            {
                try
                {
                    // クリップボードからExcelデータを取得（★既存メソッド流用）
                    var excelData = GetExcelDataFromClipboard();
                    if (excelData == null || excelData.Length == 0)
                    {
                        ErrorHandler.ExecuteSafely(() =>
                        {
                            throw new InvalidOperationException("Excelのデータをコピーしてから実行してください。");
                        }, "Excel貼り付け");
                        return;
                    }

                    // Excelデータの構造を解析
                    int excelRows = excelData.Length;
                    int excelCols = excelData[0].Length;

                    logger.Info($"Excel data structure: {excelRows} rows x {excelCols} columns");

                    // 表またはオブジェクトマトリクスを検出・処理
                    bool processed = false;

                    // 1. 表への貼り付け処理
                    var tableShapes = selectedShapes.Where(s => s.Shape.HasTable == MsoTriState.msoTrue).ToList();
                    if (tableShapes.Count > 0)
                    {
                        foreach (var tableShape in tableShapes)
                        {
                            if (PasteExcelDataToTable(tableShape.Shape.Table, excelData, excelRows, excelCols))
                            {
                                logger.Debug($"Pasted Excel data to table {tableShape.Name}");
                                processed = true;
                            }
                        }
                    }

                    // 2. オブジェクトマトリクスへの貼り付け処理
                    if (!processed)
                    {
                        var textBoxShapes = selectedShapes.Where(s =>
                            (s.HasTextFrame || s.Shape.Type == MsoShapeType.msoTextBox) &&
                            s.Shape.Type != MsoShapeType.msoLine).ToList();

                        if (textBoxShapes.Count >= 2)
                        {
                            var gridInfo = helper.DetectGridLayout(textBoxShapes); // ★既存メソッド流用
                            if (gridInfo != null)
                            {
                                if (PasteExcelDataToObjectMatrix(gridInfo, excelData, excelRows, excelCols))
                                {
                                    logger.Debug($"Pasted Excel data to object matrix ({gridInfo.Rows}x{gridInfo.Columns})");
                                    processed = true;
                                }
                            }
                        }
                    }

                    if (!processed)
                    {
                        ErrorHandler.ExecuteSafely(() =>
                        {
                            throw new InvalidOperationException(
                                "Excel データを貼り付けできる対象が見つかりません。\n" +
                                "表またはグリッド配置されたテキストボックスを選択してください。");
                        }, "Excel貼り付け");
                        return;
                    }

                    logger.Info("ExcelToPptx completed successfully");

                }
                catch (Exception ex)
                {
                    logger.Error(ex, "Failed to paste Excel data");
                    ErrorHandler.ExecuteSafely(() =>
                    {
                        throw new InvalidOperationException("Excelデータの貼り付けに失敗しました。");
                    }, "Excel貼り付け");
                }
            }, selectedShapes.Select(s => s.Shape).ToArray()); // ★既存パターン流用

            logger.Info("ExcelToPptx completed");
        }

        /// <summary>
        /// Excelデータを表に貼り付けます
        /// </summary>
        /// <param name="table">対象の表</param>
        /// <param name="excelData">Excelデータ</param>
        /// <param name="excelRows">Excelの行数</param>
        /// <param name="excelCols">Excelの列数</param>
        /// <returns>貼り付け成功時true</returns>
        private bool PasteExcelDataToTable(PowerPoint.Table table, string[][] excelData, int excelRows, int excelCols)
        {
            try
            {
                // サイズチェック
                if (table.Rows.Count < excelRows || table.Columns.Count < excelCols)
                {
                    logger.Warn($"Table size ({table.Rows.Count}x{table.Columns.Count}) is smaller than Excel data ({excelRows}x{excelCols})");

                    ErrorHandler.ExecuteSafely(() =>
                    {
                        throw new InvalidOperationException(
                            $"表のサイズ（{table.Rows.Count}行×{table.Columns.Count}列）が\n" +
                            $"Excelデータ（{excelRows}行×{excelCols}列）より小さいです。");
                    }, "Excel貼り付け");
                    return false;
                }

                int pastedCells = 0;

                // データを表に貼り付け
                for (int row = 0; row < excelRows; row++)
                {
                    for (int col = 0; col < excelCols; col++)
                    {
                        try
                        {
                            var cell = table.Cell(row + 1, col + 1); // PowerPoint表は1ベース
                            var cellText = excelData[row][col] ?? "";

                            if (cell.Shape.HasTextFrame == MsoTriState.msoTrue)
                            {
                                cell.Shape.TextFrame.TextRange.Text = cellText;
                                pastedCells++;

                                logger.Debug($"Pasted to table cell [{row + 1},{col + 1}]: '{TruncateText(cellText)}'");
                            }
                        }
                        catch (COMException comEx)
                        {
                            logger.Warn(comEx, $"COM error accessing table cell [{row + 1},{col + 1}]");
                        }
                        catch (Exception ex)
                        {
                            logger.Warn(ex, $"Failed to paste data to table cell [{row + 1},{col + 1}]");
                        }
                    }
                }

                logger.Info($"Pasted Excel data to table: {pastedCells}/{excelRows * excelCols} cells");
                return pastedCells > 0;
            }
            catch (Exception ex)
            {
                logger.Error(ex, "Failed to paste Excel data to table");
                return false;
            }
        }

        /// <summary>
        /// Excelデータをオブジェクトマトリクスに貼り付けます
        /// </summary>
        /// <param name="gridInfo">グリッド情報</param>
        /// <param name="excelData">Excelデータ</param>
        /// <param name="excelRows">Excelの行数</param>
        /// <param name="excelCols">Excelの列数</param>
        /// <returns>貼り付け成功時true</returns>
        private bool PasteExcelDataToObjectMatrix(PowerToolServiceHelper.GridInfo gridInfo, string[][] excelData, int excelRows, int excelCols)
        {
            try
            {
                // サイズチェック（一致または既存が大きい場合のみOK）
                if (gridInfo.Rows < excelRows || gridInfo.Columns < excelCols)
                {
                    logger.Warn($"Object matrix size ({gridInfo.Rows}x{gridInfo.Columns}) is smaller than Excel data ({excelRows}x{excelCols})");

                    ErrorHandler.ExecuteSafely(() =>
                    {
                        throw new InvalidOperationException(
                            $"オブジェクトマトリクス（{gridInfo.Rows}行×{gridInfo.Columns}列）が\n" +
                            $"Excelデータ（{excelRows}行×{excelCols}列）より小さいです。");
                    }, "Excel貼り付け");
                    return false;
                }

                int pastedShapes = 0;

                // データをオブジェクトマトリクスに貼り付け
                for (int row = 0; row < excelRows; row++)
                {
                    for (int col = 0; col < excelCols; col++)
                    {
                        try
                        {
                            // 該当位置に図形が存在するかチェック
                            if (row < gridInfo.ShapeGrid.Count && col < gridInfo.ShapeGrid[row].Count)
                            {
                                var shapeInfo = gridInfo.ShapeGrid[row][col];
                                var cellText = excelData[row][col] ?? "";

                                if (shapeInfo.HasTextFrame)
                                {
                                    shapeInfo.Shape.TextFrame.TextRange.Text = cellText;
                                    pastedShapes++;

                                    logger.Debug($"Pasted to object [{row},{col}] {shapeInfo.Name}: '{TruncateText(cellText)}'");
                                }
                                else
                                {
                                    logger.Warn($"Shape [{row},{col}] {shapeInfo.Name} has no text frame");
                                }
                            }
                        }
                        catch (COMException comEx)
                        {
                            logger.Warn(comEx, $"COM error accessing shape at [{row},{col}]");
                        }
                        catch (Exception ex)
                        {
                            logger.Warn(ex, $"Failed to paste data to shape at [{row},{col}]");
                        }
                    }
                }

                logger.Info($"Pasted Excel data to object matrix: {pastedShapes}/{excelRows * excelCols} shapes");
                return pastedShapes > 0;
            }
            catch (Exception ex)
            {
                logger.Error(ex, "Failed to paste Excel data to object matrix");
                return false;
            }
        }

        /// <summary>
        /// クリップボードからExcelデータを取得します
        /// </summary>
        /// <returns>二次元配列のExcelデータ</returns>
        private string[][] GetExcelDataFromClipboard()
        {
            try
            {
                if (Clipboard.ContainsText())
                {
                    var clipboardText = Clipboard.GetText();
                    return ParseExcelClipboardData(clipboardText);
                }
                else if (Clipboard.ContainsData(DataFormats.CommaSeparatedValue))
                {
                    var csvData = Clipboard.GetData(DataFormats.CommaSeparatedValue) as string;
                    return ParseExcelClipboardData(csvData);
                }
            }
            catch (Exception ex)
            {
                logger.Error(ex, "Failed to get Excel data from clipboard");
            }

            return null;
        }

        /// <summary>
        /// クリップボードのテキストデータをExcel形式として解析します
        /// </summary>
        /// <param name="clipboardData">クリップボードデータ</param>
        /// <returns>二次元配列のデータ</returns>
        private string[][] ParseExcelClipboardData(string clipboardData)
        {
            if (string.IsNullOrWhiteSpace(clipboardData))
                return null;

            try
            {
                // 行に分割
                var lines = clipboardData.Split(new[] { "\r\n", "\n", "\r" }, StringSplitOptions.None);
                var result = new List<string[]>();

                foreach (var line in lines)
                {
                    if (string.IsNullOrEmpty(line)) continue;

                    // タブ区切りで列に分割（Excelのデフォルト）
                    var cells = line.Split('\t');
                    result.Add(cells);
                }

                return result.ToArray();
            }
            catch (Exception ex)
            {
                logger.Error(ex, "Failed to parse Excel clipboard data");
                return null;
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








        /// <summary>
        /// 表の行高さ最適化（実用的なアプローチ）
        /// PowerPointの標準機能と推定計算を組み合わせた確実な方法
        /// </summary>
        public void OptimizeMatrixRowHeights()
        {
            if (!Globals.ThisAddIn.CheckFeatureAccess("OptimizeMatrixRowHeights")) return;

            logger.Info("OptimizeMatrixRowHeights operation started (practical approach)");

            var selectedShapes = helper.GetSelectedShapeInfos();
            if (!helper.ValidateSelection(selectedShapes, 1, 0, "行高さ最適化")) return;

            ComHelper.ExecuteWithComCleanup(() =>
            {
                // 表の場合
                var tableShapes = selectedShapes.Where(s => helper.IsTableShape(s.Shape)).ToList();
                if (tableShapes.Count > 0)
                {
                    int optimizedRows = 0;
                    foreach (var tableShape in tableShapes)
                    {
                        optimizedRows += OptimizeTableRowHeightsPractical(tableShape.Shape.Table);
                        logger.Debug($"Optimized table {tableShape.Name}");
                    }

                    logger.Info($"OptimizeMatrixRowHeights completed for {tableShapes.Count} table(s), {optimizedRows} rows optimized");

                    // ★ 追加：区切り線を自動再配置
                    RealignRowSeparatorsIfExists();
                    return;
                }

                // テキストボックス群の場合  
                var textBoxShapes = selectedShapes.Where(s => s.HasTextFrame || s.Shape.Type == MsoShapeType.msoTextBox).ToList();
                if (textBoxShapes.Count >= 2)
                {
                    var gridInfo = helper.DetectGridLayout(textBoxShapes);
                    if (gridInfo != null)
                    {
                        OptimizeTextBoxMatrixRowHeights(gridInfo);
                        logger.Info($"OptimizeMatrixRowHeights completed for text box grid ({gridInfo.Rows}x{gridInfo.Columns})");

                        // ★ 追加：区切り線を自動再配置
                        RealignRowSeparatorsIfExists();
                        return;
                    }
                }

                // エラー処理
                ErrorHandler.ExecuteSafely(() =>
                {
                    throw new InvalidOperationException("表またはグリッド配置されたテキストボックスを選択してください。");
                }, "行高さ最適化");

            }, selectedShapes.Select(s => s.Shape).ToArray());

            logger.Info("OptimizeMatrixRowHeights completed");
        }

        /// <summary>
        /// 表の行高さを最適化します（実用的なアプローチ）
        /// </summary>
        /// <param name="table">対象の表</param>
        /// <returns>最適化された行数</returns>
        private int OptimizeTableRowHeightsPractical(PowerPoint.Table table)
        {
            try
            {
                // 高さパターンマップ
                var heightPatterns = new Dictionary<int, float>
                {
                    [0] = 25f,  // 空白セル
                    [1] = 35f,  // 1行相当
                    [2] = 50f,  // 2行相当  
                    [3] = 65f,  // 3行相当
                    [4] = 80f,  // 4行相当
                };

                for (int row = 1; row <= table.Rows.Count; row++)
                {
                    float maxRequiredHeight = 25f; // 最小高さ

                    // その行の全セルで最大の必要高さを算出
                    for (int col = 1; col <= table.Columns.Count; col++)
                    {
                        try
                        {
                            var cell = table.Cell(row, col);
                            var estimatedLines = EstimateLinesInTableCell(cell);
                            var requiredHeight = heightPatterns.ContainsKey(estimatedLines) ? heightPatterns[estimatedLines] : 35f;
                            maxRequiredHeight = Math.Max(maxRequiredHeight, requiredHeight);

                            logger.Debug($"Cell [{row},{col}]: {estimatedLines} lines → {requiredHeight:F1}pt");
                        }
                        catch (Exception ex)
                        {
                            logger.Warn(ex, $"Failed to analyze cell [{row},{col}]");
                        }
                    }

                    // 行高さを設定
                    table.Rows[row].Height = maxRequiredHeight;
                    logger.Debug($"Row {row} height set to {maxRequiredHeight:F1}pt");
                }

                return table.Rows.Count;
            }
            catch (Exception ex)
            {
                logger.Error(ex, "Failed to optimize table row heights");
                throw;
            }
        }

        /// <summary>
        /// テキストボックス群の行高さを最適化します
        /// </summary>
        /// <param name="gridInfo">グリッド情報</param>
        private void OptimizeTextBoxMatrixRowHeights(PowerToolServiceHelper.GridInfo gridInfo)
        {
            try
            {
                // 各行の最適高さを計算
                var rowHeights = new float[gridInfo.Rows];

                for (int row = 0; row < gridInfo.Rows; row++)
                {
                    var rowShapes = gridInfo.ShapeGrid[row];
                    if (rowShapes.Count == 0)
                    {
                        rowHeights[row] = 25f; // デフォルト高さ
                        continue;
                    }

                    float maxRequiredHeight = 25f; // 最小高さ

                    // その行の全図形で最大の必要高さを算出
                    foreach (var shapeInfo in rowShapes)
                    {
                        try
                        {
                            var requiredHeight = CalculateTextBoxRequiredHeight(shapeInfo.Shape);
                            maxRequiredHeight = Math.Max(maxRequiredHeight, requiredHeight);
                        }
                        catch (Exception ex)
                        {
                            logger.Warn(ex, $"Failed to calculate height for {shapeInfo.Name}");
                        }
                    }

                    rowHeights[row] = maxRequiredHeight;
                    logger.Debug($"Row {row + 1} calculated height: {maxRequiredHeight:F1}pt");
                }

                // 高さを設定し、位置を再調整してグリッド構造を維持
                AdjustHeightsAndPositions(gridInfo, rowHeights);

                logger.Info($"Text box matrix row heights optimized with grid structure maintained");
            }
            catch (Exception ex)
            {
                logger.Error(ex, "Failed to optimize text box matrix row heights");
                throw;
            }
        }

        /// <summary>
        /// 高さを設定し、位置を再調整してグリッド構造を維持します
        /// </summary>
        /// <param name="gridInfo">グリッド情報</param>
        /// <param name="rowHeights">各行の高さ</param>
        private void AdjustHeightsAndPositions(PowerToolServiceHelper.GridInfo gridInfo, float[] rowHeights)
        {
            try
            {
                // 基準位置を取得
                var baseLeft = gridInfo.TopLeft.Left;
                var baseTop = gridInfo.TopLeft.Top;

                // 現在の列間隔を計算（最初の行から）
                var columnSpacing = 5f; // デフォルト
                if (gridInfo.ShapeGrid[0].Count > 1)
                {
                    var firstShape = gridInfo.ShapeGrid[0][0];
                    var secondShape = gridInfo.ShapeGrid[0][1];
                    columnSpacing = secondShape.Left - (firstShape.Left + firstShape.Width);
                }

                // 現在の行間隔を計算
                var rowSpacing = 5f; // デフォルト
                if (gridInfo.Rows > 1)
                {
                    var firstRowShape = gridInfo.ShapeGrid[0][0];
                    var secondRowShape = gridInfo.ShapeGrid[1][0];
                    rowSpacing = secondRowShape.Top - (firstRowShape.Top + firstRowShape.Height);
                }

                logger.Debug($"Grid spacing: Column={columnSpacing:F1}pt, Row={rowSpacing:F1}pt");

                var currentTop = baseTop;

                for (int row = 0; row < gridInfo.Rows; row++)
                {
                    var rowShapes = gridInfo.ShapeGrid[row];
                    if (rowShapes.Count == 0) continue;

                    var rowHeight = rowHeights[row];
                    var currentLeft = baseLeft;

                    for (int col = 0; col < rowShapes.Count; col++)
                    {
                        var shapeInfo = rowShapes[col];

                        try
                        {
                            // 位置と高さを設定
                            shapeInfo.Shape.Left = currentLeft;
                            shapeInfo.Shape.Top = currentTop;
                            shapeInfo.Shape.Height = rowHeight;

                            // 次の列の位置を計算
                            currentLeft += shapeInfo.Width + columnSpacing;

                            logger.Debug($"Adjusted [{row},{col}] {shapeInfo.Name}: " +
                                       $"Position=({shapeInfo.Shape.Left:F1}, {shapeInfo.Shape.Top:F1}), " +
                                       $"Height={rowHeight:F1}pt");
                        }
                        catch (Exception ex)
                        {
                            logger.Warn(ex, $"Failed to adjust shape {shapeInfo.Name}");
                        }
                    }

                    // 次の行の位置を計算
                    currentTop += rowHeight + rowSpacing;
                }

                logger.Info("Grid positions and heights adjusted successfully");
            }
            catch (Exception ex)
            {
                logger.Error(ex, "Failed to adjust heights and positions");
                throw;
            }
        }

        /// <summary>
        /// 表セル内のテキスト行数を推定します
        /// </summary>
        /// <param name="cell">対象セル</param>
        /// <returns>推定行数</returns>
        private int EstimateLinesInTableCell(PowerPoint.Cell cell)
        {
            try
            {
                var cellShape = cell.Shape;

                // テキストがない場合
                if (cellShape.TextFrame.HasText != MsoTriState.msoTrue)
                {
                    return 0;
                }

                var text = cellShape.TextFrame.TextRange.Text;
                if (string.IsNullOrWhiteSpace(text))
                {
                    return 0;
                }

                // 手動改行をカウント
                var manualBreaks = text.Split(new[] { "\r\n", "\r", "\n" }, StringSplitOptions.None).Length;

                // セル幅からの折り返し推定
                var cellWidth = cellShape.Width - cellShape.TextFrame.MarginLeft - cellShape.TextFrame.MarginRight;
                var fontSize = cellShape.TextFrame.TextRange.Font.Size;
                var avgCharWidth = fontSize * 0.7f; // 日本語考慮の平均文字幅
                var charsPerLine = Math.Max(1, (int)(cellWidth / avgCharWidth));

                var totalLines = 0;
                var lines = text.Split(new[] { "\r\n", "\r", "\n" }, StringSplitOptions.None);

                foreach (var line in lines)
                {
                    if (string.IsNullOrEmpty(line))
                    {
                        totalLines += 1;
                    }
                    else
                    {
                        var wrappedLines = Math.Max(1, (int)Math.Ceiling((double)line.Length / charsPerLine));
                        totalLines += wrappedLines;
                    }
                }

                return Math.Max(1, totalLines);
            }
            catch (Exception ex)
            {
                logger.Warn(ex, "Failed to estimate lines in table cell");
                return 1; // デフォルト
            }
        }

        /// <summary>
        /// テキストボックスの必要高さを計算します
        /// </summary>
        /// <param name="shape">テキストボックス図形</param>
        /// <returns>必要高さ（pt）</returns>
        private float CalculateTextBoxRequiredHeight(PowerPoint.Shape shape)
        {
            try
            {
                // テキストがない場合
                if (shape.TextFrame.HasText != MsoTriState.msoTrue)
                {
                    return shape.TextFrame.MarginTop + shape.TextFrame.MarginBottom + 15f;
                }

                // BoundHeightを試行（テキストボックスでは比較的信頼性高い）
                try
                {
                    var boundHeight = shape.TextFrame.TextRange.BoundHeight;
                    if (boundHeight > 0)
                    {
                        return boundHeight + shape.TextFrame.MarginTop + shape.TextFrame.MarginBottom;
                    }
                }
                catch (Exception ex)
                {
                    logger.Debug(ex, "BoundHeight failed, using line estimation");
                }

                // 行数ベースの推定
                var text = shape.TextFrame.TextRange.Text;
                var fontSize = shape.TextFrame.TextRange.Font.Size;
                var lines = text.Split(new[] { "\r\n", "\r", "\n" }, StringSplitOptions.None).Length;
                var lineHeight = fontSize * 1.2f;
                var estimatedTextHeight = lines * lineHeight;

                return estimatedTextHeight + shape.TextFrame.MarginTop + shape.TextFrame.MarginBottom;
            }
            catch (Exception ex)
            {
                logger.Warn(ex, $"Failed to calculate required height for {shape.Name}");
                return 30f; // デフォルト
            }
        }

        /// <summary>
        /// 表完全最適化（列幅と行高の同時最適化）
        /// 各列を最適幅に調整し、同時に行高も最適化して最もコンパクトな表を作成
        /// </summary>
        public void OptimizeTableComplete()
        {
            if (!Globals.ThisAddIn.CheckFeatureAccess("OptimizeTableComplete")) return;

            logger.Info("OptimizeTableComplete operation started");

            var selectedShapes = helper.GetSelectedShapeInfos();
            if (!helper.ValidateSelection(selectedShapes, 1, 0, "表最適化")) return;

            ComHelper.ExecuteWithComCleanup(() =>
            {
                // 表の場合
                var tableShapes = selectedShapes.Where(s => helper.IsTableShape(s.Shape)).ToList();
                if (tableShapes.Count > 0)
                {
                    int optimizedTables = 0;
                    foreach (var tableShape in tableShapes)
                    {
                        CompleteOptimizeTable(tableShape.Shape.Table);
                        optimizedTables++;
                        logger.Debug($"Complete optimized table {tableShape.Name}");
                    }

                    logger.Info($"OptimizeTableComplete completed for {optimizedTables} table(s)");

                    // ★ 追加：区切り線を自動再配置
                    RealignRowSeparatorsIfExists();
                    return;
                }

                // テキストボックス群の場合  
                var textBoxShapes = selectedShapes.Where(s => s.HasTextFrame || s.Shape.Type == MsoShapeType.msoTextBox).ToList();
                if (textBoxShapes.Count >= 2)
                {
                    var gridInfo = helper.DetectGridLayout(textBoxShapes);
                    if (gridInfo != null)
                    {
                        CompleteOptimizeTextBoxGrid(gridInfo);
                        logger.Info($"OptimizeTableComplete completed for text box grid ({gridInfo.Rows}x{gridInfo.Columns})");

                        // ★ 追加：区切り線を自動再配置
                        RealignRowSeparatorsIfExists();
                        return;
                    }
                }

                // エラー処理
                ErrorHandler.ExecuteSafely(() =>
                {
                    throw new InvalidOperationException("表またはグリッド配置されたテキストボックスを選択してください。");
                }, "表最適化");

            }, "表最適化", selectedShapes.Select(s => s.Shape).ToArray());

            logger.Info("OptimizeTableComplete completed");
        }

        /// <summary>
        /// 表の完全最適化を実行します（列幅と行高の同時最適化）
        /// </summary>
        /// <param name="table">対象の表</param>
        private void CompleteOptimizeTable(PowerPoint.Table table)
        {
            try
            {

                logger.Info($"=== Table Complete Optimization Started ===");
                logger.Info($"Table size: {table.Rows.Count}x{table.Columns.Count}");

                // Step 1: 各列の最適幅を計算（最長テキストベース）
                var optimalColumnWidths = CalculateOptimalColumnWidths(table);

                // Step 2: 各行の最適高さを計算
                var optimalRowHeights = CalculateOptimalRowHeights(table);

                // Step 3: 表全体幅を維持しながら列幅を調整
                var currentTotalWidth = 0f;
                for (int col = 1; col <= table.Columns.Count; col++)
                {
                    currentTotalWidth += table.Columns[col].Width;
                }

                var adjustedColumnWidths = AdjustColumnWidthsToFitTotalWidth(optimalColumnWidths, currentTotalWidth);

                // Step 4: 列幅を適用
                for (int col = 1; col <= table.Columns.Count; col++)
                {
                    table.Columns[col].Width = adjustedColumnWidths[col - 1];
                    logger.Debug($"Column {col} width: {table.Columns[col].Width:F1}pt → {adjustedColumnWidths[col - 1]:F1}pt");
                }

                // Step 5: 行高を適用
                for (int row = 1; row <= table.Rows.Count; row++)
                {
                    table.Rows[row].Height = optimalRowHeights[row - 1];
                    logger.Debug($"Row {row} height: {table.Rows[row].Height:F1}pt → {optimalRowHeights[row - 1]:F1}pt");
                }

                // Step 6: 最終結果をログ出力
                var finalTotalWidth = adjustedColumnWidths.Sum();
                var finalTotalHeight = optimalRowHeights.Sum();

                logger.Info($"=== Optimization Results ===");
                logger.Info($"Total width: {currentTotalWidth:F1}pt → {finalTotalWidth:F1}pt");
                logger.Info($"Total height: {finalTotalHeight:F1}pt");
                logger.Info($"Table optimized to most compact size while maintaining readability");

            }
            catch (Exception ex)
            {
                logger.Error(ex, "Failed to complete optimize table");
                throw;
            }
        }

        /// <summary>
        /// 各列の最適幅を計算します（その列の最長テキストベース）
        /// </summary>
        /// <param name="table">対象の表</param>
        /// <returns>各列の最適幅配列</returns>
        private float[] CalculateOptimalColumnWidths(PowerPoint.Table table)
        {
            var columnWidths = new float[table.Columns.Count];
            var minColumnWidth = 30f; // 最小列幅

            for (int col = 1; col <= table.Columns.Count; col++)
            {
                float maxRequiredWidth = minColumnWidth;

                // その列の全セルから最大必要幅を算出
                for (int row = 1; row <= table.Rows.Count; row++)
                {
                    try
                    {
                        var cell = table.Cell(row, col);
                        var requiredWidth = CalculateCellOptimalWidth(cell);
                        maxRequiredWidth = Math.Max(maxRequiredWidth, requiredWidth);

                        logger.Debug($"Cell [{row},{col}] optimal width: {requiredWidth:F1}pt");
                    }
                    catch (Exception ex)
                    {
                        logger.Warn(ex, $"Failed to calculate width for cell [{row},{col}]");
                    }
                }

                columnWidths[col - 1] = maxRequiredWidth;
                logger.Debug($"Column {col} optimal width: {maxRequiredWidth:F1}pt");
            }

            return columnWidths;
        }

        /// <summary>
        /// セルの最適幅を計算します（テキスト内容ベース）
        /// </summary>
        /// <param name="cell">対象セル</param>
        /// <returns>最適幅（ポイント）</returns>
        private float CalculateCellOptimalWidth(PowerPoint.Cell cell)
        {
            try
            {
                // テキストがない場合
                if (cell.Shape.TextFrame.HasText != MsoTriState.msoTrue)
                {
                    return 30f; // 最小幅
                }

                var textRange = cell.Shape.TextFrame.TextRange;
                var text = textRange.Text?.Trim() ?? "";

                if (string.IsNullOrEmpty(text))
                {
                    return 30f;
                }

                // フォント情報取得
                var fontSize = 12f;
                var fontName = "Arial";

                try
                {
                    fontSize = textRange.Font.Size;
                    fontName = textRange.Font.Name;
                }
                catch (Exception ex)
                {
                    logger.Debug(ex, "Using default font settings");
                }

                // 最長行を特定
                var lines = text.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
                var longestLine = lines.OrderByDescending(line => line.Length).FirstOrDefault() ?? text;

                // 日本語対応：改行なしの場合は全文を1行として扱う
                if (lines.Length == 1 && ContainsJapanese(text))
                {
                    longestLine = text;
                }

                // 文字幅計算
                var charWidth = GetCharWidthRatio(fontName) * fontSize;
                if (ContainsJapanese(longestLine))
                {
                    charWidth *= 1.1f; // 日本語は少し広めに
                }

                var textWidth = longestLine.Length * charWidth;

                // セル内余白を加算（左右各5pt）
                var totalWidth = textWidth + 10f;

                // 最小・最大制限
                return Math.Max(30f, Math.Min(200f, totalWidth));
            }
            catch (Exception ex)
            {
                logger.Warn(ex, "Failed to calculate cell optimal width");
                return 50f;
            }
        }

        /// <summary>
        /// 各行の最適高さを計算します
        /// </summary>
        /// <param name="table">対象の表</param>
        /// <returns>各行の最適高さ配列</returns>
        private float[] CalculateOptimalRowHeights(PowerPoint.Table table)
        {
            var rowHeights = new float[table.Rows.Count];
            var minRowHeight = 20f; // 最小行高

            // 高さパターンマップ（既存ロジックを再利用）
            var heightPatterns = new Dictionary<int, float>
            {
                [0] = 20f,  // 空白セル
                [1] = 30f,  // 1行相当
                [2] = 45f,  // 2行相当  
                [3] = 60f,  // 3行相当
                [4] = 75f,  // 4行相当
            };

            for (int row = 1; row <= table.Rows.Count; row++)
            {
                float maxRequiredHeight = minRowHeight;

                // その行の全セルから最大必要高さを算出
                for (int col = 1; col <= table.Columns.Count; col++)
                {
                    try
                    {
                        var cell = table.Cell(row, col);
                        var estimatedLines = EstimateLinesInTableCell(cell);
                        var requiredHeight = heightPatterns.ContainsKey(estimatedLines)
                            ? heightPatterns[estimatedLines] : 30f;

                        maxRequiredHeight = Math.Max(maxRequiredHeight, requiredHeight);

                        logger.Debug($"Cell [{row},{col}]: {estimatedLines} lines → {requiredHeight:F1}pt");
                    }
                    catch (Exception ex)
                    {
                        logger.Warn(ex, $"Failed to calculate height for cell [{row},{col}]");
                    }
                }

                rowHeights[row - 1] = maxRequiredHeight;
                logger.Debug($"Row {row} optimal height: {maxRequiredHeight:F1}pt");
            }

            return rowHeights;
        }

        /// <summary>
        /// テキストボックスグリッドの完全最適化を実行します
        /// </summary>
        /// <param name="gridInfo">グリッド情報</param>
        private void CompleteOptimizeTextBoxGrid(PowerToolServiceHelper.GridInfo gridInfo)
        {
            try
            {
                logger.Info($"=== TextBox Grid Complete Optimization Started ===");
                logger.Info($"Grid size: {gridInfo.Rows}x{gridInfo.Columns}");

                // Step 1: 各列の最適幅を計算
                var optimalColumnWidths = CalculateOptimalTextBoxColumnWidths(gridInfo);

                // Step 2: 各行の最適高さを計算
                var optimalRowHeights = CalculateOptimalTextBoxRowHeights(gridInfo);

                // Step 3: 現在のグリッド幅を維持しながら列幅を調整
                var currentTotalWidth = CalculateCurrentGridWidth(gridInfo);
                var currentSpacing = CalculateCurrentGridSpacing(gridInfo);
                var availableWidth = currentTotalWidth - currentSpacing * (gridInfo.Columns - 1);

                var adjustedColumnWidths = AdjustColumnWidthsToFitTotalWidth(optimalColumnWidths, availableWidth);

                // Step 4: グリッドに幅と高さを適用
                ApplyOptimizedDimensionsToGrid(gridInfo, adjustedColumnWidths, optimalRowHeights, currentSpacing);

                logger.Info($"TextBox grid optimized to most compact size");

            }
            catch (Exception ex)
            {
                logger.Error(ex, "Failed to complete optimize textbox grid");
                throw;
            }
        }

        /// <summary>
        /// テキストボックスグリッドの各列最適幅を計算します
        /// </summary>
        /// <param name="gridInfo">グリッド情報</param>
        /// <returns>各列の最適幅配列</returns>
        private float[] CalculateOptimalTextBoxColumnWidths(PowerToolServiceHelper.GridInfo gridInfo)
        {
            var columnWidths = new float[gridInfo.Columns];
            var minColumnWidth = 30f;

            for (int col = 0; col < gridInfo.Columns; col++)
            {
                float maxRequiredWidth = minColumnWidth;

                for (int row = 0; row < gridInfo.Rows; row++)
                {
                    if (col < gridInfo.ShapeGrid[row].Count)
                    {
                        try
                        {
                            var shapeInfo = gridInfo.ShapeGrid[row][col];
                            var requiredWidth = CalculateTextBoxOptimalWidth(shapeInfo.Shape);
                            maxRequiredWidth = Math.Max(maxRequiredWidth, requiredWidth);

                            logger.Debug($"TextBox [{row},{col}] optimal width: {requiredWidth:F1}pt");
                        }
                        catch (Exception ex)
                        {
                            logger.Warn(ex, $"Failed to calculate width for textbox [{row},{col}]");
                        }
                    }
                }

                columnWidths[col] = maxRequiredWidth;
                logger.Debug($"Column {col + 1} optimal width: {maxRequiredWidth:F1}pt");
            }

            return columnWidths;
        }

        /// <summary>
        /// テキストボックスの最適幅を計算します
        /// </summary>
        /// <param name="textBox">テキストボックス</param>
        /// <returns>最適幅（ポイント）</returns>
        private float CalculateTextBoxOptimalWidth(PowerPoint.Shape textBox)
        {
            try
            {
                if (textBox.HasTextFrame != MsoTriState.msoTrue || textBox.TextFrame.HasText != MsoTriState.msoTrue)
                {
                    return 30f;
                }

                var textRange = textBox.TextFrame.TextRange;
                var text = textRange.Text?.Trim() ?? "";

                if (string.IsNullOrEmpty(text))
                {
                    return 30f;
                }

                // フォント情報取得
                var fontSize = 12f;
                var fontName = "Arial";

                try
                {
                    fontSize = textRange.Font.Size;
                    fontName = textRange.Font.Name;
                }
                catch (Exception ex)
                {
                    logger.Debug(ex, "Using default font settings for textbox");
                }

                // 最長行を特定
                var lines = text.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
                var longestLine = lines.OrderByDescending(line => line.Length).FirstOrDefault() ?? text;

                // 日本語対応
                if (lines.Length == 1 && ContainsJapanese(text))
                {
                    longestLine = text;
                }

                // 文字幅計算
                var charWidth = GetCharWidthRatio(fontName) * fontSize;
                if (ContainsJapanese(longestLine))
                {
                    charWidth *= 1.1f;
                }

                var textWidth = longestLine.Length * charWidth;
                var totalWidth = textWidth + 6f; // テキストボックス内余白

                return Math.Max(30f, Math.Min(200f, totalWidth));
            }
            catch (Exception ex)
            {
                logger.Warn(ex, "Failed to calculate textbox optimal width");
                return 50f;
            }
        }

        /// <summary>
        /// テキストボックスグリッドの各行最適高さを計算します
        /// </summary>
        /// <param name="gridInfo">グリッド情報</param>
        /// <returns>各行の最適高さ配列</returns>
        private float[] CalculateOptimalTextBoxRowHeights(PowerToolServiceHelper.GridInfo gridInfo)
        {
            var rowHeights = new float[gridInfo.Rows];
            var minRowHeight = 20f;

            for (int row = 0; row < gridInfo.Rows; row++)
            {
                float maxRequiredHeight = minRowHeight;

                foreach (var shapeInfo in gridInfo.ShapeGrid[row])
                {
                    try
                    {
                        var requiredHeight = CalculateTextBoxRequiredHeight(shapeInfo.Shape);
                        maxRequiredHeight = Math.Max(maxRequiredHeight, requiredHeight);

                        logger.Debug($"TextBox [{row}] height: {requiredHeight:F1}pt");
                    }
                    catch (Exception ex)
                    {
                        logger.Warn(ex, $"Failed to calculate height for textbox in row {row}");
                    }
                }

                rowHeights[row] = maxRequiredHeight;
                logger.Debug($"Row {row + 1} optimal height: {maxRequiredHeight:F1}pt");
            }

            return rowHeights;
        }

        /// <summary>
        /// 現在のグリッド全体幅を計算します
        /// </summary>
        /// <param name="gridInfo">グリッド情報</param>
        /// <returns>現在の総幅</returns>
        private float CalculateCurrentGridWidth(PowerToolServiceHelper.GridInfo gridInfo)
        {
            var columnWidths = new float[gridInfo.Columns];

            for (int col = 0; col < gridInfo.Columns; col++)
            {
                float maxColumnWidth = 0f;
                for (int row = 0; row < gridInfo.Rows; row++)
                {
                    if (col < gridInfo.ShapeGrid[row].Count)
                    {
                        maxColumnWidth = Math.Max(maxColumnWidth, gridInfo.ShapeGrid[row][col].Width);
                    }
                }
                columnWidths[col] = maxColumnWidth;
            }

            var spacing = CalculateCurrentGridSpacing(gridInfo);
            return columnWidths.Sum() + spacing * (gridInfo.Columns - 1);
        }

        /// <summary>
        /// 現在のグリッド間隔を計算します
        /// </summary>
        /// <param name="gridInfo">グリッド情報</param>
        /// <returns>間隔（ポイント）</returns>
        private float CalculateCurrentGridSpacing(PowerToolServiceHelper.GridInfo gridInfo)
        {
            if (gridInfo.Columns > 1 && gridInfo.ShapeGrid[0].Count > 1)
            {
                var firstShape = gridInfo.ShapeGrid[0][0];
                var secondShape = gridInfo.ShapeGrid[0][1];
                return secondShape.Left - (firstShape.Left + firstShape.Width);
            }
            return 5f; // デフォルト間隔
        }

        /// <summary>
        /// 最適化された寸法をグリッドに適用します
        /// </summary>
        /// <param name="gridInfo">グリッド情報</param>
        /// <param name="columnWidths">列幅配列</param>
        /// <param name="rowHeights">行高配列</param>
        /// <param name="spacing">間隔</param>
        private void ApplyOptimizedDimensionsToGrid(PowerToolServiceHelper.GridInfo gridInfo, float[] columnWidths, float[] rowHeights, float spacing)
        {
            var baseLeft = gridInfo.TopLeft.Left;
            var baseTop = gridInfo.TopLeft.Top;
            var currentTop = baseTop;

            for (int row = 0; row < gridInfo.Rows; row++)
            {
                var currentLeft = baseLeft;
                var rowHeight = rowHeights[row];

                for (int col = 0; col < gridInfo.ShapeGrid[row].Count && col < columnWidths.Length; col++)
                {
                    var shapeInfo = gridInfo.ShapeGrid[row][col];
                    var columnWidth = columnWidths[col];

                    try
                    {
                        // 位置と寸法を設定
                        shapeInfo.Shape.Left = currentLeft;
                        shapeInfo.Shape.Top = currentTop;
                        shapeInfo.Shape.Width = columnWidth;
                        shapeInfo.Shape.Height = rowHeight;

                        currentLeft += columnWidth + spacing;

                        logger.Debug($"Optimized [{row},{col}] {shapeInfo.Name}: " +
                                   $"Size=({columnWidth:F1}x{rowHeight:F1}pt), " +
                                   $"Position=({shapeInfo.Shape.Left:F1}, {shapeInfo.Shape.Top:F1})");
                    }
                    catch (Exception ex)
                    {
                        logger.Warn(ex, $"Failed to apply dimensions to {shapeInfo.Name}");
                    }
                }

                currentTop += rowHeight + spacing;
            }

            logger.Info("Grid dimensions and positions optimized successfully");
        }

        /// <summary>
        /// 列幅を表全体幅に収まるよう調整します
        /// </summary>
        /// <param name="optimalWidths">最適列幅配列</param>
        /// <param name="targetTotalWidth">目標総幅</param>
        /// <returns>調整後の列幅配列</returns>
        private float[] AdjustColumnWidthsToFitTotalWidth(float[] optimalWidths, float targetTotalWidth)
        {
            var optimalTotalWidth = optimalWidths.Sum();

            logger.Debug($"Width adjustment: Optimal={optimalTotalWidth:F1}pt, Target={targetTotalWidth:F1}pt");

            // 最適幅の合計が目標幅より小さい場合は、余りを比例配分
            if (optimalTotalWidth <= targetTotalWidth)
            {
                var scaleFactor = targetTotalWidth / optimalTotalWidth;
                var result = optimalWidths.Select(w => w * scaleFactor).ToArray();
                logger.Debug($"Scaled up by factor: {scaleFactor:F3}");
                return result;
            }

            // 最適幅の合計が目標幅より大きい場合は、比例縮小
            // ただし、最小幅は保証する
            var minWidth = 30f;
            var shrinkFactor = targetTotalWidth / optimalTotalWidth;
            var adjustedWidths = new float[optimalWidths.Length];

            for (int i = 0; i < optimalWidths.Length; i++)
            {
                adjustedWidths[i] = Math.Max(minWidth, optimalWidths[i] * shrinkFactor);
            }

            logger.Debug($"Scaled down by factor: {shrinkFactor:F3}");
            return adjustedWidths;
        }

        /// <summary>
        /// テキストに日本語が含まれているかを判定します
        /// </summary>
        /// <param name="text">判定対象テキスト</param>
        /// <returns>日本語が含まれている場合true</returns>
        private bool ContainsJapanese(string text)
        {
            if (string.IsNullOrEmpty(text)) return false;

            return text.Any(c =>
                c >= 'あ' && c <= 'ん' ||     // ひらがな
                c >= 'ア' && c <= 'ン' ||     // カタカナ
                c >= '一' && c <= '龯');      // 漢字（基本範囲）
        }

        /// <summary>
        /// フォントに応じた文字幅比率を取得します
        /// </summary>
        /// <param name="fontName">フォント名</param>
        /// <returns>文字幅比率</returns>
        private float GetCharWidthRatio(string fontName)
        {
            var fontName_lower = fontName.ToLower();

            if (fontName_lower.Contains("arial") || fontName_lower.Contains("helvetica"))
                return 0.52f; // Arial系
            else if (fontName_lower.Contains("times"))
                return 0.48f; // Times系（やや狭い）
            else if (fontName_lower.Contains("courier"))
                return 0.60f; // 等幅フォント
            else if (fontName_lower.Contains("meiryo") || fontName_lower.Contains("yu gothic"))
                return 0.55f; // 日本語フォント
            else
                return 0.50f; // デフォルト
        }

        /// <summary>
        /// オブジェクトマトリクスに行間区切り線を追加
        /// グリッド配置されたテキストボックス間に水平の区切り線を描画
        /// </summary>
        public void AddMatrixRowSeparators()
        {
            if (!Globals.ThisAddIn.CheckFeatureAccess("AddMatrixRowSeparators")) return;

            logger.Info("AddMatrixRowSeparators operation started");

            var selectedShapes = helper.GetSelectedShapeInfos();
            if (!helper.ValidateSelection(selectedShapes, 2, 0, "行間区切り線")) return;

            // テキストボックス群のみを対象とする
            var textBoxShapes = selectedShapes.Where(s =>
                s.HasTextFrame || s.Shape.Type == MsoShapeType.msoTextBox).ToList();

            if (textBoxShapes.Count < 2)
            {
                ErrorHandler.ExecuteSafely(() =>
                {
                    throw new InvalidOperationException("テキストボックスを2つ以上選択してください。");
                }, "行間区切り線");
                return;
            }

            // グリッド配置を検出（既存メソッド流用）
            var gridInfo = helper.DetectGridLayout(textBoxShapes);
            if (gridInfo == null)
            {
                ErrorHandler.ExecuteSafely(() =>
                {
                    throw new InvalidOperationException("選択した図形がグリッド配置になっていません。\n" +
                        "行間区切り線を追加するには、行・列が整列している必要があります。");
                }, "行間区切り線");
                return;
            }

            if (gridInfo.Rows < 2)
            {
                ErrorHandler.ExecuteSafely(() =>
                {
                    throw new InvalidOperationException("区切り線を追加するには、2行以上のマトリクスが必要です。");
                }, "行間区切り線");
                return;
            }

            // 線設定ダイアログを表示
            LineSeparatorDialog dialog = null;
            try
            {
                dialog = new LineSeparatorDialog();
                var dialogResult = dialog.ShowDialog();

                if (dialogResult != DialogResult.OK)
                {
                    logger.Info("Line separator dialog cancelled by user");
                    return;
                }

                var lineStyle = dialog.LineStyle;
                var lineWeight = dialog.LineWeight;
                var lineColor = dialog.LineColor;

                logger.Info($"Line settings: Style={lineStyle}, Weight={lineWeight}pt, Color={lineColor.Name}");

                // COM管理下で区切り線を作成（既存パターン流用）
                ComHelper.ExecuteWithComCleanup(() =>
                {
                    var slide = helper.GetCurrentSlide(); // 既存メソッド流用
                    if (slide == null)
                    {
                        ErrorHandler.ExecuteSafely(() =>
                        {
                            throw new InvalidOperationException("アクティブなスライドが見つかりません。");
                        }, "行間区切り線");
                        return;
                    }

                    var createdLines = CreateRowSeparatorLines(slide, gridInfo, lineStyle, lineWeight, lineColor);

                    logger.Info($"Created {createdLines.Count} row separator lines");

                    // 作成した線を選択状態にする（既存パターン流用）
                    if (createdLines.Count > 0)
                    {
                        helper.SelectShapes(createdLines); // 既存メソッド流用
                    }

                }, selectedShapes.Select(s => s.Shape).ToArray());

                logger.Info("AddMatrixRowSeparators completed");
            }
            catch (Exception ex)
            {
                logger.Error(ex, "Failed to add matrix row separators");
                ErrorHandler.ExecuteSafely(() => throw ex, "行間区切り線");
            }
            finally
            {
                dialog?.Dispose();
            }
        }

        /// <summary>
        /// 行間区切り線を作成します
        /// </summary>
        /// <param name="slide">対象スライド</param>
        /// <param name="gridInfo">グリッド情報</param>
        /// <param name="lineStyle">線の種類</param>
        /// <param name="lineWeight">線の太さ</param>
        /// <param name="lineColor">線の色</param>
        /// <returns>作成された線図形のリスト</returns>
        private List<PowerPoint.Shape> CreateRowSeparatorLines(
            PowerPoint.Slide slide,
            PowerToolServiceHelper.GridInfo gridInfo,
            MsoLineDashStyle lineStyle,
            float lineWeight,
            Color lineColor)
        {
            var createdLines = new List<PowerPoint.Shape>();

            try
            {
                logger.Debug($"Creating row separators for {gridInfo.Rows}x{gridInfo.Columns} grid");

                // 区切り線の位置情報を計算
                var separatorPositions = CalculateRowSeparatorPositions(gridInfo);

                // 各行間に区切り線を作成（最後の行は除く）
                for (int i = 0; i < separatorPositions.Count; i++)
                {
                    var position = separatorPositions[i];

                    try
                    {
                        // 水平線を作成
                        var line = slide.Shapes.AddLine(
                            position.StartX, position.Y,
                            position.EndX, position.Y
                        );

                        // 線のプロパティを設定
                        line.Line.Weight = lineWeight;
                        line.Line.DashStyle = lineStyle;
                        line.Line.ForeColor.RGB = ColorTranslator.ToOle(lineColor);
                        line.Line.Visible = MsoTriState.msoTrue;

                        // 線の名前を設定（管理しやすくするため）
                        line.Name = $"RowSeparator_{i + 1}";

                        createdLines.Add(line);

                        logger.Debug($"Created separator line {i + 1}: ({position.StartX:F1}, {position.Y:F1}) to ({position.EndX:F1}, {position.Y:F1})");
                    }
                    catch (Exception ex)
                    {
                        logger.Warn(ex, $"Failed to create separator line {i + 1}");
                    }
                }

                return createdLines;
            }
            catch (Exception ex)
            {
                logger.Error(ex, "Failed to create row separator lines");
                throw;
            }
        }

        /// <summary>
        /// 行間区切り線の位置を計算します
        /// </summary>
        /// <param name="gridInfo">グリッド情報</param>
        /// <returns>区切り線位置のリスト</returns>
        private List<SeparatorLinePosition> CalculateRowSeparatorPositions(PowerToolServiceHelper.GridInfo gridInfo)
        {
            var positions = new List<SeparatorLinePosition>();

            try
            {
                // グリッドの左端と右端を計算
                var leftMost = float.MaxValue;
                var rightMost = float.MinValue;

                foreach (var row in gridInfo.ShapeGrid)
                {
                    foreach (var shape in row)
                    {
                        leftMost = Math.Min(leftMost, shape.Left);
                        rightMost = Math.Max(rightMost, shape.Left + shape.Width);
                    }
                }

                logger.Debug($"Grid boundaries: Left={leftMost:F1}pt, Right={rightMost:F1}pt");

                // 各行間の中央位置を計算
                for (int row = 0; row < gridInfo.Rows - 1; row++) // 最後の行は除く
                {
                    var currentRow = gridInfo.ShapeGrid[row];
                    var nextRow = gridInfo.ShapeGrid[row + 1];

                    if (currentRow.Count == 0 || nextRow.Count == 0) continue;

                    // 現在行の下端を計算
                    var currentRowBottom = currentRow.Max(s => s.Top + s.Height);

                    // 次行の上端を計算
                    var nextRowTop = nextRow.Min(s => s.Top);

                    // 中央位置を計算
                    var separatorY = (currentRowBottom + nextRowTop) / 2;

                    positions.Add(new SeparatorLinePosition
                    {
                        StartX = leftMost,
                        EndX = rightMost,
                        Y = separatorY
                    });

                    logger.Debug($"Separator {row + 1}: Y={separatorY:F1}pt (between row {row + 1} and {row + 2})");
                }

                return positions;
            }
            catch (Exception ex)
            {
                logger.Error(ex, "Failed to calculate row separator positions");
                throw;
            }
        }

        /// <summary>
        /// 区切り線の位置情報を表すクラス
        /// </summary>
        private class SeparatorLinePosition
        {
            public float StartX { get; set; }
            public float EndX { get; set; }
            public float Y { get; set; }
        }

        /// <summary>
        /// スライドから行区切り線を検出します
        /// </summary>
        /// <param name="slide">対象スライド</param>
        /// <returns>区切り線図形のリスト</returns>
        private List<PowerPoint.Shape> FindRowSeparators(PowerPoint.Slide slide)
        {
            var separators = new List<PowerPoint.Shape>();

            try
            {
                for (int i = 1; i <= slide.Shapes.Count; i++)
                {
                    var shape = slide.Shapes[i];

                    // 名前パターンで区切り線を識別（既存実装と同じ）
                    if (shape.Name.StartsWith("RowSeparator_") && shape.Type == MsoShapeType.msoLine)
                    {
                        separators.Add(shape);
                    }
                }

                // 名前順でソート（RowSeparator_1, RowSeparator_2, ...の順）
                separators.Sort((a, b) => string.Compare(a.Name, b.Name, StringComparison.OrdinalIgnoreCase));
            }
            catch (Exception ex)
            {
                logger.Error(ex, "Failed to find row separators");
            }

            return separators;
        }

        /// <summary>
        /// 区切り線を削除して再作成します（位置数が合わない場合）
        /// </summary>
        /// <param name="existingSeparators">既存区切り線</param>
        /// <param name="newPositions">新しい位置情報</param>
        /// <param name="slide">対象スライド</param>
        private void RecreateRowSeparators(List<PowerPoint.Shape> existingSeparators, List<SeparatorLinePosition> newPositions, PowerPoint.Slide slide)
        {
            try
            {
                // 既存区切り線の書式設定を保存（最初の線から）
                MsoLineDashStyle lineStyle = MsoLineDashStyle.msoLineSolid;
                float lineWeight = 1.0f;
                Color lineColor = Color.Black;

                if (existingSeparators.Count > 0)
                {
                    var firstSeparator = existingSeparators[0];
                    lineStyle = firstSeparator.Line.DashStyle;
                    lineWeight = firstSeparator.Line.Weight;
                    lineColor = ColorTranslator.FromOle(firstSeparator.Line.ForeColor.RGB);
                }

                // 既存区切り線を削除
                foreach (var separator in existingSeparators)
                {
                    try
                    {
                        separator.Delete();
                    }
                    catch (Exception ex)
                    {
                        logger.Warn(ex, $"Failed to delete separator {separator.Name}");
                    }
                }

                // 新しい位置に区切り線を再作成（既存ロジック流用）
                var newSeparators = CreateRowSeparatorLines(slide, null, lineStyle, lineWeight, lineColor, newPositions);
                logger.Info($"Recreated {newSeparators.Count} row separators with preserved formatting");
            }
            catch (Exception ex)
            {
                logger.Error(ex, "Failed to recreate row separators");
            }
        }

        /// <summary>
        /// 行間区切り線を作成します（位置指定版）
        /// </summary>
        private List<PowerPoint.Shape> CreateRowSeparatorLines(
            PowerPoint.Slide slide,
            PowerToolServiceHelper.GridInfo gridInfo, // nullの場合はpositionsを直接使用
            MsoLineDashStyle lineStyle,
            float lineWeight,
            Color lineColor,
            List<SeparatorLinePosition> positions = null)
        {
            var createdLines = new List<PowerPoint.Shape>();

            try
            {
                var separatorPositions = positions ?? CalculateRowSeparatorPositions(gridInfo);

                for (int i = 0; i < separatorPositions.Count; i++)
                {
                    var position = separatorPositions[i];

                    try
                    {
                        var line = slide.Shapes.AddLine(
                            position.StartX, position.Y,
                            position.EndX, position.Y
                        );

                        line.Line.Weight = lineWeight;
                        line.Line.DashStyle = lineStyle;
                        line.Line.ForeColor.RGB = ColorTranslator.ToOle(lineColor);
                        line.Line.Visible = MsoTriState.msoTrue;
                        line.Name = $"RowSeparator_{i + 1}";

                        createdLines.Add(line);
                    }
                    catch (Exception ex)
                    {
                        logger.Warn(ex, $"Failed to create separator line {i + 1}");
                    }
                }
            }
            catch (Exception ex)
            {
                logger.Error(ex, "Failed to create row separator lines");
                throw;
            }

            return createdLines;
        }

        /// <summary>
        /// 区切り線が存在する場合のみ再配置します（最小修正版）
        /// </summary>
        private void RealignRowSeparatorsIfExists()
        {
            try
            {
                var slide = helper.GetCurrentSlide();
                if (slide == null) return;

                var separatorShapes = new List<PowerPoint.Shape>();

                // 区切り線を検索
                for (int i = 1; i <= slide.Shapes.Count; i++)
                {
                    var shape = slide.Shapes[i];
                    if (shape.Name.StartsWith("RowSeparator_") && shape.Type == MsoShapeType.msoLine)
                    {
                        separatorShapes.Add(shape);
                    }
                }

                // 区切り線がない場合は何もしない
                if (separatorShapes.Count == 0) return;

                logger.Info($"Found {separatorShapes.Count} separators, realigning...");

                // 現在選択されている図形を取得（最適化対象）
                var selectedShapes = helper.GetSelectedShapeInfos();
                var matrixShapes = selectedShapes.Where(s =>
                    s.Shape.HasTable == MsoTriState.msoTrue ||
                    s.HasTextFrame ||
                    s.Shape.Type == MsoShapeType.msoTextBox).ToList();

                if (matrixShapes.Count == 0) return;

                // グリッド情報を取得
                PowerToolServiceHelper.GridInfo gridInfo = null;
                var tableShapes = matrixShapes.Where(s => s.Shape.HasTable == MsoTriState.msoTrue).ToList();

                if (tableShapes.Count > 0)
                {
                    var (gInfo, _) = helper.DetectTableMatrixLayout(tableShapes[0]);
                    gridInfo = gInfo;
                }
                else
                {
                    gridInfo = helper.DetectGridLayout(matrixShapes);
                }

                if (gridInfo == null) return;

                // 新しい位置を計算して移動
                var newPositions = CalculateRowSeparatorPositions(gridInfo);

                // 区切り線の数と計算された位置数が合わない場合の処理
                if (newPositions.Count != separatorShapes.Count)
                {
                    logger.Warn($"Separator count mismatch: found {separatorShapes.Count}, calculated {newPositions.Count}");

                    // 古い区切り線を全削除
                    foreach (var separator in separatorShapes)
                    {
                        try
                        {
                            separator.Delete();
                        }
                        catch (Exception ex)
                        {
                            logger.Warn(ex, $"Failed to delete separator {separator.Name}");
                        }
                    }

                    // 新しい区切り線を作成（既存メソッド流用）
                    CreateRowSeparatorLines(slide, null, MsoLineDashStyle.msoLineSolid, 1.0f,
                        Color.Black, newPositions);
                    return;
                }

                // 既存区切り線を新位置に移動
                for (int i = 0; i < Math.Min(separatorShapes.Count, newPositions.Count); i++)
                {
                    var separator = separatorShapes[i];
                    var newPos = newPositions[i];

                    try
                    {
                        separator.Left = newPos.StartX;
                        separator.Top = newPos.Y;
                        separator.Width = newPos.EndX - newPos.StartX;
                        separator.Height = 0; // 水平線
                    }
                    catch (Exception ex)
                    {
                        logger.Warn(ex, $"Failed to reposition separator {separator.Name}");
                    }
                }

                logger.Info($"Realigned {separatorShapes.Count} separators");
            }
            catch (Exception ex)
            {
                logger.Warn(ex, "Separator realignment failed, continuing...");
                // エラーは記録するが処理は継続
            }
        }

        /// <summary>
        /// 図形をセル位置に整列
        /// 選択されたマトリクス（表/テキストボックスグリッド）のセル中央に図形を整列する
        /// </summary>
        public void AlignShapesToCells()
        {
            if (!Globals.ThisAddIn.CheckFeatureAccess("AlignShapesToCells")) return;

            logger.Info("AlignShapesToCells operation started");

            var selectedShapes = helper.GetSelectedShapeInfos();
            if (!helper.ValidateSelection(selectedShapes, 2, 0, "図形セル整列")) return;

            ComHelper.ExecuteWithComCleanup(() =>
            {
                var slide = helper.GetCurrentSlide();
                if (slide == null) return;

                // マトリクス図形と整列対象図形を分離
                var (matrixShapes, targetShapes) = SeparateMatrixAndTargetShapes(selectedShapes);

                if (matrixShapes.Count == 0)
                {
                    ErrorHandler.ExecuteSafely(() =>
                    {
                        throw new InvalidOperationException(
                            "マトリクス（表またはテキストボックスのグリッド配置）が検出されません。\n" +
                            "整列の基準となる表またはグリッド状に配置されたテキストボックスを含めて選択してください。");
                    }, "図形セル整列");
                    return;
                }

                if (targetShapes.Count == 0)
                {
                    ErrorHandler.ExecuteSafely(() =>
                    {
                        throw new InvalidOperationException(
                            "マトリクス上に配置された整列対象図形が見つかりません。\n" +
                            "表/テキストボックスマトリクスと重なる位置にある図形を含めて選択してください。");
                    }, "図形セル整列");
                    return;
                }

                // マトリクス情報を取得
                var (gridInfo, isTable) = helper.DetectMatrixLayout(matrixShapes);
                if (gridInfo == null)
                {
                    ErrorHandler.ExecuteSafely(() =>
                    {
                        throw new InvalidOperationException(
                            "選択した図形からマトリクス配置を検出できませんでした。\n" +
                            "表または整列されたテキストボックスを選択してください。");
                    }, "図形セル整列");
                    return;
                }

                // セル情報を取得
                var cellInfos = GetCellInformations(gridInfo, isTable, matrixShapes);

                // 図形をセルにマッピング
                var shapeCellMappings = MapShapesToCells(targetShapes, cellInfos);

                if (shapeCellMappings.Count == 0)
                {
                    ErrorHandler.ExecuteSafely(() =>
                    {
                        throw new InvalidOperationException(
                            "マトリクスのセル範囲内に重なる図形が見つかりませんでした。\n" +
                            "図形がセルの範囲内に配置されていることを確認してください。");
                    }, "図形セル整列");
                    return;
                }

                // 各セル内の図形を中央整列
                int alignedShapesCount = 0;
                foreach (var mapping in shapeCellMappings)
                {
                    var cellInfo = mapping.Key;
                    var shapesInCell = mapping.Value;

                    AlignShapesToCellCenter(shapesInCell, cellInfo);
                    AdjustShapeZOrder(shapesInCell);
                    alignedShapesCount += shapesInCell.Count;

                    logger.Debug($"Aligned {shapesInCell.Count} shapes to cell at ({cellInfo.CenterX:F1}, {cellInfo.CenterY:F1})");
                }

                logger.Info($"AlignShapesToCells completed: {alignedShapesCount} shapes aligned to {shapeCellMappings.Count} cells");

            }, selectedShapes.Select(s => s.Shape).ToArray());
        }

        /// <summary>
        /// マトリクス図形と整列対象図形を分離
        /// </summary>
        /// <param name="selectedShapes">選択図形リスト</param>
        /// <returns>マトリクス図形と整列対象図形のタプル</returns>
        private (List<ShapeInfo> matrixShapes, List<ShapeInfo> targetShapes) SeparateMatrixAndTargetShapes(List<ShapeInfo> selectedShapes)
        {
            var matrixShapes = new List<ShapeInfo>();
            var targetShapes = new List<ShapeInfo>();

            foreach (var si in selectedShapes)
            {
                PowerPoint.Shape shp = si.Shape;
                MsoShapeType type = shp.Type;

                // --- マトリクス候補（テーブル / テキストボックス / 矩形系 / （任意）本文・タイトル・表のプレースホルダー）
                if (shp.HasTable == MsoTriState.msoTrue
                    || type == MsoShapeType.msoTextBox
                    || PowerToolServiceHelper.IsRectLikeAutoShape(shp)
                    || PowerToolServiceHelper.IsMatrixPlaceholder(shp))
                {
                    matrixShapes.Add(si);
                    continue;
                }

                // --- 除外（線・コネクタは完全スキップ）
                if (type == MsoShapeType.msoLine)
                {
                    continue;
                }

                // --- それ以外は整列対象
                targetShapes.Add(si);
            }

            return (matrixShapes, targetShapes);
        }

        /// <summary>
        /// セル情報
        /// </summary>
        public class CellInfo
        {
            public float Left { get; set; }
            public float Top { get; set; }
            public float Width { get; set; }
            public float Height { get; set; }
            public float CenterX => Left + Width / 2;
            public float CenterY => Top + Height / 2;
            public int Row { get; set; }
            public int Column { get; set; }
        }

        /// <summary>
        /// <summary>
        /// セル情報を取得
        /// </summary>
        /// <param name="gridInfo">グリッド情報</param>
        /// <param name="isTable">テーブルかどうか</param>
        /// <param name="matrixShapes">マトリクス図形（テーブルの場合に使用）</param>
        /// <returns>セル情報リスト</returns>
        private List<CellInfo> GetCellInformations(PowerToolServiceHelper.GridInfo gridInfo, bool isTable, List<ShapeInfo> matrixShapes)
        {
            var cellInfos = new List<CellInfo>();

            try
            {
                if (isTable)
                {
                    var tableShape = matrixShapes.FirstOrDefault(s => s.Shape.HasTable == MsoTriState.msoTrue);
                    if (tableShape != null)
                    {
                        return GetTableCellInformations(tableShape.Shape.Table);
                    }
                }
                else
                {
                    return GetTextBoxCellInformations(gridInfo.ShapeGrid);
                }
            }
            catch (Exception ex)
            {
                logger.Error(ex, "Failed to get cell informations");
            }

            return cellInfos;
        }

        /// <summary>
        /// テーブルのセル情報を取得
        /// </summary>
        /// <param name="table">テーブル</param>
        /// <returns>セル情報リスト</returns>
        private List<CellInfo> GetTableCellInformations(PowerPoint.Table table)
        {
            var cellInfos = new List<CellInfo>();

            try
            {
                for (int row = 1; row <= table.Rows.Count; row++)
                {
                    for (int col = 1; col <= table.Columns.Count; col++)
                    {
                        var cell = table.Cell(row, col);
                        var cellShape = cell.Shape;

                        cellInfos.Add(new CellInfo
                        {
                            Left = cellShape.Left,
                            Top = cellShape.Top,
                            Width = cellShape.Width,
                            Height = cellShape.Height,
                            Row = row,
                            Column = col
                        });
                    }
                }
            }
            catch (Exception ex)
            {
                logger.Error(ex, "Failed to get table cell informations");
            }

            return cellInfos;
        }

        /// <summary>
        /// テキストボックスのセル情報を取得
        /// </summary>
        /// <param name="shapeGrid">図形グリッド</param>
        /// <returns>セル情報リスト</returns>
        private List<CellInfo> GetTextBoxCellInformations(List<List<ShapeInfo>> shapeGrid)
        {
            var cellInfos = new List<CellInfo>();

            try
            {
                for (int row = 0; row < shapeGrid.Count; row++)
                {
                    for (int col = 0; col < shapeGrid[row].Count; col++)
                    {
                        var shapeInfo = shapeGrid[row][col];
                        if (shapeInfo != null)
                        {
                            cellInfos.Add(new CellInfo
                            {
                                Left = shapeInfo.Shape.Left,
                                Top = shapeInfo.Shape.Top,
                                Width = shapeInfo.Shape.Width,
                                Height = shapeInfo.Shape.Height,
                                Row = row + 1, // 1ベースに変換
                                Column = col + 1 // 1ベースに変換
                            });
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                logger.Error(ex, "Failed to get textbox cell informations");
            }

            return cellInfos;
        }

        /// <summary>
        /// 図形をセルにマッピング
        /// </summary>
        /// <param name="targetShapes">整列対象図形</param>
        /// <param name="cellInfos">セル情報リスト</param>
        /// <returns>セル-図形マッピング</returns>
        private Dictionary<CellInfo, List<ShapeInfo>> MapShapesToCells(List<ShapeInfo> targetShapes, List<CellInfo> cellInfos)
        {
            var shapeCellMappings = new Dictionary<CellInfo, List<ShapeInfo>>();

            foreach (var cellInfo in cellInfos)
            {
                shapeCellMappings[cellInfo] = new List<ShapeInfo>();
            }

            foreach (var shapeInfo in targetShapes)
            {
                var shapeCenterX = shapeInfo.Shape.Left + shapeInfo.Shape.Width / 2;
                var shapeCenterY = shapeInfo.Shape.Top + shapeInfo.Shape.Height / 2;

                // 図形中心点がセル内にあるセルを検索
                var containingCell = cellInfos.FirstOrDefault(cell =>
                    shapeCenterX >= cell.Left &&
                    shapeCenterX <= cell.Left + cell.Width &&
                    shapeCenterY >= cell.Top &&
                    shapeCenterY <= cell.Top + cell.Height);

                if (containingCell != null)
                {
                    shapeCellMappings[containingCell].Add(shapeInfo);
                    logger.Debug($"Mapped shape {shapeInfo.Name} to cell [{containingCell.Row},{containingCell.Column}]");
                }
                else
                {
                    logger.Debug($"Shape {shapeInfo.Name} is outside all cell boundaries - ignored");
                }
            }

            // 空のセルを除去
            return shapeCellMappings.Where(kvp => kvp.Value.Count > 0).ToDictionary(kvp => kvp.Key, kvp => kvp.Value);
        }

        /// <summary>
        /// 図形をセル中央に整列
        /// </summary>
        /// <param name="shapesInCell">セル内図形リスト</param>
        /// <param name="cellInfo">セル情報</param>
        private void AlignShapesToCellCenter(List<ShapeInfo> shapesInCell, CellInfo cellInfo)
        {
            foreach (var shapeInfo in shapesInCell)
            {
                try
                {
                    // セル中央座標を計算
                    var cellCenterX = cellInfo.CenterX;
                    var cellCenterY = cellInfo.CenterY;

                    // 図形を中央に移動
                    var newLeft = cellCenterX - shapeInfo.Shape.Width / 2;
                    var newTop = cellCenterY - shapeInfo.Shape.Height / 2;

                    shapeInfo.Shape.Left = newLeft;
                    shapeInfo.Shape.Top = newTop;

                    logger.Debug($"Aligned shape {shapeInfo.Name} to cell center ({cellCenterX:F1}, {cellCenterY:F1})");
                }
                catch (Exception ex)
                {
                    logger.Error(ex, $"Failed to align shape {shapeInfo.Name} to cell center");
                }
            }
        }

        /// <summary>
        /// 図形のZ-orderを調整（マトリクスの上に配置）
        /// </summary>
        /// <param name="shapes">整列図形リスト</param>
        private void AdjustShapeZOrder(List<ShapeInfo> shapes)
        {
            foreach (var shapeInfo in shapes)
            {
                try
                {
                    shapeInfo.Shape.ZOrder(MsoZOrderCmd.msoBringToFront);
                    logger.Debug($"Brought shape {shapeInfo.Name} to front");
                }
                catch (Exception ex)
                {
                    logger.Error(ex, $"Failed to adjust Z-order for shape {shapeInfo.Name}");
                }
            }
        }


        /// <summary>
        /// 見出し行を付与（超シンプル版）
        /// </summary>
        public void AddHeaderRowToMatrix()
        {
            if (!Globals.ThisAddIn.CheckFeatureAccess("AddHeaderRowToMatrix")) return;

            logger.Info("AddHeaderRowToMatrix operation started");

            var selectedShapes = helper.GetSelectedShapeInfos();
            if (!helper.ValidateSelection(selectedShapes, 1, 0, "見出し行付与")) return;

            ComHelper.ExecuteWithComCleanup(() =>
            {
                var slide = helper.GetCurrentSlide();
                if (slide == null) return;

                // テーブルかグリッドかを判定
                var tableShapes = selectedShapes.Where(s => s.Shape.HasTable == MsoTriState.msoTrue).ToList();

                if (tableShapes.Count > 0)
                {
                    // 表の処理
                    ProcessTable(tableShapes);
                }
                else
                {
                    // グリッドの処理
                    ProcessGrid(selectedShapes, slide);
                }

            }, selectedShapes.Select(s => s.Shape).ToArray());

            logger.Info("AddHeaderRowToMatrix completed");
        }

        /// <summary>
        /// 表の処理
        /// </summary>
        private void ProcessTable(List<ShapeInfo> tableShapes)
        {
            foreach (var tableShape in tableShapes)
            {
                var table = tableShape.Shape.Table;
                var newRow = table.Rows.Add(1);

                for (int col = 1; col <= table.Columns.Count; col++)
                {
                    var cell = table.Cell(1, col);
                    cell.Shape.TextFrame.TextRange.Text = $"見出し{col}";
                }

                logger.Info($"Added header row to table");
            }
        }

        /// <summary>
        /// グリッドの処理
        /// </summary>
        private void ProcessGrid(List<ShapeInfo> selectedShapes, PowerPoint.Slide slide)
        {
            // 1. 選択図形の最上端を取得
            var topMost = selectedShapes.Min(s => s.Top);
            var leftMost = selectedShapes.Min(s => s.Left);
            var rightMost = selectedShapes.Max(s => s.Left + s.Width);

            // 2. 1行目の図形を特定（最上段の図形たち）
            var tolerance = 5f; // 5pt許容誤差
            var topRowShapes = selectedShapes.Where(s => Math.Abs(s.Top - topMost) <= tolerance)
                                           .OrderBy(s => s.Left)
                                           .ToList();

            logger.Info($"Found {topRowShapes.Count} shapes in top row");

            // 3. 見出し行を作成
            var headerShapes = new List<PowerPoint.Shape>();
            foreach (var topShape in topRowShapes)
            {
                var headerBox = slide.Shapes.AddTextbox(
                    MsoTextOrientation.msoTextOrientationHorizontal,
                    topShape.Left,
                    topMost - 50f, // 仮位置
                    topShape.Width,
                    30f // 仮高さ
                );

                headerBox.TextFrame.TextRange.Text = $"見出し{topRowShapes.IndexOf(topShape) + 1}";
                headerBox.Fill.Visible = MsoTriState.msoFalse;
                headerBox.Line.Visible = MsoTriState.msoFalse;
                headerBox.TextFrame.TextRange.Font.Color.RGB = 0;

                headerShapes.Add(headerBox);
            }



            // 4. 見出し行の高さを最適化
            foreach (var header in headerShapes)
            {
                header.TextFrame.AutoSize = PowerPoint.PpAutoSize.ppAutoSizeShapeToFitText;
            }
            var maxHeaderHeight = headerShapes.Max(h => h.Height);
            foreach (var header in headerShapes)
            {
                header.TextFrame.AutoSize = PowerPoint.PpAutoSize.ppAutoSizeNone;
                header.Height = maxHeaderHeight;
            }


            // 5. 0.8mm間隔で配置
            const float SPACING_PT = 10.0f * 2.835f; // 0.8mm
            var headerTop = topMost - maxHeaderHeight - SPACING_PT;


            // 見出し行を配置
            foreach (var header in headerShapes)
            {
                header.Top = headerTop;
            }

            // 5. 見出しの最終配置がある前提（header.Top = headerTop 済み）
            float headerBottom = headerTop + maxHeaderHeight;

            // 6. 中間に区切り線
            float separatorY = (headerBottom + topMost) / 2f;

            foreach (var header in headerShapes)
            {
                var line = slide.Shapes.AddLine(header.Left, separatorY, header.Left + header.Width, separatorY);
                line.Line.Weight = 0.5f;
                line.Line.ForeColor.RGB = 0;
            }

            logger.Info($"Added header row with {headerShapes.Count} cells and separator lines");
        }

        /// <summary>
        /// セルマージン設定
        /// 選択された表のセルまたはテキストボックスのマージンを統一設定する
        /// </summary>
        public void SetCellMargins()
        {
            if (!Globals.ThisAddIn.CheckFeatureAccess("SetCellMargins")) return;

            logger.Info("SetCellMargins operation started");

            var selectedShapes = helper.GetSelectedShapeInfos();
            if (!helper.ValidateSelection(selectedShapes, 1, 0, "セルマージン設定")) return;

            // マージン設定ダイアログを表示
            using (var dialog = new MarginAdjustmentDialog("セルマージン設定"))
            {
                if (dialog.ShowDialog() != DialogResult.OK)
                {
                    logger.Info("Margin setting cancelled by user");
                    return;
                }

                var (top, bottom, left, right) = dialog.GetMarginValues();
                logger.Info($"Margin settings: top={top:F2}cm, bottom={bottom:F2}cm, left={left:F2}cm, right={right:F2}cm");

                ComHelper.ExecuteWithComCleanup(() =>
                {
                    int processedShapes = 0;
                    int processedCells = 0;

                    foreach (var shapeInfo in selectedShapes)
                    {
                        try
                        {
                            if (shapeInfo.Shape.HasTable == MsoTriState.msoTrue)
                            {
                                // 表の処理
                                var cellCount = ProcessTableMargins(shapeInfo.Shape.Table, top, bottom, left, right, shapeInfo.Name);
                                processedCells += cellCount;
                                processedShapes++;
                                logger.Debug($"Processed table {shapeInfo.Name}: {cellCount} cells updated");
                            }
                            else if (shapeInfo.HasTextFrame)
                            {
                                // テキストボックスの処理
                                ProcessTextBoxMargins(shapeInfo.Shape.TextFrame, top, bottom, left, right, shapeInfo.Name);
                                processedShapes++;
                                logger.Debug($"Processed textbox {shapeInfo.Name}");
                            }
                            else
                            {
                                logger.Debug($"Skipped shape {shapeInfo.Name}: no table or textframe");
                            }
                        }
                        catch (Exception ex)
                        {
                            logger.Error(ex, $"Failed to process shape {shapeInfo.Name}");
                        }
                    }

                    if (processedShapes == 0)
                    {
                        ErrorHandler.ExecuteSafely(() =>
                        {
                            throw new InvalidOperationException(
                                "マージンを設定できる図形が見つかりません。\n" +
                                "表またはテキストボックスを選択してください。");
                        }, "セルマージン設定");
                        return;
                    }

                    var message = processedCells > 0
                        ? $"マージン設定完了: {processedShapes}個の図形, {processedCells}個のセル"
                        : $"マージン設定完了: {processedShapes}個のテキストボックス";

                    logger.Info($"SetCellMargins completed: {message}");

                }, selectedShapes.Select(s => s.Shape).ToArray());
            }
        }

        /// <summary>
        /// 表のセルマージンを処理します
        /// </summary>
        /// <param name="table">対象の表</param>
        /// <param name="top">上マージン (cm)</param>
        /// <param name="bottom">下マージン (cm)</param>
        /// <param name="left">左マージン (cm)</param>
        /// <param name="right">右マージン (cm)</param>
        /// <param name="shapeName">図形名（ログ用）</param>
        /// <returns>処理されたセル数</returns>
        private int ProcessTableMargins(PowerPoint.Table table, float top, float bottom, float left, float right, string shapeName)
        {
            int processedCells = 0;

            try
            {
                // cmをポイントに変換 (1cm = 28.35pt)
                var topPt = top * 28.35f;
                var bottomPt = bottom * 28.35f;
                var leftPt = left * 28.35f;
                var rightPt = right * 28.35f;

                for (int row = 1; row <= table.Rows.Count; row++)
                {
                    for (int col = 1; col <= table.Columns.Count; col++)
                    {
                        try
                        {
                            var cell = table.Cell(row, col);
                            var textFrame = cell.Shape.TextFrame;

                            // マージン設定
                            textFrame.MarginTop = topPt;
                            textFrame.MarginBottom = bottomPt;
                            textFrame.MarginLeft = leftPt;
                            textFrame.MarginRight = rightPt;

                            processedCells++;
                            logger.Debug($"Set margins for cell [{row},{col}] in {shapeName}");
                        }
                        catch (Exception ex)
                        {
                            logger.Warn(ex, $"Failed to set margins for cell [{row},{col}] in {shapeName}");
                        }
                    }
                }

                logger.Info($"Processed {processedCells} cells in table {shapeName}");
            }
            catch (Exception ex)
            {
                logger.Error(ex, $"Failed to process table margins for {shapeName}");
                throw;
            }

            return processedCells;
        }

        /// <summary>
        /// テキストボックスのマージンを処理します
        /// </summary>
        /// <param name="textFrame">テキストフレーム</param>
        /// <param name="top">上マージン (cm)</param>
        /// <param name="bottom">下マージン (cm)</param>
        /// <param name="left">左マージン (cm)</param>
        /// <param name="right">右マージン (cm)</param>
        /// <param name="shapeName">図形名（ログ用）</param>
        private void ProcessTextBoxMargins(PowerPoint.TextFrame textFrame, float top, float bottom, float left, float right, string shapeName)
        {
            try
            {
                // cmをポイントに変換 (1cm = 28.35pt)
                var topPt = top * 28.35f;
                var bottomPt = bottom * 28.35f;
                var leftPt = left * 28.35f;
                var rightPt = right * 28.35f;

                // マージン設定
                textFrame.MarginTop = topPt;
                textFrame.MarginBottom = bottomPt;
                textFrame.MarginLeft = leftPt;
                textFrame.MarginRight = rightPt;

                logger.Debug($"Set margins for textbox {shapeName}: " +
                           $"top={top:F2}cm, bottom={bottom:F2}cm, left={left:F2}cm, right={right:F2}cm");
            }
            catch (Exception ex)
            {
                logger.Error(ex, $"Failed to set textbox margins for {shapeName}");
                throw;
            }
        }

        /// <summary>
        /// マトリクス行追加（Phase 1: 基本機能）
        /// 表全体・オブジェクトマトリクス全体選択時に最下端に行を追加
        /// </summary>
        public void AddMatrixRow()
        {
            if (!Globals.ThisAddIn.CheckFeatureAccess("AddMatrixRow")) return;

            logger.Info("AddMatrixRow operation started (Phase 1)");

            var selectedShapes = helper.GetSelectedShapeInfos();
            if (!helper.ValidateSelection(selectedShapes, 1, 0, "行追加")) return;

            ComHelper.ExecuteWithComCleanup(() =>
            {
                bool processed = false;

                // 表の処理
                var tableShapes = selectedShapes.Where(s => s.Shape.HasTable == MsoTriState.msoTrue).ToList();
                if (tableShapes.Count > 0)
                {
                    foreach (var tableShape in tableShapes)
                    {
                        AddRowToTable(tableShape.Shape.Table);
                        logger.Debug($"Added row to table {tableShape.Name}");
                        processed = true;
                    }
                }

                // オブジェクトマトリクスの処理
                if (!processed)
                {
                    var textBoxShapes = selectedShapes.Where(s =>
                        s.HasTextFrame || s.Shape.Type == MsoShapeType.msoTextBox).ToList();

                    if (textBoxShapes.Count >= 2)
                    {
                        var gridInfo = helper.DetectGridLayout(textBoxShapes);
                        if (gridInfo != null)
                        {
                            AddRowToObjectMatrix(gridInfo);
                            logger.Debug($"Added row to object matrix ({gridInfo.Rows}x{gridInfo.Columns})");
                            processed = true;
                        }
                    }
                }

                if (!processed)
                {
                    ErrorHandler.ExecuteSafely(() =>
                    {
                        throw new InvalidOperationException(
                            "行を追加できる対象が見つかりません。\n" +
                            "表またはグリッド配置されたテキストボックスを選択してください。");
                    }, "行追加");
                    return;
                }

                // 区切り線があれば再配置（既存機能流用）
                RealignRowSeparatorsIfExists();

                logger.Info("AddMatrixRow completed successfully");

            }, selectedShapes.Select(s => s.Shape).ToArray());

            logger.Info("AddMatrixRow completed");
        }

        /// <summary>
        /// マトリクス列追加（Phase 1: 基本機能）
        /// 表全体・オブジェクトマトリクス全体選択時に最右端に列を追加
        /// </summary>
        public void AddMatrixColumn()
        {
            if (!Globals.ThisAddIn.CheckFeatureAccess("AddMatrixColumn")) return;

            logger.Info("AddMatrixColumn operation started (Phase 1)");

            var selectedShapes = helper.GetSelectedShapeInfos();
            if (!helper.ValidateSelection(selectedShapes, 1, 0, "列追加")) return;

            ComHelper.ExecuteWithComCleanup(() =>
            {
                bool processed = false;

                // 表の処理
                var tableShapes = selectedShapes.Where(s => s.Shape.HasTable == MsoTriState.msoTrue).ToList();
                if (tableShapes.Count > 0)
                {
                    foreach (var tableShape in tableShapes)
                    {
                        AddColumnToTable(tableShape.Shape.Table);
                        logger.Debug($"Added column to table {tableShape.Name}");
                        processed = true;
                    }
                }

                // オブジェクトマトリクスの処理
                if (!processed)
                {
                    var textBoxShapes = selectedShapes.Where(s =>
                        s.HasTextFrame || s.Shape.Type == MsoShapeType.msoTextBox).ToList();

                    if (textBoxShapes.Count >= 2)
                    {
                        var gridInfo = helper.DetectGridLayout(textBoxShapes);
                        if (gridInfo != null)
                        {
                            AddColumnToObjectMatrix(gridInfo);
                            logger.Debug($"Added column to object matrix ({gridInfo.Rows}x{gridInfo.Columns})");
                            processed = true;
                        }
                    }
                }

                if (!processed)
                {
                    ErrorHandler.ExecuteSafely(() =>
                    {
                        throw new InvalidOperationException(
                            "列を追加できる対象が見つかりません。\n" +
                            "表またはグリッド配置されたテキストボックスを選択してください。");
                    }, "列追加");
                    return;
                }

                // 区切り線があれば再配置（既存機能流用）
                RealignRowSeparatorsIfExists();

                logger.Info("AddMatrixColumn completed successfully");

            }, selectedShapes.Select(s => s.Shape).ToArray());

            logger.Info("AddMatrixColumn completed");
        }

        /// <summary>
        /// 表に行を追加します
        /// </summary>
        /// <param name="table">対象の表</param>
        private void AddRowToTable(PowerPoint.Table table)
        {
            try
            {
                // 最下端に行を追加
                var newRow = table.Rows.Add();

                // 新しい行の高さを適切に設定（隣接行の高さを参考）
                if (table.Rows.Count > 1)
                {
                    var referenceRowHeight = table.Rows[table.Rows.Count - 1].Height;
                    newRow.Height = referenceRowHeight;
                }
                else
                {
                    newRow.Height = 35f; // デフォルト高さ
                }

                // 新しい行のセルに基本書式を適用
                for (int col = 1; col <= table.Columns.Count; col++)
                {
                    try
                    {
                        var newCell = table.Cell(table.Rows.Count, col);

                        // 上の行のセル書式をコピー（可能な場合）
                        if (table.Rows.Count > 1)
                        {
                            var referenceCell = table.Cell(table.Rows.Count - 1, col);
                            CopyTableCellFormatNew(referenceCell, newCell);
                        }
                    }
                    catch (Exception ex)
                    {
                        logger.Warn(ex, $"Failed to format new cell [{table.Rows.Count},{col}]");
                    }
                }

                logger.Info($"Added row to table (now {table.Rows.Count} rows x {table.Columns.Count} columns)");
            }
            catch (Exception ex)
            {
                logger.Error(ex, "Failed to add row to table");
                throw;
            }
        }

        /// <summary>
        /// 表に列を追加します
        /// </summary>
        /// <param name="table">対象の表</param>
        private void AddColumnToTable(PowerPoint.Table table)
        {
            try
            {
                // 最右端に列を追加
                var newColumn = table.Columns.Add();

                // 新しい列の幅を適切に設定（隣接列の幅を参考）
                if (table.Columns.Count > 1)
                {
                    var referenceColumnWidth = table.Columns[table.Columns.Count - 1].Width;
                    newColumn.Width = referenceColumnWidth;
                }
                else
                {
                    newColumn.Width = 120f; // デフォルト幅
                }

                // 新しい列のセルに基本書式を適用
                for (int row = 1; row <= table.Rows.Count; row++)
                {
                    try
                    {
                        var newCell = table.Cell(row, table.Columns.Count);

                        // 左の列のセル書式をコピー（可能な場合）
                        if (table.Columns.Count > 1)
                        {
                            var referenceCell = table.Cell(row, table.Columns.Count - 1);
                            CopyTableCellFormatNew(referenceCell, newCell);
                        }
                    }
                    catch (Exception ex)
                    {
                        logger.Warn(ex, $"Failed to format new cell [{row},{table.Columns.Count}]");
                    }
                }

                logger.Info($"Added column to table (now {table.Rows.Count} rows x {table.Columns.Count} columns)");
            }
            catch (Exception ex)
            {
                logger.Error(ex, "Failed to add column to table");
                throw;
            }
        }

        /// <summary>
        /// 表セル書式をコピーします（新機能専用）
        /// </summary>
        /// <param name="sourceCell">コピー元セル</param>
        /// <param name="targetCell">コピー先セル</param>
        private void CopyTableCellFormatNew(PowerPoint.Cell sourceCell, PowerPoint.Cell targetCell)
        {
            try
            {
                var sourceShape = sourceCell.Shape;
                var targetShape = targetCell.Shape;

                // 背景色をコピー
                if (sourceShape.Fill.Type != MsoFillType.msoFillMixed)
                {
                    targetShape.Fill.ForeColor.RGB = sourceShape.Fill.ForeColor.RGB;
                    targetShape.Fill.Transparency = sourceShape.Fill.Transparency;
                }

                // 線をコピー
                if (sourceShape.Line.Visible == MsoTriState.msoTrue)
                {
                    targetShape.Line.Visible = MsoTriState.msoTrue;
                    targetShape.Line.ForeColor.RGB = sourceShape.Line.ForeColor.RGB;
                    targetShape.Line.Weight = sourceShape.Line.Weight;
                    targetShape.Line.DashStyle = sourceShape.Line.DashStyle;
                }

                // テキスト書式をコピー（基本のみ）
                if (sourceShape.HasTextFrame == MsoTriState.msoTrue &&
                    targetShape.HasTextFrame == MsoTriState.msoTrue)
                {
                    var sourceTextRange = sourceShape.TextFrame.TextRange;
                    var targetTextRange = targetShape.TextFrame.TextRange;

                    targetTextRange.Font.Name = sourceTextRange.Font.Name;
                    targetTextRange.Font.Size = sourceTextRange.Font.Size;
                    targetTextRange.Font.Bold = sourceTextRange.Font.Bold;
                    targetTextRange.ParagraphFormat.Alignment = sourceTextRange.ParagraphFormat.Alignment;
                }

                logger.Debug("Cell format copied successfully");
            }
            catch (Exception ex)
            {
                logger.Debug(ex, "Failed to copy cell format (non-critical)");
                // 書式コピー失敗は致命的ではないので継続
            }
        }

        /// <summary>
        /// オブジェクトマトリクスに行を追加します
        /// </summary>
        /// <param name="gridInfo">グリッド情報</param>
        private void AddRowToObjectMatrix(PowerToolServiceHelper.GridInfo gridInfo)
        {
            try
            {
                var slide = helper.GetCurrentSlide();
                if (slide == null)
                {
                    throw new InvalidOperationException("アクティブなスライドが見つかりません。");
                }

                var createdShapes = new List<PowerPoint.Shape>();

                // 最下段の図形を参考に新しい行を作成
                var lastRowShapes = gridInfo.ShapeGrid[gridInfo.Rows - 1];
                var referenceY = lastRowShapes.Max(s => s.Top + s.Height);

                // 行間隔を計算（既存行間の平均を使用）
                float rowSpacing = 5f; // デフォルト
                if (gridInfo.Rows > 1)
                {
                    var spacings = new List<float>();
                    for (int row = 0; row < gridInfo.Rows - 1; row++)
                    {
                        var currentRowBottom = gridInfo.ShapeGrid[row].Max(s => s.Top + s.Height);
                        var nextRowTop = gridInfo.ShapeGrid[row + 1].Min(s => s.Top);
                        spacings.Add(nextRowTop - currentRowBottom);
                    }
                    rowSpacing = spacings.Average();
                }

                var newRowY = referenceY + rowSpacing;

                // 各列に新しいオブジェクトを作成
                for (int col = 0; col < gridInfo.Columns; col++)
                {
                    // 参考図形（同じ列の最下段）
                    ShapeInfo referenceShape = null;
                    for (int row = gridInfo.Rows - 1; row >= 0; row--)
                    {
                        if (col < gridInfo.ShapeGrid[row].Count)
                        {
                            referenceShape = gridInfo.ShapeGrid[row][col];
                            break;
                        }
                    }

                    if (referenceShape == null) continue;

                    // 適切なセルサイズを計算（同じ列の平均を使用）
                    var avgWidth = gridInfo.ShapeGrid.Where(row => col < row.Count)
                        .Select(row => row[col].Width).Average();
                    var avgHeight = gridInfo.ShapeGrid[gridInfo.Rows - 1]
                        .Where(s => s != null).Select(s => s.Height).DefaultIfEmpty(35f).Average();

                    // 新しいテキストボックスを作成
                    var newTextBox = slide.Shapes.AddTextbox(
                        MsoTextOrientation.msoTextOrientationHorizontal,
                        referenceShape.Left,
                        newRowY,
                        avgWidth,
                        avgHeight
                    );

                    // 書式を参考図形からコピー
                    CopyObjectShapeFormat(referenceShape.Shape, newTextBox);

                    // デフォルトテキストを設定
                    if (newTextBox.HasTextFrame == MsoTriState.msoTrue)
                    {
                        newTextBox.TextFrame.TextRange.Text = ""; // 空のテキスト
                    }

                    createdShapes.Add(newTextBox);
                    logger.Debug($"Created new cell at column {col + 1} for new row");
                }

                // 作成した図形を選択状態にする
                if (createdShapes.Count > 0)
                {
                    SelectCreatedShapes(createdShapes);
                }

                logger.Info($"Added row to object matrix (created {createdShapes.Count} new cells)");
            }
            catch (Exception ex)
            {
                logger.Error(ex, "Failed to add row to object matrix");
                throw;
            }
        }

        /// <summary>
        /// オブジェクトマトリクスに列を追加します
        /// </summary>
        /// <param name="gridInfo">グリッド情報</param>
        private void AddColumnToObjectMatrix(PowerToolServiceHelper.GridInfo gridInfo)
        {
            try
            {
                var slide = helper.GetCurrentSlide();
                if (slide == null)
                {
                    throw new InvalidOperationException("アクティブなスライドが見つかりません。");
                }

                var createdShapes = new List<PowerPoint.Shape>();

                // 列間隔を計算（既存列間の平均を使用）
                float columnSpacing = 5f; // デフォルト
                var allShapes = gridInfo.ShapeGrid.SelectMany(row => row).ToList();

                if (gridInfo.Columns > 1)
                {
                    var spacings = new List<float>();
                    foreach (var row in gridInfo.ShapeGrid)
                    {
                        for (int col = 0; col < row.Count - 1; col++)
                        {
                            var currentRight = row[col].Left + row[col].Width;
                            var nextLeft = row[col + 1].Left;
                            spacings.Add(nextLeft - currentRight);
                        }
                    }
                    if (spacings.Count > 0)
                    {
                        columnSpacing = spacings.Average();
                    }
                }

                // 新しい列のX位置を計算
                var rightmostX = allShapes.Max(s => s.Left + s.Width);
                var newColumnX = rightmostX + columnSpacing;

                // 各行に新しいオブジェクトを作成
                for (int row = 0; row < gridInfo.Rows; row++)
                {
                    var currentRow = gridInfo.ShapeGrid[row];
                    if (currentRow.Count == 0) continue;

                    // 参考図形（同じ行の最右端）から書式をコピー
                    var referenceShape = currentRow[currentRow.Count - 1];

                    // 適切なセルサイズを計算（全体の列幅統一のため最右端列の平均を使用）
                    var rightmostColumnWidths = gridInfo.ShapeGrid
                        .Where(r => r.Count > 0)
                        .Select(r => r[r.Count - 1].Width);
                    var avgWidth = rightmostColumnWidths.DefaultIfEmpty(120f).Average();
                    // 同じ行の既存オブジェクトと完全に同じ高さ・位置に統一
                    var rowTop = currentRow.Min(s => s.Top);
                    var rowHeight = currentRow.Max(s => s.Height);

                    // 新しいテキストボックスを作成
                    var newTextBox = slide.Shapes.AddTextbox(
                        MsoTextOrientation.msoTextOrientationHorizontal,
                        newColumnX,
                        rowTop,
                        avgWidth,
                        rowHeight
                    );

                    newTextBox.TextFrame.AutoSize = PowerPoint.PpAutoSize.ppAutoSizeNone;

                    // 書式を参考図形からコピー
                    CopyObjectShapeFormat(referenceShape.Shape, newTextBox);

                    // デフォルトテキストを設定
                    if (newTextBox.HasTextFrame == MsoTriState.msoTrue)
                    {
                        newTextBox.TextFrame.TextRange.Text = ""; // 空のテキスト
                    }

                    createdShapes.Add(newTextBox);
                    logger.Debug($"Created new cell at row {row + 1} for new column");
                }

                // 作成した図形を選択状態にする
                if (createdShapes.Count > 0)
                {
                    SelectCreatedShapes(createdShapes);
                }

                logger.Info($"Added column to object matrix (created {createdShapes.Count} new cells)");
            }
            catch (Exception ex)
            {
                logger.Error(ex, "Failed to add column to object matrix");
                throw;
            }
        }

        /// <summary>
        /// 図形の書式をコピーします（新機能専用）
        /// </summary>
        /// <param name="sourceShape">コピー元図形</param>
        /// <param name="targetShape">コピー先図形</param>
        private void CopyObjectShapeFormat(PowerPoint.Shape sourceShape, PowerPoint.Shape targetShape)
        {
            try
            {
                // 塗りつぶしをコピー
                if (sourceShape.Fill.Type != MsoFillType.msoFillMixed)
                {
                    targetShape.Fill.ForeColor.RGB = sourceShape.Fill.ForeColor.RGB;
                    targetShape.Fill.Transparency = sourceShape.Fill.Transparency;
                }

                // 線をコピー
                if (sourceShape.Line.Visible == MsoTriState.msoTrue)
                {
                    targetShape.Line.Visible = MsoTriState.msoTrue;
                    targetShape.Line.ForeColor.RGB = sourceShape.Line.ForeColor.RGB;
                    targetShape.Line.Weight = sourceShape.Line.Weight;
                    targetShape.Line.DashStyle = sourceShape.Line.DashStyle;
                }

                // テキスト書式をコピー
                if (sourceShape.HasTextFrame == MsoTriState.msoTrue &&
                    targetShape.HasTextFrame == MsoTriState.msoTrue)
                {
                    var sourceTextFrame = sourceShape.TextFrame;
                    var targetTextFrame = targetShape.TextFrame;

                    // マージンをコピー
                    targetTextFrame.MarginTop = sourceTextFrame.MarginTop;
                    targetTextFrame.MarginBottom = sourceTextFrame.MarginBottom;
                    targetTextFrame.MarginLeft = sourceTextFrame.MarginLeft;
                    targetTextFrame.MarginRight = sourceTextFrame.MarginRight;

                    // フォント設定をコピー（テキストがある場合のみ）
                    if (sourceShape.TextFrame.HasText == MsoTriState.msoTrue)
                    {
                        var sourceTextRange = sourceTextFrame.TextRange;
                        var targetTextRange = targetTextFrame.TextRange;

                        targetTextRange.Font.Name = sourceTextRange.Font.Name;
                        targetTextRange.Font.Size = sourceTextRange.Font.Size;
                        targetTextRange.Font.Bold = sourceTextRange.Font.Bold;
                        targetTextRange.Font.Color.RGB = sourceTextRange.Font.Color.RGB;
                        targetTextRange.ParagraphFormat.Alignment = sourceTextRange.ParagraphFormat.Alignment;
                    }
                }

                logger.Debug($"Shape format copied from {sourceShape.Name} to {targetShape.Name}");
            }
            catch (Exception ex)
            {
                logger.Debug(ex, "Failed to copy shape format (non-critical)");
                // 書式コピー失敗は致命的ではないので継続
            }
        }

        /// <summary>
        /// 複数図形を選択状態にします（新機能専用）
        /// </summary>
        /// <param name="shapes">選択する図形のリスト</param>
        private void SelectCreatedShapes(List<PowerPoint.Shape> shapes)
        {
            try
            {
                if (shapes == null || shapes.Count == 0) return;

                var application = applicationProvider.GetCurrentApplication();
                var slide = helper.GetCurrentSlide();
                if (slide == null) return;

                // 最初の図形を選択
                shapes[0].Select(MsoTriState.msoFalse);

                // 残りの図形を追加選択
                for (int i = 1; i < shapes.Count; i++)
                {
                    shapes[i].Select(MsoTriState.msoTrue);
                }

                logger.Debug($"Selected {shapes.Count} shapes");
            }
            catch (Exception ex)
            {
                logger.Warn(ex, "Failed to select shapes");
            }
        }

        /// <summary>
        /// マトリクス列幅統一（表・オブジェクト対応）
        /// 表の場合は全体幅を保持して各列を等幅に、オブジェクトマトリクスの場合は各列のテキストボックスを等幅に設定
        /// </summary>
        public void EqualizeColumnWidths()
        {
            if (!Globals.ThisAddIn.CheckFeatureAccess("EqualizeColumnWidths")) return;

            logger.Info("EqualizeColumnWidths operation started");

            var selectedShapes = helper.GetSelectedShapeInfos();
            if (!helper.ValidateSelection(selectedShapes, 1, 0, "列幅統一")) return;

            ComHelper.ExecuteWithComCleanup(() =>
            {
                bool processed = false;

                // 表の処理
                var tableShapes = selectedShapes.Where(s => s.Shape.HasTable == MsoTriState.msoTrue).ToList();
                if (tableShapes.Count > 0)
                {
                    foreach (var tableShape in tableShapes)
                    {
                        EqualizeTableColumnWidths(tableShape.Shape.Table);
                        logger.Debug($"Equalized column widths in table {tableShape.Name}");
                        processed = true;
                    }
                }

                // オブジェクトマトリクスの処理
                if (!processed)
                {
                    var textBoxShapes = selectedShapes.Where(s =>
                        (s.HasTextFrame || s.Shape.Type == MsoShapeType.msoTextBox) &&
                        s.Shape.Type != MsoShapeType.msoLine).ToList();

                    if (textBoxShapes.Count >= 2)
                    {
                        var gridInfo = helper.DetectGridLayout(textBoxShapes);
                        if (gridInfo != null)
                        {
                            // ★修正点: グリッド再配置対応版を使用
                            EqualizeObjectMatrixColumnWidths(gridInfo);
                            logger.Debug($"Equalized column widths in object matrix ({gridInfo.Rows}x{gridInfo.Columns}) with proper alignment");
                            processed = true;
                        }
                    }
                }

                if (!processed)
                {
                    ErrorHandler.ExecuteSafely(() =>
                    {
                        throw new InvalidOperationException(
                            "列幅を統一できる対象が見つかりません。\n" +
                            "表またはグリッド配置されたテキストボックスを選択してください。");
                    }, "列幅統一");
                    return;
                }

                // ★修正点: 区切り線を削除（再配置ではなく）
                DeleteRowSeparatorsIfExists();

                logger.Info("EqualizeColumnWidths completed successfully");

            }, selectedShapes.Select(s => s.Shape).ToArray());

            logger.Info("EqualizeColumnWidths completed");
        }

        /// <summary>
        /// マトリクス行高統一（表・オブジェクト対応）
        /// 一番高いセル/オブジェクトにあわせて全行を統一
        /// </summary>
        public void EqualizeRowHeights()
        {
            if (!Globals.ThisAddIn.CheckFeatureAccess("EqualizeRowHeights")) return;

            logger.Info("EqualizeRowHeights operation started");

            var selectedShapes = helper.GetSelectedShapeInfos();
            if (!helper.ValidateSelection(selectedShapes, 1, 0, "行高統一")) return;

            ComHelper.ExecuteWithComCleanup(() =>
            {
                bool processed = false;

                // 表の処理
                var tableShapes = selectedShapes.Where(s => s.Shape.HasTable == MsoTriState.msoTrue).ToList();
                if (tableShapes.Count > 0)
                {
                    foreach (var tableShape in tableShapes)
                    {
                        EqualizeTableRowHeights(tableShape.Shape.Table);
                        logger.Debug($"Equalized row heights in table {tableShape.Name}");
                        processed = true;
                    }
                }

                // オブジェクトマトリクスの処理
                if (!processed)
                {
                    var textBoxShapes = selectedShapes.Where(s =>
                        (s.HasTextFrame || s.Shape.Type == MsoShapeType.msoTextBox) &&
                        s.Shape.Type != MsoShapeType.msoLine).ToList();

                    if (textBoxShapes.Count >= 2)
                    {
                        var gridInfo = helper.DetectGridLayout(textBoxShapes);
                        if (gridInfo != null)
                        {
                            // ★修正点: グリッド再配置対応版を使用
                            EqualizeObjectMatrixRowHeights(gridInfo);
                            logger.Debug($"Equalized row heights in object matrix ({gridInfo.Rows}x{gridInfo.Columns}) with proper alignment");
                            processed = true;
                        }
                    }
                }

                if (!processed)
                {
                    ErrorHandler.ExecuteSafely(() =>
                    {
                        throw new InvalidOperationException(
                            "行高を統一できる対象が見つかりません。\n" +
                            "表またはグリッド配置されたテキストボックスを選択してください。");
                    }, "行高統一");
                    return;
                }

                // ★修正点: 区切り線を削除（再配置ではなく）
                DeleteRowSeparatorsIfExists();

                logger.Info("EqualizeRowHeights completed successfully");

            }, selectedShapes.Select(s => s.Shape).ToArray());

            logger.Info("EqualizeRowHeights completed");
        }

        /// <summary>
        /// 表の列幅を等幅に設定します（全体幅保持）
        /// </summary>
        /// <param name="table">対象の表</param>
        private void EqualizeTableColumnWidths(PowerPoint.Table table)
        {
            try
            {
                // 現在の表全体の幅を取得
                float totalWidth = 0f;
                for (int col = 1; col <= table.Columns.Count; col++)
                {
                    totalWidth += table.Columns[col].Width;
                }

                // 等分した列幅を計算
                var equalColumnWidth = totalWidth / table.Columns.Count;

                // 各列を等幅に設定
                for (int col = 1; col <= table.Columns.Count; col++)
                {
                    table.Columns[col].Width = equalColumnWidth;
                    logger.Debug($"Set column {col} width to {equalColumnWidth:F1}pt");
                }

                // 表の高さも最適化（既存機能流用）
                OptimizeTableRowHeightsPractical(table);

                logger.Info($"Equalized table column widths: {table.Columns.Count} columns at {equalColumnWidth:F1}pt each");
            }
            catch (Exception ex)
            {
                logger.Error(ex, "Failed to equalize table column widths");
                throw;
            }
        }

        /// <summary>
        /// 表の行高を等高に設定します（最大高さに統一）
        /// </summary>
        /// <param name="table">対象の表</param>
        private void EqualizeTableRowHeights(PowerPoint.Table table)
        {
            try
            {
                // 全行の中で最大の高さを取得
                float maxRowHeight = 0f;
                for (int row = 1; row <= table.Rows.Count; row++)
                {
                    maxRowHeight = Math.Max(maxRowHeight, table.Rows[row].Height);
                }

                // 最小高さを保証
                maxRowHeight = Math.Max(maxRowHeight, 25f);

                // 各行を最大高さに設定
                for (int row = 1; row <= table.Rows.Count; row++)
                {
                    table.Rows[row].Height = maxRowHeight;
                    logger.Debug($"Set row {row} height to {maxRowHeight:F1}pt");
                }

                logger.Info($"Equalized table row heights: {table.Rows.Count} rows at {maxRowHeight:F1}pt each");
            }
            catch (Exception ex)
            {
                logger.Error(ex, "Failed to equalize table row heights");
                throw;
            }
        }

        /// <summary>
        /// オブジェクトマトリクスの列幅を等幅に設定します
        /// </summary>
        /// <param name="gridInfo">グリッド情報</param>
        private void EqualizeObjectMatrixColumnWidths(PowerToolServiceHelper.GridInfo gridInfo)
        {
            try
            {
                // 全オブジェクトの平均幅を計算
                var allShapes = gridInfo.ShapeGrid.SelectMany(row => row).ToList();
                var avgWidth = allShapes.Select(s => s.Width).Average();

                // 最小幅を保証
                avgWidth = Math.Max(avgWidth, 30f);

                logger.Debug($"Target column width: {avgWidth:F1}pt");

                // 各列の統一幅配列を作成
                var columnWidths = new float[gridInfo.Columns];
                for (int col = 0; col < gridInfo.Columns; col++)
                {
                    columnWidths[col] = avgWidth;
                }

                // 現在の行高配列を作成（変更なし）
                var rowHeights = new float[gridInfo.Rows];
                for (int row = 0; row < gridInfo.Rows; row++)
                {
                    if (gridInfo.ShapeGrid[row].Count > 0)
                    {
                        rowHeights[row] = gridInfo.ShapeGrid[row][0].Height;
                    }
                }

                // 現在のグリッド間隔を保持
                var currentSpacing = CalculateCurrentGridSpacing(gridInfo);

                // 既存機能を活用してグリッド位置を再計算・適用
                ApplyOptimizedDimensionsToGrid(gridInfo, columnWidths, rowHeights, currentSpacing);

                logger.Info($"Equalized object matrix column widths: {gridInfo.Columns} columns at {avgWidth:F1}pt each with proper grid alignment");
            }
            catch (Exception ex)
            {
                logger.Error(ex, "Failed to equalize object matrix column widths");
                throw;
            }
        }

        /// <summary>
        /// オブジェクトマトリクスの行高を等高に設定します
        /// </summary>
        /// <param name="gridInfo">グリッド情報</param>
        private void EqualizeObjectMatrixRowHeights(PowerToolServiceHelper.GridInfo gridInfo)
        {
            try
            {
                // 全オブジェクトの中で最大の高さを取得
                var allShapes = gridInfo.ShapeGrid.SelectMany(row => row).ToList();
                var maxHeight = allShapes.Select(s => s.Height).Max();

                // 最小高さを保証
                maxHeight = Math.Max(maxHeight, 25f);

                logger.Debug($"Target row height: {maxHeight:F1}pt");

                // 現在の列幅配列を作成（変更なし）
                var columnWidths = new float[gridInfo.Columns];
                for (int col = 0; col < gridInfo.Columns; col++)
                {
                    float maxColumnWidth = 0f;
                    for (int row = 0; row < gridInfo.Rows; row++)
                    {
                        if (col < gridInfo.ShapeGrid[row].Count)
                        {
                            maxColumnWidth = Math.Max(maxColumnWidth, gridInfo.ShapeGrid[row][col].Width);
                        }
                    }
                    columnWidths[col] = maxColumnWidth;
                }

                // 各行の統一高配列を作成
                var rowHeights = new float[gridInfo.Rows];
                for (int row = 0; row < gridInfo.Rows; row++)
                {
                    rowHeights[row] = maxHeight;
                }

                // 現在のグリッド間隔を保持
                var currentSpacing = CalculateCurrentGridSpacing(gridInfo);

                // 既存機能を活用してグリッド位置を再計算・適用
                ApplyOptimizedDimensionsToGrid(gridInfo, columnWidths, rowHeights, currentSpacing);

                logger.Info($"Equalized object matrix row heights: {gridInfo.Rows} rows at {maxHeight:F1}pt each with proper grid alignment");
            }
            catch (Exception ex)
            {
                logger.Error(ex, "Failed to equalize object matrix row heights");
                throw;
            }
        }

        /// <summary>
        /// 区切り線が存在する場合は削除します
        /// </summary>
        private void DeleteRowSeparatorsIfExists()
        {
            try
            {
                var slide = helper.GetCurrentSlide();
                if (slide == null) return;

                var separatorShapes = new List<PowerPoint.Shape>();
                var deletedCount = 0;

                // 区切り線を検索して削除
                for (int i = slide.Shapes.Count; i >= 1; i--) // 逆順でループ（削除時のインデックス変更対策）
                {
                    try
                    {
                        var shape = slide.Shapes[i];

                        // "RowSeparator_" で始まる名前の線を対象とする
                        if (shape.Name.StartsWith("RowSeparator_") && shape.Type == MsoShapeType.msoLine)
                        {
                            shape.Delete();
                            deletedCount++;
                            logger.Debug($"Deleted separator line: {shape.Name}");
                        }
                    }
                    catch (COMException comEx)
                    {
                        // 図形が既に削除されている場合など、COMExceptionを適切にハンドリング
                        logger.Warn(comEx, $"COM error while processing shape at index {i}");
                    }
                    catch (Exception ex)
                    {
                        logger.Warn(ex, $"Failed to process shape at index {i}");
                    }
                }

                if (deletedCount > 0)
                {
                    logger.Info($"Deleted {deletedCount} row separator lines");
                }
                else
                {
                    logger.Debug("No row separator lines found to delete");
                }
            }
            catch (Exception ex)
            {
                logger.Error(ex, "Failed to delete row separators");
                // エラーは記録するが処理は継続（区切り線の削除失敗で全体が止まることを防ぐ）
            }
        }

        /// <summary>
        /// Matrix Tuner
        /// 矩形オブジェクトのマトリックス配置を高度に調整
        /// </summary>
        public void MatrixTuner()
        {
            if (!Globals.ThisAddIn.CheckFeatureAccess("MatrixTuner")) return;

            logger.Info("MatrixTuner operation started");

            var selectedShapes = helper.GetSelectedShapeInfos();
            if (!helper.ValidateSelection(selectedShapes, 2, 225, "Matrix Tuner")) return;

            // 矩形系オブジェクトのみを対象とする
            var rectangularShapes = selectedShapes.Where(s => IsRectangularShape(s)).ToList();

            if (rectangularShapes.Count < 2)
            {
                ErrorHandler.ExecuteSafely(() =>
                {
                    throw new InvalidOperationException("矩形系オブジェクトを2つ以上選択してください。\n" +
                        "（テキストボックス、長方形、画像など）");
                }, "Matrix Tuner");
                return;
            }

            // 回転チェック（±1度を超える回転は除外）
            var rotatedShapes = rectangularShapes.Where(s => Math.Abs(s.Shape.Rotation) > 1.0f).ToList();
            if (rotatedShapes.Count > 0)
            {
                rectangularShapes = rectangularShapes.Except(rotatedShapes).ToList();
                logger.Info($"Excluded {rotatedShapes.Count} rotated shapes (rotation > ±1°)");

                if (rectangularShapes.Count < 2)
                {
                    ErrorHandler.ExecuteSafely(() =>
                    {
                        throw new InvalidOperationException("回転していない矩形オブジェクトが2つ以上必要です。");
                    }, "Matrix Tuner");
                    return;
                }
            }

            // グリッド配置を検出
            var gridInfo = helper.DetectGridLayout(rectangularShapes);
            if (gridInfo == null)
            {
                ErrorHandler.ExecuteSafely(() =>
                {
                    throw new InvalidOperationException("選択した図形がグリッド配置になっていません。\n" +
                        "Matrix Tunerを使用するには、行・列が整列している必要があります。");
                }, "Matrix Tuner");
                return;
            }

            // 15×15制限チェック
            if (gridInfo.Rows > 15 || gridInfo.Columns > 15)
            {
                ErrorHandler.ExecuteSafely(() =>
                {
                    throw new InvalidOperationException($"15行×15列を超えています。\n" +
                        $"Matrix Tuner は最大15×15まで対応しています。\n" +
                        $"現在: {gridInfo.Rows}行×{gridInfo.Columns}列");
                }, "Matrix Tuner");
                return;
            }

            logger.Info($"Grid detected: {gridInfo.Rows}x{gridInfo.Columns}");

            // SmartArt、表、グラフは除外
            var invalidShapes = rectangularShapes.Where(s =>
                s.Shape.Type == MsoShapeType.msoSmartArt ||
                s.Shape.Type == MsoShapeType.msoChart ||
                s.Shape.HasTable == MsoTriState.msoTrue).ToList();

            if (invalidShapes.Count > 0)
            {
                rectangularShapes = rectangularShapes.Except(invalidShapes).ToList();
                logger.Info($"Excluded {invalidShapes.Count} SmartArt/Chart/Table shapes");
            }

            // 【新規追加】区切り線を削除
            try
            {
                logger.Info("Checking for row separator lines...");
                DeleteRowSeparatorsIfExists();
                logger.Info("Row separator deletion completed");
            }
            catch (Exception ex)
            {
                logger.Warn(ex, "Failed to delete row separators, continuing with Matrix Tuner");
                // 区切り線削除に失敗してもMatrix Tunerは継続
            }

            // Matrix Tunerダイアログを表示
            MatrixTunerDialog dialog = null;
            try
            {
                // AutoFitを無効化
                DisableAutoFitForShapes(rectangularShapes);

                dialog = new MatrixTunerDialog(rectangularShapes, gridInfo);
                var dialogResult = dialog.ShowDialog();

                if (dialogResult == DialogResult.OK)
                {
                    logger.Info("Matrix Tuner adjustments applied successfully");
                }
                else
                {
                    logger.Info("Matrix Tuner cancelled by user");
                }
            }
            catch (Exception ex)
            {
                logger.Error(ex, "Failed to apply Matrix Tuner");
                ErrorHandler.ExecuteSafely(() => throw ex, "Matrix Tuner");
            }
            finally
            {
                dialog?.Dispose();
            }
        }

        /// <summary>
        /// 矩形系の図形かどうかを判定
        /// </summary>
        private bool IsRectangularShape(ShapeInfo shape)
        {
            try
            {
                var shapeType = shape.Shape.Type;

                // テキストボックス
                if (shapeType == MsoShapeType.msoTextBox || shape.HasTextFrame)
                    return true;

                // 画像
                if (shapeType == MsoShapeType.msoPicture ||
                    shapeType == MsoShapeType.msoLinkedPicture)
                    return true;

                // 矩形系のAutoShape
                if (shapeType == MsoShapeType.msoAutoShape)
                {
                    var autoShapeType = shape.Shape.AutoShapeType;
                    return autoShapeType == MsoAutoShapeType.msoShapeRectangle ||
                           autoShapeType == MsoAutoShapeType.msoShapeRoundedRectangle ||
                           autoShapeType == MsoAutoShapeType.msoShapeSnip1Rectangle ||
                           autoShapeType == MsoAutoShapeType.msoShapeSnip2SameRectangle ||
                           autoShapeType == MsoAutoShapeType.msoShapeSnipRoundRectangle ||
                           autoShapeType == MsoAutoShapeType.msoShapeRound1Rectangle ||
                           autoShapeType == MsoAutoShapeType.msoShapeRound2SameRectangle;
                    // Office 2016では一部の定数が未定義のため除外
                    // msoShapeSnip2DiagonalRectangle
                    // msoShapeRound2DiagonalRectangle
                }

                // プレースホルダー（本文、タイトル）
                if (shapeType == MsoShapeType.msoPlaceholder)
                {
                    var placeholderType = shape.Shape.PlaceholderFormat.Type;
                    return placeholderType == PowerPoint.PpPlaceholderType.ppPlaceholderBody ||
                           placeholderType == PowerPoint.PpPlaceholderType.ppPlaceholderTitle ||
                           placeholderType == PowerPoint.PpPlaceholderType.ppPlaceholderCenterTitle ||
                           placeholderType == PowerPoint.PpPlaceholderType.ppPlaceholderSubtitle;
                }

                return false;
            }
            catch (Exception ex)
            {
                logger.Debug(ex, $"Error checking if shape is rectangular: {shape.Name}");
                return false;
            }
        }

        /// <summary>
        /// AutoFitを無効化
        /// </summary>
        private void DisableAutoFitForShapes(List<ShapeInfo> shapes)
        {
            foreach (var shape in shapes)
            {
                try
                {
                    if (shape.HasTextFrame && shape.Shape.TextFrame.HasText == MsoTriState.msoTrue)
                    {
                        shape.Shape.TextFrame.AutoSize = PowerPoint.PpAutoSize.ppAutoSizeNone;
                        shape.Shape.TextFrame.WordWrap = MsoTriState.msoTrue;
                        logger.Debug($"Disabled AutoFit for shape: {shape.Name}");
                    }
                }
                catch (Exception ex)
                {
                    logger.Warn(ex, $"Failed to disable AutoFit for shape: {shape.Name}");
                }
            }
        }

        #region 共通ヘルパーメソッド


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