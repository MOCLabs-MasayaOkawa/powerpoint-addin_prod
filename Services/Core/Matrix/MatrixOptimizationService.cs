using ImageMagick;
using Microsoft.Office.Core;
using NLog;
using PowerPointEfficiencyAddin.Models;
using PowerPointEfficiencyAddin.Services.Core.PowerTool;
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


namespace PowerPointEfficiencyAddin.Services.Core.Matrix
{
    /// <summary>
    /// マトリクス最適化・均等化機能を提供するサービス
    /// </summary>
    public class MatrixOptimizationService
    {
        private static readonly Logger logger = LogManager.GetCurrentClassLogger();
        private readonly IApplicationProvider applicationProvider;
        private readonly PowerToolServiceHelper helper;
        private readonly MatrixStructureService structureService;

        public MatrixOptimizationService(IApplicationProvider applicationProvider)
        {
            this.applicationProvider = applicationProvider ?? throw new ArgumentNullException(nameof(applicationProvider));
            logger.Debug("MatrixOptimizationService initialized");
            helper = new PowerToolServiceHelper(applicationProvider);
            structureService = new MatrixStructureService(applicationProvider);
        }

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
                    structureService.RealignRowSeparatorsIfExists();
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
                        structureService.RealignRowSeparatorsIfExists();
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
                    structureService.RealignRowSeparatorsIfExists();
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
                        structureService.RealignRowSeparatorsIfExists();
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
                structureService.DeleteRowSeparatorsIfExists();

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
                structureService.DeleteRowSeparatorsIfExists();

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
    }
}
