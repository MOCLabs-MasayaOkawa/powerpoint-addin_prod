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
    /// マトリクスExcel連携機能を提供するサービス
    /// </summary>
    public class MatrixExcelService
    {
        private static readonly Logger logger = LogManager.GetCurrentClassLogger();
        private readonly IApplicationProvider applicationProvider;
        private readonly PowerToolServiceHelper helper;

        public MatrixExcelService(IApplicationProvider applicationProvider)
        {
            this.applicationProvider = applicationProvider ?? throw new ArgumentNullException(nameof(applicationProvider));
            logger.Debug("MatrixExcelService initialized");
            helper = new PowerToolServiceHelper(applicationProvider);
        }

        public void ExcelToPptx()
        {
            if (!Globals.ThisAddIn.CheckFeatureAccess("ExcelToPptx")) return;

            logger.Info("ExcelToPptx operation started (paste to existing matrix)");

            var selectedShapes = helper.GetSelectedShapeInfos();
            if (!helper.ValidateSelection(selectedShapes, 1, 0, "Excel雋ｼ繧贋ｻ倥￠")) return;

            ComHelper.ExecuteWithComCleanup(() =>
            {
                try
                {
                    var excelData = GetExcelDataFromClipboard();
                    if (excelData == null || excelData.Length == 0)
                    {
                        ErrorHandler.ExecuteSafely(() =>
                        {
                            throw new InvalidOperationException("Excel縺ｮ繝・・繧ｿ繧偵さ繝斐・縺励※縺九ｉ螳溯｡後＠縺ｦ縺上□縺輔＞縲・);
                        }, "Excel雋ｼ繧贋ｻ倥￠");
                        return;
                    }

                    int excelRows = excelData.Length;
                    int excelCols = excelData[0].Length;

                    logger.Info($"Excel data structure: {excelRows} rows x {excelCols} columns");

                    bool processed = false;

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
                                "Excel 繝・・繧ｿ繧定ｲｼ繧贋ｻ倥￠縺ｧ縺阪ｋ蟇ｾ雎｡縺瑚ｦ九▽縺九ｊ縺ｾ縺帙ｓ縲・n" +
                                "陦ｨ縺ｾ縺溘・繧ｰ繝ｪ繝・ラ驟咲ｽｮ縺輔ｌ縺溘ユ繧ｭ繧ｹ繝医・繝・け繧ｹ繧帝∈謚槭＠縺ｦ縺上□縺輔＞縲・);
                        }, "Excel雋ｼ繧贋ｻ倥￠");
                        return;
                    }

                    logger.Info("ExcelToPptx completed successfully");
                }
                catch (Exception ex)
                {
                    logger.Error(ex, "Failed to paste Excel data");
                    ErrorHandler.ExecuteSafely(() =>
                    {
                        throw new InvalidOperationException("Excel繝・・繧ｿ縺ｮ雋ｼ繧贋ｻ倥￠縺ｫ螟ｱ謨励＠縺ｾ縺励◆縲・);
                    }, "Excel雋ｼ繧贋ｻ倥￠");
                }
            }, selectedShapes.Select(s => s.Shape).ToArray());

            logger.Info("ExcelToPptx completed");
        }

        private bool PasteExcelDataToTable(PowerPoint.Table table, string[][] excelData, int excelRows, int excelCols)
        {
            try
            {
                if (table.Rows.Count < excelRows || table.Columns.Count < excelCols)
                {
                    logger.Warn($"Table size ({table.Rows.Count}x{table.Columns.Count}) is smaller than Excel data ({excelRows}x{excelCols})");
                    ErrorHandler.ExecuteSafely(() =>
                    {
                        throw new InvalidOperationException(
                            $"陦ｨ縺ｮ繧ｵ繧､繧ｺ・・table.Rows.Count}陦古養table.Columns.Count}蛻暦ｼ峨′\n" +
                            $"Excel繝・・繧ｿ・・excelRows}陦古養excelCols}蛻暦ｼ峨ｈ繧雁ｰ上＆縺・〒縺吶・);
                    }, "Excel雋ｼ繧贋ｻ倥￠");
                    return false;
                }

                int pastedCells = 0;
                for (int row = 0; row < excelRows; row++)
                {
                    for (int col = 0; col < excelCols; col++)
                    {
                        try
                        {
                            var cell = table.Cell(row + 1, col + 1);
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

        private bool PasteExcelDataToObjectMatrix(PowerToolServiceHelper.GridInfo gridInfo, string[][] excelData, int excelRows, int excelCols)
        {
            try
            {
                if (gridInfo.Rows < excelRows || gridInfo.Columns < excelCols)
                {
                    logger.Warn($"Object matrix size ({gridInfo.Rows}x{gridInfo.Columns}) is smaller than Excel data ({excelRows}x{excelCols})");
                    ErrorHandler.ExecuteSafely(() =>
                    {
                        throw new InvalidOperationException(
                            $"繧ｪ繝悶ず繧ｧ繧ｯ繝医・繝医Μ繧ｯ繧ｹ・・gridInfo.Rows}陦古養gridInfo.Columns}蛻暦ｼ峨′\n" +
                            $"Excel繝・・繧ｿ・・excelRows}陦古養excelCols}蛻暦ｼ峨ｈ繧雁ｰ上＆縺・〒縺吶・);
                    }, "Excel雋ｼ繧贋ｻ倥￠");
                    return false;
                }

                int pastedShapes = 0;
                for (int row = 0; row < excelRows; row++)
                {
                    for (int col = 0; col < excelCols; col++)
                    {
                        try
                        {
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

        private string[][] ParseExcelClipboardData(string clipboardData)
        {
            if (string.IsNullOrWhiteSpace(clipboardData))
                return null;

            try
            {
                var lines = clipboardData.Split(new[] { "\r\n", "\n", "\r" }, StringSplitOptions.None);
                var result = new List<string[]>();

                foreach (var line in lines)
                {
                    if (string.IsNullOrEmpty(line)) continue;
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

        private string TruncateText(string text)
        {
            if (string.IsNullOrEmpty(text)) return "(empty)";
            const int maxLength = 20;
            if (text.Length <= maxLength) return text;
            return text.Substring(0, maxLength) + "...";
        }
    }
}
