using Microsoft.Office.Core;
using NLog;
using PowerPointEfficiencyAddin.Models;
using PowerPointEfficiencyAddin.Services.Core.PowerTool;
using PowerPointEfficiencyAddin.Services.Infrastructure.MultiInstance;
using PowerPointEfficiencyAddin.UI;
using PowerPointEfficiencyAddin.Utils;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PowerPointEfficiencyAddin.Services.Core.Table
{
    /// <summary>
    /// テーブルとテキストボックスの相互変換機能を提供するサービスクラス
    /// Phase 3-1a: PowerToolService.csから分離
    /// </summary>
    public class TableConversionService
    {
        private static readonly Logger logger = LogManager.GetCurrentClassLogger();
        private readonly IApplicationProvider applicationProvider;
        private readonly PowerToolServiceHelper helper;

        // DI対応コンストラクタ
        public TableConversionService(IApplicationProvider applicationProvider, PowerToolServiceHelper helper)
        {
            this.applicationProvider = applicationProvider ?? throw new ArgumentNullException(nameof(applicationProvider));
            this.helper = helper ?? throw new ArgumentNullException(nameof(helper));
            logger.Debug("TableConversionService initialized with DI");
        }

        #region Public Methods

        /// <summary>
        /// テーブルをテキストボックスに変換（20番機能）
        /// 選択されたテーブルをセル単位で独立したテキストボックスに変換
        /// </summary>
        public void ConvertTableToTextBoxes()
        {
            if (!Globals.ThisAddIn.CheckFeatureAccess("ConvertTableToTextBoxes")) return;

            logger.Info("ConvertTableToTextBoxes operation started");
            var app = Globals.ThisAddIn.Application;
            if (app == null)
            {
                logger.Error("Failed to get PowerPoint Application");
                return;
            }

            var tableShapes = new List<ShapeInfo>();
            try
            {
                var selection = app.ActiveWindow.Selection;
                if (selection.Type != PowerPoint.PpSelectionType.ppSelectionShapes)
                {
                    MessageBox.Show("図形が選択されていません。", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                int shapeIndex = 0;
                foreach (PowerPoint.Shape shape in selection.ShapeRange)
                {
                    if (shape.HasTable == MsoTriState.msoTrue)
                    {
                        tableShapes.Add(new ShapeInfo(shape, shapeIndex++));
                    }
                }

                if (tableShapes.Count == 0)
                {
                    MessageBox.Show("選択された図形の中にテーブルがありません。", "情報", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                foreach (var tableShape in tableShapes)
                {
                    ConvertSingleTableToTextBoxes(tableShape.Shape.Parent as PowerPoint.Slide, tableShape.Shape);
                }

                MessageBox.Show($"{tableShapes.Count}個のテーブルをテキストボックスに変換しました。", "完了", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                logger.Error(ex, "ConvertTableToTextBoxes failed");
                ErrorHandler.ExecuteSafely(() => { throw new InvalidOperationException("変換処理に失敗しました。"); }, "テーブル変換");
            }
            finally
            {
                ComHelper.ReleaseComObjects(tableShapes);
            }

            logger.Info($"ConvertTableToTextBoxes completed for {tableShapes.Count} tables");
        }

        /// <summary>
        /// テキストボックスをテーブルに変換（21番機能）
        /// 選択されたテキストボックスのグリッド配置を検出し、テーブルに変換
        /// </summary>
        public void ConvertTextBoxesToTable()
        {
            if (!Globals.ThisAddIn.CheckFeatureAccess("ConvertTextBoxesToTable")) return;

            logger.Info("ConvertTextBoxesToTable operation started");
            var app = Globals.ThisAddIn.Application;
            if (app == null)
            {
                logger.Error("Failed to get PowerPoint Application");
                return;
            }

            PowerPoint.Slide slide = null;
            try
            {
                var selection = app.ActiveWindow.Selection;
                if (selection.Type != PowerPoint.PpSelectionType.ppSelectionShapes)
                {
                    MessageBox.Show("図形が選択されていません。", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                slide = app.ActiveWindow.View.Slide as PowerPoint.Slide;

                var selectedShapeInfos = new List<ShapeInfo>();
                int shapeIndex = 0;
                foreach (PowerPoint.Shape shape in selection.ShapeRange)
                {
                    if (!helper.IsTableShape(shape) && shape.HasTextFrame == MsoTriState.msoTrue)
                    {
                        selectedShapeInfos.Add(new ShapeInfo(shape, shapeIndex++));
                    }
                }

                if (selectedShapeInfos.Count == 0)
                {
                    MessageBox.Show("選択された図形の中に変換可能なテキストボックスがありません。", "情報", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                var gridInfo = helper.DetectGridLayout(selectedShapeInfos);
                if (gridInfo == null)
                {
                    MessageBox.Show("選択された図形からグリッド配置を検出できませんでした。\n図形が整列していることを確認してください。", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                logger.Debug($"Grid detected: {gridInfo.Rows}x{gridInfo.Columns}");

                var spacing = 0f;
                if (gridInfo.Rows > 1 || gridInfo.Columns > 1)
                {
                    var avgRowSpacing = gridInfo.Rows > 1 ? gridInfo.ShapeGrid.Take(gridInfo.Rows - 1).Average(row => row[0].Top + row[0].Height) : 0;
                    var avgColSpacing = gridInfo.Columns > 1 ? gridInfo.ShapeGrid[0].Take(gridInfo.Columns - 1).Average(shape => shape.Left + shape.Width) : 0;
                    spacing = (avgRowSpacing + avgColSpacing) / 2f;

                    UnifyRowHeightsInGrid(gridInfo.ShapeGrid);

                    var topLeft = gridInfo.ShapeGrid[0][0];
                    AdjustPositionsAfterHeightUnification(gridInfo.ShapeGrid, topLeft.Left, topLeft.Top, spacing);
                }

                ConvertGridToTable(slide, gridInfo);

                MessageBox.Show($"選択された図形をテーブル({gridInfo.Rows}行×{gridInfo.Columns}列)に変換しました。", "完了", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                logger.Error(ex, "ConvertTextBoxesToTable failed");
                ErrorHandler.ExecuteSafely(() => { throw new InvalidOperationException("変換処理に失敗しました。"); }, "テーブル変換");
            }
            finally
            {
                ComHelper.ReleaseComObject(slide);
            }

            logger.Info("ConvertTextBoxesToTable completed");
        }

        #endregion

        #region Private Methods

        /// <summary>
        /// 単一のテーブルをテキストボックスに変換
        /// </summary>
        private void ConvertSingleTableToTextBoxes(PowerPoint.Slide slide, PowerPoint.Shape tableShape)
        {
            var table = tableShape.Table;
            var tableRows = table.Rows.Count;
            var tableColumns = table.Columns.Count;
            var baseLeft = tableShape.Left;
            var baseTop = tableShape.Top;
            var spacing = 0f;

            logger.Debug($"Converting table: {tableRows}x{tableColumns}");

            var createdTextBoxes = new List<PowerPoint.Shape>();

            try
            {
                for (int r = 1; r <= tableRows; r++)
                {
                    for (int c = 1; c <= tableColumns; c++)
                    {
                        PowerPoint.Cell cell = null;
                        PowerPoint.Shape textBox = null;

                        try
                        {
                            cell = table.Cell(r, c);
                            var cellWidth = table.Columns[c].Width;
                            var cellHeight = table.Rows[r].Height;

                            var currentLeft = baseLeft;
                            for (int i = 1; i < c; i++)
                                currentLeft += table.Columns[i].Width;

                            var currentTop = baseTop;
                            for (int i = 1; i < r; i++)
                                currentTop += table.Rows[i].Height;

                            textBox = slide.Shapes.AddTextbox(
                                MsoTextOrientation.msoTextOrientationHorizontal,
                                currentLeft, currentTop, cellWidth, cellHeight
                            );

                            textBox.TextFrame.WordWrap = MsoTriState.msoTrue;
                            textBox.TextFrame.AutoSize = PowerPoint.PpAutoSize.ppAutoSizeNone;

                            var cellText = cell.Shape.TextFrame.TextRange.Text;
                            textBox.TextFrame.TextRange.Text = cellText;

                            if (cellText.Trim().Length > 0)
                            {
                                CopyTextFormat(cell.Shape.TextFrame.TextRange, textBox.TextFrame.TextRange);
                            }

                            CopyCellFormat(cell.Shape, textBox);

                            createdTextBoxes.Add(textBox);
                        }
                        finally
                        {
                            ComHelper.ReleaseComObject(cell);
                        }
                    }
                }

                if (createdTextBoxes.Count > 0)
                {
                    var shapeGrid = ConvertToShapeGrid(createdTextBoxes, tableRows, tableColumns);
                    UnifyRowHeightsInGrid(shapeGrid);
                    AdjustPositionsAfterHeightUnification(shapeGrid, baseLeft, baseTop, spacing);
                }

                tableShape.Delete();
                logger.Debug($"Table deleted and {createdTextBoxes.Count} text boxes created");
            }
            catch (Exception ex)
            {
                logger.Error(ex, $"Failed to convert table at cell");
                foreach (var textBox in createdTextBoxes)
                {
                    try { textBox.Delete(); } catch { }
                }
                throw;
            }
            finally
            {
                ComHelper.ReleaseComObject(table);
            }
        }

        /// <summary>
        /// グリッド配置のテキストボックスをテーブルに変換
        /// </summary>
        private void ConvertGridToTable(PowerPoint.Slide slide, PowerToolServiceHelper.GridInfo gridInfo)
        {
            PowerPoint.Shape tableShape = null;
            PowerPoint.Table table = null;

            try
            {
                var topLeft = gridInfo.TopLeft;
                var bottomRightShape = gridInfo.ShapeGrid[gridInfo.Rows - 1][gridInfo.Columns - 1];
                var tableWidth = bottomRightShape.Left + bottomRightShape.Width - topLeft.Left;
                var tableHeight = bottomRightShape.Top + bottomRightShape.Height - topLeft.Top;

                tableShape = slide.Shapes.AddTable(gridInfo.Rows, gridInfo.Columns, topLeft.Left, topLeft.Top, tableWidth, tableHeight);
                table = tableShape.Table;

                for (int r = 0; r < gridInfo.Rows; r++)
                {
                    for (int c = 0; c < gridInfo.Columns; c++)
                    {
                        PowerPoint.Cell cell = null;
                        try
                        {
                            cell = table.Cell(r + 1, c + 1);
                            var sourceShape = gridInfo.ShapeGrid[r][c].Shape;

                            cell.Shape.TextFrame.TextRange.Text = sourceShape.TextFrame.TextRange.Text;

                            if (sourceShape.TextFrame.TextRange.Text.Trim().Length > 0)
                            {
                                CopyTextFormat(sourceShape.TextFrame.TextRange, cell.Shape.TextFrame.TextRange);
                            }

                            CopyCellFormat(sourceShape, cell.Shape);
                        }
                        finally
                        {
                            ComHelper.ReleaseComObject(cell);
                        }
                    }
                }

                foreach (var row in gridInfo.ShapeGrid)
                {
                    foreach (var shapeInfo in row)
                    {
                        try { shapeInfo.Shape.Delete(); } catch { }
                    }
                }

                logger.Debug($"Grid converted to table: {gridInfo.Rows}x{gridInfo.Columns}");
            }
            finally
            {
                ComHelper.ReleaseComObject(table);
                ComHelper.ReleaseComObject(tableShape);
            }
        }

        /// <summary>
        /// テキスト書式をコピー
        /// </summary>
        private void CopyTextFormat(PowerPoint.TextRange source, PowerPoint.TextRange target)
        {
            try
            {
                if (source.Font != null && target.Font != null)
                {
                    target.Font.Name = source.Font.Name;
                    target.Font.Size = source.Font.Size;
                    target.Font.Bold = source.Font.Bold;
                    target.Font.Italic = source.Font.Italic;
                    target.Font.Underline = source.Font.Underline;
                    target.Font.Color.RGB = source.Font.Color.RGB;
                }

                target.ParagraphFormat.Alignment = source.ParagraphFormat.Alignment;
            }
            catch (Exception ex)
            {
                logger.Warn(ex, "Failed to copy text format");
            }
        }

        /// <summary>
        /// セル書式をコピー
        /// </summary>
        private void CopyCellFormat(PowerPoint.Shape sourceCell, PowerPoint.Shape targetTextBox)
        {
            try
            {
                if (sourceCell.Fill.Visible == MsoTriState.msoTrue && sourceCell.Fill.ForeColor != null)
                {
                    targetTextBox.Fill.Visible = MsoTriState.msoTrue;
                    targetTextBox.Fill.ForeColor.RGB = sourceCell.Fill.ForeColor.RGB;
                    targetTextBox.Fill.Transparency = sourceCell.Fill.Transparency;
                }

                if (sourceCell.Line.Visible == MsoTriState.msoTrue)
                {
                    targetTextBox.Line.Visible = MsoTriState.msoTrue;
                    targetTextBox.Line.ForeColor.RGB = sourceCell.Line.ForeColor.RGB;
                    targetTextBox.Line.Weight = sourceCell.Line.Weight;
                }

                if (sourceCell.TextFrame != null && targetTextBox.TextFrame != null)
                {
                    targetTextBox.TextFrame.MarginLeft = sourceCell.TextFrame.MarginLeft;
                    targetTextBox.TextFrame.MarginRight = sourceCell.TextFrame.MarginRight;
                    targetTextBox.TextFrame.MarginTop = sourceCell.TextFrame.MarginTop;
                    targetTextBox.TextFrame.MarginBottom = sourceCell.TextFrame.MarginBottom;
                }
            }
            catch (Exception ex)
            {
                logger.Warn(ex, "Failed to copy cell format");
            }
        }

        /// <summary>
        /// グリッド内の各行の高さを統一
        /// </summary>
        private void UnifyRowHeightsInGrid(List<List<ShapeInfo>> shapeGrid)
        {
            foreach (var row in shapeGrid)
            {
                if (row.Count == 0) continue;

                float maxHeight = row.Max(s => s.Height);

                foreach (var shapeInfo in row)
                {
                    if (Math.Abs(shapeInfo.Height - maxHeight) > 0.1f)
                    {
                        try
                        {
                            shapeInfo.Shape.Height = maxHeight;
                            shapeInfo.Height = maxHeight;
                            logger.Debug($"Unified row height: {shapeInfo.Shape.Name} -> {maxHeight}pt");
                        }
                        catch (Exception ex)
                        {
                            logger.Warn(ex, $"Failed to unify height for shape {shapeInfo.Shape.Name}");
                        }
                    }
                }
            }
        }

        /// <summary>
        /// 図形リストを2次元グリッドに変換
        /// </summary>
        private List<List<ShapeInfo>> ConvertToShapeGrid(List<PowerPoint.Shape> shapes, int rows, int columns)
        {
            var shapeInfos = shapes.Select((s, index) => new ShapeInfo(s, index)).OrderBy(s => s.Top).ThenBy(s => s.Left).ToList();

            var grid = new List<List<ShapeInfo>>();
            for (int r = 0; r < rows; r++)
            {
                var row = new List<ShapeInfo>();
                for (int c = 0; c < columns; c++)
                {
                    var index = r * columns + c;
                    if (index < shapeInfos.Count)
                        row.Add(shapeInfos[index]);
                }
                grid.Add(row);
            }

            return grid;
        }

        /// <summary>
        /// 行の高さ統一後、グリッド内の図形位置を再調整
        /// </summary>
        private void AdjustPositionsAfterHeightUnification(List<List<ShapeInfo>> shapeGrid, float baseLeft, float baseTop, float spacing)
        {
            float currentTop = baseTop;

            for (int r = 0; r < shapeGrid.Count; r++)
            {
                float currentLeft = baseLeft;
                float rowHeight = shapeGrid[r].Count > 0 ? shapeGrid[r][0].Height : 0;

                for (int c = 0; c < shapeGrid[r].Count; c++)
                {
                    var shapeInfo = shapeGrid[r][c];
                    shapeInfo.Shape.Left = currentLeft;
                    shapeInfo.Shape.Top = currentTop;

                    currentLeft += shapeInfo.Width + spacing;
                }

                currentTop += rowHeight + spacing;
            }
        }

        #endregion
    }
}
