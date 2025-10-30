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
    /// マトリクス配置・整列・余白・MatrixTuner機能を提供するサービス
    /// </summary>
    public class MatrixAlignmentService
    {
        private static readonly Logger logger = LogManager.GetCurrentClassLogger();
        private readonly IApplicationProvider applicationProvider;
        private readonly PowerToolServiceHelper helper;

        public MatrixAlignmentService(IApplicationProvider applicationProvider)
        {
            this.applicationProvider = applicationProvider ?? throw new ArgumentNullException(nameof(applicationProvider));
            logger.Debug("MatrixAlignmentService initialized");
            helper = new PowerToolServiceHelper(applicationProvider);
        }

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

                var cellInfos = GetCellInformations(gridInfo, isTable, matrixShapes);
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

        private (List<ShapeInfo> matrixShapes, List<ShapeInfo> targetShapes) SeparateMatrixAndTargetShapes(List<ShapeInfo> selectedShapes)
        {
            var matrixShapes = new List<ShapeInfo>();
            var targetShapes = new List<ShapeInfo>();

            foreach (var si in selectedShapes)
            {
                PowerPoint.Shape shp = si.Shape;
                MsoShapeType type = shp.Type;

                if (shp.HasTable == MsoTriState.msoTrue
                    || type == MsoShapeType.msoTextBox
                    || PowerToolServiceHelper.IsRectLikeAutoShape(shp)
                    || PowerToolServiceHelper.IsMatrixPlaceholder(shp))
                {
                    matrixShapes.Add(si);
                    continue;
                }

                if (type == MsoShapeType.msoLine)
                {
                    continue;
                }

                targetShapes.Add(si);
            }

            return (matrixShapes, targetShapes);
        }

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
                                Row = row + 1,
                                Column = col + 1
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

            return shapeCellMappings.Where(kvp => kvp.Value.Count > 0).ToDictionary(kvp => kvp.Key, kvp => kvp.Value);
        }

        private void AlignShapesToCellCenter(List<ShapeInfo> shapesInCell, CellInfo cellInfo)
        {
            foreach (var shapeInfo in shapesInCell)
            {
                try
                {
                    var cellCenterX = cellInfo.CenterX;
                    var cellCenterY = cellInfo.CenterY;

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

        public void SetCellMargins()
        {
            if (!Globals.ThisAddIn.CheckFeatureAccess("SetCellMargins")) return;

            logger.Info("SetCellMargins operation started");

            var selectedShapes = helper.GetSelectedShapeInfos();
            if (!helper.ValidateSelection(selectedShapes, 1, 0, "セルマージン設定")) return;

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
                                var cellCount = ProcessTableMargins(shapeInfo.Shape.Table, top, bottom, left, right, shapeInfo.Name);
                                processedCells += cellCount;
                                processedShapes++;
                                logger.Debug($"Processed table {shapeInfo.Name}: {cellCount} cells updated");
                            }
                            else if (shapeInfo.HasTextFrame)
                            {
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

        private int ProcessTableMargins(PowerPoint.Table table, float top, float bottom, float left, float right, string shapeName)
        {
            int processedCells = 0;

            try
            {
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

        private void ProcessTextBoxMargins(PowerPoint.TextFrame textFrame, float top, float bottom, float left, float right, string shapeName)
        {
            try
            {
                var topPt = top * 28.35f;
                var bottomPt = bottom * 28.35f;
                var leftPt = left * 28.35f;
                var rightPt = right * 28.35f;

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

        public void MatrixTuner()
        {
            if (!Globals.ThisAddIn.CheckFeatureAccess("MatrixTuner")) return;

            logger.Info("MatrixTuner operation started");

            var selectedShapes = helper.GetSelectedShapeInfos();
            if (!helper.ValidateSelection(selectedShapes, 2, 225, "Matrix Tuner")) return;

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

            var invalidShapes = rectangularShapes.Where(s =>
                s.Shape.Type == MsoShapeType.msoSmartArt ||
                s.Shape.Type == MsoShapeType.msoChart ||
                s.Shape.HasTable == MsoTriState.msoTrue).ToList();

            if (invalidShapes.Count > 0)
            {
                rectangularShapes = rectangularShapes.Except(invalidShapes).ToList();
                logger.Info($"Excluded {invalidShapes.Count} SmartArt/Chart/Table shapes");
            }

            try
            {
                logger.Info("Checking for row separator lines...");
                DeleteRowSeparatorsIfExists();
                logger.Info("Row separator deletion completed");
            }
            catch (Exception ex)
            {
                logger.Warn(ex, "Failed to delete row separators, continuing with Matrix Tuner");
            }

            MatrixTunerDialog dialog = null;
            try
            {
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

        private bool IsRectangularShape(ShapeInfo shape)
        {
            try
            {
                var shapeType = shape.Shape.Type;

                if (shapeType == MsoShapeType.msoTextBox || shape.HasTextFrame)
                    return true;

                if (shapeType == MsoShapeType.msoPicture ||
                    shapeType == MsoShapeType.msoLinkedPicture)
                    return true;

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
                }

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

        private void DeleteRowSeparatorsIfExists()
        {
            // TODO: This method will be implemented in MatrixStructureService
            // Temporary stub to avoid compilation errors
            logger.Debug("DeleteRowSeparatorsIfExists called (stub implementation)");
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
