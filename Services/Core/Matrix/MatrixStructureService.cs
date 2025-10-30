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
    /// マトリクス構造変更（行列追加・区切り線・ヘッダー）を提供するサービス
    /// </summary>
    public class MatrixStructureService
    {
        private static readonly Logger logger = LogManager.GetCurrentClassLogger();
        private readonly IApplicationProvider applicationProvider;
        private readonly PowerToolServiceHelper helper;

        public MatrixStructureService(IApplicationProvider applicationProvider)
        {
            this.applicationProvider = applicationProvider ?? throw new ArgumentNullException(nameof(applicationProvider));
            logger.Debug("MatrixStructureService initialized");
            helper = new PowerToolServiceHelper(applicationProvider);
        }

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
        public void RealignRowSeparatorsIfExists()
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
        public void DeleteRowSeparatorsIfExists()
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
    }
}
