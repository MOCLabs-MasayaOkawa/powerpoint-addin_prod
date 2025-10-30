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
    /// �}�g���N�X�\���ύX�i�s��ǉ��E��؂���E�w�b�_�[�j��񋟂���T�[�r�X
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
            if (!helper.ValidateSelection(selectedShapes, 2, 0, "�s�ԋ�؂��")) return;

            // �e�L�X�g�{�b�N�X�Q�݂̂�ΏۂƂ���
            var textBoxShapes = selectedShapes.Where(s =>
                s.HasTextFrame || s.Shape.Type == MsoShapeType.msoTextBox).ToList();

            if (textBoxShapes.Count < 2)
            {
                ErrorHandler.ExecuteSafely(() =>
                {
                    throw new InvalidOperationException("�e�L�X�g�{�b�N�X��2�ȏ�I�����Ă��������B");
                }, "�s�ԋ�؂��");
                return;
            }

            // �O���b�h�z�u�����o�i�������\�b�h���p�j
            var gridInfo = helper.DetectGridLayout(textBoxShapes);
            if (gridInfo == null)
            {
                ErrorHandler.ExecuteSafely(() =>
                {
                    throw new InvalidOperationException("�I�������}�`���O���b�h�z�u�ɂȂ��Ă��܂���B\n" +
                        "�s�ԋ�؂����ǉ�����ɂ́A�s�E�񂪐��񂵂Ă���K�v������܂��B");
                }, "�s�ԋ�؂��");
                return;
            }

            if (gridInfo.Rows < 2)
            {
                ErrorHandler.ExecuteSafely(() =>
                {
                    throw new InvalidOperationException("��؂����ǉ�����ɂ́A2�s�ȏ�̃}�g���N�X���K�v�ł��B");
                }, "�s�ԋ�؂��");
                return;
            }

            // ���ݒ�_�C�A���O��\��
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

                // COM�Ǘ����ŋ�؂�����쐬�i�����p�^�[�����p�j
                ComHelper.ExecuteWithComCleanup(() =>
                {
                    var slide = helper.GetCurrentSlide(); // �������\�b�h���p
                    if (slide == null)
                    {
                        ErrorHandler.ExecuteSafely(() =>
                        {
                            throw new InvalidOperationException("�A�N�e�B�u�ȃX���C�h��������܂���B");
                        }, "�s�ԋ�؂��");
                        return;
                    }

                    var createdLines = CreateRowSeparatorLines(slide, gridInfo, lineStyle, lineWeight, lineColor);

                    logger.Info($"Created {createdLines.Count} row separator lines");

                    // �쐬��������I����Ԃɂ���i�����p�^�[�����p�j
                    if (createdLines.Count > 0)
                    {
                        helper.SelectShapes(createdLines); // �������\�b�h���p
                    }

                }, selectedShapes.Select(s => s.Shape).ToArray());

                logger.Info("AddMatrixRowSeparators completed");
            }
            catch (Exception ex)
            {
                logger.Error(ex, "Failed to add matrix row separators");
                ErrorHandler.ExecuteSafely(() => throw ex, "�s�ԋ�؂��");
            }
            finally
            {
                dialog?.Dispose();
            }
        }

        /// <summary>
        /// �s�ԋ�؂�����쐬���܂�
        /// </summary>
        /// <param name="slide">�ΏۃX���C�h</param>
        /// <param name="gridInfo">�O���b�h���</param>
        /// <param name="lineStyle">���̎��</param>
        /// <param name="lineWeight">���̑���</param>
        /// <param name="lineColor">���̐F</param>
        /// <returns>�쐬���ꂽ���}�`�̃��X�g</returns>
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

                // ��؂���̈ʒu�����v�Z
                var separatorPositions = CalculateRowSeparatorPositions(gridInfo);

                // �e�s�Ԃɋ�؂�����쐬�i�Ō�̍s�͏����j
                for (int i = 0; i < separatorPositions.Count; i++)
                {
                    var position = separatorPositions[i];

                    try
                    {
                        // ���������쐬
                        var line = slide.Shapes.AddLine(
                            position.StartX, position.Y,
                            position.EndX, position.Y
                        );

                        // ���̃v���p�e�B��ݒ�
                        line.Line.Weight = lineWeight;
                        line.Line.DashStyle = lineStyle;
                        line.Line.ForeColor.RGB = ColorTranslator.ToOle(lineColor);
                        line.Line.Visible = MsoTriState.msoTrue;

                        // ���̖��O��ݒ�i�Ǘ����₷�����邽�߁j
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
        /// �s�ԋ�؂���̈ʒu���v�Z���܂�
        /// </summary>
        /// <param name="gridInfo">�O���b�h���</param>
        /// <returns>��؂���ʒu�̃��X�g</returns>
        private List<SeparatorLinePosition> CalculateRowSeparatorPositions(PowerToolServiceHelper.GridInfo gridInfo)
        {
            var positions = new List<SeparatorLinePosition>();

            try
            {
                // �O���b�h�̍��[�ƉE�[���v�Z
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

                // �e�s�Ԃ̒����ʒu���v�Z
                for (int row = 0; row < gridInfo.Rows - 1; row++) // �Ō�̍s�͏���
                {
                    var currentRow = gridInfo.ShapeGrid[row];
                    var nextRow = gridInfo.ShapeGrid[row + 1];

                    if (currentRow.Count == 0 || nextRow.Count == 0) continue;

                    // ���ݍs�̉��[���v�Z
                    var currentRowBottom = currentRow.Max(s => s.Top + s.Height);

                    // ���s�̏�[���v�Z
                    var nextRowTop = nextRow.Min(s => s.Top);

                    // �����ʒu���v�Z
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
        /// ��؂���̈ʒu����\���N���X
        /// </summary>
        private class SeparatorLinePosition
        {
            public float StartX { get; set; }
            public float EndX { get; set; }
            public float Y { get; set; }
        }

        /// <summary>
        /// �X���C�h����s��؂�������o���܂�
        /// </summary>
        /// <param name="slide">�ΏۃX���C�h</param>
        /// <returns>��؂���}�`�̃��X�g</returns>
        private List<PowerPoint.Shape> FindRowSeparators(PowerPoint.Slide slide)
        {
            var separators = new List<PowerPoint.Shape>();

            try
            {
                for (int i = 1; i <= slide.Shapes.Count; i++)
                {
                    var shape = slide.Shapes[i];

                    // ���O�p�^�[���ŋ�؂�������ʁi���������Ɠ����j
                    if (shape.Name.StartsWith("RowSeparator_") && shape.Type == MsoShapeType.msoLine)
                    {
                        separators.Add(shape);
                    }
                }

                // ���O���Ń\�[�g�iRowSeparator_1, RowSeparator_2, ...�̏��j
                separators.Sort((a, b) => string.Compare(a.Name, b.Name, StringComparison.OrdinalIgnoreCase));
            }
            catch (Exception ex)
            {
                logger.Error(ex, "Failed to find row separators");
            }

            return separators;
        }

        /// <summary>
        /// ��؂�����폜���čč쐬���܂��i�ʒu��������Ȃ��ꍇ�j
        /// </summary>
        /// <param name="existingSeparators">������؂��</param>
        /// <param name="newPositions">�V�����ʒu���</param>
        /// <param name="slide">�ΏۃX���C�h</param>
        private void RecreateRowSeparators(List<PowerPoint.Shape> existingSeparators, List<SeparatorLinePosition> newPositions, PowerPoint.Slide slide)
        {
            try
            {
                // ������؂���̏����ݒ��ۑ��i�ŏ��̐�����j
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

                // ������؂�����폜
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

                // �V�����ʒu�ɋ�؂�����č쐬�i�������W�b�N���p�j
                var newSeparators = CreateRowSeparatorLines(slide, null, lineStyle, lineWeight, lineColor, newPositions);
                logger.Info($"Recreated {newSeparators.Count} row separators with preserved formatting");
            }
            catch (Exception ex)
            {
                logger.Error(ex, "Failed to recreate row separators");
            }
        }

        /// <summary>
        /// �s�ԋ�؂�����쐬���܂��i�ʒu�w��Łj
        /// </summary>
        private List<PowerPoint.Shape> CreateRowSeparatorLines(
            PowerPoint.Slide slide,
            PowerToolServiceHelper.GridInfo gridInfo, // null�̏ꍇ��positions�𒼐ڎg�p
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
        /// ��؂�������݂���ꍇ�̂ݍĔz�u���܂��i�ŏ��C���Łj
        /// </summary>
        public void RealignRowSeparatorsIfExists()
        {
            try
            {
                var slide = helper.GetCurrentSlide();
                if (slide == null) return;

                var separatorShapes = new List<PowerPoint.Shape>();

                // ��؂��������
                for (int i = 1; i <= slide.Shapes.Count; i++)
                {
                    var shape = slide.Shapes[i];
                    if (shape.Name.StartsWith("RowSeparator_") && shape.Type == MsoShapeType.msoLine)
                    {
                        separatorShapes.Add(shape);
                    }
                }

                // ��؂�����Ȃ��ꍇ�͉������Ȃ�
                if (separatorShapes.Count == 0) return;

                logger.Info($"Found {separatorShapes.Count} separators, realigning...");

                // ���ݑI������Ă���}�`���擾�i�œK���Ώہj
                var selectedShapes = helper.GetSelectedShapeInfos();
                var matrixShapes = selectedShapes.Where(s =>
                    s.Shape.HasTable == MsoTriState.msoTrue ||
                    s.HasTextFrame ||
                    s.Shape.Type == MsoShapeType.msoTextBox).ToList();

                if (matrixShapes.Count == 0) return;

                // �O���b�h�����擾
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

                // �V�����ʒu���v�Z���Ĉړ�
                var newPositions = CalculateRowSeparatorPositions(gridInfo);

                // ��؂���̐��ƌv�Z���ꂽ�ʒu��������Ȃ��ꍇ�̏���
                if (newPositions.Count != separatorShapes.Count)
                {
                    logger.Warn($"Separator count mismatch: found {separatorShapes.Count}, calculated {newPositions.Count}");

                    // �Â���؂����S�폜
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

                    // �V������؂�����쐬�i�������\�b�h���p�j
                    CreateRowSeparatorLines(slide, null, MsoLineDashStyle.msoLineSolid, 1.0f,
                        Color.Black, newPositions);
                    return;
                }

                // ������؂����V�ʒu�Ɉړ�
                for (int i = 0; i < Math.Min(separatorShapes.Count, newPositions.Count); i++)
                {
                    var separator = separatorShapes[i];
                    var newPos = newPositions[i];

                    try
                    {
                        separator.Left = newPos.StartX;
                        separator.Top = newPos.Y;
                        separator.Width = newPos.EndX - newPos.StartX;
                        separator.Height = 0; // ������
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
                // �G���[�͋L�^���邪�����͌p��
            }
        }

        /// <summary>
        /// �}�`���Z���ʒu�ɐ���
        /// �I�����ꂽ�}�g���N�X�i�\/�e�L�X�g�{�b�N�X�O���b�h�j�̃Z�������ɐ}�`�𐮗񂷂�
        public void AddHeaderRowToMatrix()
        {
            if (!Globals.ThisAddIn.CheckFeatureAccess("AddHeaderRowToMatrix")) return;

            logger.Info("AddHeaderRowToMatrix operation started");

            var selectedShapes = helper.GetSelectedShapeInfos();
            if (!helper.ValidateSelection(selectedShapes, 1, 0, "���o���s�t�^")) return;

            ComHelper.ExecuteWithComCleanup(() =>
            {
                var slide = helper.GetCurrentSlide();
                if (slide == null) return;

                // �e�[�u�����O���b�h���𔻒�
                var tableShapes = selectedShapes.Where(s => s.Shape.HasTable == MsoTriState.msoTrue).ToList();

                if (tableShapes.Count > 0)
                {
                    // �\�̏���
                    ProcessTable(tableShapes);
                }
                else
                {
                    // �O���b�h�̏���
                    ProcessGrid(selectedShapes, slide);
                }

            }, selectedShapes.Select(s => s.Shape).ToArray());

            logger.Info("AddHeaderRowToMatrix completed");
        }

        /// <summary>
        /// �\�̏���
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
                    cell.Shape.TextFrame.TextRange.Text = $"���o��{col}";
                }

                logger.Info($"Added header row to table");
            }
        }

        /// <summary>
        /// �O���b�h�̏���
        /// </summary>
        private void ProcessGrid(List<ShapeInfo> selectedShapes, PowerPoint.Slide slide)
        {
            // 1. �I��}�`�̍ŏ�[���擾
            var topMost = selectedShapes.Min(s => s.Top);
            var leftMost = selectedShapes.Min(s => s.Left);
            var rightMost = selectedShapes.Max(s => s.Left + s.Width);

            // 2. 1�s�ڂ̐}�`�����i�ŏ�i�̐}�`�����j
            var tolerance = 5f; // 5pt���e�덷
            var topRowShapes = selectedShapes.Where(s => Math.Abs(s.Top - topMost) <= tolerance)
                                           .OrderBy(s => s.Left)
                                           .ToList();

            logger.Info($"Found {topRowShapes.Count} shapes in top row");

            // 3. ���o���s���쐬
            var headerShapes = new List<PowerPoint.Shape>();
            foreach (var topShape in topRowShapes)
            {
                var headerBox = slide.Shapes.AddTextbox(
                    MsoTextOrientation.msoTextOrientationHorizontal,
                    topShape.Left,
                    topMost - 50f, // ���ʒu
                    topShape.Width,
                    30f // ������
                );

                headerBox.TextFrame.TextRange.Text = $"���o��{topRowShapes.IndexOf(topShape) + 1}";
                headerBox.Fill.Visible = MsoTriState.msoFalse;
                headerBox.Line.Visible = MsoTriState.msoFalse;
                headerBox.TextFrame.TextRange.Font.Color.RGB = 0;

                headerShapes.Add(headerBox);
            }



            // 4. ���o���s�̍������œK��
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


            // 5. 0.8mm�Ԋu�Ŕz�u
            const float SPACING_PT = 10.0f * 2.835f; // 0.8mm
            var headerTop = topMost - maxHeaderHeight - SPACING_PT;


            // ���o���s��z�u
            foreach (var header in headerShapes)
            {
                header.Top = headerTop;
            }

            // 5. ���o���̍ŏI�z�u������O��iheader.Top = headerTop �ς݁j
            float headerBottom = headerTop + maxHeaderHeight;

            // 6. ���Ԃɋ�؂��
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
        /// �Z���}�[�W���ݒ�
        /// �I�����ꂽ�\�̃Z���܂��̓e�L�X�g�{�b�N�X�̃}�[�W���𓝈�ݒ肷��
        /// </summary>
        public void AddMatrixRow()
        {
            if (!Globals.ThisAddIn.CheckFeatureAccess("AddMatrixRow")) return;

            logger.Info("AddMatrixRow operation started (Phase 1)");

            var selectedShapes = helper.GetSelectedShapeInfos();
            if (!helper.ValidateSelection(selectedShapes, 1, 0, "�s�ǉ�")) return;

            ComHelper.ExecuteWithComCleanup(() =>
            {
                bool processed = false;

                // �\�̏���
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

                // �I�u�W�F�N�g�}�g���N�X�̏���
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
                            "�s��ǉ��ł���Ώۂ�������܂���B\n" +
                            "�\�܂��̓O���b�h�z�u���ꂽ�e�L�X�g�{�b�N�X��I�����Ă��������B");
                    }, "�s�ǉ�");
                    return;
                }

                // ��؂��������΍Ĕz�u�i�����@�\���p�j
                RealignRowSeparatorsIfExists();

                logger.Info("AddMatrixRow completed successfully");

            }, selectedShapes.Select(s => s.Shape).ToArray());

            logger.Info("AddMatrixRow completed");
        }

        /// <summary>
        /// �}�g���N�X��ǉ��iPhase 1: ��{�@�\�j
        /// �\�S�́E�I�u�W�F�N�g�}�g���N�X�S�̑I�����ɍŉE�[�ɗ��ǉ�
        /// </summary>
        public void AddMatrixColumn()
        {
            if (!Globals.ThisAddIn.CheckFeatureAccess("AddMatrixColumn")) return;

            logger.Info("AddMatrixColumn operation started (Phase 1)");

            var selectedShapes = helper.GetSelectedShapeInfos();
            if (!helper.ValidateSelection(selectedShapes, 1, 0, "��ǉ�")) return;

            ComHelper.ExecuteWithComCleanup(() =>
            {
                bool processed = false;

                // �\�̏���
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

                // �I�u�W�F�N�g�}�g���N�X�̏���
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
                            "���ǉ��ł���Ώۂ�������܂���B\n" +
                            "�\�܂��̓O���b�h�z�u���ꂽ�e�L�X�g�{�b�N�X��I�����Ă��������B");
                    }, "��ǉ�");
                    return;
                }

                // ��؂��������΍Ĕz�u�i�����@�\���p�j
                RealignRowSeparatorsIfExists();

                logger.Info("AddMatrixColumn completed successfully");

            }, selectedShapes.Select(s => s.Shape).ToArray());

            logger.Info("AddMatrixColumn completed");
        }

        /// <summary>
        /// �\�ɍs��ǉ����܂�
        /// </summary>
        /// <param name="table">�Ώۂ̕\</param>
        private void AddRowToTable(PowerPoint.Table table)
        {
            try
            {
                // �ŉ��[�ɍs��ǉ�
                var newRow = table.Rows.Add();

                // �V�����s�̍�����K�؂ɐݒ�i�אڍs�̍������Q�l�j
                if (table.Rows.Count > 1)
                {
                    var referenceRowHeight = table.Rows[table.Rows.Count - 1].Height;
                    newRow.Height = referenceRowHeight;
                }
                else
                {
                    newRow.Height = 35f; // �f�t�H���g����
                }

                // �V�����s�̃Z���Ɋ�{������K�p
                for (int col = 1; col <= table.Columns.Count; col++)
                {
                    try
                    {
                        var newCell = table.Cell(table.Rows.Count, col);

                        // ��̍s�̃Z���������R�s�[�i�\�ȏꍇ�j
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
        /// �\�ɗ��ǉ����܂�
        /// </summary>
        /// <param name="table">�Ώۂ̕\</param>
        private void AddColumnToTable(PowerPoint.Table table)
        {
            try
            {
                // �ŉE�[�ɗ��ǉ�
                var newColumn = table.Columns.Add();

                // �V������̕���K�؂ɐݒ�i�אڗ�̕����Q�l�j
                if (table.Columns.Count > 1)
                {
                    var referenceColumnWidth = table.Columns[table.Columns.Count - 1].Width;
                    newColumn.Width = referenceColumnWidth;
                }
                else
                {
                    newColumn.Width = 120f; // �f�t�H���g��
                }

                // �V������̃Z���Ɋ�{������K�p
                for (int row = 1; row <= table.Rows.Count; row++)
                {
                    try
                    {
                        var newCell = table.Cell(row, table.Columns.Count);

                        // ���̗�̃Z���������R�s�[�i�\�ȏꍇ�j
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
        /// �\�Z���������R�s�[���܂��i�V�@�\��p�j
        /// </summary>
        /// <param name="sourceCell">�R�s�[���Z��</param>
        /// <param name="targetCell">�R�s�[��Z��</param>
        private void CopyTableCellFormatNew(PowerPoint.Cell sourceCell, PowerPoint.Cell targetCell)
        {
            try
            {
                var sourceShape = sourceCell.Shape;
                var targetShape = targetCell.Shape;

                // �w�i�F���R�s�[
                if (sourceShape.Fill.Type != MsoFillType.msoFillMixed)
                {
                    targetShape.Fill.ForeColor.RGB = sourceShape.Fill.ForeColor.RGB;
                    targetShape.Fill.Transparency = sourceShape.Fill.Transparency;
                }

                // �����R�s�[
                if (sourceShape.Line.Visible == MsoTriState.msoTrue)
                {
                    targetShape.Line.Visible = MsoTriState.msoTrue;
                    targetShape.Line.ForeColor.RGB = sourceShape.Line.ForeColor.RGB;
                    targetShape.Line.Weight = sourceShape.Line.Weight;
                    targetShape.Line.DashStyle = sourceShape.Line.DashStyle;
                }

                // �e�L�X�g�������R�s�[�i��{�̂݁j
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
                // �����R�s�[���s�͒v���I�ł͂Ȃ��̂Ōp��
            }
        }

        /// <summary>
        /// �I�u�W�F�N�g�}�g���N�X�ɍs��ǉ����܂�
        /// </summary>
        /// <param name="gridInfo">�O���b�h���</param>
        private void AddRowToObjectMatrix(PowerToolServiceHelper.GridInfo gridInfo)
        {
            try
            {
                var slide = helper.GetCurrentSlide();
                if (slide == null)
                {
                    throw new InvalidOperationException("�A�N�e�B�u�ȃX���C�h��������܂���B");
                }

                var createdShapes = new List<PowerPoint.Shape>();

                // �ŉ��i�̐}�`���Q�l�ɐV�����s���쐬
                var lastRowShapes = gridInfo.ShapeGrid[gridInfo.Rows - 1];
                var referenceY = lastRowShapes.Max(s => s.Top + s.Height);

                // �s�Ԋu���v�Z�i�����s�Ԃ̕��ς��g�p�j
                float rowSpacing = 5f; // �f�t�H���g
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

                // �e��ɐV�����I�u�W�F�N�g���쐬
                for (int col = 0; col < gridInfo.Columns; col++)
                {
                    // �Q�l�}�`�i������̍ŉ��i�j
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

                    // �K�؂ȃZ���T�C�Y���v�Z�i������̕��ς��g�p�j
                    var avgWidth = gridInfo.ShapeGrid.Where(row => col < row.Count)
                        .Select(row => row[col].Width).Average();
                    var avgHeight = gridInfo.ShapeGrid[gridInfo.Rows - 1]
                        .Where(s => s != null).Select(s => s.Height).DefaultIfEmpty(35f).Average();

                    // �V�����e�L�X�g�{�b�N�X���쐬
                    var newTextBox = slide.Shapes.AddTextbox(
                        MsoTextOrientation.msoTextOrientationHorizontal,
                        referenceShape.Left,
                        newRowY,
                        avgWidth,
                        avgHeight
                    );

                    // �������Q�l�}�`����R�s�[
                    CopyObjectShapeFormat(referenceShape.Shape, newTextBox);

                    // �f�t�H���g�e�L�X�g��ݒ�
                    if (newTextBox.HasTextFrame == MsoTriState.msoTrue)
                    {
                        newTextBox.TextFrame.TextRange.Text = ""; // ��̃e�L�X�g
                    }

                    createdShapes.Add(newTextBox);
                    logger.Debug($"Created new cell at column {col + 1} for new row");
                }

                // �쐬�����}�`��I����Ԃɂ���
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
        /// �I�u�W�F�N�g�}�g���N�X�ɗ��ǉ����܂�
        /// </summary>
        /// <param name="gridInfo">�O���b�h���</param>
        private void AddColumnToObjectMatrix(PowerToolServiceHelper.GridInfo gridInfo)
        {
            try
            {
                var slide = helper.GetCurrentSlide();
                if (slide == null)
                {
                    throw new InvalidOperationException("�A�N�e�B�u�ȃX���C�h��������܂���B");
                }

                var createdShapes = new List<PowerPoint.Shape>();

                // ��Ԋu���v�Z�i������Ԃ̕��ς��g�p�j
                float columnSpacing = 5f; // �f�t�H���g
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

                // �V�������X�ʒu���v�Z
                var rightmostX = allShapes.Max(s => s.Left + s.Width);
                var newColumnX = rightmostX + columnSpacing;

                // �e�s�ɐV�����I�u�W�F�N�g���쐬
                for (int row = 0; row < gridInfo.Rows; row++)
                {
                    var currentRow = gridInfo.ShapeGrid[row];
                    if (currentRow.Count == 0) continue;

                    // �Q�l�}�`�i�����s�̍ŉE�[�j���珑�����R�s�[
                    var referenceShape = currentRow[currentRow.Count - 1];

                    // �K�؂ȃZ���T�C�Y���v�Z�i�S�̗̂񕝓���̂��ߍŉE�[��̕��ς��g�p�j
                    var rightmostColumnWidths = gridInfo.ShapeGrid
                        .Where(r => r.Count > 0)
                        .Select(r => r[r.Count - 1].Width);
                    var avgWidth = rightmostColumnWidths.DefaultIfEmpty(120f).Average();
                    // �����s�̊����I�u�W�F�N�g�Ɗ��S�ɓ��������E�ʒu�ɓ���
                    var rowTop = currentRow.Min(s => s.Top);
                    var rowHeight = currentRow.Max(s => s.Height);

                    // �V�����e�L�X�g�{�b�N�X���쐬
                    var newTextBox = slide.Shapes.AddTextbox(
                        MsoTextOrientation.msoTextOrientationHorizontal,
                        newColumnX,
                        rowTop,
                        avgWidth,
                        rowHeight
                    );

                    newTextBox.TextFrame.AutoSize = PowerPoint.PpAutoSize.ppAutoSizeNone;

                    // �������Q�l�}�`����R�s�[
                    CopyObjectShapeFormat(referenceShape.Shape, newTextBox);

                    // �f�t�H���g�e�L�X�g��ݒ�
                    if (newTextBox.HasTextFrame == MsoTriState.msoTrue)
                    {
                        newTextBox.TextFrame.TextRange.Text = ""; // ��̃e�L�X�g
                    }

                    createdShapes.Add(newTextBox);
                    logger.Debug($"Created new cell at row {row + 1} for new column");
                }

                // �쐬�����}�`��I����Ԃɂ���
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
        /// �}�`�̏������R�s�[���܂��i�V�@�\��p�j
        /// </summary>
        /// <param name="sourceShape">�R�s�[���}�`</param>
        /// <param name="targetShape">�R�s�[��}�`</param>
        private void CopyObjectShapeFormat(PowerPoint.Shape sourceShape, PowerPoint.Shape targetShape)
        {
            try
            {
                // �h��Ԃ����R�s�[
                if (sourceShape.Fill.Type != MsoFillType.msoFillMixed)
                {
                    targetShape.Fill.ForeColor.RGB = sourceShape.Fill.ForeColor.RGB;
                    targetShape.Fill.Transparency = sourceShape.Fill.Transparency;
                }

                // �����R�s�[
                if (sourceShape.Line.Visible == MsoTriState.msoTrue)
                {
                    targetShape.Line.Visible = MsoTriState.msoTrue;
                    targetShape.Line.ForeColor.RGB = sourceShape.Line.ForeColor.RGB;
                    targetShape.Line.Weight = sourceShape.Line.Weight;
                    targetShape.Line.DashStyle = sourceShape.Line.DashStyle;
                }

                // �e�L�X�g�������R�s�[
                if (sourceShape.HasTextFrame == MsoTriState.msoTrue &&
                    targetShape.HasTextFrame == MsoTriState.msoTrue)
                {
                    var sourceTextFrame = sourceShape.TextFrame;
                    var targetTextFrame = targetShape.TextFrame;

                    // �}�[�W�����R�s�[
                    targetTextFrame.MarginTop = sourceTextFrame.MarginTop;
                    targetTextFrame.MarginBottom = sourceTextFrame.MarginBottom;
                    targetTextFrame.MarginLeft = sourceTextFrame.MarginLeft;
                    targetTextFrame.MarginRight = sourceTextFrame.MarginRight;

                    // �t�H���g�ݒ���R�s�[�i�e�L�X�g������ꍇ�̂݁j
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
                // �����R�s�[���s�͒v���I�ł͂Ȃ��̂Ōp��
            }
        }

        /// <summary>
        /// �����}�`��I����Ԃɂ��܂��i�V�@�\��p�j
        /// </summary>
        /// <param name="shapes">�I������}�`�̃��X�g</param>
        private void SelectCreatedShapes(List<PowerPoint.Shape> shapes)
        {
            try
            {
                if (shapes == null || shapes.Count == 0) return;

                var application = applicationProvider.GetCurrentApplication();
                var slide = helper.GetCurrentSlide();
                if (slide == null) return;

                // �ŏ��̐}�`��I��
                shapes[0].Select(MsoTriState.msoFalse);

                // �c��̐}�`��ǉ��I��
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

                // ��؂�����������č폜
                for (int i = slide.Shapes.Count; i >= 1; i--) // �t���Ń��[�v�i�폜���̃C���f�b�N�X�ύX�΍�j
                {
                    try
                    {
                        var shape = slide.Shapes[i];

                        // "RowSeparator_" �Ŏn�܂閼�O�̐���ΏۂƂ���
                        if (shape.Name.StartsWith("RowSeparator_") && shape.Type == MsoShapeType.msoLine)
                        {
                            shape.Delete();
                            deletedCount++;
                            logger.Debug($"Deleted separator line: {shape.Name}");
                        }
                    }
                    catch (COMException comEx)
                    {
                        // �}�`�����ɍ폜����Ă���ꍇ�ȂǁACOMException��K�؂Ƀn���h�����O
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
                // �G���[�͋L�^���邪�����͌p���i��؂���̍폜���s�őS�̂��~�܂邱�Ƃ�h���j
            }
        }

        /// <summary>
        /// Matrix Tuner
        /// ��`�I�u�W�F�N�g�̃}�g���b�N�X�z�u�����x�ɒ���
        /// </summary>
    }
}
