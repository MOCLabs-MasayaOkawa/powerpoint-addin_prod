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
    /// �p���[�c�[���E����@�\��񋟂���T�[�r�X�N���X
    /// </summary>
    public class PowerToolService
    {
        private static readonly Logger logger = LogManager.GetCurrentClassLogger();
        private readonly IApplicationProvider applicationProvider;
        private readonly PowerToolServiceHelper helper;

        // DI�Ή��R���X�g���N�^�i���p���x���j
        public PowerToolService(IApplicationProvider applicationProvider)
        {
            this.applicationProvider = applicationProvider ?? throw new ArgumentNullException(nameof(applicationProvider));
            logger.Debug("PowerToolService initialized with DI application provider");
            helper = new PowerToolServiceHelper(applicationProvider);
        }

        // �����R���X�g���N�^�i����݊����ێ��j
        public PowerToolService() : this(new DefaultApplicationProvider())
        {
            logger.Debug("PowerToolService initialized with default application provider");
        }

        #region �p���[�c�[���O���[�v (16-23)

        /// <summary>
        /// �e�L�X�g�����i16�ԋ@�\�j
        /// �I�������}�`�̃e�L�X�g�����s��؂�ō������A��}�`�ɐݒ�B���̐}�`���폜
        /// </summary>
        public void MergeText()
        {

            if (!Globals.ThisAddIn.CheckFeatureAccess("MergeText")) return;

            logger.Info("MergeText operation started");

            var selectedShapes = helper.GetSelectedShapeInfos();
            if (!helper.ValidateSelection(selectedShapes, 2, 0, "�e�L�X�g����")) return;

            ComHelper.ExecuteWithComCleanup(() =>
            {
                var textParts = new List<string>();

                // �I�����Ƀe�L�X�g�����W
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
                        throw new InvalidOperationException("�I�������}�`�Ƀe�L�X�g���܂܂�Ă��܂���B");
                    }, "�e�L�X�g����");
                    return;
                }

                // ���s��؂�Ńe�L�X�g������
                var mergedText = string.Join(Environment.NewLine, textParts);

                // �ŏ��̐}�`�i��}�`�j�Ƀe�L�X�g��ݒ�
                var referenceShape = selectedShapes.OrderBy(s => s.SelectionOrder).First();
                var targetShapes = selectedShapes.Skip(1).ToList(); // ��}�`�ȊO

                try
                {
                    // ��}�`�Ƀe�L�X�g��ݒ�
                    if (referenceShape.HasTextFrame)
                    {
                        referenceShape.Shape.TextFrame.TextRange.Text = mergedText;
                    }
                    else
                    {
                        // �e�L�X�g�t���[�����Ȃ��ꍇ�́A�e�L�X�g�{�b�N�X�ɕϊ�
                        referenceShape.Shape.TextFrame.TextRange.Text = mergedText;
                    }

                    // �T�C�Y�𒲐��i�K�v�ɉ����č������g���j
                    var lineCount = textParts.Count;
                    var currentHeight = referenceShape.Height;
                    var estimatedHeight = currentHeight * lineCount * 0.8f; // �T�Z
                    if (estimatedHeight > currentHeight)
                    {
                        referenceShape.Shape.Height = estimatedHeight;
                    }

                    logger.Debug($"Merged text set to reference shape: {referenceShape.Name}");

                    // ��}�`�ȊO���폜
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
                        throw new InvalidOperationException("�e�L�X�g�����Ɏ��s���܂����B");
                    }, "�e�L�X�g����");
                }
            }, selectedShapes.Select(s => s.Shape).ToArray());

            logger.Info("MergeText completed");
        }

        /// <summary>
        /// ���𐅕��ɂ���i18�ԋ@�\�j
        /// �I���������̊p�x�𐅕��i0�x�j�ɂ��A���̒�����ێ�����
        /// </summary>
        public void MakeLineHorizontal()
        {
            if (!Globals.ThisAddIn.CheckFeatureAccess("MakeLineHorizontal")) return;

            logger.Info("MakeLineHorizontal operation started");

            var selectedShapes = helper.GetSelectedShapeInfos();
            if (!helper.ValidateSelection(selectedShapes, 1, 0, "���𐅕��ɂ���")) return;

            var lineShapes = selectedShapes.Where(s => helper.IsLineShape(s.Shape)).ToList();
            if (lineShapes.Count == 0)
            {
                ErrorHandler.ExecuteSafely(() =>
                {
                    throw new InvalidOperationException("���}�`��I�����Ă��������B");
                }, "���𐅕��ɂ���");
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
        /// ���𐂒��ɂ���i19�ԋ@�\�j
        /// �I���������̊p�x�𐂒��i90�x�j�ɂ��A���̒�����ێ�����
        /// </summary>
        public void MakeLineVertical()
        {
            if (!Globals.ThisAddIn.CheckFeatureAccess("MakeLineVertical")) return;

            logger.Info("MakeLineVertical operation started");

            var selectedShapes = helper.GetSelectedShapeInfos();
            if (!helper.ValidateSelection(selectedShapes, 1, 0, "���𐂒��ɂ���")) return;

            var lineShapes = selectedShapes.Where(s => helper.IsLineShape(s.Shape)).ToList();
            if (lineShapes.Count == 0)
            {
                ErrorHandler.ExecuteSafely(() =>
                {
                    throw new InvalidOperationException("���}�`��I�����Ă��������B");
                }, "���𐂒��ɂ���");
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
        /// �}�`�ʒu�̌����i20�ԋ@�\�j
        /// 2�̑I�������}�`�̈ʒu����������
        /// </summary>
        public void SwapPositions()
        {
            if (!Globals.ThisAddIn.CheckFeatureAccess("SwapPositions")) return;

            logger.Info("SwapPositions operation started");

            var selectedShapes = helper.GetSelectedShapeInfos();
            if (!helper.ValidateSelection(selectedShapes, 2, 2, "�}�`�ʒu�̌���")) return;

            var shape1 = selectedShapes[0];
            var shape2 = selectedShapes[1];

            ComHelper.ExecuteWithComCleanup(() =>
            {
                try
                {
                    // �ʒu��ۑ�
                    var temp1Left = shape1.Left;
                    var temp1Top = shape1.Top;

                    // �ʒu������
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
        /// ����}�`�Ɉꊇ�I���i21�ԋ@�\�j
        /// �I�������}�`�Ɠ���̐}�`���X���C�h���ňꊇ�I��
        /// </summary>
        public void SelectSimilarShapes()
        {
            if (!Globals.ThisAddIn.CheckFeatureAccess("SelectSimilarShapes")) return;

            logger.Info("SelectSimilarShapes operation started");

            var selectedShapes = helper.GetSelectedShapeInfos();
            if (!helper.ValidateSelection(selectedShapes, 1, 1, "����}�`�Ɉꊇ�I��")) return;

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
                    // �V�����I�����쐬
                    helper.SelectShapes(similarShapes);
                    logger.Info($"Selected {similarShapes.Count} similar shapes");
                }
                else
                {
                    ErrorHandler.ExecuteSafely(() =>
                    {
                        throw new InvalidOperationException("����̐}�`��������܂���ł����B");
                    }, "����}�`�Ɉꊇ�I��");
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
                    // ���̌��݂̒����ƒ��S�_���v�Z
                    var currentLength = GetLineLength(lineShape);
                    var centerX = lineShape.Left + lineShape.Width / 2;
                    var centerY = lineShape.Top + lineShape.Height / 2;

                    // �������Ƃ��Đݒ�i������ێ��j
                    lineShape.Left = centerX - currentLength / 2;
                    lineShape.Top = centerY;
                    lineShape.Width = currentLength;
                    lineShape.Height = 0;

                    logger.Debug($"Converted line to horizontal with length {currentLength}");
                }
                else if (lineShape.Type == MsoShapeType.msoFreeform && lineShape.Connector == MsoTriState.msoTrue)
                {
                    // �R�l�N�^�̏ꍇ�̏���
                    var currentLength = GetLineLength(lineShape);
                    var centerX = lineShape.Left + lineShape.Width / 2;
                    var centerY = lineShape.Top + lineShape.Height / 2;

                    // �R�l�N�^�𐅕����Ƃ��čĔz�u
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
        /// ���𐂒��ɂ��܂��i������ێ��j
        /// </summary>
        private void MakeLineVerticalInternal(PowerPoint.Shape lineShape)
        {
            try
            {
                if (lineShape.Type == MsoShapeType.msoLine)
                {
                    // ���̌��݂̒����ƒ��S�_���v�Z
                    var currentLength = GetLineLength(lineShape);
                    var centerX = lineShape.Left + lineShape.Width / 2;
                    var centerY = lineShape.Top + lineShape.Height / 2;

                    // �������Ƃ��Đݒ�i������ێ��j
                    lineShape.Left = centerX;
                    lineShape.Top = centerY - currentLength / 2;
                    lineShape.Width = 0;
                    lineShape.Height = currentLength;

                    logger.Debug($"Converted line to vertical with length {currentLength}");
                }
                else if (lineShape.Type == MsoShapeType.msoFreeform && lineShape.Connector == MsoTriState.msoTrue)
                {
                    // �R�l�N�^�̏ꍇ�̏���
                    var currentLength = GetLineLength(lineShape);
                    var centerX = lineShape.Left + lineShape.Width / 2;
                    var centerY = lineShape.Top + lineShape.Height / 2;

                    // �R�l�N�^�𐂒����Ƃ��čĔz�u
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
        /// �t�H���g�I���_�C�A���O��\�����܂�
        /// </summary>
        /// <returns>�I�����ꂽ�t�H���g���A�L�����Z�����͋󕶎�</returns>
        private string ShowFontSelectionDialog()
        {
            string selectedFont = "";

            try
            {
                using (var form = new Form())
                {
                    form.Text = "�t�H���g�ꊇ����";
                    form.Size = new Size(380, 250);
                    form.StartPosition = FormStartPosition.CenterScreen;
                    form.FormBorderStyle = FormBorderStyle.FixedDialog;
                    form.MaximizeBox = false;
                    form.MinimizeBox = false;

                    var label = new Label()
                    {
                        Text = "�v���[���e�[�V�����S�̂ɓK�p����t�H���g��I�����Ă�������:",
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

                    // �����t�H���g���ŏ��ɒǉ�
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

                    // �V�X�e���̑S�t�H���g���擾
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

                    // �f�t�H���g��Meiryo UI��I��
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
                        Text = "����: ���̑���͌��ɖ߂��܂���B\n�K�v�ɉ����Ď��O�Ƀt�@�C����ۑ����Ă��������B",
                        Location = new Point(20, 110),
                        Size = new Size(320, 40),
                        ForeColor = Color.DarkRed,
                        AutoSize = false
                    };

                    var okButton = new Button()
                    {
                        Text = "���s",
                        Location = new Point(180, 170),
                        Size = new Size(75, 25),
                        DialogResult = DialogResult.OK
                    };

                    var cancelButton = new Button()
                    {
                        Text = "�L�����Z��",
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
                    "�t�H���g�I���_�C�A���O�̕\���Ɏ��s���܂����B",
                    "�G���[",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error
                );
            }

            return selectedFont;
        }

        /// <summary>
        /// ���̒������擾���܂�
        /// </summary>
        /// <param name="lineShape">���}�`</param>
        /// <returns>���̒���</returns>
        private float GetLineLength(PowerPoint.Shape lineShape)
        {
            try
            {
                if (lineShape.Type == MsoShapeType.msoLine)
                {
                    // �����̏ꍇ�A���ƍ�������Εӂ��v�Z
                    var width = Math.Abs(lineShape.Width);
                    var height = Math.Abs(lineShape.Height);
                    return (float)Math.Sqrt(width * width + height * height);
                }
                else if (lineShape.Type == MsoShapeType.msoFreeform)
                {
                    // �t���[�t�H�[���i�R�l�N�^���j�̏ꍇ�͕��ƍ����̍ő�l���g�p
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
        /// ���̒����𒲐����܂��i���S�_�Œ�j
        /// </summary>
        /// <param name="lineShape">���}�`</param>
        /// <param name="targetLength">�ڕW�̒���</param>
        private void AdjustLineLength(PowerPoint.Shape lineShape, float targetLength)
        {
            try
            {
                if (lineShape.Type == MsoShapeType.msoLine)
                {
                    var currentLength = GetLineLength(lineShape);
                    if (currentLength <= 0) return;

                    // ���݂̒��S�_��ۑ�
                    var centerX = lineShape.Left + lineShape.Width / 2;
                    var centerY = lineShape.Top + lineShape.Height / 2;

                    var ratio = targetLength / currentLength;

                    // �V�����T�C�Y���v�Z
                    var newWidth = lineShape.Width * ratio;
                    var newHeight = lineShape.Height * ratio;

                    // ���S�_���ێ����ĐV�����T�C�Y��ݒ�
                    lineShape.Left = centerX - newWidth / 2;
                    lineShape.Top = centerY - newHeight / 2;
                    lineShape.Width = newWidth;
                    lineShape.Height = newHeight;

                    logger.Debug($"Adjusted line {lineShape.Name}: length {currentLength:F1} �� {targetLength:F1}, center maintained at ({centerX:F1}, {centerY:F1})");
                }
                else if (lineShape.Type == MsoShapeType.msoFreeform)
                {
                    // �t���[�t�H�[���i�R�l�N�^���j�̏ꍇ�����S�_���ێ�
                    var centerX = lineShape.Left + lineShape.Width / 2;
                    var centerY = lineShape.Top + lineShape.Height / 2;

                    if (Math.Abs(lineShape.Width) > Math.Abs(lineShape.Height))
                    {
                        // ������������̂̏ꍇ
                        var newWidth = lineShape.Width > 0 ? targetLength : -targetLength;
                        lineShape.Left = centerX - newWidth / 2;
                        lineShape.Width = newWidth;
                    }
                    else
                    {
                        // ������������̂̏ꍇ
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

        #region ����@�\�O���[�v (24-27)



        /// <summary>
        /// �t�H���g�ꊇ����i�C���Łj
        /// �S�y�[�W�̂��ׂẴe�L�X�g���w��t�H���g�Ɋ��S����
        /// </summary>
        public void UnifyFont()
        {
            logger.Info("UnifyFont operation started (improved version)");

            // �t�H���g�I���_�C�A���O��\��
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
                        throw new InvalidOperationException("�A�N�e�B�u�ȃv���[���e�[�V������������܂���B");
                    }, "�t�H���g�ꊇ����");
                    return;
                }

                int changedCount = 0;
                int errorCount = 0;

                logger.Info($"Processing {activePresentation.Slides.Count} slides for font unification to '{selectedFont}'");

                // ���ׂẴX���C�h������
                for (int i = 1; i <= activePresentation.Slides.Count; i++)
                {
                    var slide = activePresentation.Slides[i];
                    var slideChangedCount = 0;

                    try
                    {
                        logger.Debug($"Processing slide {i}");

                        // 1. �X���C�h���̂��ׂĂ̐}�`�������i�ʏ�̐}�`�j
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

                        // 2. �X���C�h�̃v���[�X�z���_�[������
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

                // ���ʕ\��
                var message = errorCount > 0
                    ? $"�t�H���g���u{selectedFont}�v�ɓ��ꂵ�܂����B\n�ύX���ꂽ�e�L�X�g��: {changedCount}\n�����G���[: {errorCount}��"
                    : $"�t�H���g���u{selectedFont}�v�ɓ��ꂵ�܂����B\n�ύX���ꂽ�e�L�X�g��: {changedCount}";

                MessageBox.Show(
                    message,
                    "�t�H���g�ꊇ���� ����",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information
                );
            });

            logger.Info("UnifyFont completed (improved version)");
        }

        /// <summary>
        /// �}�`�̃t�H���g���������܂�
        /// </summary>
        /// <param name="shape">�����Ώۂ̐}�`</param>
        /// <param name="targetFont">�ݒ肷��t�H���g��</param>
        /// <returns>�ύX���ꂽ�e�L�X�g�͈͐�</returns>
        private int ProcessShapeFont(PowerPoint.Shape shape, string targetFont)
        {
            int changedCount = 0;

            try
            {
                // 1. �ʏ�̃e�L�X�g�t���[������
                if (shape.HasTextFrame == MsoTriState.msoTrue)
                {
                    changedCount += ProcessTextFrameFont(shape.TextFrame, targetFont, shape.Name);
                }

                // 2. �\�̏���
                if (shape.HasTable == MsoTriState.msoTrue)
                {
                    changedCount += ProcessTableFont(shape.Table, targetFont, shape.Name);
                }

                // 3. �O���[�v�}�`�̏���
                if (shape.Type == MsoShapeType.msoGroup)
                {
                    for (int i = 1; i <= shape.GroupItems.Count; i++)
                    {
                        var groupItem = shape.GroupItems[i];
                        changedCount += ProcessShapeFont(groupItem, targetFont);
                    }
                }

                // 4. SmartArt�A�O���t�Ȃǂ̓���}�`�̏���
                if (shape.Type == MsoShapeType.msoChart ||
                    shape.Type == MsoShapeType.msoSmartArt ||
                    shape.Type == MsoShapeType.msoDiagram)
                {
                    // ��{�I�ȃe�L�X�g�t���[���̂ݏ����i�ڍׂ�SmartArt�����͕��G�����邽�ߏȗ��j
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
        /// �v���[�X�z���_�[�̃t�H���g���������܂�
        /// </summary>
        /// <param name="placeholder">�v���[�X�z���_�[</param>
        /// <param name="targetFont">�ݒ肷��t�H���g��</param>
        /// <returns>�ύX���ꂽ�e�L�X�g�͈͐�</returns>
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
        /// �e�L�X�g�t���[���̃t�H���g���ڍ׏������܂�
        /// </summary>
        /// <param name="textFrame">�e�L�X�g�t���[��</param>
        /// <param name="targetFont">�ݒ肷��t�H���g��</param>
        /// <param name="shapeName">�}�`���i���O�p�j</param>
        /// <returns>�ύX���ꂽ�e�L�X�g�͈͐�</returns>
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

                // ���@1: �S�̂̃t�H���g���ꊇ�ύX
                try
                {
                    textRange.Font.Name = targetFont;
                    changedCount++;
                    logger.Debug($"Changed font for entire text range in {shapeName}");
                }
                catch (Exception ex)
                {
                    logger.Warn(ex, $"Failed to change font for entire text range in {shapeName}, trying character-by-character");

                    // ���@2: �����P�ʂł̕ύX�i�t�H�[���o�b�N�j
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

                // ���@3: �i���P�ʂł̕ύX�i�ǉ��̕ی��j
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
        /// �\�̃t�H���g���������܂�
        /// </summary>
        /// <param name="table">�\</param>
        /// <param name="targetFont">�ݒ肷��t�H���g��</param>
        /// <param name="shapeName">�}�`���i���O�p�j</param>
        /// <returns>�ύX���ꂽ�e�L�X�g�͈͐�</returns>
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
        /// ���̒����𑵂���i�V�@�\C�j
        /// �I���������̒��ōŏ��ɑI���������̂���ɒ����𑵂��A��[�𑵂���
        /// </summary>
        public void AlignLineLength()
        {
            if (!Globals.ThisAddIn.CheckFeatureAccess("AlignLineLength")) return;

            logger.Info("AlignLineLength operation started");

            var selectedShapes = helper.GetSelectedShapeInfos();
            if (!helper.ValidateSelection(selectedShapes, 2, 0, "���̒����𑵂���")) return;

            var lineShapes = selectedShapes.Where(s => helper.IsLineShape(s.Shape)).ToList();
            if (lineShapes.Count < 2)
            {
                ErrorHandler.ExecuteSafely(() =>
                {
                    throw new InvalidOperationException("�Œ�2�̐��}�`��I�����Ă��������B");
                }, "���̒����𑵂���");
                return;
            }

            ComHelper.ExecuteWithComCleanup(() =>
            {
                // �ŏ��ɑI������������Ƃ��Ď擾
                var referenceLine = lineShapes.First();
                var referenceLength = GetLineLength(referenceLine.Shape);

                logger.Debug($"Reference line: {referenceLine.Name}, Length: {referenceLength} (�ŏ��ɑI��)");

                // ���̐��𒲐��i�ʒu�͈ړ����������̂ݒ����j
                foreach (var lineInfo in lineShapes.Skip(1))
                {
                    try
                    {
                        AdjustLineLength(lineInfo.Shape, referenceLength);

                        logger.Debug($"Adjusted line {lineInfo.Name} to length {referenceLength} (�: {referenceLine.Name}, �ʒu�ێ�)");
                    }
                    catch (Exception ex)
                    {
                        logger.Error(ex, $"Failed to adjust line {lineInfo.Name}");
                    }
                }

                logger.Info($"AlignLineLength completed for {lineShapes.Count} lines (�: �ŏ��I��, �ʒu�ړ��Ȃ�)");
            }, lineShapes.Select(l => l.Shape).ToArray());

            logger.Info("AlignLineLength completed (�: �ŏ��I��, �ʒu�ړ��Ȃ�)");
        }

        /// <summary>
        /// �}�`�ɘA�ԕt�^�i�V�@�\D�j
        /// �I��}�`�ɍ�����1����̘A�Ԃ������e�L�X�g�̌�ɒǉ�
        /// </summary>
        public void AddSequentialNumbers()
        {
            if (!Globals.ThisAddIn.CheckFeatureAccess("AddSequentialNumbers")) return;

            logger.Info("AddSequentialNumbers operation started");

            var selectedShapes = helper.GetSelectedShapeInfos();
            if (!helper.ValidateSelection(selectedShapes, 1, 0, "�}�`�ɘA�ԕt�^")) return;

            ComHelper.ExecuteWithComCleanup(() =>
            {
                // �����Ń\�[�g�i�ォ�牺�A������E�j
                var sortedShapes = selectedShapes.OrderBy(s => s.Top).ThenBy(s => s.Left).ToList();

                for (int i = 0; i < sortedShapes.Count; i++)
                {
                    var shapeInfo = sortedShapes[i];
                    var sequenceNumber = (i + 1).ToString();

                    try
                    {
                        // �e�L�X�g�t���[�����Ȃ��ꍇ�͍쐬
                        if (shapeInfo.Shape.HasTextFrame != MsoTriState.msoTrue)
                        {
                            // �}�`�Ƀe�L�X�g�t���[����ǉ�������@�����s
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
                            // �����e�L�X�g�̌�ɔԍ���ǉ�
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



        #region �}�`�X�^�C���ݒ�@�\�i�V�@�\�j

        /// <summary>
        /// �}�`�X�^�C���ݒ�_�C�A���O��\��
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
                            "�}�`�X�^�C���ݒ肪�ۑ�����܂����B\n�V�����쐬����}�`�ɐݒ肪�K�p����܂��B",
                            "�}�`�X�^�C���ݒ�",
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
                ErrorHandler.ExecuteSafely(() => throw ex, "�}�`�X�^�C���ݒ�");
            }
        }

        /// <summary>
        /// �쐬���ꂽ�}�`�ɃX�^�C����K�p
        /// </summary>
        /// <param name="shape">�Ώې}�`</param>
        /// <param name="commandName">�R�}���h���i�}�`��ʂ̔���p�j</param>
        public void ApplyShapeStyle(PowerPoint.Shape shape, string commandName)
        {
            if (shape == null) return;

            // �y�C���z�e�L�X�g�{�b�N�X�ɂ͈ꊇ�X�^�C���ݒ��K�p���Ȃ�
            if (IsTextBoxCommand(commandName))
            {
                logger.Debug($"Skipping style application for TextBox: {shape.Name}");
                return;
            }

            try
            {
                // ���݂̐ݒ��ǂݍ���
                var settings = SettingsService.Instance.LoadShapeStyleSettings();

                if (!settings.IsApplicable())
                {
                    logger.Debug("Shape styling is disabled or invalid settings");
                    return;
                }

                logger.Debug($"Applying shape style to {shape.Name} (command: {commandName})");

                // �F�ݒ��K�p
                ApplyFillStyle(shape, settings, commandName);
                ApplyLineStyle(shape, settings, commandName);

                // �t�H���g�F�݂̂�K�p
                ApplyFontColorStyle(shape, settings);

                logger.Info($"Shape style applied successfully to {shape.Name}");
            }
            catch (Exception ex)
            {
                logger.Error(ex, $"Failed to apply shape style to {shape.Name}");
                // �X�^�C���K�p�̎��s�͐}�`�쐬��W���Ȃ�
            }
        }

        /// <summary>
        /// �t�H���g�F�݂̂�K�p
        /// </summary>
        /// <param name="shape">�Ώې}�`</param>
        /// <param name="settings">�X�^�C���ݒ�</param>
        private void ApplyFontColorStyle(PowerPoint.Shape shape, ShapeStyleSettings settings)
        {
            if (shape == null) return;

            try
            {
                // �e�L�X�g�����}�`�̃t�H���g�F��ύX
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
        /// �h��Ԃ��X�^�C����K�p
        /// </summary>
        /// <param name="shape">�Ώې}�`</param>
        /// <param name="settings">�X�^�C���ݒ�</param>
        /// <param name="commandName">�R�}���h��</param>
        private void ApplyFillStyle(PowerPoint.Shape shape, ShapeStyleSettings settings, string commandName)
        {
            try
            {
                // ���n�}�`�͓h��Ԃ��s�v
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
        /// ���X�^�C����K�p
        /// </summary>
        /// <param name="shape">�Ώې}�`</param>
        /// <param name="settings">�X�^�C���ݒ�</param>
        /// <param name="commandName">�R�}���h��</param>
        private void ApplyLineStyle(PowerPoint.Shape shape, ShapeStyleSettings settings, string commandName)
        {
            try
            {
                if (shape.Line != null)
                {
                    shape.Line.Visible = MsoTriState.msoTrue;
                    shape.Line.ForeColor.RGB = ColorTranslator.ToOle(settings.LineColor);

                    // ���n�}�`�̏ꍇ�͐�����K�؂ɐݒ�
                    if (IsLineCommand(commandName))
                    {
                        shape.Line.Weight = 2.0f; // �f�t�H���g����
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
        /// �e�L�X�g�{�b�N�X�n�R�}���h���ǂ����𔻒肵�܂�
        /// </summary>
        /// <param name="commandName">�R�}���h��</param>
        /// <returns>�e�L�X�g�{�b�N�X�n�̏ꍇtrue</returns>
        private bool IsTextBoxCommand(string commandName)
        {
            // �e�L�X�g�{�b�N�X�n�R�}���h�̃��X�g
            var textBoxCommands = new HashSet<string>
    {
        "TextBox"
        // �����I�ɑ��̃e�L�X�g�n�}�`���ǉ������ꍇ�͂����ɒǉ�
    };

            return textBoxCommands.Contains(commandName);
        }

        /// <summary>
        /// ���n�R�}���h���ǂ����𔻒�
        /// </summary>
        /// <param name="commandName">�R�}���h��</param>
        /// <returns>���n�R�}���h�̏ꍇtrue</returns>
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
        /// �X���C�h��̐}�`�z�u�󋵂����O�o�́i�f�o�b�O�p�j
        /// </summary>
        /// <param name="slide">�ΏۃX���C�h</param>
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

                // �񂲂Ƃɕ��ނ��ă��O�o��
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
        /// �}�`����R�}���h���𐄒�
        /// </summary>
        /// <param name="shape">�}�`</param>
        /// <returns>���肳�ꂽ�R�}���h��</returns>
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
                                return "ShapeRectangle"; // �f�t�H���g
                        }
                    case MsoShapeType.msoTextBox:
                        return "TextBox";
                    case MsoShapeType.msoLine:
                        return "ShapeLine";
                    default:
                        return "ShapeRectangle"; // �f�t�H���g
                }
            }
            catch (Exception ex)
            {
                logger.Warn(ex, $"Failed to estimate command name for shape {shape.Name}");
                return "ShapeRectangle"; // �t�H�[���o�b�N
            }
        }


        /// <summary>
        /// �e�L�X�g��\���p�ɏȗ����܂��i���O�o�͗p�j
        /// ��TextFormatService.cs��TruncateText()�Ɠ����p�^�[���𗬗p
        /// </summary>
        /// <param name="text">���e�L�X�g</param>
        /// <returns>�ȗ����ꂽ�e�L�X�g</returns>
        private string TruncateText(string text)
        {
            if (string.IsNullOrEmpty(text)) return "(empty)";

            const int maxLength = 20;
            if (text.Length <= maxLength) return text;

            return text.Substring(0, maxLength) + "...";
        }

    }
}