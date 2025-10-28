using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using PowerPointEfficiencyAddin.Models;
using PowerPointEfficiencyAddin.Utils;
using NLog;
using PowerPointEfficiencyAddin.Services.Infrastructure.MultiInstance;

namespace PowerPointEfficiencyAddin.Services.Core.Shape
{
    /// <summary>
    /// 図形整形機能を提供するサービスクラス
    /// </summary>
    public class ShapeService
    {
        private static readonly Logger logger = LogManager.GetCurrentClassLogger();
        private readonly IApplicationProvider applicationProvider;

        // DI対応コンストラクタ
        public ShapeService(IApplicationProvider applicationProvider)
        {
            this.applicationProvider = applicationProvider ?? throw new ArgumentNullException(nameof(applicationProvider));
            logger.Debug("ShapeService initialized with DI application provider");
        }

        // 既存コンストラクタ（後方互換性維持）
        public ShapeService() : this(new DefaultApplicationProvider())
        {
            logger.Debug("ShapeService initialized with default application provider");
        }

        /// <summary>
        /// 幅を合わせる（1番機能）
        /// 最初に選択した図形の幅に他の図形を合わせる
        /// </summary>
        public void MatchWidth()
        {
            if (!Globals.ThisAddIn.CheckFeatureAccess("MatchWidth")) return;

            logger.Info("MatchWidth operation started");

            var selectedShapes = GetSelectedShapeInfos();
            if (!ValidateSelection(selectedShapes, 2, 0, "幅を合わせる")) return;

            var referenceShape = selectedShapes.First(); // 最初に選択した図形を基準
            var targetShapes = selectedShapes.Skip(1).ToList();

            ComHelper.ExecuteWithComCleanup(() =>
            {
                foreach (var shapeInfo in targetShapes)
                {
                    try
                    {
                        shapeInfo.Shape.Width = referenceShape.Width;
                        logger.Debug($"Set width of {shapeInfo.Name} to {referenceShape.Width} (基準: {referenceShape.Name})");
                    }
                    catch (Exception ex)
                    {
                        logger.Error(ex, $"Failed to set width for shape {shapeInfo.Name}");
                    }
                }
            }, selectedShapes.Select(s => s.Shape).ToArray());

            logger.Info($"MatchWidth completed for {targetShapes.Count} shapes (基準: 最初選択)");
        }

        /// <summary>
        /// 高さを合わせる（2番機能）
        /// 最初に選択した図形の高さに他の図形を合わせる
        /// </summary>
        public void MatchHeight()
        {
            if (!Globals.ThisAddIn.CheckFeatureAccess("MatchHeight")) return;

            logger.Info("MatchHeight operation started");

            var selectedShapes = GetSelectedShapeInfos();
            if (!ValidateSelection(selectedShapes, 2, 0, "高さを合わせる")) return;

            var referenceShape = selectedShapes.First(); // 最初に選択した図形を基準
            var targetShapes = selectedShapes.Skip(1).ToList();

            ComHelper.ExecuteWithComCleanup(() =>
            {
                foreach (var shapeInfo in targetShapes)
                {
                    try
                    {
                        shapeInfo.Shape.Height = referenceShape.Height;
                        logger.Debug($"Set height of {shapeInfo.Name} to {referenceShape.Height} (基準: {referenceShape.Name})");
                    }
                    catch (Exception ex)
                    {
                        logger.Error(ex, $"Failed to set height for shape {shapeInfo.Name}");
                    }
                }
            }, selectedShapes.Select(s => s.Shape).ToArray());

            logger.Info($"MatchHeight completed for {targetShapes.Count} shapes (基準: 最初選択)");
        }

        /// <summary>
        /// 幅・高さを合わせる（3番機能）
        /// 最初に選択した図形の幅と高さに他の図形を合わせる
        /// </summary>
        public void MatchSize()
        {
            if (!Globals.ThisAddIn.CheckFeatureAccess("MatchSize")) return;

            logger.Info("MatchSize operation started");

            var selectedShapes = GetSelectedShapeInfos();
            if (!ValidateSelection(selectedShapes, 2, 0, "幅・高さを合わせる")) return;

            var referenceShape = selectedShapes.First(); // 最初に選択した図形を基準
            var targetShapes = selectedShapes.Skip(1).ToList();

            ComHelper.ExecuteWithComCleanup(() =>
            {
                foreach (var shapeInfo in targetShapes)
                {
                    try
                    {
                        shapeInfo.Shape.Width = referenceShape.Width;
                        shapeInfo.Shape.Height = referenceShape.Height;
                        logger.Debug($"Set size of {shapeInfo.Name} to {referenceShape.Width}x{referenceShape.Height} (基準: {referenceShape.Name})");
                    }
                    catch (Exception ex)
                    {
                        logger.Error(ex, $"Failed to set size for shape {shapeInfo.Name}");
                    }
                }
            }, selectedShapes.Select(s => s.Shape).ToArray());

            logger.Info($"MatchSize completed for {targetShapes.Count} shapes (基準: 最初選択)");
        }

        /// <summary>
        /// 書式を合わせる（4番機能）
        /// 最初に選択した図形の書式に他の図形を合わせる
        /// </summary>
        public void MatchFormat()
        {
            if (!Globals.ThisAddIn.CheckFeatureAccess("MatchFormat")) return;

            logger.Info("MatchFormat operation started");

            var selectedShapes = GetSelectedShapeInfos();
            if (!ValidateSelection(selectedShapes, 2, 0, "書式を合わせる")) return;

            var referenceShape = selectedShapes.First(); // 最初に選択した図形を基準
            var targetShapes = selectedShapes.Skip(1).ToList();

            ComHelper.ExecuteWithComCleanup(() =>
            {
                try
                {
                    // 基準図形の書式を取得
                    referenceShape.Shape.PickUp();

                    // 他の図形に書式を適用
                    foreach (var shapeInfo in targetShapes)
                    {
                        try
                        {
                            shapeInfo.Shape.Apply();
                            logger.Debug($"Applied format to {shapeInfo.Name} (基準: {referenceShape.Name})");
                        }
                        catch (Exception ex)
                        {
                            logger.Error(ex, $"Failed to apply format to shape {shapeInfo.Name}");
                        }
                    }
                }
                catch (Exception ex)
                {
                    logger.Error(ex, "Failed to pickup format from reference shape");

                    // 代替方法: 個別に書式要素をコピー
                    CopyIndividualFormatProperties(referenceShape, targetShapes);
                }
            }, selectedShapes.Select(s => s.Shape).ToArray());

            logger.Info($"MatchFormat completed for {targetShapes.Count} shapes (基準: 最初選択)");
        }

        /// <summary>
        /// 図形環境を合わせる（5番機能）
        /// ハンドル設定のある図形のハンドル位置を同じにする
        /// </summary>
        public void MatchEnvironment()
        {
            if (!Globals.ThisAddIn.CheckFeatureAccess("MatchEnvironment")) return;

            logger.Info("MatchEnvironment operation started");

            var selectedShapes = GetSelectedShapeInfos();
            if (!ValidateSelection(selectedShapes, 2, 0, "図形環境を合わせる")) return;

            var referenceShape = selectedShapes.First(); // 最初に選択した図形を基準
            var targetShapes = selectedShapes.Skip(1).ToList();

            ComHelper.ExecuteWithComCleanup(() =>
            {
                // 基準図形のハンドル設定を取得
                var referenceAdjustments = GetShapeAdjustments(referenceShape.Shape);

                if (referenceAdjustments.Count == 0)
                {
                    logger.Warn($"Reference shape {referenceShape.Name} has no adjustments");
                    ErrorHandler.ExecuteSafely(() =>
                    {
                        throw new InvalidOperationException("基準図形にハンドル設定がありません。");
                    }, "図形環境を合わせる");
                    return;
                }

                foreach (var shapeInfo in targetShapes)
                {
                    try
                    {
                        ApplyShapeAdjustments(shapeInfo.Shape, referenceAdjustments);
                        logger.Debug($"Applied adjustments to {shapeInfo.Name} (基準: {referenceShape.Name})");
                    }
                    catch (Exception ex)
                    {
                        logger.Error(ex, $"Failed to apply adjustments to shape {shapeInfo.Name}");
                    }
                }
            }, selectedShapes.Select(s => s.Shape).ToArray());

            logger.Info($"MatchEnvironment completed for {targetShapes.Count} shapes (基準: 最初選択)");
        }


        /// <summary>
        /// 角丸の設定を合わせる（6番機能）
        /// 角丸具合のある図形の角丸位置を同じにする
        /// </summary>
        public void MatchRoundCorner()
        {
            if (!Globals.ThisAddIn.CheckFeatureAccess("MatchRoundCorner")) return;

            logger.Info("MatchRoundCorner operation started");

            var selectedShapes = GetSelectedShapeInfos();
            if (!ValidateSelection(selectedShapes, 2, 0, "角丸の設定を合わせる")) return;

            var referenceShape = selectedShapes.First(); // 最初に選択した図形を基準
            var targetShapes = selectedShapes.Skip(1).ToList();

            ComHelper.ExecuteWithComCleanup(() =>
            {
                // 基準図形の角丸設定を取得
                var roundCornerRadius = GetRoundCornerRadius(referenceShape.Shape);

                if (roundCornerRadius == null)
                {
                    logger.Warn($"Reference shape {referenceShape.Name} has no round corner settings");
                    ErrorHandler.ExecuteSafely(() =>
                    {
                        throw new InvalidOperationException("基準図形に角丸設定がありません。");
                    }, "角丸の設定を合わせる");
                    return;
                }

                foreach (var shapeInfo in targetShapes)
                {
                    try
                    {
                        SetRoundCornerRadius(shapeInfo.Shape, roundCornerRadius.Value);
                        logger.Debug($"Applied round corner to {shapeInfo.Name} (基準: {referenceShape.Name})");
                    }
                    catch (Exception ex)
                    {
                        logger.Error(ex, $"Failed to apply round corner to shape {shapeInfo.Name}");
                    }
                }
            }, selectedShapes.Select(s => s.Shape).ToArray());

            logger.Info($"MatchRoundCorner completed for {targetShapes.Count} shapes (基準: 最初選択)");
        }

        /// <summary>
        /// マトリクス生成（新機能B）
        /// ポップアップで設定したマトリクスを四角形で生成
        /// </summary>
        public void GenerateMatrix()
        {
            if (!Globals.ThisAddIn.CheckFeatureAccess("GenerateMatrix")) return;

            logger.Info("GenerateMatrix operation started");

            // マトリクス設定ダイアログを表示
            var matrixSettings = ShowMatrixSettingsDialog();
            if (matrixSettings == null)
            {
                logger.Info("Matrix generation cancelled by user");
                return;
            }

            ComHelper.ExecuteWithComCleanup(() =>
            {
                // Undo対応のためのアクション開始
                var slide = GetCurrentSlide();
                if (slide == null)
                {
                    ErrorHandler.ExecuteSafely(() =>
                    {
                        throw new InvalidOperationException("アクティブなスライドが見つかりません。");
                    }, "マトリクス生成");
                    return;
                }

                var createdShapes = new List<PowerPoint.Shape>();

                // 基準位置（現在選択位置、なければデフォルト）
                var startX = 100f;
                var startY = 100f;

                // 現在選択されている図形があれば、その位置を基準にする
                var selectedShapes = GetSelectedShapeInfos();
                if (selectedShapes.Count > 0)
                {
                    startX = selectedShapes.First().Left;
                    startY = selectedShapes.First().Top;
                }

                // cmをポイントに変換（1cm = 28.35ポイント）
                var cellWidth = matrixSettings.CellWidth * 28.35f;
                var cellHeight = matrixSettings.CellHeight * 28.35f;
                var spacing = matrixSettings.Spacing * 28.35f;

                // マトリクス生成
                for (int row = 0; row < matrixSettings.Rows; row++)
                {
                    for (int col = 0; col < matrixSettings.Columns; col++)
                    {
                        var x = startX + col * (cellWidth + spacing);
                        var y = startY + row * (cellHeight + spacing);

                        var rectangle = slide.Shapes.AddShape(
                            MsoAutoShapeType.msoShapeRectangle,
                            x, y, cellWidth, cellHeight
                        );

                        // 基本書式設定
                        rectangle.Fill.ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.White);
                        rectangle.Line.Visible = MsoTriState.msoTrue;
                        rectangle.Line.Weight = 1.0f;
                        rectangle.Line.ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);

                        createdShapes.Add(rectangle);
                        logger.Debug($"Created matrix cell [{row},{col}] at ({x}, {y})");
                    }
                }

                // 作成した図形を選択状態にする
                if (createdShapes.Count > 0)
                {
                    SelectShapes(createdShapes);
                    logger.Info($"Generated matrix: {matrixSettings.Rows}x{matrixSettings.Columns} = {createdShapes.Count} shapes");
                }
            }, "マトリクス生成" /* comObjects */);

            logger.Info("GenerateMatrix completed");
        }

        /// <summary>
        /// 図形等間隔調整（新機能E）
        /// 選択図形を表形式に整頓し、指定間隔で配置
        /// </summary>
        public void AdjustEqualSpacing()
        {
            if (!Globals.ThisAddIn.CheckFeatureAccess("AdjustEqualSpacing")) return;

            logger.Info("AdjustEqualSpacing operation started");

            var selectedShapes = GetSelectedShapeInfos();
            if (!ValidateSelection(selectedShapes, 2, 0, "図形等間隔調整")) return;

            // 間隔設定ダイアログを表示
            var spacing = ShowSpacingSettingsDialog();
            if (spacing == null)
            {
                logger.Info("Spacing adjustment cancelled by user");
                return;
            }

            ComHelper.ExecuteWithComCleanup(() =>
            {
                // 図形を行・列にグループ化
                var grid = OrganizeShapesIntoGrid(selectedShapes);

                // 左上の図形を基準として設定
                var baseX = grid.SelectMany(row => row).Min(s => s.Left);
                var baseY = grid.Min(row => row.Min(s => s.Top));

                var spacingPoints = spacing.Value * 28.35f; // cmをポイントに変換

                // グリッド配置で間隔調整
                for (int row = 0; row < grid.Count; row++)
                {
                    for (int col = 0; col < grid[row].Count; col++)
                    {
                        var shape = grid[row][col];

                        // 新しい位置を計算
                        var newX = baseX;
                        var newY = baseY;

                        // 前の列の図形の右端 + 間隔を加算
                        for (int prevCol = 0; prevCol < col; prevCol++)
                        {
                            newX += grid[row][prevCol].Width + spacingPoints;
                        }

                        // 前の行の図形の下端 + 間隔を加算
                        for (int prevRow = 0; prevRow < row; prevRow++)
                        {
                            if (prevRow == 0)
                            {
                                newY += grid[prevRow].Max(s => s.Height) + spacingPoints;
                            }
                            else
                            {
                                newY += grid[prevRow].Max(s => s.Height) + spacingPoints;
                            }
                        }

                        // 図形を移動
                        shape.Shape.Left = newX;
                        shape.Shape.Top = newY;

                        logger.Debug($"Moved shape {shape.Name} to ({newX}, {newY})");
                    }
                }

                logger.Info($"AdjustEqualSpacing completed for {selectedShapes.Count} shapes with {spacing.Value}cm spacing");
            }, selectedShapes.Select(s => s.Shape).ToArray());

            logger.Info("AdjustEqualSpacing completed");
        }

        #region Private Helper Methods

        /// <summary>
        /// 現在選択されている図形の情報を取得します
        /// </summary>
        /// <returns>選択された図形の情報リスト</returns>
        private List<ShapeInfo> GetSelectedShapeInfos()
        {
            var shapeInfos = new List<ShapeInfo>();

            try
            {
                var addin = Globals.ThisAddIn;
                var selectedShapes = addin.GetSelectedShapes();

                if (selectedShapes != null)
                {
                    for (int i = 1; i <= selectedShapes.Count; i++)
                    {
                        var shape = selectedShapes[i];
                        shapeInfos.Add(new ShapeInfo(shape, i - 1));
                    }
                }
            }
            catch (Exception ex)
            {
                logger.Error(ex, "Failed to get selected shape infos");
            }

            return shapeInfos;
        }

        /// <summary>
        /// 選択状態を検証します
        /// </summary>
        /// <param name="shapeInfos">図形情報リスト</param>
        /// <param name="minRequired">最小必要数</param>
        /// <param name="maxAllowed">最大許可数（0は無制限）</param>
        /// <param name="operationName">操作名</param>
        /// <returns>検証結果</returns>
        private bool ValidateSelection(List<ShapeInfo> shapeInfos, int minRequired, int maxAllowed, string operationName)
        {
            return ErrorHandler.ValidateSelection(shapeInfos.Count, minRequired, maxAllowed, operationName);
        }

        /// <summary>
        /// 個別に書式プロパティをコピーします（PickUp/Applyの代替）
        /// </summary>
        /// <param name="referenceShape">基準図形</param>
        /// <param name="targetShapes">対象図形リスト</param>
        private void CopyIndividualFormatProperties(ShapeInfo referenceShape, List<ShapeInfo> targetShapes)
        {
            try
            {
                var refShape = referenceShape.Shape;

                foreach (var shapeInfo in targetShapes)
                {
                    var targetShape = shapeInfo.Shape;

                    try
                    {
                        // 塗りつぶし
                        if (refShape.Fill.Type != MsoFillType.msoFillMixed)
                        {
                            targetShape.Fill.ForeColor.RGB = refShape.Fill.ForeColor.RGB;
                            targetShape.Fill.BackColor.RGB = refShape.Fill.BackColor.RGB;
                            targetShape.Fill.Transparency = refShape.Fill.Transparency;
                        }

                        // 線
                        if (refShape.Line.Visible == MsoTriState.msoTrue)
                        {
                            targetShape.Line.ForeColor.RGB = refShape.Line.ForeColor.RGB;
                            targetShape.Line.Weight = refShape.Line.Weight;
                            targetShape.Line.DashStyle = refShape.Line.DashStyle;
                        }

                        // 影
                        if (refShape.Shadow.Type != MsoShadowType.msoShadowMixed)
                        {
                            targetShape.Shadow.Type = refShape.Shadow.Type;
                            if (refShape.Shadow.Visible == MsoTriState.msoTrue)
                            {
                                targetShape.Shadow.ForeColor.RGB = refShape.Shadow.ForeColor.RGB;
                                targetShape.Shadow.OffsetX = refShape.Shadow.OffsetX;
                                targetShape.Shadow.OffsetY = refShape.Shadow.OffsetY;
                            }
                        }

                        logger.Debug($"Copied individual format properties to {shapeInfo.Name}");
                    }
                    catch (Exception ex)
                    {
                        logger.Error(ex, $"Failed to copy format properties to {shapeInfo.Name}");
                    }
                }
            }
            catch (Exception ex)
            {
                logger.Error(ex, "Failed to copy individual format properties");
            }
        }

        /// <summary>
        /// 図形のハンドル調整値を取得します
        /// </summary>
        /// <param name="shape">図形</param>
        /// <returns>調整値のディクショナリ</returns>
        private Dictionary<int, float> GetShapeAdjustments(PowerPoint.Shape shape)
        {
            var adjustments = new Dictionary<int, float>();

            try
            {
                if (shape.Type == MsoShapeType.msoAutoShape)
                {
                    for (int i = 1; i <= shape.Adjustments.Count; i++)
                    {
                        adjustments[i] = shape.Adjustments[i];
                    }
                }
            }
            catch (Exception ex)
            {
                logger.Error(ex, $"Failed to get adjustments for shape {shape.Name}");
            }

            return adjustments;
        }

        /// <summary>
        /// 図形にハンドル調整値を適用します
        /// </summary>
        /// <param name="shape">図形</param>
        /// <param name="adjustments">調整値のディクショナリ</param>
        private void ApplyShapeAdjustments(PowerPoint.Shape shape, Dictionary<int, float> adjustments)
        {
            try
            {
                if (shape.Type == MsoShapeType.msoAutoShape && adjustments.Count > 0)
                {
                    foreach (var adjustment in adjustments)
                    {
                        if (adjustment.Key <= shape.Adjustments.Count)
                        {
                            shape.Adjustments[adjustment.Key] = adjustment.Value;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                logger.Error(ex, $"Failed to apply adjustments to shape {shape.Name}");
            }
        }

        /// <summary>
        /// 図形の角丸半径を取得します
        /// </summary>
        /// <param name="shape">図形</param>
        /// <returns>角丸半径（設定されていない場合はnull）</returns>
        private float? GetRoundCornerRadius(PowerPoint.Shape shape)
        {
            try
            {
                // 角丸四角形の場合
                if (shape.AutoShapeType == MsoAutoShapeType.msoShapeRoundedRectangle)
                {
                    if (shape.Adjustments.Count > 0)
                    {
                        return shape.Adjustments[1];
                    }
                }

                // その他の角丸対応図形
                // TODO: 必要に応じて他の図形タイプに対応
            }
            catch (Exception ex)
            {
                logger.Error(ex, $"Failed to get round corner radius for shape {shape.Name}");
            }

            return null;
        }

        /// <summary>
        /// 図形に角丸半径を設定します
        /// </summary>
        /// <param name="shape">図形</param>
        /// <param name="radius">角丸半径</param>
        private void SetRoundCornerRadius(PowerPoint.Shape shape, float radius)
        {
            try
            {
                // 角丸四角形の場合
                if (shape.AutoShapeType == MsoAutoShapeType.msoShapeRoundedRectangle)
                {
                    if (shape.Adjustments.Count > 0)
                    {
                        shape.Adjustments[1] = radius;
                    }
                }

                // その他の角丸対応図形
                // TODO: 必要に応じて他の図形タイプに対応
            }
            catch (Exception ex)
            {
                logger.Error(ex, $"Failed to set round corner radius for shape {shape.Name}");
            }
        }

        #endregion

        #region New Feature Helper Methods

        /// <summary>
        /// マトリクス設定ダイアログを表示します
        /// </summary>
        /// <returns>マトリクス設定、キャンセル時はnull</returns>
        private MatrixSettings ShowMatrixSettingsDialog()
        {
            var result = new MatrixSettings();

            using (var form = new System.Windows.Forms.Form())
            {
                form.Text = "マトリクス生成";
                form.Size = new System.Drawing.Size(300, 280);
                form.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
                form.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
                form.MaximizeBox = false;
                form.MinimizeBox = false;

                // 行数
                var labelRows = new System.Windows.Forms.Label()
                {
                    Text = "行数:",
                    Location = new System.Drawing.Point(20, 20),
                    Size = new System.Drawing.Size(80, 20)
                };
                var numRows = new System.Windows.Forms.NumericUpDown()
                {
                    Location = new System.Drawing.Point(110, 18),
                    Size = new System.Drawing.Size(120, 20),
                    Minimum = 1,
                    Maximum = 20,
                    Value = 2,
                    Increment = 1
                };

                // 列数
                var labelCols = new System.Windows.Forms.Label()
                {
                    Text = "列数:",
                    Location = new System.Drawing.Point(20, 50),
                    Size = new System.Drawing.Size(80, 20)
                };
                var numCols = new System.Windows.Forms.NumericUpDown()
                {
                    Location = new System.Drawing.Point(110, 48),
                    Size = new System.Drawing.Size(120, 20),
                    Minimum = 1,
                    Maximum = 20,
                    Value = 2,
                    Increment = 1
                };

                // 横幅
                var labelWidth = new System.Windows.Forms.Label()
                {
                    Text = "横幅(cm):",
                    Location = new System.Drawing.Point(20, 80),
                    Size = new System.Drawing.Size(80, 20)
                };
                var numWidth = new System.Windows.Forms.NumericUpDown()
                {
                    Location = new System.Drawing.Point(110, 78),
                    Size = new System.Drawing.Size(120, 20),
                    Minimum = 0.01M,
                    Maximum = 50.0M,
                    Value = 3.0M,
                    DecimalPlaces = 2,
                    Increment = 0.1M
                };

                // 縦幅
                var labelHeight = new System.Windows.Forms.Label()
                {
                    Text = "縦幅(cm):",
                    Location = new System.Drawing.Point(20, 110),
                    Size = new System.Drawing.Size(80, 20)
                };
                var numHeight = new System.Windows.Forms.NumericUpDown()
                {
                    Location = new System.Drawing.Point(110, 108),
                    Size = new System.Drawing.Size(120, 20),
                    Minimum = 0.01M,
                    Maximum = 50.0M,
                    Value = 1.0M,
                    DecimalPlaces = 2,
                    Increment = 0.1M
                };

                // 間隔
                var labelSpacing = new System.Windows.Forms.Label()
                {
                    Text = "間隔(cm):",
                    Location = new System.Drawing.Point(20, 140),
                    Size = new System.Drawing.Size(80, 20)
                };
                var numSpacing = new System.Windows.Forms.NumericUpDown()
                {
                    Location = new System.Drawing.Point(110, 138),
                    Size = new System.Drawing.Size(120, 20),
                    Minimum = 0.0M,
                    Maximum = 10.0M,
                    Value = 0.2M,
                    DecimalPlaces = 2,
                    Increment = 0.1M
                };

                // ボタン
                var okButton = new System.Windows.Forms.Button()
                {
                    Text = "OK",
                    Location = new System.Drawing.Point(110, 180),
                    Size = new System.Drawing.Size(75, 30),
                    DialogResult = System.Windows.Forms.DialogResult.OK
                };

                var cancelButton = new System.Windows.Forms.Button()
                {
                    Text = "キャンセル",
                    Location = new System.Drawing.Point(195, 180),
                    Size = new System.Drawing.Size(75, 30),
                    DialogResult = System.Windows.Forms.DialogResult.Cancel
                };

                form.Controls.AddRange(new System.Windows.Forms.Control[]
                {
                    labelRows, numRows, labelCols, numCols, labelWidth, numWidth,
                    labelHeight, numHeight, labelSpacing, numSpacing, okButton, cancelButton
                });

                form.AcceptButton = okButton;
                form.CancelButton = cancelButton;

                if (form.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    result.Rows = (int)numRows.Value;
                    result.Columns = (int)numCols.Value;
                    result.CellWidth = (float)numWidth.Value;
                    result.CellHeight = (float)numHeight.Value;
                    result.Spacing = (float)numSpacing.Value;
                    return result;
                }
            }

            return null;
        }

        /// <summary>
        /// 間隔設定ダイアログを表示します
        /// </summary>
        /// <returns>間隔(cm)、キャンセル時はnull</returns>
        private float? ShowSpacingSettingsDialog()
        {
            using (var form = new System.Windows.Forms.Form())
            {
                form.Text = "図形等間隔調整";
                form.Size = new System.Drawing.Size(300, 180);
                form.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
                form.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
                form.MaximizeBox = false;
                form.MinimizeBox = false;

                var label = new System.Windows.Forms.Label()
                {
                    Text = "図形間の間隔(cm):",
                    Location = new System.Drawing.Point(20, 30),
                    Size = new System.Drawing.Size(150, 20)
                };

                var numSpacing = new System.Windows.Forms.NumericUpDown()
                {
                    Location = new System.Drawing.Point(20, 60),
                    Size = new System.Drawing.Size(150, 20),
                    Minimum = 0.0M,
                    Maximum = 10.0M,
                    Value = 0.2M,
                    DecimalPlaces = 2,
                    Increment = 0.1M
                };

                var okButton = new System.Windows.Forms.Button()
                {
                    Text = "OK",
                    Location = new System.Drawing.Point(100, 110),
                    Size = new System.Drawing.Size(75, 30),
                    DialogResult = System.Windows.Forms.DialogResult.OK
                };

                var cancelButton = new System.Windows.Forms.Button()
                {
                    Text = "キャンセル",
                    Location = new System.Drawing.Point(185, 110),
                    Size = new System.Drawing.Size(75, 30),
                    DialogResult = System.Windows.Forms.DialogResult.Cancel
                };

                form.Controls.AddRange(new System.Windows.Forms.Control[] { label, numSpacing, okButton, cancelButton });
                form.AcceptButton = okButton;
                form.CancelButton = cancelButton;

                if (form.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    return (float)numSpacing.Value;
                }
            }

            return null;
        }

        /// <summary>
        /// 図形を行・列のグリッドに整理します
        /// </summary>
        /// <param name="shapes">図形リスト</param>
        /// <returns>行・列に整理された図形グリッド</returns>
        private List<List<ShapeInfo>> OrganizeShapesIntoGrid(List<ShapeInfo> shapes)
        {
            // 動的な許容誤差を計算
            var tolerance = CalculateDynamicTolerance(shapes, true); // true = Y座標用
            var rows = new List<List<ShapeInfo>>();

            var sortedByY = shapes.OrderBy(s => s.Top).ToList();

            foreach (var shape in sortedByY)
            {
                var assignedToRow = false;

                // 既存の行に追加可能かチェック
                foreach (var row in rows)
                {
                    var avgY = row.Average(s => s.Top);
                    if (Math.Abs(shape.Top - avgY) <= tolerance)
                    {
                        row.Add(shape);
                        assignedToRow = true;
                        break;
                    }
                }

                // 新しい行を作成
                if (!assignedToRow)
                {
                    rows.Add(new List<ShapeInfo> { shape });
                }
            }

            // 各行内でX座標順にソート
            foreach (var row in rows)
            {
                row.Sort((a, b) => a.Left.CompareTo(b.Left));
            }

            logger.Debug($"Grid organized: {rows.Count} rows with tolerance {tolerance:F1}pt");
            return rows;
        }

        /// <summary>
        /// 動的な許容誤差を計算します
        /// </summary>
        /// <param name="shapes">図形リスト</param>
        /// <param name="isVertical">true=Y座標用（高さベース）、false=X座標用（幅ベース）</param>
        /// <returns>計算された許容誤差（ポイント）</returns>
        private float CalculateDynamicTolerance(List<ShapeInfo> shapes, bool isVertical)
        {
            if (!shapes.Any()) return 10f; // デフォルト値

            // 図形の平均サイズを計算
            var averageSize = isVertical
                ? shapes.Average(s => s.Height)  // 行判定用：高さの平均
                : shapes.Average(s => s.Width);  // 列判定用：幅の平均

            // 平均サイズの30%を許容誤差とする（調整可能）
            var calculatedTolerance = averageSize * 0.3f;

            // 最小・最大の制限を設ける
            const float MIN_TOLERANCE = 3f;   // 最小3pt
            const float MAX_TOLERANCE = 25f;  // 最大25pt

            var tolerance = Math.Max(MIN_TOLERANCE, Math.Min(MAX_TOLERANCE, calculatedTolerance));

            logger.Debug($"Dynamic tolerance calculated: {tolerance:F1}pt (avg {(isVertical ? "height" : "width")}: {averageSize:F1}pt)");
            return tolerance;
        }

        /// <summary>
        /// 現在のスライドを取得します
        /// </summary>
        /// <returns>アクティブなスライド</returns>
        private PowerPoint.Slide GetCurrentSlide()
        {
            try
            {
                var application = applicationProvider.GetCurrentApplication();
                var activeWindow = application.ActiveWindow;

                if (activeWindow.ViewType == PowerPoint.PpViewType.ppViewSlide ||
                    activeWindow.ViewType == PowerPoint.PpViewType.ppViewNormal)
                {
                    return activeWindow.View.Slide;
                }
            }
            catch (Exception ex)
            {
                logger.Error(ex, "Failed to get current slide");
            }

            return null;
        }

        /// <summary>
        /// 図形を選択状態にします
        /// </summary>
        /// <param name="shapes">選択する図形リスト</param>
        private void SelectShapes(List<PowerPoint.Shape> shapes)
        {
            try
            {
                if (shapes.Count == 1)
                {
                    shapes[0].Select();
                }
                else if (shapes.Count > 1)
                {
                    shapes[0].Select();
                    for (int i = 1; i < shapes.Count; i++)
                    {
                        shapes[i].Select(MsoTriState.msoFalse);
                    }
                }
            }
            catch (Exception ex)
            {
                logger.Error(ex, "Failed to select shapes");
            }
        }

        #endregion

        /// <summary>
        /// マトリクス設定クラス
        /// </summary>
        public class MatrixSettings
        {
            public int Rows { get; set; }
            public int Columns { get; set; }
            public float CellWidth { get; set; }
            public float CellHeight { get; set; }
            public float Spacing { get; set; }
        }

        #region 図形分割・複製機能

        /// <summary>
        /// 図形分割（新機能F）
        /// 選択した図形を指定したグリッドに分割し、元図形を削除
        /// </summary>
        public void SplitShape()
        {
            if (!Globals.ThisAddIn.CheckFeatureAccess("SplitShape")) return;

            logger.Info("SplitShape operation started");

            var selectedShapes = GetSelectedShapeInfos();
            if (!ValidateSelection(selectedShapes, 1, 1, "図形分割")) return;

            var targetShape = selectedShapes.First();

            // 分割設定ダイアログを表示
            var splitSettings = ShowSplitShapeDialog();
            if (splitSettings == null)
            {
                logger.Info("Shape split cancelled by user");
                return;
            }

            ComHelper.ExecuteWithComCleanup(() =>
            {
                var slide = GetCurrentSlide();
                if (slide == null)
                {
                    ErrorHandler.ExecuteSafely(() =>
                    {
                        throw new InvalidOperationException("アクティブなスライドが見つかりません。");
                    }, "図形分割");
                    return;
                }

                var createdShapes = new List<PowerPoint.Shape>();

                try
                {
                    // 元図形の情報を保存
                    var originalLeft = targetShape.Left;
                    var originalTop = targetShape.Top;
                    var originalWidth = targetShape.Width;
                    var originalHeight = targetShape.Height;

                    // 分割後の各セルサイズを計算
                    var spacingPoints = splitSettings.Spacing * 28.35f; // cmをポイントに変換
                    var totalSpacingWidth = (splitSettings.Columns - 1) * spacingPoints;
                    var totalSpacingHeight = (splitSettings.Rows - 1) * spacingPoints;

                    var cellWidth = (originalWidth - totalSpacingWidth) / splitSettings.Columns;
                    var cellHeight = (originalHeight - totalSpacingHeight) / splitSettings.Rows;

                    // 元図形の書式を取得
                    targetShape.Shape.PickUp();

                    // 分割図形を作成
                    for (int row = 0; row < splitSettings.Rows; row++)
                    {
                        for (int col = 0; col < splitSettings.Columns; col++)
                        {
                            var x = originalLeft + col * (cellWidth + spacingPoints);
                            var y = originalTop + row * (cellHeight + spacingPoints);

                            // 元図形と同じタイプの図形を作成
                            var newShape = CreateSimilarShape(slide, targetShape.Shape, x, y, cellWidth, cellHeight);

                            if (newShape != null)
                            {
                                // 書式を適用
                                newShape.Apply();

                                // テキストがある場合は空にする（分割時は個別テキストなし）
                                if (newShape.HasTextFrame == MsoTriState.msoTrue)
                                {
                                    newShape.TextFrame.TextRange.Text = "";
                                }

                                createdShapes.Add(newShape);
                                logger.Debug($"Created split shape [{row},{col}] at ({x:F1}, {y:F1})");
                            }
                        }
                    }

                    // 元図形を削除
                    targetShape.Shape.Delete();
                    logger.Debug($"Deleted original shape: {targetShape.Name}");

                    // 作成した図形を選択状態にする
                    if (createdShapes.Count > 0)
                    {
                        SelectShapes(createdShapes);
                        logger.Info($"Split shape into {splitSettings.Rows}x{splitSettings.Columns} = {createdShapes.Count} shapes");
                    }
                }
                catch (Exception ex)
                {
                    logger.Error(ex, "Failed to split shape");
                    ErrorHandler.ExecuteSafely(() =>
                    {
                        throw new InvalidOperationException("図形の分割に失敗しました。");
                    }, "図形分割");
                }
            }, selectedShapes.Select(s => s.Shape).ToArray());

            logger.Info("SplitShape completed");
        }

        /// <summary>
        /// 図形複製（新機能G）
        /// 選択した図形を指定したグリッドに複製
        /// </summary>
        public void DuplicateShape()
        {
            if (!Globals.ThisAddIn.CheckFeatureAccess("DuplicateShape")) return;

            logger.Info("DuplicateShape operation started");

            var selectedShapes = GetSelectedShapeInfos();
            if (!ValidateSelection(selectedShapes, 1, 1, "図形複製")) return;

            var targetShape = selectedShapes.First();

            // 複製設定ダイアログを表示
            var duplicateSettings = ShowDuplicateShapeDialog();
            if (duplicateSettings == null)
            {
                logger.Info("Shape duplication cancelled by user");
                return;
            }

            ComHelper.ExecuteWithComCleanup(() =>
            {
                var slide = GetCurrentSlide();
                if (slide == null)
                {
                    ErrorHandler.ExecuteSafely(() =>
                    {
                        throw new InvalidOperationException("アクティブなスライドが見つかりません。");
                    }, "図形複製");
                    return;
                }

                var createdShapes = new List<PowerPoint.Shape>();

                try
                {
                    // 元図形の情報
                    var baseLeft = targetShape.Left;
                    var baseTop = targetShape.Top;
                    var shapeWidth = targetShape.Width;
                    var shapeHeight = targetShape.Height;
                    var spacingPoints = duplicateSettings.Spacing * 28.35f; // cmをポイントに変換

                    // 元図形のテキスト
                    var originalText = duplicateSettings.IncludeText && targetShape.HasTextFrame &&
                                     !string.IsNullOrEmpty(targetShape.Text) ? targetShape.Text : "";

                    // 複製図形を作成（最初の1,1は元図形なのでスキップ）
                    for (int row = 0; row < duplicateSettings.Rows; row++)
                    {
                        for (int col = 0; col < duplicateSettings.Columns; col++)
                        {
                            // 元図形の位置（row=0, col=0）はスキップ
                            if (row == 0 && col == 0) continue;

                            var x = baseLeft + col * (shapeWidth + spacingPoints);
                            var y = baseTop + row * (shapeHeight + spacingPoints);

                            // 図形を複製（Duplicate()はShapeRangeを返すため、[1]で最初の図形を取得）
                            var duplicatedShapeRange = targetShape.Shape.Duplicate();
                            var duplicatedShape = duplicatedShapeRange[1];
                            duplicatedShape.Left = x;
                            duplicatedShape.Top = y;

                            // テキスト設定
                            if (duplicatedShape.HasTextFrame == MsoTriState.msoTrue)
                            {
                                if (duplicateSettings.IncludeText)
                                {
                                    duplicatedShape.TextFrame.TextRange.Text = originalText;
                                }
                                else
                                {
                                    duplicatedShape.TextFrame.TextRange.Text = "";
                                }
                            }

                            createdShapes.Add(duplicatedShape);
                            logger.Debug($"Created duplicate shape [{row},{col}] at ({x:F1}, {y:F1})");
                        }
                    }

                    // 元図形も含めて選択状態にする
                    var allShapes = new List<PowerPoint.Shape> { targetShape.Shape };
                    allShapes.AddRange(createdShapes);
                    SelectShapes(allShapes);

                    var totalShapes = duplicateSettings.Rows * duplicateSettings.Columns;
                    logger.Info($"Duplicated shape to {duplicateSettings.Rows}x{duplicateSettings.Columns} = {totalShapes} shapes total");
                }
                catch (Exception ex)
                {
                    logger.Error(ex, "Failed to duplicate shape");
                    ErrorHandler.ExecuteSafely(() =>
                    {
                        throw new InvalidOperationException("図形の複製に失敗しました。");
                    }, "図形複製");
                }
            }, selectedShapes.Select(s => s.Shape).ToArray());

            logger.Info("DuplicateShape completed");
        }

        #endregion

        #region 分割・複製 Helper Methods

        /// <summary>
        /// 図形分割設定ダイアログを表示します
        /// </summary>
        /// <returns>分割設定、キャンセル時はnull</returns>
        private SplitShapeSettings ShowSplitShapeDialog()
        {
            using (var form = new System.Windows.Forms.Form())
            {
                form.Text = "図形分割";
                form.Size = new System.Drawing.Size(320, 230); // 高さを280に増加
                form.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
                form.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
                form.MaximizeBox = false;
                form.MinimizeBox = false;
                form.MinimumSize = new System.Drawing.Size(320, 230); // 最小サイズ設定

                // 行数
                var labelRows = new System.Windows.Forms.Label()
                {
                    Text = "行数:",
                    Location = new System.Drawing.Point(20, 20),
                    Size = new System.Drawing.Size(80, 20)
                };
                var numRows = new System.Windows.Forms.NumericUpDown()
                {
                    Location = new System.Drawing.Point(110, 18),
                    Size = new System.Drawing.Size(120, 20),
                    Minimum = 1,
                    Maximum = 30,
                    Value = 3,
                    Increment = 1
                };

                // 列数
                var labelCols = new System.Windows.Forms.Label()
                {
                    Text = "列数:",
                    Location = new System.Drawing.Point(20, 50),
                    Size = new System.Drawing.Size(80, 20)
                };
                var numCols = new System.Windows.Forms.NumericUpDown()
                {
                    Location = new System.Drawing.Point(110, 48),
                    Size = new System.Drawing.Size(120, 20),
                    Minimum = 1,
                    Maximum = 30,
                    Value = 3,
                    Increment = 1
                };

                // 間隔
                var labelSpacing = new System.Windows.Forms.Label()
                {
                    Text = "間隔(cm):",
                    Location = new System.Drawing.Point(20, 80),
                    Size = new System.Drawing.Size(80, 20)
                };
                var numSpacing = new System.Windows.Forms.NumericUpDown()
                {
                    Location = new System.Drawing.Point(110, 78),
                    Size = new System.Drawing.Size(120, 20),
                    Minimum = 0.0M,
                    Maximum = 10.0M,
                    Value = 0.1M,
                    DecimalPlaces = 2,
                    Increment = 0.1M
                };

                // エラーメッセージラベル（位置調整）
                var errorLabel = new System.Windows.Forms.Label()
                {
                    Text = "※ 1つの図形を選択してください",
                    Location = new System.Drawing.Point(20, 110), // Y位置を調整
                    Size = new System.Drawing.Size(250, 20),
                    ForeColor = System.Drawing.Color.Red,
                    Visible = false
                };

                // OKボタン（位置調整）
                var btnOK = new System.Windows.Forms.Button()
                {
                    Text = "OK",
                    Location = new System.Drawing.Point(110, 140), // Y位置を調整
                    Size = new System.Drawing.Size(75, 25),
                    DialogResult = System.Windows.Forms.DialogResult.OK
                };

                // キャンセルボタン（位置調整）
                var btnCancel = new System.Windows.Forms.Button()
                {
                    Text = "キャンセル",
                    Location = new System.Drawing.Point(200, 140), // Y位置を調整
                    Size = new System.Drawing.Size(75, 25),
                    DialogResult = System.Windows.Forms.DialogResult.Cancel
                };

                // コントロール追加
                form.Controls.AddRange(new System.Windows.Forms.Control[]
                {
            labelRows, numRows,
            labelCols, numCols,
            labelSpacing, numSpacing,
            errorLabel,
            btnOK, btnCancel
                });

                form.AcceptButton = btnOK;
                form.CancelButton = btnCancel;

                if (form.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    return new SplitShapeSettings
                    {
                        Rows = (int)numRows.Value,
                        Columns = (int)numCols.Value,
                        Spacing = (float)numSpacing.Value
                    };
                }

                return null;
            }
        }

        /// <summary>
        /// 図形複製設定ダイアログを表示します
        /// </summary>
        /// <returns>複製設定、キャンセル時はnull</returns>
        private DuplicateShapeSettings ShowDuplicateShapeDialog()
        {
            using (var form = new System.Windows.Forms.Form())
            {
                form.Text = "図形複製";
                form.Size = new System.Drawing.Size(300, 260);
                form.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
                form.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
                form.MaximizeBox = false;
                form.MinimizeBox = false;

                // 行数
                var labelRows = new System.Windows.Forms.Label()
                {
                    Text = "行数:",
                    Location = new System.Drawing.Point(20, 20),
                    Size = new System.Drawing.Size(80, 20)
                };
                var numRows = new System.Windows.Forms.NumericUpDown()
                {
                    Location = new System.Drawing.Point(110, 18),
                    Size = new System.Drawing.Size(120, 20),
                    Minimum = 1,
                    Maximum = 30,
                    Value = 3,
                    Increment = 1
                };

                // 列数
                var labelCols = new System.Windows.Forms.Label()
                {
                    Text = "列数:",
                    Location = new System.Drawing.Point(20, 50),
                    Size = new System.Drawing.Size(80, 20)
                };
                var numCols = new System.Windows.Forms.NumericUpDown()
                {
                    Location = new System.Drawing.Point(110, 48),
                    Size = new System.Drawing.Size(120, 20),
                    Minimum = 1,
                    Maximum = 30,
                    Value = 3,
                    Increment = 1
                };

                // 間隔
                var labelSpacing = new System.Windows.Forms.Label()
                {
                    Text = "間隔(cm):",
                    Location = new System.Drawing.Point(20, 80),
                    Size = new System.Drawing.Size(80, 20)
                };
                var numSpacing = new System.Windows.Forms.NumericUpDown()
                {
                    Location = new System.Drawing.Point(110, 78),
                    Size = new System.Drawing.Size(120, 20),
                    Minimum = 0.0M,
                    Maximum = 10.0M,
                    Value = 0.1M,
                    DecimalPlaces = 2,
                    Increment = 0.1M
                };

                // 文字を含める
                var checkIncludeText = new System.Windows.Forms.CheckBox()
                {
                    Text = "文字を含める",
                    Location = new System.Drawing.Point(20, 110),
                    Size = new System.Drawing.Size(150, 20),
                    Checked = true
                };

                // エラーメッセージラベル
                var errorLabel = new System.Windows.Forms.Label()
                {
                    Text = "※ 1つの図形のみ選択してください",
                    Location = new System.Drawing.Point(20, 140),
                    Size = new System.Drawing.Size(250, 20),
                    ForeColor = System.Drawing.Color.Red,
                    Font = new System.Drawing.Font("メイリオ", 8f)
                };

                // ボタン
                var okButton = new System.Windows.Forms.Button()
                {
                    Text = "OK",
                    Location = new System.Drawing.Point(110, 180),
                    Size = new System.Drawing.Size(75, 30),
                    DialogResult = System.Windows.Forms.DialogResult.OK
                };

                var cancelButton = new System.Windows.Forms.Button()
                {
                    Text = "キャンセル",
                    Location = new System.Drawing.Point(195, 180),
                    Size = new System.Drawing.Size(75, 30),
                    DialogResult = System.Windows.Forms.DialogResult.Cancel
                };

                form.Controls.AddRange(new System.Windows.Forms.Control[]
                {
                    labelRows, numRows, labelCols, numCols, labelSpacing, numSpacing,
                    checkIncludeText, errorLabel, okButton, cancelButton
                });

                form.AcceptButton = okButton;
                form.CancelButton = cancelButton;

                if (form.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    return new DuplicateShapeSettings
                    {
                        Rows = (int)numRows.Value,
                        Columns = (int)numCols.Value,
                        Spacing = (float)numSpacing.Value,
                        IncludeText = checkIncludeText.Checked
                    };
                }
            }

            return null;
        }

        /// <summary>
        /// 元図形と似た図形を作成します
        /// </summary>
        /// <param name="slide">スライド</param>
        /// <param name="originalShape">元図形</param>
        /// <param name="left">左位置</param>
        /// <param name="top">上位置</param>
        /// <param name="width">幅</param>
        /// <param name="height">高さ</param>
        /// <returns>作成された図形</returns>
        private PowerPoint.Shape CreateSimilarShape(PowerPoint.Slide slide, PowerPoint.Shape originalShape,
            float left, float top, float width, float height)
        {
            try
            {
                PowerPoint.Shape newShape = null;

                switch (originalShape.Type)
                {
                    case MsoShapeType.msoAutoShape:
                        newShape = slide.Shapes.AddShape(originalShape.AutoShapeType, left, top, width, height);
                        break;

                    case MsoShapeType.msoTextBox:
                        newShape = slide.Shapes.AddTextbox(MsoTextOrientation.msoTextOrientationHorizontal,
                            left, top, width, height);
                        break;

                    case MsoShapeType.msoFreeform:
                    case MsoShapeType.msoGroup:
                    case MsoShapeType.msoPicture:
                        // 複雑な図形の場合は複製してサイズ変更
                        var duplicatedShapeRange = originalShape.Duplicate();
                        newShape = duplicatedShapeRange[1];
                        newShape.Left = left;
                        newShape.Top = top;
                        newShape.Width = width;
                        newShape.Height = height;
                        break;

                    default:
                        // その他の場合は四角形で代替
                        newShape = slide.Shapes.AddShape(MsoAutoShapeType.msoShapeRectangle, left, top, width, height);
                        break;
                }

                return newShape;
            }
            catch (Exception ex)
            {
                logger.Error(ex, $"Failed to create similar shape for {originalShape.Name}");

                // フォールバック：四角形を作成
                try
                {
                    return slide.Shapes.AddShape(MsoAutoShapeType.msoShapeRectangle, left, top, width, height);
                }
                catch (Exception fallbackEx)
                {
                    logger.Error(fallbackEx, "Failed to create fallback rectangle shape");
                    return null;
                }
            }
        }

        #endregion

        #region 分割・複製設定クラス

        /// <summary>
        /// 図形分割設定クラス
        /// </summary>
        public class SplitShapeSettings
        {
            public int Rows { get; set; }
            public int Columns { get; set; }
            public float Spacing { get; set; }
        }

        /// <summary>
        /// 図形複製設定クラス
        /// </summary>
        public class DuplicateShapeSettings
        {
            public int Rows { get; set; }
            public int Columns { get; set; }
            public float Spacing { get; set; }
            public bool IncludeText { get; set; }
        }


        #endregion

        #region オブジェクト調整機能

        /// <summary>
        /// サイズアップトグル
        /// 選択した図形のサイズを5%大きくする（累積効果）
        /// </summary>
        public void SizeUpToggle()
        {
            if (!Globals.ThisAddIn.CheckFeatureAccess("SizeUpToggle")) return;

            logger.Info("SizeUpToggle operation started");

            var selectedShapes = GetSelectedShapeInfos();
            if (!ValidateSelection(selectedShapes, 1, 0, "サイズアップ")) return;

            ComHelper.ExecuteWithComCleanup(() =>
            {
                foreach (var shapeInfo in selectedShapes)
                {
                    try
                    {
                        // 5%サイズアップ（1.05倍）
                        var newWidth = shapeInfo.Width * 1.05f;
                        var newHeight = shapeInfo.Height * 1.05f;

                        // 中心位置を保持してサイズ変更
                        var centerX = shapeInfo.CenterX;
                        var centerY = shapeInfo.CenterY;

                        shapeInfo.Shape.Width = newWidth;
                        shapeInfo.Shape.Height = newHeight;
                        shapeInfo.Shape.Left = centerX - newWidth / 2;
                        shapeInfo.Shape.Top = centerY - newHeight / 2;

                        logger.Debug($"Size up: {shapeInfo.Name} to {newWidth:F1}x{newHeight:F1}");
                    }
                    catch (Exception ex)
                    {
                        logger.Error(ex, $"Failed to size up shape {shapeInfo.Name}");
                    }
                }
            }, selectedShapes.Select(s => s.Shape).ToArray());

            logger.Info($"SizeUpToggle completed for {selectedShapes.Count} shapes");
        }

        /// <summary>
        /// サイズダウントグル
        /// 選択した図形のサイズを5%小さくする（累積効果）
        /// </summary>
        public void SizeDownToggle()
        {
            if (!Globals.ThisAddIn.CheckFeatureAccess("SizeDownToggle")) return;

            logger.Info("SizeDownToggle operation started");

            var selectedShapes = GetSelectedShapeInfos();
            if (!ValidateSelection(selectedShapes, 1, 0, "サイズダウン")) return;

            ComHelper.ExecuteWithComCleanup(() =>
            {
                foreach (var shapeInfo in selectedShapes)
                {
                    try
                    {
                        // 5%サイズダウン（0.95倍）
                        var newWidth = shapeInfo.Width * 0.95f;
                        var newHeight = shapeInfo.Height * 0.95f;

                        // 最小サイズ制限（1pt）
                        if (newWidth < 1f || newHeight < 1f)
                        {
                            logger.Warn($"Shape {shapeInfo.Name} cannot be smaller than 1pt");
                            continue;
                        }

                        // 中心位置を保持してサイズ変更
                        var centerX = shapeInfo.CenterX;
                        var centerY = shapeInfo.CenterY;

                        shapeInfo.Shape.Width = newWidth;
                        shapeInfo.Shape.Height = newHeight;
                        shapeInfo.Shape.Left = centerX - newWidth / 2;
                        shapeInfo.Shape.Top = centerY - newHeight / 2;

                        logger.Debug($"Size down: {shapeInfo.Name} to {newWidth:F1}x{newHeight:F1}");
                    }
                    catch (Exception ex)
                    {
                        logger.Error(ex, $"Failed to size down shape {shapeInfo.Name}");
                    }
                }
            }, selectedShapes.Select(s => s.Shape).ToArray());

            logger.Info($"SizeDownToggle completed for {selectedShapes.Count} shapes");
        }

        /// <summary>
        /// 線太さ変更アップトグル
        /// 選択した図形の枠線の太さを0.25pt太くする（累積効果）
        /// </summary>
        public void LineWeightUpToggle()
        {
            if (!Globals.ThisAddIn.CheckFeatureAccess("LineWeightUpToggle")) return;

            logger.Info("LineWeightUpToggle operation started");

            var selectedShapes = GetSelectedShapeInfos();
            if (!ValidateSelection(selectedShapes, 1, 0, "線太さアップ")) return;

            ComHelper.ExecuteWithComCleanup(() =>
            {
                foreach (var shapeInfo in selectedShapes)
                {
                    try
                    {
                        var shape = shapeInfo.Shape;

                        // 線が非表示の場合は表示にする
                        if (shape.Line.Visible != MsoTriState.msoTrue)
                        {
                            shape.Line.Visible = MsoTriState.msoTrue;
                        }

                        // 現在の線の太さを取得
                        var currentWeight = shape.Line.Weight;
                        var newWeight = currentWeight + 0.25f;

                        // PowerPointの最大線幅制限（1584pt）を適用
                        if (newWeight > 1584f)
                        {
                            logger.Warn($"Shape {shapeInfo.Name} line weight cannot exceed 1584pt");
                            newWeight = 1584f;
                        }

                        shape.Line.Weight = newWeight;

                        logger.Debug($"Line weight up: {shapeInfo.Name} from {currentWeight:F2}pt to {newWeight:F2}pt");
                    }
                    catch (Exception ex)
                    {
                        logger.Error(ex, $"Failed to increase line weight for shape {shapeInfo.Name}");
                    }
                }
            }, selectedShapes.Select(s => s.Shape).ToArray());

            logger.Info($"LineWeightUpToggle completed for {selectedShapes.Count} shapes");
        }

        /// <summary>
        /// 線太さ変更ダウントグル
        /// 選択した図形の枠線の太さを0.25pt細くする（累積効果）
        /// </summary>
        public void LineWeightDownToggle()
        {
            if (!Globals.ThisAddIn.CheckFeatureAccess("LineWeightDownToggle")) return;

            logger.Info("LineWeightDownToggle operation started");

            var selectedShapes = GetSelectedShapeInfos();
            if (!ValidateSelection(selectedShapes, 1, 0, "線太さダウン")) return;

            ComHelper.ExecuteWithComCleanup(() =>
            {
                foreach (var shapeInfo in selectedShapes)
                {
                    try
                    {
                        var shape = shapeInfo.Shape;

                        // 線が非表示の場合は何もしない
                        if (shape.Line.Visible != MsoTriState.msoTrue)
                        {
                            logger.Debug($"Shape {shapeInfo.Name} has no visible line, skipping");
                            continue;
                        }

                        // 現在の線の太さを取得
                        var currentWeight = shape.Line.Weight;

                        // 現在が0ptの場合は何もしない
                        if (currentWeight <= 0f)
                        {
                            logger.Debug($"Shape {shapeInfo.Name} line weight is already 0pt, skipping");
                            continue;
                        }

                        var newWeight = Math.Max(0f, currentWeight - 0.25f);
                        shape.Line.Weight = newWeight;

                        logger.Debug($"Line weight down: {shapeInfo.Name} from {currentWeight:F2}pt to {newWeight:F2}pt");
                    }
                    catch (Exception ex)
                    {
                        logger.Error(ex, $"Failed to decrease line weight for shape {shapeInfo.Name}");
                    }
                }
            }, selectedShapes.Select(s => s.Shape).ToArray());

            logger.Info($"LineWeightDownToggle completed for {selectedShapes.Count} shapes");
        }

        /// <summary>
        /// 点線変更トグル
        /// 選択した図形の枠線を順次変更する
        /// 実線 → 点線（丸） → 点線（角） → 破線 → 1点鎖線 → 長破線 → 長鎖線 → 長二点鎖線 → 実線に戻る
        /// </summary>
        public void DashStyleToggle()
        {
            if (!Globals.ThisAddIn.CheckFeatureAccess("DashStyleToggle")) return;

            logger.Info("DashStyleToggle operation started - using corrected order based on MsoLineDashStyle values");

            var selectedShapes = GetSelectedShapeInfos();
            if (!ValidateSelection(selectedShapes, 1, 0, "点線変更")) return;

            ComHelper.ExecuteWithComCleanup(() =>
            {
                foreach (var shapeInfo in selectedShapes)
                {
                    try
                    {
                        var shape = shapeInfo.Shape;

                        // 線が非表示の場合は表示にする
                        if (shape.Line.Visible != MsoTriState.msoTrue)
                        {
                            shape.Line.Visible = MsoTriState.msoTrue;
                            // 線を表示した場合は、デフォルトの線幅を設定
                            if (shape.Line.Weight < 0.25f)
                            {
                                shape.Line.Weight = 0.75f; // デフォルト線幅
                            }
                            logger.Debug($"Made line visible for {shapeInfo.Name}");
                        }

                        // 現在の線種を取得
                        var currentDashStyle = shape.Line.DashStyle;
                        var currentStyleName = GetDashStyleName(currentDashStyle);
                        var currentValue = (int)currentDashStyle;

                        // 次の線種に変更
                        var nextDashStyle = GetNextDashStyle(currentDashStyle);
                        var nextStyleName = GetDashStyleName(nextDashStyle);
                        var nextValue = (int)nextDashStyle;

                        shape.Line.DashStyle = nextDashStyle;

                        // 詳細ログ（値も含む）
                        logger.Info($"Dash style changed for {shapeInfo.Name}: {currentStyleName}({currentValue}) → {nextStyleName}({nextValue})");
                    }
                    catch (Exception ex)
                    {
                        logger.Error(ex, $"Failed to change dash style for shape {shapeInfo.Name}");
                    }
                }
            }, selectedShapes.Select(s => s.Shape).ToArray());

            logger.Info($"DashStyleToggle completed for {selectedShapes.Count} shapes");
        }

        /// <summary>
        /// 次の線種を取得します
        /// </summary>
        /// <param name="currentDashStyle">現在の線種</param>
        /// <returns>次の線種</returns>
        private MsoLineDashStyle GetNextDashStyle(MsoLineDashStyle currentDashStyle)
        {
            
            switch (currentDashStyle)
            {
                case MsoLineDashStyle.msoLineSolid:               // 1: 実線
                    return MsoLineDashStyle.msoLineSquareDot;     // → 2: 点線(角)

                case MsoLineDashStyle.msoLineSquareDot:           // 2: 点線(角)
                    return MsoLineDashStyle.msoLineRoundDot;      // → 3: 点線(丸)

                case MsoLineDashStyle.msoLineRoundDot:            // 3: 点線(丸)
                    return MsoLineDashStyle.msoLineDash;          // → 4: 破線

                case MsoLineDashStyle.msoLineDash:                // 4: 破線
                    return MsoLineDashStyle.msoLineLongDash;      // → 7: 長破線

                case MsoLineDashStyle.msoLineLongDash:            // 7: 長破線
                    return MsoLineDashStyle.msoLineDashDot;       // → 5: 一点鎖線

                case MsoLineDashStyle.msoLineDashDot:             // 5: 一点鎖線
                    return MsoLineDashStyle.msoLineDashDotDot;    // → 6: 二点鎖線

                case MsoLineDashStyle.msoLineDashDotDot:          // 6: 二点鎖線
                    return MsoLineDashStyle.msoLineLongDashDot;   // → 8: 長鎖線

                case MsoLineDashStyle.msoLineLongDashDot:         // 8: 長鎖線
                    return MsoLineDashStyle.msoLineSolid;         // → 1: 実線に戻る

                case MsoLineDashStyle.msoLineDashStyleMixed:      // -2: サポートされていません
                default:
                    return MsoLineDashStyle.msoLineSolid;         // → 1: 実線から開始
            }
        }

        /// <summary>
        /// 線種名を取得（デバッグ用）
        /// </summary>
        /// <param name="dashStyle">線種</param>
        /// <returns>線種名</returns>
        private string GetDashStyleName(MsoLineDashStyle dashStyle)
        {
            switch (dashStyle)
            {
                case MsoLineDashStyle.msoLineSolid:              // 1
                    return "実線";
                case MsoLineDashStyle.msoLineSquareDot:          // 2
                    return "点線(角)";
                case MsoLineDashStyle.msoLineRoundDot:           // 3
                    return "点線(丸)";
                case MsoLineDashStyle.msoLineDash:               // 4
                    return "破線";
                case MsoLineDashStyle.msoLineDashDot:            // 5
                    return "一点鎖線";
                case MsoLineDashStyle.msoLineDashDotDot:         // 6
                    return "二点鎖線";
                case MsoLineDashStyle.msoLineLongDash:           // 7
                    return "長破線";
                case MsoLineDashStyle.msoLineLongDashDot:        // 8
                    return "長鎖線";
                case MsoLineDashStyle.msoLineDashStyleMixed:     // -2
                    return "混合(サポート外)";
                default:
                    return $"不明({(int)dashStyle})";
            }
        }

        #endregion

    }
}