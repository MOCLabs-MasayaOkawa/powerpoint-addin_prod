using Microsoft.Office.Core;
using NLog;
using PowerPointEfficiencyAddin.Services.Infrastructure.MultiInstance;
using PowerPointEfficiencyAddin.Utils;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PowerPointEfficiencyAddin.Services.Core.Selection
{
    /// <summary>
    /// 図形選択および透過率調整機能を提供するサービスクラス
    /// </summary>
    public class ShapeSelectionService
    {
        private static readonly Logger logger = LogManager.GetCurrentClassLogger();
        private readonly IApplicationProvider applicationProvider;

        /// <summary>
        /// DI対応コンストラクタ
        /// </summary>
        /// <param name="applicationProvider">アプリケーションプロバイダー</param>
        public ShapeSelectionService(IApplicationProvider applicationProvider)
        {
            this.applicationProvider = applicationProvider ?? throw new ArgumentNullException(nameof(applicationProvider));
            logger.Debug("ShapeSelectionService initialized with DI application provider");
        }

        #region Public Methods

        /// <summary>
        /// 同色図形の一括選択
        /// 同スライド内の同じ塗りつぶし色の図形を一括選択する
        /// </summary>
        public void SelectSameColorShapes()
        {
            if (!Globals.ThisAddIn.CheckFeatureAccess("SelectSameColorShapes")) return;

            logger.Info("SelectSameColorShapes operation started");

            var selectedShapes = GetSelectedShapeInfos();
            if (!ValidateSelection(selectedShapes, 1, 1, "同色図形の一括選択")) return;

            var referenceShape = selectedShapes.First();

            ComHelper.ExecuteWithComCleanup(() =>
            {
                var slide = GetCurrentSlide();
                if (slide == null)
                {
                    ErrorHandler.ExecuteSafely(() =>
                    {
                        throw new InvalidOperationException("アクティブなスライドが見つかりません。");
                    }, "同色図形の一括選択");
                    return;
                }

                try
                {
                    // 基準図形の塗りつぶし色を取得
                    var referenceColor = GetShapeFillColor(referenceShape.Shape);
                    if (referenceColor == null)
                    {
                        MessageBox.Show(
                            "選択した図形に塗りつぶし色が設定されていません。",
                            "同色図形の一括選択",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Warning
                        );
                        return;
                    }

                    // 同じ色の図形を検索
                    var sameColorShapes = new List<PowerPoint.Shape>();

                    for (int i = 1; i <= slide.Shapes.Count; i++)
                    {
                        var shape = slide.Shapes[i];
                        try
                        {
                            var shapeColor = GetShapeFillColor(shape);
                            if (shapeColor.HasValue && shapeColor.Value == referenceColor.Value)
                            {
                                sameColorShapes.Add(shape);
                            }
                        }
                        catch (Exception ex)
                        {
                            logger.Warn(ex, $"Failed to check color for shape {shape.Name}");
                        }
                    }

                    if (sameColorShapes.Count == 0)
                    {
                        MessageBox.Show(
                            "同じ色の図形が見つかりませんでした。",
                            "同色図形の一括選択",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Information
                        );
                        return;
                    }

                    // 図形を選択
                    SelectShapes(sameColorShapes);

                    logger.Info($"Selected {sameColorShapes.Count} shapes with same color");
                }
                catch (Exception ex)
                {
                    logger.Error(ex, "Failed to select same color shapes");
                    ErrorHandler.ExecuteSafely(() => throw ex, "同色図形の一括選択");
                }
            });

            logger.Info("SelectSameColorShapes completed");
        }

        /// <summary>
        /// 同サイズ図形の一括選択
        /// 同スライド内の同じサイズ（幅・高さ完全一致）の図形を一括選択する
        /// </summary>
        public void SelectSameSizeShapes()
        {
            if (!Globals.ThisAddIn.CheckFeatureAccess("SelectSameSizeShapes")) return;

            logger.Info("SelectSameSizeShapes operation started");

            var selectedShapes = GetSelectedShapeInfos();
            if (!ValidateSelection(selectedShapes, 1, 1, "同サイズ図形の一括選択")) return;

            var referenceShape = selectedShapes.First();

            ComHelper.ExecuteWithComCleanup(() =>
            {
                var slide = GetCurrentSlide();
                if (slide == null)
                {
                    ErrorHandler.ExecuteSafely(() =>
                    {
                        throw new InvalidOperationException("アクティブなスライドが見つかりません。");
                    }, "同サイズ図形の一括選択");
                    return;
                }

                try
                {
                    // 基準図形のサイズを取得
                    var referenceWidth = referenceShape.Width;
                    var referenceHeight = referenceShape.Height;

                    // 同じサイズの図形を検索
                    var sameSizeShapes = new List<PowerPoint.Shape>();

                    for (int i = 1; i <= slide.Shapes.Count; i++)
                    {
                        var shape = slide.Shapes[i];
                        try
                        {
                            if (Math.Abs(shape.Width - referenceWidth) < 0.01f &&
                                Math.Abs(shape.Height - referenceHeight) < 0.01f)
                            {
                                sameSizeShapes.Add(shape);
                            }
                        }
                        catch (Exception ex)
                        {
                            logger.Warn(ex, $"Failed to check size for shape {shape.Name}");
                        }
                    }

                    if (sameSizeShapes.Count == 0)
                    {
                        MessageBox.Show(
                            "同じサイズの図形が見つかりませんでした。",
                            "同サイズ図形の一括選択",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Information
                        );
                        return;
                    }

                    // 図形を選択
                    SelectShapes(sameSizeShapes);

                    logger.Info($"Selected {sameSizeShapes.Count} shapes with same size ({referenceWidth}x{referenceHeight})");
                }
                catch (Exception ex)
                {
                    logger.Error(ex, "Failed to select same size shapes");
                    ErrorHandler.ExecuteSafely(() => throw ex, "同サイズ図形の一括選択");
                }
            });

            logger.Info("SelectSameSizeShapes completed");
        }

        /// <summary>
        /// 図形背景の透過率Upトグル
        /// 選択した図形の透過率を10%ずつあげていく
        /// </summary>
        public void TransparencyUpToggle()
        {
            if (!Globals.ThisAddIn.CheckFeatureAccess("TransparencyUpToggle")) return;

            logger.Info("TransparencyUpToggle operation started");

            var selectedShapes = GetSelectedShapeInfos();
            if (!ValidateSelection(selectedShapes, 1, 0, "図形背景の透過率Up")) return;

            ComHelper.ExecuteWithComCleanup(() =>
            {
                int adjustedCount = 0;

                foreach (var shapeInfo in selectedShapes)
                {
                    try
                    {
                        var shape = shapeInfo.Shape;

                        // 塗りつぶしの透過率を調整
                        float currentFillTransparency = shape.Fill.Transparency;
                        float newFillTransparency = currentFillTransparency + 0.1f; // 10%アップ

                        if (newFillTransparency <= 1.0f) // 100%以下の場合のみ適用
                        {
                            shape.Fill.Transparency = newFillTransparency;
                            adjustedCount++;

                            logger.Debug($"Adjusted fill transparency for {shape.Name}: {currentFillTransparency:P0} -> {newFillTransparency:P0}");
                        }

                        // 線の透過率も同時に調整
                        if (shape.Line.Visible == MsoTriState.msoTrue)
                        {
                            float currentLineTransparency = shape.Line.Transparency;
                            float newLineTransparency = currentLineTransparency + 0.1f; // 10%アップ

                            if (newLineTransparency <= 1.0f) // 100%以下の場合のみ適用
                            {
                                shape.Line.Transparency = newLineTransparency;
                                logger.Debug($"Adjusted line transparency for {shape.Name}: {currentLineTransparency:P0} -> {newLineTransparency:P0}");
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        logger.Error(ex, $"Failed to adjust transparency for shape {shapeInfo.Name}");
                    }
                }

                logger.Info($"TransparencyUpToggle completed for {adjustedCount}/{selectedShapes.Count} shapes");

            }, selectedShapes.Select(s => s.Shape).ToArray());

            logger.Info("TransparencyUpToggle completed");
        }

        /// <summary>
        /// 図形背景の透過率Downトグル
        /// 選択した図形の透過率を10%ずつ下げていく
        /// </summary>
        public void TransparencyDownToggle()
        {
            if (!Globals.ThisAddIn.CheckFeatureAccess("TransparencyDownToggle")) return;

            logger.Info("TransparencyDownToggle operation started");

            var selectedShapes = GetSelectedShapeInfos();
            if (!ValidateSelection(selectedShapes, 1, 0, "図形背景の透過率Down")) return;

            ComHelper.ExecuteWithComCleanup(() =>
            {
                int adjustedCount = 0;

                foreach (var shapeInfo in selectedShapes)
                {
                    try
                    {
                        var shape = shapeInfo.Shape;

                        // 塗りつぶしの透過率を調整
                        float currentFillTransparency = shape.Fill.Transparency;
                        float newFillTransparency = currentFillTransparency - 0.1f; // 10%ダウン

                        if (newFillTransparency >= 0.0f) // 0%以上の場合のみ適用
                        {
                            shape.Fill.Transparency = newFillTransparency;
                            adjustedCount++;

                            logger.Debug($"Adjusted fill transparency for {shape.Name}: {currentFillTransparency:P0} -> {newFillTransparency:P0}");
                        }

                        // 線の透過率も同時に調整
                        if (shape.Line.Visible == MsoTriState.msoTrue)
                        {
                            float currentLineTransparency = shape.Line.Transparency;
                            float newLineTransparency = currentLineTransparency - 0.1f; // 10%ダウン

                            if (newLineTransparency >= 0.0f) // 0%以上の場合のみ適用
                            {
                                shape.Line.Transparency = newLineTransparency;
                                logger.Debug($"Adjusted line transparency for {shape.Name}: {currentLineTransparency:P0} -> {newLineTransparency:P0}");
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        logger.Error(ex, $"Failed to adjust transparency for shape {shapeInfo.Name}");
                    }
                }

                logger.Info($"TransparencyDownToggle completed for {adjustedCount}/{selectedShapes.Count} shapes");

            }, selectedShapes.Select(s => s.Shape).ToArray());

            logger.Info("TransparencyDownToggle completed");
        }

        #endregion

        #region Private Helper Methods

        /// <summary>
        /// 選択されている図形の情報を取得します
        /// </summary>
        private List<Models.ShapeInfo> GetSelectedShapeInfos()
        {
            var shapeInfos = new List<Models.ShapeInfo>();

            try
            {
                var application = applicationProvider.GetCurrentApplication();
                var activeWindow = application.ActiveWindow;

                if (activeWindow?.Selection == null)
                {
                    logger.Debug("No active window or selection");
                    return shapeInfos;
                }

                var selection = activeWindow.Selection;
                logger.Debug($"Selection type: {selection.Type}");

                switch (selection.Type)
                {
                    case PowerPoint.PpSelectionType.ppSelectionShapes:
                        var normalShapeRange = selection.ShapeRange;
                        if (normalShapeRange != null)
                        {
                            for (int i = 1; i <= normalShapeRange.Count; i++)
                            {
                                var shape = normalShapeRange[i];
                                shapeInfos.Add(new Models.ShapeInfo(shape, i - 1));
                            }
                        }
                        break;
                }

                logger.Debug($"Retrieved {shapeInfos.Count} shape(s)");
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
        private bool ValidateSelection(List<Models.ShapeInfo> shapeInfos, int minRequired, int maxAllowed, string operationName)
        {
            return ErrorHandler.ValidateSelection(shapeInfos.Count, minRequired, maxAllowed, operationName);
        }

        /// <summary>
        /// 現在のスライドを取得します
        /// </summary>
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
        /// 図形の塗りつぶし色を取得します
        /// </summary>
        /// <param name="shape">対象図形</param>
        /// <returns>RGB色値、取得できない場合はnull</returns>
        private int? GetShapeFillColor(PowerPoint.Shape shape)
        {
            try
            {
                if (shape.Fill.Type == MsoFillType.msoFillSolid)
                {
                    return shape.Fill.ForeColor.RGB;
                }
            }
            catch (Exception ex)
            {
                logger.Warn(ex, $"Failed to get fill color for shape {shape.Name}");
            }
            return null;
        }

        /// <summary>
        /// 図形を選択します
        /// </summary>
        private void SelectShapes(List<PowerPoint.Shape> shapes)
        {
            try
            {
                var application = applicationProvider.GetCurrentApplication();

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
    }
}
