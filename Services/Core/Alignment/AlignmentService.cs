using Microsoft.Office.Core;
using NLog;
using PowerPointEfficiencyAddin.Models;
using PowerPointEfficiencyAddin.Models.Licensing;
using PowerPointEfficiencyAddin.Services.Infrastructure.MultiInstance;
using PowerPointEfficiencyAddin.Services.Infrastructure.Licensing;
using PowerPointEfficiencyAddin.Utils;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PowerPointEfficiencyAddin.Services.Core.Alignment
{
    /// <summary>
    /// 整列・配置機能を提供するサービスクラス
    /// </summary>
    public class AlignmentService
    {
        private static readonly Logger logger = LogManager.GetCurrentClassLogger();
        private readonly IApplicationProvider applicationProvider;

        // DI対応コンストラクタ（商用レベル）
        public AlignmentService(IApplicationProvider applicationProvider)
        {
            this.applicationProvider = applicationProvider ?? throw new ArgumentNullException(nameof(applicationProvider));
            logger.Debug("AlignmentService initialized with DI application provider");
        }

        // 既存コンストラクタ（後方互換性維持）
        public AlignmentService() : this(new DefaultApplicationProvider())
        {
            logger.Debug("AlignmentService initialized with default application provider");
        }

        #region 伸縮グループ (7-10)

        /// <summary>
        /// 左端を揃える（7番機能）
        /// 選択した図形の左端を最初に選択した図形（基準図形）に伸縮して揃える
        /// </summary>
        public void AlignSizeLeft()
        {
            if (!Globals.ThisAddIn.CheckFeatureAccess("AlignSizeLeft")) return;

            logger.Info("AlignSizeLeft operation started");

            var selectedShapes = GetSelectedShapeInfos();
            if (!ValidateSelection(selectedShapes, 2, 0, "左端を揃える")) return;

            // 最初に選択した図形を基準として取得
            var referenceShape = selectedShapes.First();
            var targetShapes = selectedShapes.Skip(1).ToList();

            ComHelper.ExecuteWithComCleanup(() =>
            {
                foreach (var shapeInfo in targetShapes)
                {
                    try
                    {
                        // 右端位置を保持して、左端を基準図形の左端に合わせる（幅を調整）
                        float rightPosition = shapeInfo.Right;
                        shapeInfo.Shape.Left = referenceShape.Left;
                        shapeInfo.Shape.Width = rightPosition - referenceShape.Left;

                        logger.Debug($"Stretched left edge of {shapeInfo.Name} to {referenceShape.Left} (基準: {referenceShape.Name})");
                    }
                    catch (Exception ex)
                    {
                        logger.Error(ex, $"Failed to stretch left edge of shape {shapeInfo.Name}");
                    }
                }
            }, selectedShapes.Select(s => s.Shape).ToArray());

            logger.Info($"AlignSizeLeft completed for {targetShapes.Count} shapes (基準: 最初選択)");
        }

        /// <summary>
        /// 右端を揃える（8番機能）
        /// 選択した図形の右端を最初に選択した図形（基準図形）に伸縮して揃える
        /// </summary>
        public void AlignSizeRight()
        {
            if (!Globals.ThisAddIn.CheckFeatureAccess("AlignSizeRight")) return;

            logger.Info("AlignSizeRight operation started");

            var selectedShapes = GetSelectedShapeInfos();
            if (!ValidateSelection(selectedShapes, 2, 0, "右端を揃える")) return;

            // 最初に選択した図形を基準として取得
            var referenceShape = selectedShapes.First();
            var targetShapes = selectedShapes.Skip(1).ToList();

            ComHelper.ExecuteWithComCleanup(() =>
            {
                foreach (var shapeInfo in targetShapes)
                {
                    try
                    {
                        // 左端位置を保持して、右端を基準図形の右端に合わせる（幅を調整）
                        float leftPosition = shapeInfo.Left;
                        shapeInfo.Shape.Width = referenceShape.Right - leftPosition;

                        logger.Debug($"Stretched right edge of {shapeInfo.Name} to {referenceShape.Right} (基準: {referenceShape.Name})");
                    }
                    catch (Exception ex)
                    {
                        logger.Error(ex, $"Failed to stretch right edge of shape {shapeInfo.Name}");
                    }
                }
            }, selectedShapes.Select(s => s.Shape).ToArray());

            logger.Info($"AlignSizeRight completed for {targetShapes.Count} shapes (基準: 最初選択)");
        }

        /// <summary>
        /// 上端を揃える（9番機能）
        /// 選択した図形の上端を最初に選択した図形（基準図形）に伸縮して揃える
        /// </summary>
        public void AlignSizeTop()
        {
            if (!Globals.ThisAddIn.CheckFeatureAccess("AlignSizeTop")) return;

            logger.Info("AlignSizeTop operation started");

            var selectedShapes = GetSelectedShapeInfos();
            if (!ValidateSelection(selectedShapes, 2, 0, "上端を揃える")) return;

            // 最初に選択した図形を基準として取得
            var referenceShape = selectedShapes.First();
            var targetShapes = selectedShapes.Skip(1).ToList();

            ComHelper.ExecuteWithComCleanup(() =>
            {
                foreach (var shapeInfo in targetShapes)
                {
                    try
                    {
                        // 下端位置を保持して、上端を基準図形の上端に合わせる（高さを調整）
                        float bottomPosition = shapeInfo.Bottom;
                        shapeInfo.Shape.Top = referenceShape.Top;
                        shapeInfo.Shape.Height = bottomPosition - referenceShape.Top;

                        logger.Debug($"Stretched top edge of {shapeInfo.Name} to {referenceShape.Top} (基準: {referenceShape.Name})");
                    }
                    catch (Exception ex)
                    {
                        logger.Error(ex, $"Failed to stretch top edge of shape {shapeInfo.Name}");
                    }
                }
            }, selectedShapes.Select(s => s.Shape).ToArray());

            logger.Info($"AlignSizeTop completed for {targetShapes.Count} shapes (基準: 最初選択)");
        }

        /// <summary>
        /// 下端を揃える（10番機能）
        /// 選択した図形の下端を最初に選択した図形（基準図形）に伸縮して揃える
        /// </summary>
        public void AlignSizeBottom()
        {
            if (!Globals.ThisAddIn.CheckFeatureAccess("AlignSizeBottom")) return;

            logger.Info("AlignSizeBottom operation started");

            var selectedShapes = GetSelectedShapeInfos();
            if (!ValidateSelection(selectedShapes, 2, 0, "下端を揃える")) return;

            // 最初に選択した図形を基準として取得
            var referenceShape = selectedShapes.First();
            var targetShapes = selectedShapes.Skip(1).ToList();

            ComHelper.ExecuteWithComCleanup(() =>
            {
                foreach (var shapeInfo in targetShapes)
                {
                    try
                    {
                        // 上端位置を保持して、下端を基準図形の下端に合わせる（高さを調整）
                        float topPosition = shapeInfo.Top;
                        shapeInfo.Shape.Height = referenceShape.Bottom - topPosition;

                        logger.Debug($"Stretched bottom edge of {shapeInfo.Name} to {referenceShape.Bottom} (基準: {referenceShape.Name})");
                    }
                    catch (Exception ex)
                    {
                        logger.Error(ex, $"Failed to stretch bottom edge of shape {shapeInfo.Name}");
                    }
                }
            }, selectedShapes.Select(s => s.Shape).ToArray());

            logger.Info($"AlignSizeBottom completed for {targetShapes.Count} shapes (基準: 最初選択)");
        }

        #endregion

        #region 接着グループ (11-14)

        /// <summary>
        /// 右端を左端へ（11番機能）
        /// 2つの図形で、基準図形（最初に選択）以外の図形の右端を基準図形の左端に接着
        /// </summary>
        public void PlaceRightToLeft()
        {
            if (!Globals.ThisAddIn.CheckFeatureAccess("PlaceRightToLeft")) return;

            logger.Info("PlaceRightToLeft operation started");

            var selectedShapes = GetSelectedShapeInfos();
            if (!ValidateSelection(selectedShapes, 2, 2, "右端を左端へ")) return;

            // 最初に選択した図形を基準として取得
            var referenceShape = selectedShapes.First();
            var targetShape = selectedShapes.Skip(1).First();

            ComHelper.ExecuteWithComCleanup(() =>
            {
                try
                {
                    // 対象図形の右端を基準図形の左端に配置
                    targetShape.Shape.Left = referenceShape.Left - targetShape.Width;
                    logger.Debug($"Placed {targetShape.Name} right edge to {referenceShape.Name} left edge (基準: 最初選択)");
                }
                catch (Exception ex)
                {
                    logger.Error(ex, $"Failed to place right to left");
                }
            }, selectedShapes.Select(s => s.Shape).ToArray());

            logger.Info("PlaceRightToLeft completed (基準: 最初選択)");
        }

        /// <summary>
        /// 左端を右端へ（12番機能）
        /// 2つの図形で、基準図形（最初に選択）以外の図形の左端を基準図形の右端に接着
        /// </summary>
        public void PlaceLeftToRight()
        {
            if (!Globals.ThisAddIn.CheckFeatureAccess("PlaceLeftToRight")) return;

            logger.Info("PlaceLeftToRight operation started");

            var selectedShapes = GetSelectedShapeInfos();
            if (!ValidateSelection(selectedShapes, 2, 2, "左端を右端へ")) return;

            // 最初に選択した図形を基準として取得
            var referenceShape = selectedShapes.First();
            var targetShape = selectedShapes.Skip(1).First();

            ComHelper.ExecuteWithComCleanup(() =>
            {
                try
                {
                    // 対象図形の左端を基準図形の右端に配置
                    targetShape.Shape.Left = referenceShape.Right;
                    logger.Debug($"Placed {targetShape.Name} left edge to {referenceShape.Name} right edge (基準: 最初選択)");
                }
                catch (Exception ex)
                {
                    logger.Error(ex, $"Failed to place left to right");
                }
            }, selectedShapes.Select(s => s.Shape).ToArray());

            logger.Info("PlaceLeftToRight completed (基準: 最初選択)");
        }

        /// <summary>
        /// 上端を下端へ（13番機能）
        /// 2つの図形で、基準図形（最初に選択）以外の図形の上端を基準図形の下端に接着
        /// </summary>
        public void PlaceTopToBottom()
        {
            if (!Globals.ThisAddIn.CheckFeatureAccess("PlaceTopToBottom")) return;

            logger.Info("PlaceTopToBottom operation started");

            var selectedShapes = GetSelectedShapeInfos();
            if (!ValidateSelection(selectedShapes, 2, 2, "上端を下端へ")) return;

            // 最初に選択した図形を基準として取得
            var referenceShape = selectedShapes.First();
            var targetShape = selectedShapes.Skip(1).First();

            ComHelper.ExecuteWithComCleanup(() =>
            {
                try
                {
                    // 対象図形の上端を基準図形の下端に配置
                    targetShape.Shape.Top = referenceShape.Bottom;
                    logger.Debug($"Placed {targetShape.Name} top edge to {referenceShape.Name} bottom edge (基準: 最初選択)");
                }
                catch (Exception ex)
                {
                    logger.Error(ex, $"Failed to place top to bottom");
                }
            }, selectedShapes.Select(s => s.Shape).ToArray());

            logger.Info("PlaceTopToBottom completed (基準: 最初選択)");
        }

        /// <summary>
        /// 下端を上端へ（14番機能）
        /// 2つの図形で、基準図形（最初に選択）以外の図形の下端を基準図形の上端に接着
        /// </summary>
        public void PlaceBottomToTop()
        {
            if (!Globals.ThisAddIn.CheckFeatureAccess("PlaceBottomToTop")) return;

            logger.Info("PlaceBottomToTop operation started");

            var selectedShapes = GetSelectedShapeInfos();
            if (!ValidateSelection(selectedShapes, 2, 2, "下端を上端へ")) return;

            // 最初に選択した図形を基準として取得
            var referenceShape = selectedShapes.First();
            var targetShape = selectedShapes.Skip(1).First();

            ComHelper.ExecuteWithComCleanup(() =>
            {
                try
                {
                    // 対象図形の下端を基準図形の上端に配置
                    targetShape.Shape.Top = referenceShape.Top - targetShape.Height;
                    logger.Debug($"Placed {targetShape.Name} bottom edge to {referenceShape.Name} top edge (基準: 最初選択)");
                }
                catch (Exception ex)
                {
                    logger.Error(ex, $"Failed to place bottom to top");
                }
            }, selectedShapes.Select(s => s.Shape).ToArray());

            logger.Info("PlaceBottomToTop completed (基準: 最初選択)");
        }

        #endregion

        #region 水平垂直中央揃えグループ (15)

        /// <summary>
        /// 水平垂直中央揃え（15番機能）
        /// 選択した図形を水平・垂直中央に配置
        /// </summary>
        public void CenterAlign()
        {
            if (!Globals.ThisAddIn.CheckFeatureAccess("CenterAlign")) return;

            logger.Info("CenterAlign operation started");

            var selectedShapes = GetSelectedShapeInfos();
            if (!ValidateSelection(selectedShapes, 1, 0, "水平垂直中央揃え")) return;

            // スライドの中央位置を取得
            var slideCenter = GetSlideCenterPosition();
            if (slideCenter == null)
            {
                logger.Error("Failed to get slide center position");
                return;
            }

            ComHelper.ExecuteWithComCleanup(() =>
            {
                foreach (var shapeInfo in selectedShapes)
                {
                    try
                    {
                        // 水平中央に配置
                        shapeInfo.Shape.Left = slideCenter.Value.centerX - shapeInfo.Width / 2;

                        // 垂直中央に配置
                        shapeInfo.Shape.Top = slideCenter.Value.centerY - shapeInfo.Height / 2;

                        logger.Debug($"Centered {shapeInfo.Name} at ({slideCenter.Value.centerX}, {slideCenter.Value.centerY})");
                    }
                    catch (Exception ex)
                    {
                        logger.Error(ex, $"Failed to center align shape {shapeInfo.Name}");
                    }
                }
            }, selectedShapes.Select(s => s.Shape).ToArray());

            logger.Info($"CenterAlign completed for {selectedShapes.Count} shapes");
        }

        #endregion

        #region グループ化機能

        /// <summary>
        /// 行ごとにグループ化
        /// 選択したオブジェクトを行別にグループ化する
        /// </summary>
        public void GroupByRows()
        {
            if (!Globals.ThisAddIn.CheckFeatureAccess("GroupByRows")) return;

            logger.Info("GroupByRows operation started");

            var selectedShapes = GetSelectedShapeInfos();
            if (!ValidateSelection(selectedShapes, 2, 0, "行ごとにグループ化")) return;

            ComHelper.ExecuteWithComCleanup(() =>
            {
                var slide = GetCurrentSlide();
                if (slide == null)
                {
                    ErrorHandler.ExecuteSafely(() =>
                    {
                        throw new InvalidOperationException("アクティブなスライドが見つかりません。");
                    }, "行ごとにグループ化");
                    return;
                }

                // 行別にグループ化
                var rowGroups = GroupShapesByRows(selectedShapes);

                int groupIndex = 1;
                foreach (var rowGroup in rowGroups)
                {
                    try
                    {
                        if (rowGroup.Count > 1)
                        {
                            // 図形をグループ化
                            var shapeNames = rowGroup.Select(s => s.Shape.Name).ToArray();
                            var groupRange = slide.Shapes.Range(shapeNames);
                            var group = groupRange.Group();

                            // グループ名を設定
                            group.Name = $"行グループ{groupIndex}";

                            logger.Debug($"Created row group '{group.Name}' with {rowGroup.Count} shapes");
                            groupIndex++;
                        }
                    }
                    catch (Exception ex)
                    {
                        logger.Error(ex, $"Failed to create row group {groupIndex}");
                    }
                }

                logger.Info($"GroupByRows completed: created {groupIndex - 1} row groups");
            }, selectedShapes.Select(s => s.Shape).ToArray());

            logger.Info("GroupByRows completed");
        }

        /// <summary>
        /// 列ごとにグループ化
        /// 選択したオブジェクトを列別にグループ化する
        /// </summary>
        public void GroupByColumns()
        {
            if (!Globals.ThisAddIn.CheckFeatureAccess("GroupByColumns")) return;

            logger.Info("GroupByColumns operation started");

            var selectedShapes = GetSelectedShapeInfos();
            if (!ValidateSelection(selectedShapes, 2, 0, "列ごとにグループ化")) return;

            ComHelper.ExecuteWithComCleanup(() =>
            {
                var slide = GetCurrentSlide();
                if (slide == null)
                {
                    ErrorHandler.ExecuteSafely(() =>
                    {
                        throw new InvalidOperationException("アクティブなスライドが見つかりません。");
                    }, "列ごとにグループ化");
                    return;
                }

                // 列別にグループ化
                var columnGroups = GroupShapesByColumns(selectedShapes);

                int groupIndex = 1;
                foreach (var columnGroup in columnGroups)
                {
                    try
                    {
                        if (columnGroup.Count > 1)
                        {
                            // 図形をグループ化
                            var shapeNames = columnGroup.Select(s => s.Shape.Name).ToArray();
                            var groupRange = slide.Shapes.Range(shapeNames);
                            var group = groupRange.Group();

                            // グループ名を設定
                            group.Name = $"列グループ{groupIndex}";

                            logger.Debug($"Created column group '{group.Name}' with {columnGroup.Count} shapes");
                            groupIndex++;
                        }
                    }
                    catch (Exception ex)
                    {
                        logger.Error(ex, $"Failed to create column group {groupIndex}");
                    }
                }

                logger.Info($"GroupByColumns completed: created {groupIndex - 1} column groups");
            }, selectedShapes.Select(s => s.Shape).ToArray());

            logger.Info("GroupByColumns completed");
        }

        #endregion

        #region 間隔調整機能

        /// <summary>
        /// 間隔をなくす
        /// グリッド配置されたオブジェクトのすべての間隔を削除して密着させる
        /// </summary>
        public void RemoveSpacing()
        {
            if (!Globals.ThisAddIn.CheckFeatureAccess("RemoveSpacing")) return;

            logger.Info("RemoveSpacing operation started");

            var selectedShapes = GetSelectedShapeInfos();
            if (!ValidateSelection(selectedShapes, 2, 0, "間隔をなくす")) return;

            ComHelper.ExecuteWithComCleanup(() =>
            {
                // グリッド配置を検出して間隔を削除
                RemoveGridSpacing(selectedShapes);
                logger.Info($"RemoveSpacing completed for {selectedShapes.Count} shapes");
            }, selectedShapes.Select(s => s.Shape).ToArray());

            logger.Info("RemoveSpacing completed");
        }

        /// <summary>
        /// 垂直間隔調整
        /// 垂直に均等配置してから詳細調整ダイアログを表示
        /// </summary>
        public void AdjustVerticalSpacing()
        {
            if (!Globals.ThisAddIn.CheckFeatureAccess("AdjustVerticalSpacing")) return;

            logger.Info("AdjustVerticalSpacing operation started");

            var selectedShapes = GetSelectedShapeInfos();
            if (!ValidateSelection(selectedShapes, 2, 0, "垂直間隔調整")) return;

            ComHelper.ExecuteWithComCleanup(() =>
            {
                // まず垂直均等配置 + 垂直揃え（左端揃え）を実行
                var sortedShapes = selectedShapes.OrderBy(s => s.Top).ToList();
                var topShape = sortedShapes.First();
                var bottomShape = sortedShapes.Last();

                // 垂直揃え（左端を最も左の図形に合わせる）
                var leftmostX = sortedShapes.Min(s => s.Left);
                foreach (var shape in sortedShapes)
                {
                    shape.Shape.Left = leftmostX;
                }

                // 均等配置
                DistributeVertically(sortedShapes, topShape.Top, bottomShape.Bottom);

                // 現在の間隔を計算
                var currentSpacing = CalculateAverageVerticalSpacing(sortedShapes);

                // 間隔調整ダイアログを表示
                var spacingSettings = ShowVerticalSpacingDialog(currentSpacing, sortedShapes);
                if (spacingSettings != null)
                {
                    ApplyVerticalSpacing(sortedShapes, spacingSettings, topShape.Top, bottomShape.Bottom);
                    logger.Info($"Applied vertical spacing: {spacingSettings.Spacing}cm, method: {spacingSettings.AdjustmentMethod}");
                }

                logger.Info("AdjustVerticalSpacing completed");
            }, selectedShapes.Select(s => s.Shape).ToArray());

            logger.Info("AdjustVerticalSpacing completed");
        }

        /// <summary>
        /// 水平間隔調整
        /// 水平に均等配置してから詳細調整ダイアログを表示
        /// </summary>
        public void AdjustHorizontalSpacing()
        {
            if (!Globals.ThisAddIn.CheckFeatureAccess("AdjustHorizontalSpacing")) return;

            logger.Info("AdjustHorizontalSpacing operation started");

            var selectedShapes = GetSelectedShapeInfos();
            if (!ValidateSelection(selectedShapes, 2, 0, "水平間隔調整")) return;

            ComHelper.ExecuteWithComCleanup(() =>
            {
                // まず水平均等配置 + 水平揃え（上端揃え）を実行
                var sortedShapes = selectedShapes.OrderBy(s => s.Left).ToList();
                var leftShape = sortedShapes.First();
                var rightShape = sortedShapes.Last();

                // 水平揃え（上端を最も上の図形に合わせる）
                var topmostY = sortedShapes.Min(s => s.Top);
                foreach (var shape in sortedShapes)
                {
                    shape.Shape.Top = topmostY;
                }

                // 均等配置
                DistributeHorizontally(sortedShapes, leftShape.Left, rightShape.Right);

                // 現在の間隔を計算
                var currentSpacing = CalculateAverageHorizontalSpacing(sortedShapes);

                // 間隔調整ダイアログを表示
                var spacingSettings = ShowHorizontalSpacingDialog(currentSpacing, sortedShapes);
                if (spacingSettings != null)
                {
                    ApplyHorizontalSpacing(sortedShapes, spacingSettings, leftShape.Left, rightShape.Right);
                    logger.Info($"Applied horizontal spacing: {spacingSettings.Spacing}cm, method: {spacingSettings.AdjustmentMethod}");
                }

                logger.Info("AdjustHorizontalSpacing completed");
            }, selectedShapes.Select(s => s.Shape).ToArray());

            logger.Info("AdjustHorizontalSpacing completed");
        }

        #endregion

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
        /// スライドの中央位置を取得します
        /// </summary>
        /// <returns>スライド中央位置（centerX, centerY）</returns>
        private (float centerX, float centerY)? GetSlideCenterPosition()
        {
            try
            {
                // 新実装（DI経由）
                var application = applicationProvider.GetCurrentApplication();
                var activePresentation = GetActivePresentationFromApplication(application);

                if (activePresentation != null)
                {
                    var activeWindow = application.ActiveWindow;
                    if (activeWindow.ViewType == PowerPoint.PpViewType.ppViewSlide ||
                        activeWindow.ViewType == PowerPoint.PpViewType.ppViewNormal)
                    {
                        var slide = activeWindow.View.Slide;
                        var slideWidth = activePresentation.PageSetup.SlideWidth;
                        var slideHeight = activePresentation.PageSetup.SlideHeight;

                        return (slideWidth / 2, slideHeight / 2);
                    }
                }
            }
            catch (Exception ex)
            {
                logger.Error(ex, "Failed to get slide center position");
            }

            return null;
        }

        // アプリケーション固有のプレゼンテーション取得
        private PowerPoint.Presentation GetActivePresentationFromApplication(PowerPoint.Application application)
        {
            try
            {
                if (application.Presentations.Count > 0)
                {
                    return application.ActivePresentation;
                }
            }
            catch (Exception ex)
            {
                logger.Error(ex, "Failed to get active presentation from application");
            }

            return null;
        }

        /// <summary>
        /// 現在のスライドを取得します
        /// </summary>
        /// <returns>アクティブなスライド</returns>
        private PowerPoint.Slide GetCurrentSlide()
        {
            try
            {
                // ✅ 新実装（DI経由）
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

        #endregion

        #region グループ化ヘルパーメソッド

        /// <summary>
        /// 図形を行別にグループ化します
        /// </summary>
        /// <param name="shapes">図形リスト</param>
        /// <param name="tolerance">Y座標の許容誤差</param>
        /// <returns>行別グループのリスト</returns>
        private List<List<ShapeInfo>> GroupShapesByRows(List<ShapeInfo> shapes)
        {
            var tolerance = CalculateDynamicTolerance(shapes, true); // Y座標用
            var rows = new List<List<ShapeInfo>>();

            foreach (var shape in shapes.OrderBy(s => s.Top))
            {
                var assignedToRow = false;
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
                if (!assignedToRow)
                {
                    rows.Add(new List<ShapeInfo> { shape });
                }
            }

            foreach (var row in rows)
            {
                row.Sort((a, b) => a.Left.CompareTo(b.Left));
            }

            return rows;
        }

        /// <summary>
        /// 図形を列別にグループ化します
        /// </summary>
        /// <param name="shapes">図形リスト</param>
        /// <param name="tolerance">X座標の許容誤差</param>
        /// <returns>列別グループのリスト</returns>
        private List<List<ShapeInfo>> GroupShapesByColumns(List<ShapeInfo> shapes)
        {
            var tolerance = CalculateDynamicTolerance(shapes, false); // X座標用
            var columns = new List<List<ShapeInfo>>();

            foreach (var shape in shapes.OrderBy(s => s.Left))
            {
                var assignedToColumn = false;
                foreach (var column in columns)
                {
                    var avgX = column.Average(s => s.Left);
                    if (Math.Abs(shape.Left - avgX) <= tolerance)
                    {
                        column.Add(shape);
                        assignedToColumn = true;
                        break;
                    }
                }
                if (!assignedToColumn)
                {
                    columns.Add(new List<ShapeInfo> { shape });
                }
            }

            foreach (var column in columns)
            {
                column.Sort((a, b) => a.Top.CompareTo(b.Top));
            }

            return columns;
        }

        /// <summary>
        /// 動的な許容誤差を計算します（AlignmentService用）
        /// </summary>
        /// <param name="shapes">図形リスト</param>
        /// <param name="isVertical">true=Y座標用（高さベース）、false=X座標用（幅ベース）</param>
        /// <returns>計算された許容誤差（ポイント）</returns>
        private float CalculateDynamicTolerance(List<ShapeInfo> shapes, bool isVertical)
        {
            if (!shapes.Any()) return 10f;

            var averageSize = isVertical
                ? shapes.Average(s => s.Height)
                : shapes.Average(s => s.Width);

            var calculatedTolerance = averageSize * 0.3f;
            const float MIN_TOLERANCE = 3f;
            const float MAX_TOLERANCE = 25f;

            return Math.Max(MIN_TOLERANCE, Math.Min(MAX_TOLERANCE, calculatedTolerance));
        }

        #endregion

        #region 間隔調整ヘルパーメソッド

        /// <summary>
        /// グリッド配置の間隔を削除します
        /// </summary>
        /// <param name="shapes">図形リスト</param>
        private void RemoveGridSpacing(List<ShapeInfo> shapes)
        {
            try
            {
                // 図形を行別にグループ化
                var rowGroups = GroupShapesByRows(shapes);

                if (rowGroups.Count == 0) return;

                // 最も左上の図形を基準点として設定
                var topLeftShape = shapes.OrderBy(s => s.Top).ThenBy(s => s.Left).First();
                var baseX = topLeftShape.Left;
                var baseY = topLeftShape.Top;

                var currentY = baseY;

                // 各行を処理
                foreach (var row in rowGroups.OrderBy(r => r.Min(s => s.Top)))
                {
                    if (row.Count == 0) continue;

                    // 行内の図形を左から順にソート
                    var sortedRowShapes = row.OrderBy(s => s.Left).ToList();

                    // 行内の図形を水平方向で密着配置
                    var currentX = baseX;
                    var rowHeight = 0f;

                    foreach (var shape in sortedRowShapes)
                    {
                        // 図形を配置
                        shape.Shape.Left = currentX;
                        shape.Shape.Top = currentY;

                        // 次の位置を計算
                        currentX += shape.Width;
                        rowHeight = Math.Max(rowHeight, shape.Height);

                        logger.Debug($"Positioned {shape.Name} at ({currentX - shape.Width}, {currentY})");
                    }

                    // 次の行の開始Y座標を設定
                    currentY += rowHeight;

                    logger.Debug($"Completed row with height {rowHeight}, next row starts at Y={currentY}");
                }

                logger.Info($"Removed grid spacing for {rowGroups.Count} rows with total {shapes.Count} shapes");
            }
            catch (Exception ex)
            {
                logger.Error(ex, "Failed to remove grid spacing");

                // フォールバック：従来の方法を試行
                var isHorizontalLayout = IsHorizontalLayout(shapes);
                var isVerticalLayout = IsVerticalLayout(shapes);

                if (isHorizontalLayout)
                {
                    RemoveHorizontalSpacing(shapes);
                    logger.Debug("Fallback: Removed horizontal spacing");
                }
                else if (isVerticalLayout)
                {
                    RemoveVerticalSpacing(shapes);
                    logger.Debug("Fallback: Removed vertical spacing");
                }
                else
                {
                    RemoveNearestSpacing(shapes);
                    logger.Debug("Fallback: Removed nearest spacing");
                }
            }
        }

        /// <summary>
        /// 水平レイアウトかどうかを判定します
        /// </summary>
        private bool IsHorizontalLayout(List<ShapeInfo> shapes)
        {
            if (shapes.Count < 2) return false;

            var sortedByLeft = shapes.OrderBy(s => s.Left).ToList();
            var maxTopDiff = 0f;

            for (int i = 0; i < sortedByLeft.Count - 1; i++)
            {
                var topDiff = Math.Abs(sortedByLeft[i + 1].Top - sortedByLeft[i].Top);
                maxTopDiff = Math.Max(maxTopDiff, topDiff);
            }

            return maxTopDiff <= 40f; // 40pt以内なら水平配置とみなす
        }

        /// <summary>
        /// 垂直レイアウトかどうかを判定します
        /// </summary>
        private bool IsVerticalLayout(List<ShapeInfo> shapes)
        {
            if (shapes.Count < 2) return false;

            var sortedByTop = shapes.OrderBy(s => s.Top).ToList();
            var maxLeftDiff = 0f;

            for (int i = 0; i < sortedByTop.Count - 1; i++)
            {
                var leftDiff = Math.Abs(sortedByTop[i + 1].Left - sortedByTop[i].Left);
                maxLeftDiff = Math.Max(maxLeftDiff, leftDiff);
            }

            return maxLeftDiff <= 40f; // 40pt以内なら垂直配置とみなす
        }

        /// <summary>
        /// 水平方向の間隔を削除します
        /// </summary>
        private void RemoveHorizontalSpacing(List<ShapeInfo> shapes)
        {
            var sortedShapes = shapes.OrderBy(s => s.Left).ToList();

            for (int i = 1; i < sortedShapes.Count; i++)
            {
                var prevShape = sortedShapes[i - 1];
                var currentShape = sortedShapes[i];

                // 前の図形の右端に隣接させる
                currentShape.Shape.Left = prevShape.Right;
            }
        }

        /// <summary>
        /// 垂直方向の間隔を削除します
        /// </summary>
        private void RemoveVerticalSpacing(List<ShapeInfo> shapes)
        {
            var sortedShapes = shapes.OrderBy(s => s.Top).ToList();

            for (int i = 1; i < sortedShapes.Count; i++)
            {
                var prevShape = sortedShapes[i - 1];
                var currentShape = sortedShapes[i];

                // 前の図形の下端に隣接させる
                currentShape.Shape.Top = prevShape.Bottom;
            }
        }

        /// <summary>
        /// 散在する図形の最近接間隔を削除します
        /// </summary>
        private void RemoveNearestSpacing(List<ShapeInfo> shapes)
        {
            // 最も近い図形ペアを見つけて隣接させる
            for (int i = 0; i < shapes.Count; i++)
            {
                var currentShape = shapes[i];
                var nearestShape = FindNearestShape(currentShape, shapes.Where((s, idx) => idx != i).ToList());

                if (nearestShape != null)
                {
                    // 最も近い方向に隣接させる
                    var horizontalDistance = Math.Abs(currentShape.CenterX - nearestShape.CenterX);
                    var verticalDistance = Math.Abs(currentShape.CenterY - nearestShape.CenterY);

                    if (horizontalDistance < verticalDistance)
                    {
                        // 水平方向に隣接
                        if (currentShape.CenterX < nearestShape.CenterX)
                        {
                            currentShape.Shape.Left = nearestShape.Left - currentShape.Width;
                        }
                        else
                        {
                            currentShape.Shape.Left = nearestShape.Right;
                        }
                    }
                    else
                    {
                        // 垂直方向に隣接
                        if (currentShape.CenterY < nearestShape.CenterY)
                        {
                            currentShape.Shape.Top = nearestShape.Top - currentShape.Height;
                        }
                        else
                        {
                            currentShape.Shape.Top = nearestShape.Bottom;
                        }
                    }
                }
            }
        }

        /// <summary>
        /// 最も近い図形を見つけます
        /// </summary>
        private ShapeInfo FindNearestShape(ShapeInfo targetShape, List<ShapeInfo> otherShapes)
        {
            if (!otherShapes.Any()) return null;

            return otherShapes.OrderBy(s =>
            {
                var dx = targetShape.CenterX - s.CenterX;
                var dy = targetShape.CenterY - s.CenterY;
                return Math.Sqrt(dx * dx + dy * dy);
            }).First();
        }

        /// <summary>
        /// 垂直均等配置を実行します
        /// </summary>
        private void DistributeVertically(List<ShapeInfo> shapes, float topBound, float bottomBound)
        {
            if (shapes.Count < 2) return;

            var totalHeight = shapes.Sum(s => s.Height);
            var availableSpace = bottomBound - topBound - totalHeight;
            var spacing = availableSpace / (shapes.Count - 1);

            var currentTop = topBound;
            foreach (var shape in shapes)
            {
                shape.Shape.Top = currentTop;
                currentTop += shape.Height + spacing;
            }
        }

        /// <summary>
        /// 水平均等配置を実行します
        /// </summary>
        private void DistributeHorizontally(List<ShapeInfo> shapes, float leftBound, float rightBound)
        {
            if (shapes.Count < 2) return;

            var totalWidth = shapes.Sum(s => s.Width);
            var availableSpace = rightBound - leftBound - totalWidth;
            var spacing = availableSpace / (shapes.Count - 1);

            var currentLeft = leftBound;
            foreach (var shape in shapes)
            {
                shape.Shape.Left = currentLeft;
                currentLeft += shape.Width + spacing;
            }
        }

        /// <summary>
        /// 垂直方向の平均間隔を計算します
        /// </summary>
        private float CalculateAverageVerticalSpacing(List<ShapeInfo> shapes)
        {
            if (shapes.Count < 2) return 0f;

            var spacings = new List<float>();
            for (int i = 0; i < shapes.Count - 1; i++)
            {
                var spacing = shapes[i + 1].Top - shapes[i].Bottom;
                spacings.Add(spacing);
            }

            return spacings.Average() / 28.35f; // ポイントからcmに変換
        }

        /// <summary>
        /// 水平方向の平均間隔を計算します
        /// </summary>
        private float CalculateAverageHorizontalSpacing(List<ShapeInfo> shapes)
        {
            if (shapes.Count < 2) return 0f;

            var spacings = new List<float>();
            for (int i = 0; i < shapes.Count - 1; i++)
            {
                var spacing = shapes[i + 1].Left - shapes[i].Right;
                spacings.Add(spacing);
            }

            return spacings.Average() / 28.35f; // ポイントからcmに変換
        }

        /// <summary>
        /// 垂直間隔調整ダイアログを表示します
        /// </summary>
        private SpacingSettings ShowVerticalSpacingDialog(float currentSpacing, List<ShapeInfo> shapes)
        {
            using (var dialog = new UI.Dialogs.SpacingAdjustmentDialog("垂直間隔調整", currentSpacing, shapes, true))
            {
                if (dialog.ShowDialog() == DialogResult.OK)
                {
                    return dialog.GetSettings();
                }
            }
            return null;
        }

        /// <summary>
        /// 水平間隔調整ダイアログを表示します
        /// </summary>
        private SpacingSettings ShowHorizontalSpacingDialog(float currentSpacing, List<ShapeInfo> shapes)
        {
            using (var dialog = new UI.Dialogs.SpacingAdjustmentDialog("水平間隔調整", currentSpacing, shapes, false))
            {
                if (dialog.ShowDialog() == DialogResult.OK)
                {
                    return dialog.GetSettings();
                }
            }
            return null;
        }

        /// <summary>
        /// 垂直間隔を適用します
        /// </summary>
        private void ApplyVerticalSpacing(List<ShapeInfo> shapes, SpacingSettings settings, float topBound, float bottomBound)
        {
            var spacingPoints = settings.Spacing * 28.35f; // cmをポイントに変換

            // 垂直揃え（左端揃え）を維持
            var leftmostX = shapes.Min(s => s.Left);

            if (settings.AdjustmentMethod == SpacingAdjustmentMethod.MoveObjects)
            {
                // オブジェクトを移動して間隔調整
                var currentTop = topBound;
                foreach (var shape in shapes)
                {
                    shape.Shape.Left = leftmostX; // 左端揃えを維持
                    shape.Shape.Top = currentTop;
                    currentTop += shape.Height + spacingPoints;
                }
            }
            else
            {
                // オブジェクトサイズを変更して間隔調整
                var totalSpacing = spacingPoints * (shapes.Count - 1);
                var availableSpace = bottomBound - topBound - totalSpacing;
                var newHeight = availableSpace / shapes.Count;

                var currentTop = topBound;
                foreach (var shape in shapes)
                {
                    shape.Shape.Left = leftmostX; // 左端揃えを維持
                    shape.Shape.Top = currentTop;
                    shape.Shape.Height = newHeight;
                    currentTop += newHeight + spacingPoints;
                }
            }
        }

        /// <summary>
        /// 水平間隔を適用します
        /// </summary>
        private void ApplyHorizontalSpacing(List<ShapeInfo> shapes, SpacingSettings settings, float leftBound, float rightBound)
        {
            var spacingPoints = settings.Spacing * 28.35f; // cmをポイントに変換

            // 水平揃え（上端揃え）を維持
            var topmostY = shapes.Min(s => s.Top);

            if (settings.AdjustmentMethod == SpacingAdjustmentMethod.MoveObjects)
            {
                // オブジェクトを移動して間隔調整
                var currentLeft = leftBound;
                foreach (var shape in shapes)
                {
                    shape.Shape.Top = topmostY; // 上端揃えを維持
                    shape.Shape.Left = currentLeft;
                    currentLeft += shape.Width + spacingPoints;
                }
            }
            else
            {
                // オブジェクトサイズを変更して間隔調整
                var totalSpacing = spacingPoints * (shapes.Count - 1);
                var availableSpace = rightBound - leftBound - totalSpacing;
                var newWidth = availableSpace / shapes.Count;

                var currentLeft = leftBound;
                foreach (var shape in shapes)
                {
                    shape.Shape.Top = topmostY; // 上端揃えを維持
                    shape.Shape.Left = currentLeft;
                    shape.Shape.Width = newWidth;
                    currentLeft += newWidth + spacingPoints;
                }
            }
        }

        #endregion

    }

    /// <summary>
    /// 配置方向を表す列挙型
    /// </summary>
    public enum DistributionDirection
    {
        /// <summary>水平方向</summary>
        Horizontal,
        /// <summary>垂直方向</summary>
        Vertical
    }
}