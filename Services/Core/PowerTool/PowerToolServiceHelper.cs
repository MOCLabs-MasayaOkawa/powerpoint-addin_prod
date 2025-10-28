using Microsoft.Office.Core;
using NLog;
using PowerPointEfficiencyAddin.Models;
using PowerPointEfficiencyAddin.Services.Infrastructure.MultiInstance;
using PowerPointEfficiencyAddin.Utils;
using System;
using System.Collections.Generic;
using System.Linq;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PowerPointEfficiencyAddin.Services.Core.PowerTool
{
    /// <summary>
    /// PowerToolServiceの共通ヘルパーメソッドを提供するクラス
    /// 図形選択、検証、グリッド検出などの汎用処理を集約
    /// </summary>
    public class PowerToolServiceHelper
    {
        private static readonly Logger logger = LogManager.GetCurrentClassLogger();
        private readonly IApplicationProvider applicationProvider;

        /// <summary>
        /// DI対応コンストラクタ
        /// </summary>
        /// <param name="applicationProvider">アプリケーションプロバイダー</param>
        public PowerToolServiceHelper(IApplicationProvider applicationProvider)
        {
            this.applicationProvider = applicationProvider ?? throw new ArgumentNullException(nameof(applicationProvider));
            logger.Debug("PowerToolServiceHelper initialized with DI application provider");
        }

        #region 図形選択・取得の共通処理

        /// <summary>
        /// 選択されている図形の情報を取得します（テキスト編集モード対応）
        /// </summary>
        /// <returns>選択図形情報のリスト</returns>
        public List<ShapeInfo> GetSelectedShapeInfos()
        {
            var shapeInfos = new List<ShapeInfo>();

            try
            {
                var application = applicationProvider.GetCurrentApplication();
                var selectedShapes = GetSelectedShapesFromApplication(application);

                var activeWindow = application.ActiveWindow;

                if (activeWindow?.Selection == null)
                {
                    logger.Debug("No active window or selection");
                    return shapeInfos;
                }

                var selection = activeWindow.Selection;

                // 選択状態の種類を確認
                logger.Debug($"Selection type: {selection.Type}");

                switch (selection.Type)
                {
                    case PowerPoint.PpSelectionType.ppSelectionShapes:
                        // 通常の図形選択状態
                        logger.Debug("Normal shape selection mode");
                        var normalShapeRange = selection.ShapeRange;
                        if (normalShapeRange != null)
                        {
                            for (int i = 1; i <= normalShapeRange.Count; i++)
                            {
                                var shape = normalShapeRange[i];
                                shapeInfos.Add(new ShapeInfo(shape, i - 1));
                            }
                        }
                        break;

                    case PowerPoint.PpSelectionType.ppSelectionText:
                        // テキスト編集モード
                        logger.Debug("Text editing mode detected");
                        try
                        {
                            // テキスト編集中の図形を取得
                            var textRange = selection.TextRange;
                            if (textRange?.Parent != null)
                            {
                                var parentShape = textRange.Parent.Parent; // TextFrame -> Shape
                                if (parentShape is PowerPoint.Shape shape)
                                {
                                    shapeInfos.Add(new ShapeInfo(shape, 0));
                                    logger.Debug($"Added text editing shape: {shape.Name}");
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            logger.Debug(ex, "Failed to get shape from text editing mode, trying alternative method");

                            // 代替方法：ShapeRangeを試行
                            try
                            {
                                var textModeShapeRange = selection.ShapeRange;
                                if (textModeShapeRange?.Count > 0)
                                {
                                    var shape = textModeShapeRange[1];
                                    shapeInfos.Add(new ShapeInfo(shape, 0));
                                    logger.Debug($"Added shape via alternative method: {shape.Name}");
                                }
                            }
                            catch (Exception ex2)
                            {
                                logger.Warn(ex2, "Alternative method also failed to get text editing shape");
                            }
                        }
                        break;

                    case PowerPoint.PpSelectionType.ppSelectionNone:
                        logger.Debug("No selection");
                        break;

                    default:
                        logger.Debug($"Other selection type: {selection.Type}");
                        break;
                }

                logger.Debug($"Retrieved {shapeInfos.Count} shape(s) including text editing mode");
            }
            catch (Exception ex)
            {
                logger.Error(ex, "Failed to get selected shape infos including text editing mode");
            }

            return shapeInfos;
        }

        /// <summary>
        /// アプリケーション固有の図形選択取得メソッド
        /// </summary>
        /// <param name="application">PowerPointアプリケーション</param>
        /// <returns>選択図形範囲</returns>
        public PowerPoint.ShapeRange GetSelectedShapesFromApplication(PowerPoint.Application application)
        {
            try
            {
                var activeWindow = application.ActiveWindow;
                if (activeWindow == null) return null;

                var selection = activeWindow.Selection;

                switch (selection.Type)
                {
                    case PowerPoint.PpSelectionType.ppSelectionShapes:
                        return selection.ShapeRange;

                    case PowerPoint.PpSelectionType.ppSelectionText:
                        try
                        {
                            var textModeShapeRange = selection.ShapeRange;
                            if (textModeShapeRange?.Count > 0)
                            {
                                return textModeShapeRange;
                            }
                        }
                        catch (Exception ex)
                        {
                            logger.Debug(ex, "Failed to get ShapeRange from text editing mode");
                        }
                        break;
                }
            }
            catch (Exception ex)
            {
                logger.Error(ex, "Failed to get selected shapes from application");
            }

            return null;
        }

        /// <summary>
        /// 現在のスライドを取得します
        /// </summary>
        /// <returns>現在のスライド</returns>
        public PowerPoint.Slide GetCurrentSlide()
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
        /// 図形を選択します
        /// </summary>
        /// <param name="shapes">選択する図形のリスト</param>
        public void SelectShapes(List<PowerPoint.Shape> shapes)
        {
            try
            {
                if (shapes == null || shapes.Count == 0)
                {
                    logger.Debug("No shapes to select");
                    return;
                }

                // 最初の図形を選択
                shapes[0].Select();

                // 残りの図形を追加選択
                for (int i = 1; i < shapes.Count; i++)
                {
                    shapes[i].Select(MsoTriState.msoFalse);
                }

                logger.Debug($"Selected {shapes.Count} shapes");
            }
            catch (Exception ex)
            {
                logger.Error(ex, "Failed to select shapes");
            }
        }

        #endregion

        #region 図形検証・判定ヘルパー

        /// <summary>
        /// 選択状態を検証します
        /// </summary>
        /// <param name="shapeInfos">図形情報リスト</param>
        /// <param name="minRequired">最小必要数</param>
        /// <param name="maxAllowed">最大許容数（0=無制限）</param>
        /// <param name="operationName">操作名</param>
        /// <returns>検証結果</returns>
        public bool ValidateSelection(List<ShapeInfo> shapeInfos, int minRequired, int maxAllowed, string operationName)
        {
            return ErrorHandler.ValidateSelection(shapeInfos.Count, minRequired, maxAllowed, operationName);
        }

        /// <summary>
        /// 図形が線かどうかを判定します
        /// </summary>
        /// <param name="shape">図形</param>
        /// <returns>線図形の場合true</returns>
        public bool IsLineShape(PowerPoint.Shape shape)
        {
            try
            {
                // 線図形の判定
                if (shape.Type == MsoShapeType.msoLine)
                    return true;

                // フリーフォーム（コネクタを含む）の判定
                if (shape.Type == MsoShapeType.msoFreeform)
                {
                    // Connectorプロパティでコネクタかどうかを判定
                    try
                    {
                        return shape.Connector == MsoTriState.msoTrue;
                    }
                    catch
                    {
                        // Connectorプロパティがない場合は、パスポイント数で判定
                        return shape.Nodes.Count == 2;
                    }
                }

                // オートシェイプの線タイプの判定
                if (shape.Type == MsoShapeType.msoAutoShape)
                {
                    return shape.AutoShapeType == MsoAutoShapeType.msoShapeMixed;
                }

                return false;
            }
            catch (Exception ex)
            {
                logger.Error(ex, $"Failed to determine if shape {shape.Name} is a line");
                return false;
            }
        }

        /// <summary>
        /// 図形が類似しているかを判定します
        /// </summary>
        /// <param name="reference">基準図形</param>
        /// <param name="target">対象図形</param>
        /// <returns>類似している場合true</returns>
        public bool IsSimilarShape(PowerPoint.Shape reference, PowerPoint.Shape target)
        {
            try
            {
                return reference.Type == target.Type &&
                       reference.AutoShapeType == target.AutoShapeType;
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// 図形がテーブルかどうかを判定します
        /// </summary>
        /// <param name="shape">図形</param>
        /// <returns>テーブルの場合true</returns>
        public bool IsTableShape(PowerPoint.Shape shape)
        {
            try
            {
                return shape.HasTable == MsoTriState.msoTrue;
            }
            catch (Exception ex)
            {
                logger.Error(ex, $"Failed to determine if shape {shape.Name} is a table");
                return false;
            }
        }

        /// <summary>
        /// 矩形系のオートシェイプかどうかを判定します（静的メソッド）
        /// </summary>
        /// <param name="shp">図形</param>
        /// <returns>矩形系オートシェイプの場合true</returns>
        public static bool IsRectLikeAutoShape(PowerPoint.Shape shp)
        {
            return shp.Type == MsoShapeType.msoAutoShape &&
                   (shp.AutoShapeType == MsoAutoShapeType.msoShapeRectangle ||
                    shp.AutoShapeType == MsoAutoShapeType.msoShapeRoundedRectangle);
        }

        /// <summary>
        /// マトリクスプレースホルダーかどうかを判定します（静的メソッド）
        /// </summary>
        /// <param name="shp">図形</param>
        /// <returns>マトリクスプレースホルダーの場合true</returns>
        public static bool IsMatrixPlaceholder(PowerPoint.Shape shp)
        {
            return shp.Type == MsoShapeType.msoPlaceholder &&
                   (shp.PlaceholderFormat.Type == PowerPoint.PpPlaceholderType.ppPlaceholderObject ||
                    shp.PlaceholderFormat.Type == PowerPoint.PpPlaceholderType.ppPlaceholderBody);
        }

        #endregion

        #region グリッド検出ロジック

        /// <summary>
        /// グリッドレイアウトを検出します
        /// </summary>
        /// <param name="shapes">図形リスト</param>
        /// <returns>グリッド情報</returns>
        public GridInfo DetectGridLayout(List<ShapeInfo> shapes)
        {
            try
            {
                // 動的な許容誤差を計算
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

                // 各行をX座標でソート
                foreach (var row in rows)
                {
                    row.Sort((a, b) => a.Left.CompareTo(b.Left));
                }

                // グリッド情報を構築
                var gridInfo = new GridInfo
                {
                    Rows = rows.Count,
                    Columns = rows.Max(r => r.Count),
                    ShapeGrid = rows,
                    TopLeft = rows.First().First() // 左上の図形
                };

                logger.Debug($"Grid detected: {gridInfo.Rows}x{gridInfo.Columns} with tolerance {tolerance:F1}pt");
                return gridInfo;
            }
            catch (Exception ex)
            {
                logger.Error(ex, "Failed to detect grid layout");
                return null;
            }
        }

        /// <summary>
        /// 動的な許容誤差を計算します
        /// </summary>
        /// <param name="shapes">図形リスト</param>
        /// <param name="isVertical">true=Y座標用（高さベース）、false=X座標用（幅ベース）</param>
        /// <returns>計算された許容誤差（ポイント）</returns>
        public float CalculateDynamicTolerance(List<ShapeInfo> shapes, bool isVertical)
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

        /// <summary>
        /// マトリクスレイアウトを検出します
        /// </summary>
        /// <param name="matrixShapes">マトリクス図形リスト</param>
        /// <returns>グリッド情報とテーブルフラグ</returns>
        public (GridInfo gridInfo, bool isTable) DetectMatrixLayout(List<ShapeInfo> matrixShapes)
        {
            try
            {
                // 単一のテーブル図形の場合
                if (matrixShapes.Count == 1 && IsTableShape(matrixShapes[0].Shape))
                {
                    return DetectTableMatrixLayout(matrixShapes[0]);
                }

                // 複数図形からグリッドを検出
                var gridInfo = DetectGridLayout(matrixShapes);
                return (gridInfo, false);
            }
            catch (Exception ex)
            {
                logger.Error(ex, "Failed to detect matrix layout");
                return (null, false);
            }
        }

        /// <summary>
        /// テーブルマトリクスレイアウトを検出します
        /// </summary>
        /// <param name="tableShape">テーブル図形情報</param>
        /// <returns>グリッド情報とテーブルフラグ</returns>
        public (GridInfo gridInfo, bool isTable) DetectTableMatrixLayout(ShapeInfo tableShape)
        {
            try
            {
                var table = tableShape.Shape.Table;
                var gridInfo = new GridInfo
                {
                    Rows = table.Rows.Count,
                    Columns = table.Columns.Count,
                    ShapeGrid = null, // テーブルの場合はnull
                    TopLeft = tableShape
                };

                logger.Debug($"Table matrix detected: {gridInfo.Rows}x{gridInfo.Columns}");
                return (gridInfo, true);
            }
            catch (Exception ex)
            {
                logger.Error(ex, "Failed to detect table matrix layout");
                return (null, false);
            }
        }

        #endregion

        #region GridInfoクラス

        /// <summary>
        /// グリッド情報を保持するクラス
        /// </summary>
        public class GridInfo
        {
            public int Rows { get; set; }
            public int Columns { get; set; }
            public List<List<ShapeInfo>> ShapeGrid { get; set; }
            public ShapeInfo TopLeft { get; set; }
        }

        #endregion
    }
}
