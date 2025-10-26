using System;
using System.Collections.Generic;
using System.Linq;
using NLog;
using PowerPointEfficiencyAddin.Models;
using PowerPointEfficiencyAddin.Models.Licensing;

namespace PowerPointEfficiencyAddin.Services.Licensing
{
    /// <summary>
    /// 機能とプランのマッピング管理
    /// </summary>
    public class FeatureRegistry
    {
        private static readonly Logger logger = LogManager.GetCurrentClassLogger();
        private static FeatureRegistry instance;
        private readonly Dictionary<string, FeatureDefinition> features;

        public static FeatureRegistry Instance
        {
            get
            {
                if (instance == null)
                {
                    instance = new FeatureRegistry();
                }
                return instance;
            }
        }

        private FeatureRegistry()
        {
            features = new Dictionary<string, FeatureDefinition>(StringComparer.OrdinalIgnoreCase);
            InitializeFeatures();
        }

        /// <summary>
        /// 機能定義を初期化
        /// </summary>
        private void InitializeFeatures()
        {
            int order = 0;

            // ========== Selection（選択）カテゴリ ==========
            Register("SelectSameColorShapes", "色で選択", FunctionCategory.Selection, FeatureAccessLevel.Pro, order++); //OK
            Register("SelectSameSizeShapes", "サイズで選択", FunctionCategory.Selection, FeatureAccessLevel.Pro, order++); //OK
            Register("SelectSimilarShapes", "同種図形で選択", FunctionCategory.Selection, FeatureAccessLevel.Pro, order++); //OK

            // ========== Text（テキスト）カテゴリ ==========
            Register("ToggleTextWrap", "折り返しトグル", FunctionCategory.Text, FeatureAccessLevel.Pro, order++); //OK
            Register("AdjustMarginUp", "余白Up", FunctionCategory.Text, FeatureAccessLevel.Pro, order++); //OK
            Register("AdjustMarginDown", "余白Down", FunctionCategory.Text, FeatureAccessLevel.Pro, order++); //OK
            Register("ShowMarginAdjustDialog", "余白調整", FunctionCategory.Text, FeatureAccessLevel.Pro, order++); //OK
            Register("TextBox", "Text Box", FunctionCategory.Text, FeatureAccessLevel.Pro, order++);
            Register("ClearTextsFromSelectedShapes", "テキストクリア", FunctionCategory.Text, FeatureAccessLevel.Pro, order++); //OK

            // ========== Shape（図形）カテゴリ ==========
            Register("ShapeRectangle", "四角", FunctionCategory.Shape, FeatureAccessLevel.Pro, order++); 
            Register("ShapeRoundedRectangle", "角丸", FunctionCategory.Shape, FeatureAccessLevel.Pro, order++); 
            Register("ShapeOval", "丸", FunctionCategory.Shape, FeatureAccessLevel.Pro, order++); 
            Register("ShapeIsoscelesTriangle", "三角", FunctionCategory.Shape, FeatureAccessLevel.Pro, order++); 
            Register("ShapeRectangularCallout", "吹き出し", FunctionCategory.Shape, FeatureAccessLevel.Pro, order++); 
            Register("ShapeRightArrow", "矢印（右）", FunctionCategory.Shape, FeatureAccessLevel.Pro, order++); 
            Register("ShapeDownArrow", "矢印（下）", FunctionCategory.Shape, FeatureAccessLevel.Pro, order++); 
            Register("ShapeLine", "線", FunctionCategory.Shape, FeatureAccessLevel.Pro, order++); 
            Register("ShapeLineArrow", "矢印線", FunctionCategory.Shape, FeatureAccessLevel.Pro, order++); 
            Register("ShapeElbowConnector", "鍵線", FunctionCategory.Shape, FeatureAccessLevel.Pro, order++); 
            Register("ShapeElbowArrowConnector", "鍵線矢印", FunctionCategory.Shape, FeatureAccessLevel.Pro, order++); 
            Register("ShapeLeftBrace", "中括弧", FunctionCategory.Shape, FeatureAccessLevel.Pro, order++); 
            Register("ShapePentagon", "五角形", FunctionCategory.Shape, FeatureAccessLevel.Pro, order++); 
            Register("ShapeChevron", "シェブロン", FunctionCategory.Shape, FeatureAccessLevel.Pro, order++); 
            Register("ShapeStyleSettings", "図形スタイル設定", FunctionCategory.Shape, FeatureAccessLevel.Pro, order++); //OK

            // ========== Format（整形）カテゴリ ==========
            Register("SizeUpToggle", "サイズUp", FunctionCategory.Format, FeatureAccessLevel.Pro, order++); //OK
            Register("SizeDownToggle", "サイズDown", FunctionCategory.Format, FeatureAccessLevel.Pro, order++); //OK
            Register("LineWeightUpToggle", "枠線の太さUp", FunctionCategory.Format, FeatureAccessLevel.Pro, order++); //OK
            Register("LineWeightDownToggle", "枠線の太さDown", FunctionCategory.Format, FeatureAccessLevel.Pro, order++); //OK
            Register("DashStyleToggle", "枠線の種類変更トグル", FunctionCategory.Format, FeatureAccessLevel.Pro, order++); //OK
            Register("TransparencyUpToggle", "透過率Upトグル", FunctionCategory.Format, FeatureAccessLevel.Pro, order++); //OK
            Register("TransparencyDownToggle", "透過率Downトグル", FunctionCategory.Format, FeatureAccessLevel.Pro, order++); //OK
            Register("MatchHeight", "縦幅を揃える", FunctionCategory.Format, FeatureAccessLevel.Pro, order++); //OK
            Register("MatchWidth", "横幅を揃える", FunctionCategory.Format, FeatureAccessLevel.Pro, order++); //OK
            Register("MatchSize", "横幅縦幅を揃える", FunctionCategory.Format, FeatureAccessLevel.Pro, order++); //OK
            Register("MatchFormat", "書式を揃える", FunctionCategory.Format, FeatureAccessLevel.Pro, order++); //OK
            Register("AlignSizeLeft", "左端を揃える", FunctionCategory.Format, FeatureAccessLevel.Pro, order++); //OK
            Register("AlignSizeRight", "右端を揃える", FunctionCategory.Format, FeatureAccessLevel.Pro, order++); //OK
            Register("AlignSizeTop", "上端を揃える", FunctionCategory.Format, FeatureAccessLevel.Pro, order++); //OK
            Register("AlignSizeBottom", "下端を揃える", FunctionCategory.Format, FeatureAccessLevel.Pro, order++); //OK
            Register("AlignLineLength", "線の長さを揃える", FunctionCategory.Format, FeatureAccessLevel.Pro, order++); //OK

            // ========== Grouping（グループ化）カテゴリ ==========
            Register("GroupShapes", "グループ化", FunctionCategory.Grouping, FeatureAccessLevel.Pro, order++);
            Register("UngroupShapes", "グループ解除", FunctionCategory.Grouping, FeatureAccessLevel.Pro, order++);
            Register("GroupByRows", "行でグループ化", FunctionCategory.Grouping, FeatureAccessLevel.Pro, order++); //OK
            Register("GroupByColumns", "列でグループ化", FunctionCategory.Grouping, FeatureAccessLevel.Pro, order++); //OK

            // ========== Alignment（整列）カテゴリ ==========
            Register("AlignLeft", "左揃え", FunctionCategory.Alignment, FeatureAccessLevel.Free, order++); 
            Register("AlignCenterHorizontal", "中央揃え", FunctionCategory.Alignment, FeatureAccessLevel.Free, order++); 
            Register("AlignRight", "右揃え", FunctionCategory.Alignment, FeatureAccessLevel.Free, order++);
            Register("AlignTop", "上揃え", FunctionCategory.Alignment, FeatureAccessLevel.Free, order++);
            Register("AlignCenterVertical", "水平揃え", FunctionCategory.Alignment, FeatureAccessLevel.Free, order++);
            Register("AlignBottom", "下揃え", FunctionCategory.Alignment, FeatureAccessLevel.Free, order++);
            Register("PlaceLeftToRight", "左端を右端へ", FunctionCategory.Alignment, FeatureAccessLevel.Pro, order++); //OK
            Register("PlaceRightToLeft", "右端を左端へ", FunctionCategory.Alignment, FeatureAccessLevel.Pro, order++); //OK
            Register("PlaceTopToBottom", "上端を下端へ", FunctionCategory.Alignment, FeatureAccessLevel.Pro, order++); //OK
            Register("PlaceBottomToTop", "下端を上端へ", FunctionCategory.Alignment, FeatureAccessLevel.Pro, order++); //OK
            Register("CenterAlign", "水平垂直中央揃え", FunctionCategory.Alignment, FeatureAccessLevel.Pro, order++); //OK
            Register("MakeLineHorizontal", "水平にする", FunctionCategory.Alignment, FeatureAccessLevel.Pro, order++); //OK
            Register("MakeLineVertical", "垂直にする", FunctionCategory.Alignment, FeatureAccessLevel.Pro, order++); //OK
            Register("MatchRoundCorner", "角丸統一", FunctionCategory.Alignment, FeatureAccessLevel.Pro, order++); //OK
            Register("MatchEnvironment", "矢羽統一", FunctionCategory.Alignment, FeatureAccessLevel.Pro, order++); //OK

            // ========== ShapeOperation（図形操作プロ）カテゴリ ==========
            Register("SplitShape", "図形分割", FunctionCategory.ShapeOperation, FeatureAccessLevel.Pro, order++); //OK
            Register("DuplicateShape", "図形複製", FunctionCategory.ShapeOperation, FeatureAccessLevel.Pro, order++); //OK
            Register("GenerateMatrix", "マトリクス生成", FunctionCategory.ShapeOperation, FeatureAccessLevel.Pro, order++); //OK
            Register("AddSequentialNumbers", "連番付与", FunctionCategory.ShapeOperation, FeatureAccessLevel.Pro, order++); //OK
            Register("MergeText", "テキスト図形統合", FunctionCategory.ShapeOperation, FeatureAccessLevel.Pro, order++); //OK
            Register("SwapPositions", "図形位置の交換", FunctionCategory.ShapeOperation, FeatureAccessLevel.Pro, order++); //OK

            // ========== TableOperation（表操作）カテゴリ ==========
            Register("ConvertTableToTextBoxes", "表→オブジェクト", FunctionCategory.TableOperation, FeatureAccessLevel.Pro, order++); //OK
            Register("ConvertTextBoxesToTable", "オブジェクト→表", FunctionCategory.TableOperation, FeatureAccessLevel.Pro, order++); //OK
            Register("OptimizeMatrixRowHeights", "行高さ最適化", FunctionCategory.TableOperation, FeatureAccessLevel.Pro, order++); //OK
            Register("OptimizeTableComplete", "表最適化", FunctionCategory.TableOperation, FeatureAccessLevel.Pro, order++); //OK
            Register("EqualizeRowHeights", "行高統一", FunctionCategory.TableOperation, FeatureAccessLevel.Pro, order++); //OK
            Register("EqualizeColumnWidths", "列幅統一", FunctionCategory.TableOperation, FeatureAccessLevel.Pro, order++); //OK
            Register("ExcelToPptx", "ExcelToPPT", FunctionCategory.TableOperation, FeatureAccessLevel.Pro, order++); //OK
            Register("AddMatrixRowSeparators", "行間区切り線", FunctionCategory.TableOperation, FeatureAccessLevel.Pro, order++); //OK

            Register("AlignShapesToCells", "図形セル整列", FunctionCategory.TableOperation, FeatureAccessLevel.Pro, order++); //OK
            Register("AddHeaderRowToMatrix", "見出し行付与", FunctionCategory.TableOperation, FeatureAccessLevel.Pro, order++); //OK
            Register("SetCellMargins", "セルマージン設定", FunctionCategory.TableOperation, FeatureAccessLevel.Pro, order++); //OK
            Register("AddMatrixRow", "行追加", FunctionCategory.TableOperation, FeatureAccessLevel.Pro, order++); //OK
            Register("AddMatrixColumn", "列追加", FunctionCategory.TableOperation, FeatureAccessLevel.Pro, order++); //OK

            Register("MatrixTuner", "Matrix Tuner", FunctionCategory.TableOperation, FeatureAccessLevel.Pro, order++); //OK


            // ========== Spacing（間隔）カテゴリ ==========
            Register("RemoveSpacing", "間隔をなくす", FunctionCategory.Spacing, FeatureAccessLevel.Pro, order++); //OK
            Register("AdjustHorizontalSpacing", "水平間隔調整", FunctionCategory.Spacing, FeatureAccessLevel.Pro, order++); //OK
            Register("AdjustVerticalSpacing", "垂直間隔調整", FunctionCategory.Spacing, FeatureAccessLevel.Pro, order++); //OK
            Register("AdjustEqualSpacing", "間隔調整", FunctionCategory.Spacing, FeatureAccessLevel.Pro, order++); //OK

            // ========== PowerTool（パワーツール）カテゴリ ==========
            Register("UnifyFont", "テキスト一括置換", FunctionCategory.PowerTool, FeatureAccessLevel.Pro, order++);
            Register("CompressImages", "画像圧縮", FunctionCategory.PowerTool, FeatureAccessLevel.Pro, order++); //OK

            logger.Info($"Feature registry initialized with {features.Count} features");
        }

        /// <summary>
        /// 機能を登録
        /// </summary>
        private void Register(string id, string name, FunctionCategory category,
                             FeatureAccessLevel requiredLevel, int order)
        {
            features[id] = new FeatureDefinition
            {
                FeatureId = id,
                DisplayName = name,
                Category = category,
                RequiredLevel = requiredLevel,
                Order = order,
                IsEnabled = true
            };
        }

        /// <summary>
        /// 機能が利用可能かチェック
        /// </summary>
        public bool IsFeatureAvailable(string featureId, FeatureAccessLevel currentLevel)
        {
            if (currentLevel == FeatureAccessLevel.Development) return true;
            if (currentLevel == FeatureAccessLevel.Blocked) return false;

            if (!features.TryGetValue(featureId, out var feature))
            {
                logger.Warn($"Unknown feature: {featureId}, defaulting to Pro requirement");
                return currentLevel >= FeatureAccessLevel.Pro;
            }

            return currentLevel >= feature.RequiredLevel;
        }

        /// <summary>
        /// 必要なレベルを取得
        /// </summary>
        public FeatureAccessLevel GetRequiredLevel(string featureId)
        {
            if (features.TryGetValue(featureId, out var feature))
            {
                return feature.RequiredLevel;
            }
            return FeatureAccessLevel.Pro;
        }

        /// <summary>
        /// プラン別の機能数を取得
        /// </summary>
        public Dictionary<FeatureAccessLevel, int> GetFeatureCountByLevel()
        {
            return new Dictionary<FeatureAccessLevel, int>
            {
                { FeatureAccessLevel.Free, features.Count(f => f.Value.RequiredLevel <= FeatureAccessLevel.Free) },
                { FeatureAccessLevel.Starter, features.Count(f => f.Value.RequiredLevel <= FeatureAccessLevel.Starter) },
                { FeatureAccessLevel.Growth, features.Count(f => f.Value.RequiredLevel <= FeatureAccessLevel.Growth) },
                { FeatureAccessLevel.Pro, features.Count }
            };
        }
    }
}
