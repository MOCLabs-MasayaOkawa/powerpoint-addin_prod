using Microsoft.Office.Core;
using NLog;
using PowerPointEfficiencyAddin.Models;
using PowerPointEfficiencyAddin.Services.Core.PowerTool;
using PowerPointEfficiencyAddin.Services.Infrastructure.MultiInstance;
using PowerPointEfficiencyAddin.Services.Infrastructure.Settings;
using PowerPointEfficiencyAddin.UI;
using PowerPointEfficiencyAddin.Utils;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PowerPointEfficiencyAddin.Services.Core.BuiltIn
{
    /// <summary>
    /// PowerPoint標準コマンド実行とBuilt-in図形作成機能を提供するサービスクラス
    /// </summary>
    public class BuiltInShapeService
    {
        private static readonly Logger logger = LogManager.GetCurrentClassLogger();
        private readonly IApplicationProvider applicationProvider;
        private readonly PowerToolServiceHelper helper;
        private readonly PowerToolService powerToolService;

        // DI対応コンストラクタ（商用レベル）
        public BuiltInShapeService(
            IApplicationProvider applicationProvider, 
            PowerToolServiceHelper helper,
            PowerToolService powerToolService)
        {
            this.applicationProvider = applicationProvider ?? throw new ArgumentNullException(nameof(applicationProvider));
            this.helper = helper ?? throw new ArgumentNullException(nameof(helper));
            this.powerToolService = powerToolService ?? throw new ArgumentNullException(nameof(powerToolService));
            logger.Debug("BuiltInShapeService initialized with DI dependencies");
        }

        /// <summary>
        /// PowerPoint標準コマンドを実行します
        /// </summary>
        /// <param name="commandName">コマンド名</param>
        public void ExecutePowerPointCommand(string commandName)
        {

            if (!Globals.ThisAddIn.CheckFeatureAccess(commandName)) return;

            logger.Info($"ExecutePowerPointCommand operation started: {commandName}");

            try
            {
                // コマンドタイプに応じて処理を分岐
                if (IsShapeCreationCommand(commandName))
                {
                    logger.Debug($"Executing as shape creation command: {commandName}");
                    ExecuteShapeCreationCommand(commandName);
                }
                else
                {
                    logger.Debug($"Executing as standard PowerPoint command: {commandName}");
                    ExecuteStandardCommand(commandName);
                }

                logger.Info($"ExecutePowerPointCommand completed: {commandName}");
            }
            catch (Exception ex)
            {
                logger.Error(ex, $"Error in ExecutePowerPointCommand: {commandName}");
                ErrorHandler.ExecuteSafely(() => throw ex, $"コマンド実行: {GetCommandDisplayName(commandName)}");
            }
        }

        /// <summary>
        /// 図形作成系コマンドかどうかを判定します
        /// </summary>
        /// <param name="commandName">コマンド名</param>
        /// <returns>図形作成系の場合true</returns>
        private bool IsShapeCreationCommand(string commandName)
        {
            // 図形作成系コマンドのリスト
            var shapeCreationCommands = new HashSet<string>
    {
        // Shape関連
        "ShapeRectangle",
        "ShapeRoundedRectangle",
        "ShapeOval",
        "ShapeIsoscelesTriangle",
        "ShapeRectangularCallout",
        "ShapeRightArrow",
        "ShapeDownArrow",
        "ShapeLeftArrow",
        "ShapeUpArrow",
        "ShapeLine",
        "ShapeLineArrow",
        "ShapeElbowConnector",
        "ShapeElbowArrowConnector",
        "ShapeLeftBrace",
        "ShapePentagon",
        "ShapeChevron",
        
        // Text関連
        "TextBox"
    };

            return shapeCreationCommands.Contains(commandName);
        }

        /// <summary>
        /// 図形作成系コマンドを実行します
        /// </summary>
        /// <param name="commandName">コマンド名</param>
        private void ExecuteShapeCreationCommand(string commandName)
        {
            ComHelper.ExecuteWithComCleanup(() =>
            {
                var slide = helper.GetCurrentSlide();
                if (slide == null)
                {
                    throw new InvalidOperationException("アクティブなスライドが見つかりません。");
                }

                // デバッグ:配置前の状況をログ
                LogShapePlacementStatus(slide, "Before placement");

                // スライド左上基準で図形を配置
                var shape = CreateShapeAtOptimalPosition(slide, commandName);
                if (shape != null)
                {
                    shape.Select();
                    logger.Info($"Successfully placed {GetCommandDisplayName(commandName)} at position ({shape.Left:F1}, {shape.Top:F1})");

                    // デバッグ:配置後の状況をログ
                    LogShapePlacementStatus(slide, "After placement");
                }
                else
                {
                    throw new InvalidOperationException($"図形の作成に失敗しました: {commandName}");
                }
            });
        }

        /// <summary>
        /// PowerPoint標準機能コマンドを実行します
        /// </summary>
        /// <param name="commandName">コマンド名</param>
        private void ExecuteStandardCommand(string commandName)
        {
            ComHelper.ExecuteWithComCleanup(() =>
            {
                var application = applicationProvider.GetCurrentApplication();

                try
                {
                    switch (commandName)
                    {
                        // 整列系コマンド
                        case "AlignLeft":
                            ExecuteAlignmentCommand(MsoAlignCmd.msoAlignLefts);
                            break;

                        case "AlignCenterHorizontal":
                            ExecuteAlignmentCommand(MsoAlignCmd.msoAlignCenters);
                            break;

                        case "AlignRight":
                            ExecuteAlignmentCommand(MsoAlignCmd.msoAlignRights);
                            break;

                        case "AlignTop":
                            ExecuteAlignmentCommand(MsoAlignCmd.msoAlignTops);
                            break;

                        case "AlignCenterVertical":
                            ExecuteAlignmentCommand(MsoAlignCmd.msoAlignMiddles);
                            break;

                        case "AlignBottom":
                            ExecuteAlignmentCommand(MsoAlignCmd.msoAlignBottoms);
                            break;

                        // グループ化系コマンド
                        case "GroupObjects":
                            ExecuteGroupCommand();
                            break;

                        case "UngroupObjects":
                            ExecuteUngroupCommand();
                            break;

                        default:
                            logger.Warn($"Unknown standard command: {commandName}");
                            // フォールバック: ExecuteMsoを試行
                            try
                            {
                                application.CommandBars.ExecuteMso(commandName);
                                logger.Info($"Executed via ExecuteMso: {commandName}");
                            }
                            catch (Exception msoEx)
                            {
                                logger.Error(msoEx, $"Failed to execute via ExecuteMso: {commandName}");
                                throw new InvalidOperationException($"サポートされていないコマンドです: {commandName}");
                            }
                            break;
                    }

                    logger.Info($"Standard command executed successfully: {commandName}");
                }
                catch (COMException comEx)
                {
                    logger.Error(comEx, $"COM error executing standard command: {commandName}");
                    throw new InvalidOperationException($"PowerPoint標準機能の実行に失敗しました: {commandName}", comEx);
                }
            });
        }

        /// <summary>
        /// 整列コマンドを実行します
        /// </summary>
        /// <param name="alignCommand">整列タイプ</param>
        private void ExecuteAlignmentCommand(MsoAlignCmd alignCommand)
        {
            var application = applicationProvider.GetCurrentApplication();

            try
            {
                // 選択されている図形を取得
                var selection = application.ActiveWindow.Selection;

                if (selection.Type != PowerPoint.PpSelectionType.ppSelectionShapes)
                {
                    throw new InvalidOperationException("図形が選択されていません。整列を行うには図形を選択してください。");
                }

                if (selection.ShapeRange.Count < 2)
                {
                    throw new InvalidOperationException("整列を行うには2つ以上の図形を選択してください。");
                }

                // 整列実行
                selection.ShapeRange.Align(alignCommand, MsoTriState.msoFalse);
                logger.Debug($"Alignment executed: {alignCommand} on {selection.ShapeRange.Count} shapes");
            }
            catch (COMException comEx)
            {
                logger.Error(comEx, $"Failed to execute alignment command: {alignCommand}");
                throw new InvalidOperationException($"整列コマンドの実行に失敗しました: {alignCommand}", comEx);
            }
        }

        /// <summary>
        /// グループ化コマンドを実行します
        /// </summary>
        private void ExecuteGroupCommand()
        {
            var application = applicationProvider.GetCurrentApplication();

            try
            {
                var selection = application.ActiveWindow.Selection;

                if (selection.Type != PowerPoint.PpSelectionType.ppSelectionShapes)
                {
                    throw new InvalidOperationException("図形が選択されていません。グループ化を行うには図形を選択してください。");
                }

                if (selection.ShapeRange.Count < 2)
                {
                    throw new InvalidOperationException("グループ化を行うには2つ以上の図形を選択してください。");
                }

                // グループ化実行
                selection.ShapeRange.Group();
                logger.Debug($"Grouped {selection.ShapeRange.Count} shapes");
            }
            catch (COMException comEx)
            {
                logger.Error(comEx, "Failed to execute group command");
                throw new InvalidOperationException("グループ化の実行に失敗しました", comEx);
            }
        }

        /// <summary>
        /// グループ解除コマンドを実行します
        /// </summary>
        private void ExecuteUngroupCommand()
        {
            var application = applicationProvider.GetCurrentApplication();

            try
            {
                var selection = application.ActiveWindow.Selection;

                if (selection.Type != PowerPoint.PpSelectionType.ppSelectionShapes)
                {
                    throw new InvalidOperationException("図形が選択されていません。グループ解除を行うには図形を選択してください。");
                }

                // グループ解除実行
                selection.ShapeRange.Ungroup();
                logger.Debug("Ungrouped selected shapes");
            }
            catch (COMException comEx)
            {
                logger.Error(comEx, "Failed to execute ungroup command");
                throw new InvalidOperationException("グループ解除の実行に失敗しました", comEx);
            }
        }

        /// <summary>
        /// 最適な位置に図形を作成します（2列斜め配置）
        /// </summary>
        /// <param name="slide">対象スライド</param>
        /// <param name="commandName">コマンド名（図形種別判定用）</param>
        /// <returns>作成された図形</returns>
        private PowerPoint.Shape CreateShapeAtOptimalPosition(PowerPoint.Slide slide, string commandName)
        {
            try
            {
                // 次の配置位置を計算（2列斜め配置）
                var position = CalculateDiagonalColumnPosition(slide);

                // 図形の標準サイズを取得
                var size = GetStandardShapeSize(commandName);

                // 図形を作成
                var shape = CreateShapeAtPosition(slide, commandName, position.Left, position.Top, size.Width, size.Height);

                // スタイル適用（PowerToolServiceのpublicメソッドを呼び出し）
                powerToolService.ApplyShapeStyle(shape, commandName);

                return shape;
            }
            catch (Exception ex)
            {
                logger.Error(ex, $"Failed to create shape at optimal position: {commandName}");
                throw;
            }
        }

        /// <summary>
        /// 2列斜め配置の次の位置を計算します
        /// </summary>
        /// <param name="slide">対象スライド</param>
        /// <returns>配置位置（Left, Top）</returns>
        private (float Left, float Top) CalculateDiagonalColumnPosition(PowerPoint.Slide slide)
        {
            const float column1Left = 50f;   // 第1列の左端
            const float column2Left = 200f;  // 第2列の左端
            const float topMargin = 100f;    // 上マージン
            const float verticalSpacing = 100f; // 縦方向の間隔

            try
            {
                // 既存の図形数を取得
                int existingShapeCount = slide.Shapes.Count;

                // 列を交互に使う（0,2,4... → 列1、1,3,5... → 列2）
                bool isColumn1 = (existingShapeCount % 2 == 0);
                float left = isColumn1 ? column1Left : column2Left;

                // 同じ列内での縦位置（行番号）を計算
                int rowInColumn = existingShapeCount / 2;
                float top = topMargin + (rowInColumn * verticalSpacing);

                logger.Debug($"Calculated position for shape #{existingShapeCount + 1}: Column {(isColumn1 ? 1 : 2)}, Row {rowInColumn + 1}, Position ({left:F1}, {top:F1})");

                return (left, top);
            }
            catch (Exception ex)
            {
                logger.Error(ex, "Failed to calculate diagonal column position");
                // フォールバック: デフォルト位置
                return (column1Left, topMargin);
            }
        }

        /// <summary>
        /// 図形の標準サイズを取得します
        /// </summary>
        /// <param name="commandName">コマンド名</param>
        /// <returns>標準サイズ（Width, Height）</returns>
        private (float Width, float Height) GetStandardShapeSize(string commandName)
        {
            // 線系図形は特殊処理
            if (commandName.Contains("Line") || commandName.Contains("Connector"))
            {
                return (100f, 0f); // 線の場合は幅のみ（高さは0）
            }

            // その他の図形はデフォルトサイズ
            return (100f, 100f);
        }

        /// <summary>
        /// 指定位置に図形を作成します
        /// </summary>
        /// <param name="slide">対象スライド</param>
        /// <param name="commandName">コマンド名</param>
        /// <param name="left">左端位置</param>
        /// <param name="top">上端位置</param>
        /// <param name="width">幅</param>
        /// <param name="height">高さ</param>
        /// <returns>作成された図形</returns>
        private PowerPoint.Shape CreateShapeAtPosition(
            PowerPoint.Slide slide,
            string commandName,
            float left,
            float top,
            float width,
            float height)
        {
            PowerPoint.Shape shape = null;

            try
            {
                switch (commandName)
                {
                    // 基本図形
                    case "ShapeRectangle":
                        shape = slide.Shapes.AddShape(MsoAutoShapeType.msoShapeRectangle, left, top, width, height);
                        break;

                    case "ShapeRoundedRectangle":
                        shape = slide.Shapes.AddShape(MsoAutoShapeType.msoShapeRoundedRectangle, left, top, width, height);
                        break;

                    case "ShapeOval":
                        shape = slide.Shapes.AddShape(MsoAutoShapeType.msoShapeOval, left, top, width, height);
                        break;

                    case "ShapeIsoscelesTriangle":
                        shape = slide.Shapes.AddShape(MsoAutoShapeType.msoShapeIsoscelesTriangle, left, top, width, height);
                        break;

                    case "ShapeRectangularCallout":
                        shape = slide.Shapes.AddShape(MsoAutoShapeType.msoShapeRectangularCallout, left, top, width, height);
                        break;

                    // 矢印系図形
                    case "ShapeRightArrow":
                        shape = slide.Shapes.AddShape(MsoAutoShapeType.msoShapeRightArrow, left, top, width, height);
                        break;

                    case "ShapeDownArrow":
                        shape = slide.Shapes.AddShape(MsoAutoShapeType.msoShapeDownArrow, left, top, width, height);
                        break;

                    case "ShapeLeftArrow":
                        shape = slide.Shapes.AddShape(MsoAutoShapeType.msoShapeLeftArrow, left, top, width, height);
                        break;

                    case "ShapeUpArrow":
                        shape = slide.Shapes.AddShape(MsoAutoShapeType.msoShapeUpArrow, left, top, width, height);
                        break;

                    // 線系図形
                    case "ShapeLine":
                        shape = slide.Shapes.AddLine(left, top, left + width, top);
                        break;

                    case "ShapeLineArrow":
                        shape = slide.Shapes.AddLine(left, top, left + width, top);
                        if (shape.Line != null)
                        {
                            shape.Line.EndArrowheadStyle = MsoArrowheadStyle.msoArrowheadTriangle;
                        }
                        break;

                    case "ShapeElbowConnector":
                        shape = slide.Shapes.AddConnector(MsoConnectorType.msoConnectorElbow, left, top, left + width, top + height);
                        break;

                    case "ShapeElbowArrowConnector":
                        shape = slide.Shapes.AddConnector(MsoConnectorType.msoConnectorElbow, left, top, left + width, top + height);
                        if (shape.Line != null)
                        {
                            shape.Line.EndArrowheadStyle = MsoArrowheadStyle.msoArrowheadTriangle;
                        }
                        break;

                    // その他の図形
                    case "ShapeLeftBrace":
                        shape = slide.Shapes.AddShape(MsoAutoShapeType.msoShapeLeftBrace, left, top, width, height);
                        break;

                    case "ShapePentagon":
                        shape = slide.Shapes.AddShape(MsoAutoShapeType.msoShapePentagon, left, top, width, height);
                        break;

                    case "ShapeChevron":
                        shape = slide.Shapes.AddShape(MsoAutoShapeType.msoShapeChevron, left, top, width, height);
                        break;

                    // テキストボックス
                    case "TextBox":
                        shape = slide.Shapes.AddTextbox(
                            MsoTextOrientation.msoTextOrientationHorizontal,
                            left, top, width, height);
                        break;

                    default:
                        logger.Warn($"Unknown shape command: {commandName}");
                        // デフォルトで四角形を作成
                        shape = slide.Shapes.AddShape(MsoAutoShapeType.msoShapeRectangle, left, top, width, height);
                        break;
                }

                if (shape != null)
                {
                    logger.Debug($"Created {commandName} at ({left:F1}, {top:F1}) with size ({width:F1} x {height:F1})");
                }

                return shape;
            }
            catch (Exception ex)
            {
                logger.Error(ex, $"Failed to create shape: {commandName}");
                throw;
            }
        }

        /// <summary>
        /// コマンド名から表示名を取得します
        /// </summary>
        /// <param name="commandName">コマンド名</param>
        /// <returns>表示名</returns>
        private string GetCommandDisplayName(string commandName)
        {
            switch (commandName)
            {
                // Shape関連
                case "ShapeRectangle":
                    return "四角形";
                case "ShapeRoundedRectangle":
                    return "角丸四角形";
                case "ShapeOval":
                    return "楕円";
                case "ShapeIsoscelesTriangle":
                    return "二等辺三角形";
                case "ShapeRectangularCallout":
                    return "吹き出し（四角形）";
                case "ShapeRightArrow":
                    return "右矢印";
                case "ShapeDownArrow":
                    return "下矢印";
                case "ShapeLeftArrow":
                    return "左矢印";
                case "ShapeUpArrow":
                    return "上矢印";
                case "ShapeLine":
                    return "直線";
                case "ShapeLineArrow":
                    return "矢印付き直線";
                case "ShapeElbowConnector":
                    return "L字コネクタ";
                case "ShapeElbowArrowConnector":
                    return "矢印付きL字コネクタ";
                // Text関連
                case "TextBox":
                    return "テキストボックス";
                case "ShapeLeftBrace":
                    return "中括弧";
                case "ShapePentagon":
                    return "五角形";
                case "ShapeChevron":
                    return "シェブロン";
                default:
                    return commandName;
            }
        }

        /// <summary>
        /// スライド上の図形配置状況をログ出力（デバッグ用）
        /// </summary>
        /// <param name="slide">対象スライド</param>
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

                // 列ごとに分類してログ出力
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
    }
}
