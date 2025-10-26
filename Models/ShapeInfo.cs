using Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PowerPointEfficiencyAddin.Models
{
    /// <summary>
    /// 図形情報を格納するモデルクラス
    /// </summary>
    public class ShapeInfo
    {
        /// <summary>
        /// 図形参照
        /// </summary>
        public PowerPoint.Shape Shape { get; set; }

        /// <summary>
        /// 図形名
        /// </summary>
        public string Name { get; set; }

        /// <summary>
        /// 選択順序（0が最初）
        /// </summary>
        public int SelectionOrder { get; set; }

        /// <summary>
        /// 左端位置
        /// </summary>
        public float Left { get; set; }

        /// <summary>
        /// 上端位置
        /// </summary>
        public float Top { get; set; }

        /// <summary>
        /// 幅
        /// </summary>
        public float Width { get; set; }

        /// <summary>
        /// 高さ
        /// </summary>
        public float Height { get; set; }

        /// <summary>
        /// 右端位置
        /// </summary>
        public float Right => Left + Width;

        /// <summary>
        /// 下端位置
        /// </summary>
        public float Bottom => Top + Height;

        /// <summary>
        /// 中央X座標
        /// </summary>
        public float CenterX => Left + (Width / 2);

        /// <summary>
        /// 中央Y座標
        /// </summary>
        public float CenterY => Top + (Height / 2);

        /// <summary>
        /// 図形の種類
        /// </summary>
        public MsoShapeType ShapeType { get; set; }

        /// <summary>
        /// テキストフレームを持つかどうか
        /// </summary>
        public bool HasTextFrame { get; set; }

        /// <summary>
        /// テキスト内容
        /// </summary>
        public string Text { get; set; }

        /// <summary>
        /// コンストラクタ
        /// </summary>
        /// <param name="shape">PowerPoint図形オブジェクト</param>
        /// <param name="selectionOrder">選択順序</param>
        public ShapeInfo(PowerPoint.Shape shape, int selectionOrder)
        {
            Shape = shape;
            SelectionOrder = selectionOrder;

            // 基本プロパティの取得
            Name = shape.Name;
            Left = shape.Left;
            Top = shape.Top;
            Width = shape.Width;
            Height = shape.Height;
            ShapeType = shape.Type;

            // テキスト情報の取得
            try
            {
                HasTextFrame = shape.HasTextFrame == MsoTriState.msoTrue;
                if (HasTextFrame && shape.TextFrame.HasText == MsoTriState.msoTrue)
                {
                    Text = shape.TextFrame.TextRange.Text;
                }
                else
                {
                    Text = string.Empty;
                }
            }
            catch
            {
                HasTextFrame = false;
                Text = string.Empty;
            }
        }

        /// <summary>
        /// 図形の位置とサイズを更新します
        /// </summary>
        public void UpdateDimensions()
        {
            if (Shape != null)
            {
                Left = Shape.Left;
                Top = Shape.Top;
                Width = Shape.Width;
                Height = Shape.Height;
            }
        }

        /// <summary>
        /// 図形のテキストを更新します
        /// </summary>
        public void UpdateText()
        {
            if (Shape != null && HasTextFrame)
            {
                try
                {
                    if (Shape.TextFrame.HasText == MsoTriState.msoTrue)
                    {
                        Text = Shape.TextFrame.TextRange.Text;
                    }
                    else
                    {
                        Text = string.Empty;
                    }
                }
                catch
                {
                    Text = string.Empty;
                }
            }
        }

        /// <summary>
        /// デバッグ用文字列表現
        /// </summary>
        /// <returns>図形情報の文字列</returns>
        public override string ToString()
        {
            return $"ShapeInfo: {Name} ({Left}, {Top}, {Width}x{Height}) Order: {SelectionOrder}";
        }
    }

    /// <summary>
    /// 整列基準を表す列挙型
    /// </summary>
    public enum AlignmentReference
    {
        /// <summary>最左端の図形を基準</summary>
        LeftMost,
        /// <summary>最右端の図形を基準</summary>
        RightMost,
        /// <summary>最上端の図形を基準</summary>
        TopMost,
        /// <summary>最下端の図形を基準</summary>
        BottomMost,
        /// <summary>最初に選択された図形を基準</summary>
        FirstSelected,
        /// <summary>選択範囲全体を基準</summary>
        SelectionBounds,
        /// <summary>スライド全体を基準</summary>
        SlideBounds
    }

    /// <summary>
    /// 配置方向を表す列挙型
    /// </summary>
    public enum PlacementDirection
    {
        /// <summary>右端を左端へ</summary>
        RightToLeft,
        /// <summary>左端を右端へ</summary>
        LeftToRight,
        /// <summary>上端を下端へ</summary>
        TopToBottom,
        /// <summary>下端を上端へ</summary>
        BottomToTop
    }
}