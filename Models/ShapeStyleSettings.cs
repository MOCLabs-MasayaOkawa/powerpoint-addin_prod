using System;
using System.Drawing;

namespace PowerPointEfficiencyAddin.Models
{
    /// <summary>
    /// 図形スタイル設定クラス
    /// 新規作成図形に適用するデフォルトスタイルを管理
    /// </summary>
    public class ShapeStyleSettings
    {
        #region プロパティ

        /// <summary>
        /// スタイリング機能の有効/無効
        /// </summary>
        public bool EnableStyling { get; set; }

        /// <summary>
        /// 塗りつぶし色
        /// </summary>
        public Color FillColor { get; set; }

        /// <summary>
        /// 枠線色
        /// </summary>
        public Color LineColor { get; set; }

        /// <summary>
        /// フォント名
        /// </summary>
        public string FontName { get; set; }

        /// <summary>
        /// フォント色
        /// </summary>
        public Color FontColor { get; set; }

        /// <summary>
        /// フォントサイズ（ポイント）
        /// </summary>
        public float FontSize { get; set; }

        /// <summary>
        /// 太字設定
        /// </summary>
        public bool FontBold { get; set; }

        /// <summary>
        /// 斜体設定
        /// </summary>
        public bool FontItalic { get; set; }

        #endregion

        #region コンストラクタ

        /// <summary>
        /// デフォルト値でインスタンスを初期化
        /// </summary>
        public ShapeStyleSettings()
        {
            SetDefaults();
        }

        /// <summary>
        /// 指定値でインスタンスを初期化
        /// </summary>
        public ShapeStyleSettings(bool enableStyling, Color fillColor, Color lineColor,
            string fontName, Color fontColor, float fontSize, bool fontBold, bool fontItalic)
        {
            EnableStyling = enableStyling;
            FillColor = fillColor;
            LineColor = lineColor;
            FontName = fontName ?? GetDefaultFontName();
            FontColor = fontColor;
            FontSize = fontSize;
            FontBold = fontBold;
            FontItalic = fontItalic;
        }

        #endregion

        #region パブリックメソッド

        /// <summary>
        /// デフォルト値を設定
        /// </summary>
        public void SetDefaults()
        {
            EnableStyling = false; // 初期状態は無効
            FillColor = Color.FromArgb(68, 114, 196); // PowerPoint標準の青
            LineColor = Color.FromArgb(68, 114, 196); // PowerPoint標準の青
            FontName = GetDefaultFontName();
            FontColor = Color.Black; // 標準の黒
            FontSize = 11.0f; // PowerPoint標準サイズ
            FontBold = false;
            FontItalic = false;
        }

        /// <summary>
        /// 設定値を複製
        /// </summary>
        /// <returns>複製された設定</returns>
        public ShapeStyleSettings Clone()
        {
            return new ShapeStyleSettings(
                EnableStyling,
                FillColor,
                LineColor,
                FontName,
                FontColor,
                FontSize,
                FontBold,
                FontItalic
            );
        }

        /// <summary>
        /// 他の設定から値をコピー
        /// </summary>
        /// <param name="source">コピー元設定</param>
        public void CopyFrom(ShapeStyleSettings source)
        {
            if (source == null) return;

            EnableStyling = source.EnableStyling;
            FillColor = source.FillColor;
            LineColor = source.LineColor;
            FontName = source.FontName;
            FontColor = source.FontColor;
            FontSize = source.FontSize;
            FontBold = source.FontBold;
            FontItalic = source.FontItalic;
        }

        /// <summary>
        /// 設定が有効かつ適用可能かを判定
        /// </summary>
        /// <returns>適用可能な場合true</returns>
        public bool IsApplicable()
        {
            return EnableStyling && !string.IsNullOrEmpty(FontName) && FontSize > 0;
        }

        #endregion

        #region プライベートメソッド

        /// <summary>
        /// デフォルトフォント名を取得
        /// 環境に応じた標準フォントを返す
        /// </summary>
        /// <returns>フォント名</returns>
        private static string GetDefaultFontName()
        {
            // 日本語環境では游ゴシック、その他はCalibriを使用
            // PowerPoint COM APIの複雑性を避けて固定値で対応
            try
            {
                return System.Globalization.CultureInfo.CurrentCulture.Name.StartsWith("ja") ?
                    "游ゴシック" : "Calibri";
            }
            catch
            {
                // 最終フォールバック
                return "Calibri";
            }
        }

        #endregion

        #region 等価性比較

        /// <summary>
        /// オブジェクトの等価性を判定
        /// </summary>
        /// <param name="obj">比較対象</param>
        /// <returns>等価な場合true</returns>
        public override bool Equals(object obj)
        {
            if (obj is ShapeStyleSettings other)
            {
                return EnableStyling == other.EnableStyling &&
                       FillColor.Equals(other.FillColor) &&
                       LineColor.Equals(other.LineColor) &&
                       string.Equals(FontName, other.FontName, StringComparison.OrdinalIgnoreCase) &&
                       FontColor.Equals(other.FontColor) &&
                       Math.Abs(FontSize - other.FontSize) < 0.01f &&
                       FontBold == other.FontBold &&
                       FontItalic == other.FontItalic;
            }
            return false;
        }

        /// <summary>
        /// ハッシュコードを取得
        /// </summary>
        /// <returns>ハッシュコード</returns>
        public override int GetHashCode()
        {
            unchecked
            {
                int hash = 17;
                hash = hash * 23 + EnableStyling.GetHashCode();
                hash = hash * 23 + FillColor.GetHashCode();
                hash = hash * 23 + LineColor.GetHashCode();
                hash = hash * 23 + (FontName?.GetHashCode() ?? 0);
                hash = hash * 23 + FontColor.GetHashCode();
                hash = hash * 23 + FontSize.GetHashCode();
                hash = hash * 23 + FontBold.GetHashCode();
                hash = hash * 23 + FontItalic.GetHashCode();
                return hash;
            }
        }

        #endregion

        #region 文字列表現

        /// <summary>
        /// 設定の文字列表現を取得（デバッグ用）
        /// </summary>
        /// <returns>設定内容の文字列</returns>
        public override string ToString()
        {
            return $"ShapeStyleSettings: Enable={EnableStyling}, " +
                   $"FillColor={FillColor.Name}, LineColor={LineColor.Name}, " +
                   $"FontName={FontName}, FontColor={FontColor.Name}, " +
                   $"FontSize={FontSize}, Bold={FontBold}, Italic={FontItalic}";
        }

        #endregion
    }
}