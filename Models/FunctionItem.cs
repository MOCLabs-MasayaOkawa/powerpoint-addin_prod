using Microsoft.Office.Interop.PowerPoint;
using System;
using System.Drawing;

namespace PowerPointEfficiencyAddin.Models
{
    /// <summary>
    /// カスタムペイン機能項目定義
    /// </summary>
    public class FunctionItem
    {
        public string Id { get; set; }
        public string Name { get; set; }
        public string Description { get; set; }
        public string IconPath { get; set; }
        public Action Action { get; set; }
        public FunctionCategory Category { get; set; }
        public bool IsEnabled { get; set; } = true;
        public int RowPosition { get; set; } = 1; // PDF配置表の行位置
        public int Order { get; set; } = 0; // 行内での順序（0から開始）
        public bool IsBuiltIn { get; set; } = false; // 標準機能かどうか

        public FunctionItem(string id, string name, string description,
            string iconPath, Action action, FunctionCategory category, int rowPosition = 1, int order = 0, bool isBuiltIn = false)
        {
            Id = id;
            Name = name;
            Description = description;
            IconPath = iconPath;
            Action = action;
            Category = category;
            RowPosition = rowPosition;
            Order = order;
            IsBuiltIn = isBuiltIn;
        }

        /// <summary>
        /// アイコンを取得（埋め込みリソース方式）
        /// </summary>
        public Bitmap GetIcon()
        {
            try
            {
                // 埋め込みリソースから読み込み
                var assembly = System.Reflection.Assembly.GetExecutingAssembly();
                var resourceName = $"PowerPointEfficiencyAddin.Resources.Icons.{IconPath}";

                NLog.LogManager.GetCurrentClassLogger().Debug($"Looking for embedded resource: {resourceName}");

                using (var stream = assembly.GetManifestResourceStream(resourceName))
                {
                    if (stream != null)
                    {
                        var bitmap = new Bitmap(stream);
                        NLog.LogManager.GetCurrentClassLogger().Info($"Successfully loaded embedded icon: {IconPath}");
                        return bitmap;
                    }
                    else
                    {
                        // リソース名の形式を変更して再試行
                        var alternativeResourceName = $"PowerPointEfficiencyAddin.Resources.icons.{IconPath}";
                        NLog.LogManager.GetCurrentClassLogger().Debug($"Retrying with alternative resource name: {alternativeResourceName}");

                        using (var altStream = assembly.GetManifestResourceStream(alternativeResourceName))
                        {
                            if (altStream != null)
                            {
                                var bitmap = new Bitmap(altStream);
                                NLog.LogManager.GetCurrentClassLogger().Info($"Successfully loaded alternative embedded icon: {IconPath}");
                                return bitmap;
                            }
                        }

                        // デフォルトアイコンを返す
                        NLog.LogManager.GetCurrentClassLogger().Warn($"Icon resource not found: {IconPath}, using default");
                        return CreateDefaultIcon();
                    }
                }
            }
            catch (Exception ex)
            {
                NLog.LogManager.GetCurrentClassLogger().Error(ex, $"Failed to load icon: {IconPath}");
                return CreateDefaultIcon();
            }
        }

        /// <summary>
        /// デフォルトアイコンを作成
        /// </summary>
        private Bitmap CreateDefaultIcon()
        {
            var bitmap = new Bitmap(16, 16);
            using (var g = Graphics.FromImage(bitmap))
            {
                // カテゴリ別の色でデフォルトアイコンを作成
                var color = GetCategoryColor();
                using (var brush = new SolidBrush(color))
                {
                    g.FillRectangle(brush, 0, 0, 16, 16);
                }

                // 機能名の頭文字を描画
                using (var font = new System.Drawing.Font("Arial", 8, System.Drawing.FontStyle.Bold))
                using (var textBrush = new SolidBrush(Color.White))
                {
                    var text = GetShortName();
                    var size = g.MeasureString(text, font);
                    var x = (16 - size.Width) / 2;
                    var y = (16 - size.Height) / 2;
                    g.DrawString(text, font, textBrush, x, y);
                }
            }
            return bitmap;
        }

        /// <summary>
        /// 機能名の短縮形を取得
        /// </summary>
        public string GetShortName()
        {
            try
            {
                if (string.IsNullOrEmpty(Name))
                    return "？";

                if (Name.Length >= 2)
                {
                    return Name.Substring(0, 2);
                }
                else
                {
                    return Name.Length > 0 ? Name.Substring(0, 1) : "？";
                }
            }
            catch
            {
                return "？";
            }
        }

        /// <summary>
        /// カテゴリ別の色を取得
        /// </summary>
        public Color GetCategoryColor()
        {
            switch (Category)
            {
                case FunctionCategory.Selection:
                    return Color.FromArgb(52, 152, 219); // 青
                case FunctionCategory.Text:
                    return Color.FromArgb(241, 196, 15); // 黄
                case FunctionCategory.Shape:
                    return Color.FromArgb(46, 204, 113); // 緑
                case FunctionCategory.Format:
                    return Color.FromArgb(155, 89, 182); // 紫
                case FunctionCategory.Grouping:
                    return Color.FromArgb(230, 126, 34); // オレンジ
                case FunctionCategory.Alignment:
                    return Color.FromArgb(231, 76, 60); // 赤
                case FunctionCategory.ShapeOperation:
                    return Color.FromArgb(52, 73, 94); // ダークブルー
                case FunctionCategory.TableOperation:
                    return Color.FromArgb(26, 188, 156); // ターコイズ
                case FunctionCategory.Spacing:
                    return Color.FromArgb(142, 68, 173); // パープル
                case FunctionCategory.PowerTool:
                    return Color.FromArgb(192, 57, 43); // ダークレッド
                default:
                    return Color.FromArgb(127, 140, 141); // デフォルトグレー
            }
        }
    }

    /// <summary>
    /// 機能カテゴリ（PDF配置表対応）
    /// </summary>
    public enum FunctionCategory
    {
        /// <summary>
        /// 選択
        /// </summary>
        Selection,

        /// <summary>
        /// テキスト
        /// </summary>
        Text,

        /// <summary>
        /// 図形
        /// </summary>
        Shape,

        /// <summary>
        /// 整形
        /// </summary>
        Format,

        /// <summary>
        /// グループ化
        /// </summary>
        Grouping,

        /// <summary>
        /// 整列
        /// </summary>
        Alignment,

        /// <summary>
        /// 図形操作プロ
        /// </summary>
        ShapeOperation,

        /// <summary>
        /// 表操作
        /// </summary>
        TableOperation,

        /// <summary>
        /// 間隔
        /// </summary>
        Spacing,

        /// <summary>
        /// PowerTool
        /// </summary>
        PowerTool
    }

    /// <summary>
    /// カテゴリ情報
    /// </summary>
    public class CategoryInfo
    {
        public FunctionCategory Category { get; set; }
        public string DisplayName { get; set; }
        public Color HeaderColor { get; set; }

        public static CategoryInfo[] GetAllCategories()
        {
            return new[]
            {
                new CategoryInfo
                {
                    Category = FunctionCategory.Selection,
                    DisplayName = "選択",
                    HeaderColor = Color.FromArgb(234, 246, 253)
                },
                new CategoryInfo
                {
                    Category = FunctionCategory.Text,
                    DisplayName = "テキスト",
                    HeaderColor = Color.FromArgb(234, 246, 253)
                },
                new CategoryInfo
                {
                    Category = FunctionCategory.Shape,
                    DisplayName = "図形",
                    HeaderColor = Color.FromArgb(234, 246, 253)
                },
                new CategoryInfo
                {
                    Category = FunctionCategory.Format,
                    DisplayName = "整形",
                    HeaderColor = Color.FromArgb(234, 246, 253)
                },
                new CategoryInfo
                {
                    Category = FunctionCategory.Grouping,
                    DisplayName = "グループ化",
                    HeaderColor = Color.FromArgb(234, 246, 253)
                },
                new CategoryInfo
                {
                    Category = FunctionCategory.Alignment,
                    DisplayName = "整列",
                    HeaderColor = Color.FromArgb(234, 246, 253)
                },
                new CategoryInfo
                {
                    Category = FunctionCategory.ShapeOperation,
                    DisplayName = "図形操作プロ",
                    HeaderColor = Color.FromArgb(234, 246, 253)
                },
                new CategoryInfo
                {
                    Category = FunctionCategory.TableOperation,
                    DisplayName = "表操作",
                    HeaderColor = Color.FromArgb(234, 246, 253)
                },
                new CategoryInfo
                {
                    Category = FunctionCategory.Spacing,
                    DisplayName = "間隔",
                    HeaderColor = Color.FromArgb(234, 246, 253)
                },
                new CategoryInfo
                {
                    Category = FunctionCategory.PowerTool,
                    DisplayName = "PowerTool",
                    HeaderColor = Color.FromArgb(234, 246, 253)
                }
            };
        }
    }
}