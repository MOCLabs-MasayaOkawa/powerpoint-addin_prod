using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PowerPointEfficiencyAddin.Models.Licensing
{
    /// <summary>
    /// 機能アクセスレベル（プラン別）
    /// </summary>
    public enum FeatureAccessLevel
    {
        /// <summary>完全ブロック（ライセンス無効）</summary>
        Blocked = 0,

        /// <summary>無料プラン（基本機能のみ）</summary>
        Free = 1,

        /// <summary>スタータープラン</summary>
        Starter = 2,

        /// <summary>グロースプラン</summary>
        Growth = 3,

        /// <summary>プロプラン（全機能）</summary>
        Pro = 4,

        /// <summary>開発モード（全機能利用可）</summary>
        Development = 99
    }

    /// <summary>
    /// FeatureAccessLevel拡張メソッド
    /// </summary>
    public static class FeatureAccessLevelExtensions
    {
        /// <summary>
        /// プラン表示名を取得
        /// </summary>
        public static string GetDisplayName(this FeatureAccessLevel level)
        {
            switch (level)
            {
                case FeatureAccessLevel.Blocked: return "ライセンスなし";
                case FeatureAccessLevel.Free: return "無料版";
                case FeatureAccessLevel.Starter: return "スターター";
                case FeatureAccessLevel.Growth: return "グロース";
                case FeatureAccessLevel.Pro: return "プロ";
                case FeatureAccessLevel.Development: return "開発版";
                default: return "不明";
            }
        }

        /// <summary>
        /// 指定レベル以上かチェック
        /// </summary>
        public static bool IsAtLeast(this FeatureAccessLevel current, FeatureAccessLevel required)
        {
            if (current == FeatureAccessLevel.Development) return true;
            return current >= required;
        }

        /// <summary>
        /// 旧バージョンとの互換性（Limited/Full）
        /// </summary>
        public static FeatureAccessLevel FromLegacyLevel(string legacyLevel)
        {
            switch (legacyLevel?.ToLower())
            {
                case "limited": return FeatureAccessLevel.Free;
                case "full": return FeatureAccessLevel.Pro;
                case "blocked": return FeatureAccessLevel.Blocked;
                default: return FeatureAccessLevel.Free;
            }
        }
    }
}