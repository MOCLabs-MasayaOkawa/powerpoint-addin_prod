using System;
using System.Collections.Generic;
using System.Drawing.Drawing2D;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PowerPointEfficiencyAddin.Models.Licensing
{
    /// <summary>
    /// 現在のライセンス状態
    /// </summary>
    public class LicenseStatus
    {
        /// <summary>
        /// ライセンスが有効かどうか
        /// </summary>
        public bool IsValid { get; set; }

        /// <summary>
        /// アクセスレベル
        /// </summary>
        public FeatureAccessLevel AccessLevel { get; set; }

        /// <summary>
        /// 有効期限
        /// </summary>
        public DateTime? ExpiryDate { get; set; }

        /// <summary>
        /// 最終検証日時
        /// </summary>
        public DateTime? LastValidation { get; set; }

        /// <summary>
        /// ステータスメッセージ
        /// </summary>
        public string Message { get; set; }

        /// <summary>
        /// オフラインモードかどうか
        /// </summary>
        public bool IsOfflineMode => LastValidation.HasValue &&
            (DateTime.Now - LastValidation.Value).TotalHours > 1;

        // PlanTypeプロパティを修正（文字列から自動変換）
        private string planType;

        public string PlanType
        {
            get => AccessLevel.GetDisplayName();
            set
            {
                planType = value;
                // APIからの文字列を自動変換
                AccessLevel = ParsePlanType(value);
            }
        }

        /// <summary>
        /// 残りのオフライン猶予日数
        /// </summary>
        public int GetOfflineGraceDaysRemaining()
        {
            if (!LastValidation.HasValue) return 0;

            var daysSinceLastValidation = (DateTime.Now - LastValidation.Value).TotalDays;

            if (daysSinceLastValidation <= 3)
                return (int)(3 - daysSinceLastValidation);
            else if (daysSinceLastValidation <= 7)
                return (int)(7 - daysSinceLastValidation);
            else
                return 0;
        }

        private FeatureAccessLevel ParsePlanType(string plan)
        {
            if (string.IsNullOrEmpty(plan)) return FeatureAccessLevel.Free;

            switch (plan.ToLower())
            {
                case "free": return FeatureAccessLevel.Free;
                case "starter": return FeatureAccessLevel.Starter;
                case "growth": return FeatureAccessLevel.Growth;
                case "pro":
                case "premium": return FeatureAccessLevel.Pro;
                case "development": return FeatureAccessLevel.Development;

                // 旧バージョン互換
                case "basic": return FeatureAccessLevel.Free;
                case "limited": return FeatureAccessLevel.Free;
                case "full": return FeatureAccessLevel.Pro;

                default: return FeatureAccessLevel.Free;
            }
        }
    }
}