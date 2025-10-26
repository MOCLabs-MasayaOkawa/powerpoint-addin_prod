using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PowerPointEfficiencyAddin.Models.Licensing
{
    /// <summary>
    /// ライセンス情報を保持するモデルクラス
    /// </summary>
    public class LicenseInfo
    {
        /// <summary>
        /// ライセンスキー
        /// </summary>
        public string LicenseKey { get; set; }

        /// <summary>
        /// ユーザーID
        /// </summary>
        public string UserId { get; set; }

        /// <summary>
        /// プランタイプ (Basic/Premium/Enterprise)
        /// </summary>
        public string PlanType { get; set; }

        /// <summary>
        /// 有効期限
        /// </summary>
        public DateTime? ExpiryDate { get; set; }

        /// <summary>
        /// 最終検証日時
        /// </summary>
        public DateTime? LastValidation { get; set; }

        /// <summary>
        /// ライセンス開始日
        /// </summary>
        public DateTime? StartDate { get; set; }

        /// <summary>
        /// ライセンスが有効かどうか
        /// </summary>
        public bool IsExpired => ExpiryDate.HasValue && ExpiryDate.Value < DateTime.Now;

        /// <summary>
        /// 残り日数を取得
        /// </summary>
        public int? GetRemainingDays()
        {
            if (!ExpiryDate.HasValue) return null;
            var remaining = (ExpiryDate.Value - DateTime.Now).TotalDays;
            return remaining > 0 ? (int)remaining : 0;
        }
    }
}