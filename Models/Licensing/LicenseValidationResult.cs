using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PowerPointEfficiencyAddin.Models.Licensing
{
    /// <summary>
    /// ライセンス検証結果
    /// </summary>
    public class LicenseValidationResult
    {
        /// <summary>
        /// 検証が成功したかどうか
        /// </summary>
        public bool IsSuccess { get; set; }

        /// <summary>
        /// エラーメッセージまたはステータスメッセージ
        /// </summary>
        public string Message { get; set; }

        /// <summary>
        /// アクセスレベル
        /// </summary>
        public FeatureAccessLevel AccessLevel { get; set; }

        /// <summary>
        /// ユーザーID
        /// </summary>
        public string UserId { get; set; }

        /// <summary>
        /// プランタイプ
        /// </summary>
        public string PlanType { get; set; }

        /// <summary>
        /// 有効期限
        /// </summary>
        public DateTime? ExpiryDate { get; set; }

        /// <summary>
        /// 検証タイプ
        /// </summary>
        public ValidationType Type { get; set; }

        // 静的ファクトリメソッド

        public static LicenseValidationResult Success(string message, string userId = null,
            string planType = null, DateTime? expiryDate = null)
        {
            return new LicenseValidationResult
            {
                IsSuccess = true,
                Message = message,
                AccessLevel = FeatureAccessLevel.Pro,
                UserId = userId,
                PlanType = planType,
                ExpiryDate = expiryDate,
                Type = ValidationType.Online
            };
        }

        public static LicenseValidationResult Invalid(string message)
        {
            return new LicenseValidationResult
            {
                IsSuccess = false,
                Message = message,
                AccessLevel = FeatureAccessLevel.Blocked,
                Type = ValidationType.Invalid
            };
        }

        public static LicenseValidationResult Expired(string message)
        {
            return new LicenseValidationResult
            {
                IsSuccess = false,
                Message = message,
                AccessLevel = FeatureAccessLevel.Blocked,
                Type = ValidationType.Expired
            };
        }

        public static LicenseValidationResult NetworkError()
        {
            return new LicenseValidationResult
            {
                IsSuccess = false,
                Message = "ネットワークエラーが発生しました",
                AccessLevel = FeatureAccessLevel.Blocked,
                Type = ValidationType.NetworkError
            };
        }

        public static LicenseValidationResult OfflineGrace(FeatureAccessLevel level, string message)
        {
            return new LicenseValidationResult
            {
                IsSuccess = level == FeatureAccessLevel.Pro,
                Message = message,
                AccessLevel = level,
                Type = ValidationType.OfflineGrace
            };
        }

        public static LicenseValidationResult NoLicense()
        {
            return new LicenseValidationResult
            {
                IsSuccess = false,
                Message = "ライセンスが登録されていません",
                AccessLevel = FeatureAccessLevel.Blocked,
                Type = ValidationType.NoLicense
            };
        }

        public static LicenseValidationResult Error(string message)
        {
            return new LicenseValidationResult
            {
                IsSuccess = false,
                Message = message,
                AccessLevel = FeatureAccessLevel.Blocked,
                Type = ValidationType.Error
            };
        }
    }

    /// <summary>
    /// 検証タイプ
    /// </summary>
    public enum ValidationType
    {
        /// <summary>オンライン検証成功</summary>
        Online,
        /// <summary>オフライン猶予期間</summary>
        OfflineGrace,
        /// <summary>無効なライセンス</summary>
        Invalid,
        /// <summary>期限切れ</summary>
        Expired,
        /// <summary>ネットワークエラー</summary>
        NetworkError,
        /// <summary>ライセンスなし</summary>
        NoLicense,
        /// <summary>その他のエラー</summary>
        Error
    }
}