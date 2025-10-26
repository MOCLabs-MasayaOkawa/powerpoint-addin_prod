using System;
using Newtonsoft.Json;
using NLog;
using PowerPointEfficiencyAddin.Models.Licensing;
using PowerPointEfficiencyAddin.Services.Security;

namespace PowerPointEfficiencyAddin.Services.Licensing
{
    /// <summary>
    /// ライセンス情報のローカルキャッシュ管理
    /// </summary>
    public class LicenseCache
    {
        private static readonly Logger logger = LogManager.GetCurrentClassLogger();
        private readonly RegistryManager registryManager;

        // レジストリキー名
        private const string KEY_LICENSE_DATA = "LicenseData";
        private const string KEY_LAST_VALIDATION = "LastValidation";
        private const string KEY_LICENSE_KEY = "LicenseKey";

        public LicenseCache(RegistryManager registryManager)
        {
            this.registryManager = registryManager ?? throw new ArgumentNullException(nameof(registryManager));
        }

        /// <summary>
        /// ライセンス情報を保存
        /// </summary>
        public bool SaveLicense(LicenseInfo license)
        {
            if (license == null)
            {
                logger.Warn("Attempted to save null license");
                return false;
            }

            try
            {
                // ライセンス情報をJSON形式でシリアライズ
                var jsonData = JsonConvert.SerializeObject(license, new JsonSerializerSettings
                {
                    DateFormatHandling = DateFormatHandling.IsoDateFormat,
                    NullValueHandling = NullValueHandling.Ignore
                });

                // レジストリに保存
                bool success = registryManager.SaveSecureString(KEY_LICENSE_DATA, jsonData);

                if (success)
                {
                    // ライセンスキーは別途保存（高速アクセス用）
                    registryManager.SaveSecureString(KEY_LICENSE_KEY, license.LicenseKey);

                    // 最終検証日時を更新
                    if (license.LastValidation.HasValue)
                    {
                        registryManager.SaveDateTime(KEY_LAST_VALIDATION, license.LastValidation.Value);
                    }

                    logger.Info("License information saved to cache");
                }

                return success;
            }
            catch (Exception ex)
            {
                logger.Error(ex, "Failed to save license to cache");
                return false;
            }
        }

        /// <summary>
        /// ライセンス情報を読み込み
        /// </summary>
        public LicenseInfo LoadLicense()
        {
            try
            {
                // JSONデータを読み込み
                string jsonData = registryManager.LoadSecureString(KEY_LICENSE_DATA);

                if (string.IsNullOrEmpty(jsonData))
                {
                    logger.Debug("No license data found in cache");
                    return null;
                }

                // デシリアライズ
                var license = JsonConvert.DeserializeObject<LicenseInfo>(jsonData);

                if (license != null)
                {
                    // 最終検証日時を別途読み込み（互換性のため）
                    var lastValidation = registryManager.LoadDateTime(KEY_LAST_VALIDATION);
                    if (lastValidation.HasValue)
                    {
                        license.LastValidation = lastValidation.Value;
                    }

                    logger.Debug("License information loaded from cache");
                }

                return license;
            }
            catch (JsonException ex)
            {
                logger.Error(ex, "Failed to deserialize license data");
                return null;
            }
            catch (Exception ex)
            {
                logger.Error(ex, "Failed to load license from cache");
                return null;
            }
        }

        /// <summary>
        /// ライセンスキーのみを高速読み込み
        /// </summary>
        public string GetCachedLicenseKey()
        {
            try
            {
                return registryManager.LoadSecureString(KEY_LICENSE_KEY);
            }
            catch (Exception ex)
            {
                logger.Error(ex, "Failed to load cached license key");
                return null;
            }
        }

        /// <summary>
        /// 最終検証日時を更新
        /// </summary>
        public bool UpdateLastValidation(DateTime validationTime)
        {
            try
            {
                // 既存のライセンス情報を読み込み
                var license = LoadLicense();
                if (license != null)
                {
                    license.LastValidation = validationTime;
                    return SaveLicense(license);
                }

                // ライセンス情報がない場合は日時のみ保存
                return registryManager.SaveDateTime(KEY_LAST_VALIDATION, validationTime);
            }
            catch (Exception ex)
            {
                logger.Error(ex, "Failed to update last validation time");
                return false;
            }
        }

        /// <summary>
        /// キャッシュをクリア
        /// </summary>
        public bool ClearCache()
        {
            try
            {
                return registryManager.ClearAllLicenseData();
            }
            catch (Exception ex)
            {
                logger.Error(ex, "Failed to clear license cache");
                return false;
            }
        }

        /// <summary>
        /// キャッシュが存在するか確認
        /// </summary>
        public bool HasCachedLicense()
        {
            try
            {
                return registryManager.HasLicenseData();
            }
            catch (Exception ex)
            {
                logger.Error(ex, "Failed to check cached license existence");
                return false;
            }
        }
    }
}