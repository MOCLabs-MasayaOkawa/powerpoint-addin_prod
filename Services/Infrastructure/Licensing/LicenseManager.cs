using NLog;
using PowerPointEfficiencyAddin.Models.Licensing;
using PowerPointEfficiencyAddin.Services.Infrastructure.Security;
using PowerPointEfficiencyAddin.Services.Infrastructure.Update;
using System;
using System.Configuration;
using System.Net.Http;
using System.Threading.Tasks;
using System.Timers;

namespace PowerPointEfficiencyAddin.Services.Infrastructure.Licensing
{
    /// <summary>
    /// ライセンス管理の中核クラス（商用レベル実装）
    /// Phase 1: シンプルオンライン認証対応
    /// </summary>
    public sealed class LicenseManager : IDisposable
    {
        private static readonly Logger logger = LogManager.GetCurrentClassLogger();
        private static LicenseManager instance;
        private static readonly object lockObject = new object();

        private readonly LicenseApiClient apiClient;
        private readonly LicenseCache cache;
        private readonly RegistryManager registryManager;
        private readonly Timer validationTimer;

        private static readonly int validationIntervalHours;
        private static readonly int offlineGracePeriodFull;
        private static readonly int offlineGracePeriodLimited;

        private LicenseStatus currentStatus;
        private bool disposed = false;

        private UpdateService updateService;

        /// <summary>
        /// 開発モードフラグ（本番環境ではfalseに設定）
        /// </summary>
        public static bool DevelopmentMode { get; set; } = true; // TODO: App.configから読み込み

        /// <summary>
        /// シングルトンインスタンス取得
        /// </summary>
        public static LicenseManager Instance
        {
            get
            {
                if (instance == null)
                {
                    lock (lockObject)
                    {
                        if (instance == null)
                        {
                            instance = new LicenseManager();
                        }
                    }
                }
                return instance;
            }
        }

        /// <summary>
        /// 現在のライセンス状態
        /// </summary>
        public LicenseStatus CurrentStatus => currentStatus ?? new LicenseStatus();

        // 静的コンストラクタを追加
        static LicenseManager()
        {
            // App.configから設定を読み込み
            try
            {
                var mode = ConfigurationManager.AppSettings["LicenseMode"];
                DevelopmentMode = string.Equals(mode, "Development", StringComparison.OrdinalIgnoreCase);

                // 検証間隔の読み込み
                if (int.TryParse(ConfigurationManager.AppSettings["ValidationIntervalHours"], out int hours))
                {
                    validationIntervalHours = hours;
                }
                else
                {
                    validationIntervalHours = 24; // デフォルト値
                }

                // オフライン猶予期間の読み込み
                if (int.TryParse(ConfigurationManager.AppSettings["OfflineGracePeriodFull"], out int fullDays))
                {
                    offlineGracePeriodFull = fullDays;
                }
                else
                {
                    offlineGracePeriodFull = 3; // デフォルト値
                }

                if (int.TryParse(ConfigurationManager.AppSettings["OfflineGracePeriodLimited"], out int limitedDays))
                {
                    offlineGracePeriodLimited = limitedDays;
                }
                else
                {
                    offlineGracePeriodLimited = 7; // デフォルト値
                }

                logger.Info($"License configuration loaded - Mode: {(DevelopmentMode ? "Development" : "Production")}, " +
                           $"Validation Interval: {validationIntervalHours}h, " +
                           $"Grace Period: {offlineGracePeriodFull}/{offlineGracePeriodLimited} days");
            }
            catch (Exception ex)
            {
                logger.Error(ex, "Failed to load configuration from App.config, using defaults");
                DevelopmentMode = true; // エラー時は安全のため開発モードに
                validationIntervalHours = 24;
                offlineGracePeriodFull = 3;
                offlineGracePeriodLimited = 7;
            }
        }

        /// <summary>
        /// プライベートコンストラクタ（シングルトン）
        /// </summary>
        private LicenseManager()
        {
            try
            {
                logger.Info("Initializing LicenseManager");

                registryManager = new RegistryManager();
                cache = new LicenseCache(registryManager);
                apiClient = new LicenseApiClient();

                // UpdateServiceの初期化
                updateService = UpdateService.Instance;

                // タイマー間隔を設定値から取得
                validationTimer = new Timer(validationIntervalHours * 60 * 60 * 1000);
                validationTimer.Elapsed += async (s, e) => await PerformBackgroundValidation();
                validationTimer.AutoReset = true;

                logger.Info("LicenseManager initialized successfully");
            }
            catch (Exception ex)
            {
                logger.Error(ex, "Failed to initialize LicenseManager");
                throw new LicenseInitializationException("ライセンス管理システムの初期化に失敗しました", ex);
            }
        }

        /// <summary>
        /// ライセンス初期化と検証
        /// </summary>
        public async Task<LicenseValidationResult> InitializeAsync()
        {
            try
            {
                logger.Info("Starting license initialization");

                // 開発モードの場合は常に成功
                if (DevelopmentMode)
                {
                    logger.Info("Running in DEVELOPMENT MODE - all features enabled");
                    currentStatus = new LicenseStatus
                    {
                        IsValid = true,
                        AccessLevel = FeatureAccessLevel.Pro,
                        PlanType = "Development",
                        ExpiryDate = DateTime.MaxValue,
                        Message = "開発モード"
                    };
                    return LicenseValidationResult.Success("開発モードで動作中");
                }

                // キャッシュからライセンス情報を読み込み
                var cachedLicense = cache.LoadLicense();
                if (cachedLicense == null)
                {
                    logger.Warn("No license found in cache");
                    currentStatus = CreateRestrictedStatus("ライセンスが見つかりません");
                    return LicenseValidationResult.NoLicense();
                }

                // オンライン検証を試行
                var validationResult = await ValidateOnlineAsync(cachedLicense.LicenseKey);

                if (validationResult.IsSuccess)
                {
                    // 成功時はキャッシュを更新
                    cache.UpdateLastValidation(DateTime.Now);
                    validationTimer.Start();
                    logger.Info("License validated successfully online");
                }
                else
                {
                    // オフライン時は猶予期間をチェック
                    validationResult = CheckOfflineGracePeriod(cachedLicense);
                }

                currentStatus = ConvertToStatus(validationResult, cachedLicense);
                return validationResult;
            }
            catch (Exception ex)
            {
                logger.Error(ex, "License initialization failed");
                currentStatus = CreateRestrictedStatus("ライセンス検証中にエラーが発生しました");
                return LicenseValidationResult.Error(ex.Message);
            }
        }

        // 必要なアクセスレベルを取得（新規追加）
        public FeatureAccessLevel GetRequiredLevel(string featureName)
        {
            try
            {
                return FeatureRegistry.Instance.GetRequiredLevel(featureName);
            }
            catch (Exception ex)
            {
                logger.Error(ex, $"Error getting required level for '{featureName}'");
                return FeatureAccessLevel.Pro;
            }
        }

        /// <summary>
        /// ライセンスキーの設定と検証
        /// </summary>
        public async Task<LicenseValidationResult> SetLicenseKeyAsync(string licenseKey)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(licenseKey))
                {
                    return LicenseValidationResult.Invalid("ライセンスキーが入力されていません");
                }

                logger.Info($"Setting new license key: {MaskLicenseKey(licenseKey)}");

                // オンライン検証
                var result = await ValidateOnlineAsync(licenseKey);

                if (result.IsSuccess)
                {
                    // 成功時はキャッシュに保存
                    var licenseInfo = new LicenseInfo
                    {
                        LicenseKey = licenseKey,
                        UserId = result.UserId,
                        PlanType = result.PlanType,
                        ExpiryDate = result.ExpiryDate,
                        LastValidation = DateTime.Now
                    };

                    cache.SaveLicense(licenseInfo);
                    currentStatus = ConvertToStatus(result, licenseInfo);

                    // 定期検証タイマー開始
                    validationTimer.Start();

                    logger.Info("License key set and validated successfully");
                }

                return result;
            }
            catch (Exception ex)
            {
                logger.Error(ex, "Failed to set license key");
                return LicenseValidationResult.Error(ex.Message);
            }
        }

        /// <summary>
        /// 機能の利用可否をチェック
        /// </summary>
        public bool IsFeatureAllowed(string featureName)
        {
            try
            {
                // 開発モードは常に許可
                if (DevelopmentMode) return true;

                // ライセンス状態チェック
                if (currentStatus == null || !currentStatus.IsValid)
                {
                    logger.Debug($"Feature '{featureName}' blocked - invalid license");
                    return false;
                }

                // FeatureRegistryによるチェック
                var registry = FeatureRegistry.Instance;
                bool allowed = registry.IsFeatureAvailable(featureName, currentStatus.AccessLevel);

                if (!allowed)
                {
                    var requiredLevel = registry.GetRequiredLevel(featureName);
                    logger.Info($"Feature '{featureName}' requires {requiredLevel}, current: {currentStatus.AccessLevel}");
                }

                return allowed;
            }
            catch (Exception ex)
            {
                logger.Error(ex, $"Error checking feature access for '{featureName}'");
                return false;
            }
        }

        /// <summary>
        /// 処理可能なオブジェクト数をチェック
        /// </summary>
        public bool IsWithinObjectLimit(int objectCount)
        {
            if (DevelopmentMode) return true;

            if (currentStatus?.AccessLevel == FeatureAccessLevel.Free)
            {
                // 設定値から読み込み
                int limitedModeMaxObjects = 10; // デフォルト値
                if (int.TryParse(ConfigurationManager.AppSettings["LimitedModeMaxObjects"], out int configValue))
                {
                    limitedModeMaxObjects = configValue;
                }

                return objectCount <= limitedModeMaxObjects;
            }

            return currentStatus?.AccessLevel == FeatureAccessLevel.Pro;
        }

        /// <summary>
        /// ライセンス状態の文字列表現を取得
        /// </summary>
        public string GetStatusMessage()
        {
            if (DevelopmentMode)
            {
                return "開発モードで動作中（全機能利用可能）";
            }

            if (currentStatus == null)
            {
                return "ライセンス未確認";
            }

            return currentStatus.Message ?? "ライセンス状態不明";
        }

        #region Private Methods

        /// <summary>
        /// オンラインでライセンスを検証
        /// </summary>
        private async Task<LicenseValidationResult> ValidateOnlineAsync(string licenseKey)
        {
            try
            {
                logger.Debug("Attempting online validation");
                return await apiClient.ValidateLicenseAsync(licenseKey);
            }
            catch (Exception ex)
            {
                logger.Warn(ex, "Online validation failed, will check offline grace period");
                return LicenseValidationResult.NetworkError();
            }
        }

        /// <summary>
        /// オフライン猶予期間をチェック
        /// </summary>
        private LicenseValidationResult CheckOfflineGracePeriod(LicenseInfo license)
        {
            try
            {
                if (license.LastValidation == null)
                {
                    return LicenseValidationResult.Invalid("前回の認証記録がありません");
                }

                var daysSinceLastValidation = (DateTime.Now - license.LastValidation.Value).TotalDays;

                // 設定値を使用
                if (daysSinceLastValidation <= offlineGracePeriodFull)
                {
                    logger.Info($"Offline grace period: {daysSinceLastValidation:F1} days - full access");
                    return LicenseValidationResult.OfflineGrace(
                        FeatureAccessLevel.Pro,
                        $"オフラインモード（残り{offlineGracePeriodFull - (int)daysSinceLastValidation}日間）"
                    );
                }
                else if (daysSinceLastValidation <= offlineGracePeriodLimited)
                {
                    logger.Warn($"Offline grace period: {daysSinceLastValidation:F1} days - limited access");
                    return LicenseValidationResult.OfflineGrace(
                        FeatureAccessLevel.Free,
                        $"オフライン制限モード（残り{offlineGracePeriodLimited - (int)daysSinceLastValidation}日間）"
                    );
                }
                else
                {
                    logger.Error($"Offline grace period exceeded: {daysSinceLastValidation:F1} days");
                    return LicenseValidationResult.Expired("オフライン猶予期間が終了しました");
                }
            }
            catch (Exception ex)
            {
                logger.Error(ex, "Error checking offline grace period");
                return LicenseValidationResult.Error("オフライン期間の確認に失敗しました");
            }
        }

        /// <summary>
        /// バックグラウンドでの定期検証
        /// </summary>
        private async Task PerformBackgroundValidation()
        {
            try
            {
                logger.Debug("Performing background license validation");

                var license = cache.LoadLicense();
                if (license == null) return; // 先にnullチェック
                var response = await apiClient.ValidateLicenseWithUpdate(license.LicenseKey);

                if (license != null)
                {
                    var result = await ValidateOnlineAsync(license.LicenseKey);
                    if (result.IsSuccess)
                    {
                        cache.UpdateLastValidation(DateTime.Now);
                        currentStatus = ConvertToStatus(response.LicenseResult, license);
                        logger.Debug("Background validation successful");

                        if (response.UpdateInfo != null && !DevelopmentMode)
                        {
                            await ProcessUpdateAsync(response.UpdateInfo);
                        }

                    }
                    else
                    {
                        // オフライン猶予期間をチェック
                        var offlineResult = CheckOfflineGracePeriod(license);
                        currentStatus = ConvertToStatus(offlineResult, license);
                        logger.Warn($"Background validation failed: {result.Message}");
                    }
                }
            }
            catch (Exception ex)
            {
                logger.Error(ex, "Error during background validation");
            }
        }

        // 更新情報の処理
        private async Task ProcessUpdateAsync(UpdateInfo updateInfo)
        {
            try
            {
                var updateResult = await updateService.CheckForUpdateAsync(updateInfo);

                if (updateResult.UpdateAvailable)
                {
                    logger.Info($"Update available: {updateInfo.Version}");

                    // 重要更新の場合は自動ダウンロード
                    if (updateInfo.IsCritical)
                    {
                        logger.Info("Critical update detected, starting auto-download");
                        _ = Task.Run(async () =>
                        {
                            await updateService.DownloadUpdateAsync(updateInfo);
                        });
                    }
                }
            }
            catch (Exception ex)
            {
                logger.Error(ex, "Failed to process update information");
            }
        }

        /// <summary>
        /// 検証結果をステータスに変換
        /// </summary>
        private LicenseStatus ConvertToStatus(LicenseValidationResult result, LicenseInfo license)
        {
            return new LicenseStatus
            {
                IsValid = result.IsSuccess || result.AccessLevel != FeatureAccessLevel.Blocked,
                AccessLevel = result.AccessLevel,
                PlanType = result.PlanType ?? license?.PlanType ?? "Unknown",
                ExpiryDate = result.ExpiryDate ?? license?.ExpiryDate,
                Message = result.Message,
                LastValidation = license?.LastValidation
            };
        }

        /// <summary>
        /// 制限付きステータスを作成
        /// </summary>
        private LicenseStatus CreateRestrictedStatus(string message)
        {
            return new LicenseStatus
            {
                IsValid = false,
                AccessLevel = FeatureAccessLevel.Blocked,
                Message = message
            };
        }

        /// <summary>
        /// ライセンスキーをマスク表示
        /// </summary>
        private string MaskLicenseKey(string key)
        {
            if (string.IsNullOrEmpty(key) || key.Length < 8)
                return "****";

            return key.Substring(0, 5) + "****" + key.Substring(key.Length - 3);
        }

        // 新規メソッド追加：更新情報の取得
        /// <summary>
        /// 保留中の更新があるかチェック
        /// </summary>
        public bool HasPendingUpdate()
        {
            if (DevelopmentMode) return false;
            return updateService?.HasPendingUpdate() ?? false;
        }

        /// <summary>
        /// 保留中の更新情報を取得
        /// </summary>
        public UpdateInfo GetPendingUpdate()
        {
            if (DevelopmentMode) return null;
            return updateService?.GetPendingUpdate();
        }

        /// <summary>
        /// 手動で更新をダウンロード
        /// </summary>
        public async Task<bool> DownloadUpdateAsync()
        {
            if (DevelopmentMode) return false;

            var update = GetPendingUpdate();
            if (update == null) return false;

            return await updateService.DownloadUpdateAsync(update);
        }

        #endregion

        #region IDisposable

        public void Dispose()
        {
            if (!disposed)
            {
                try
                {
                    validationTimer?.Stop();
                    validationTimer?.Dispose();
                    apiClient?.Dispose();
                    updateService?.Dispose();

                    instance = null;
                    disposed = true;

                    logger.Info("LicenseManager disposed");
                }
                catch (Exception ex)
                {
                    logger.Error(ex, "Error disposing LicenseManager");
                }
            }
        }

        #endregion
    }

    /// <summary>
    /// ライセンス初期化例外
    /// </summary>
    public class LicenseInitializationException : Exception
    {
        public LicenseInitializationException(string message) : base(message) { }
        public LicenseInitializationException(string message, Exception inner) : base(message, inner) { }
    }
}