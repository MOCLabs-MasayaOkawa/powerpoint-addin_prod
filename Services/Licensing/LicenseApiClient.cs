using Newtonsoft.Json;
using NLog;
using PowerPointEfficiencyAddin.Models.Licensing;
using System;
using System.Configuration;
using System.Net;
using System.Net.Http;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace PowerPointEfficiencyAddin.Services.Licensing
{
    /// <summary>
    /// ライセンス認証APIクライアント
    /// </summary>
    public class LicenseApiClient : IDisposable
    {
        private static readonly Logger logger = LogManager.GetCurrentClassLogger();
        private HttpClient httpClient;
        private readonly string baseUrl;
        private readonly int timeoutSeconds;
        private readonly int maxRetryCount;
        private bool disposed = false;

        // API設定（本番環境では App.config から読み込む）
        private const string DEFAULT_BASE_URL = "https://api.yourcompany.com"; // TODO: 実際のURLに変更
        private const int TIMEOUT_SECONDS = 30;
        private const int MAX_RETRY_COUNT = 3;

        // 更新情報を含むレスポンス構造体
        public class LicenseValidationWithUpdateResponse
        {
            public LicenseValidationResult LicenseResult { get; set; }
            public UpdateInfo UpdateInfo { get; set; }
        }

        public LicenseApiClient(string baseUrl = null)
        {
            // App.configから設定を読み込み
            this.baseUrl = baseUrl ?? GetApiBaseUrl();

            // タイムアウト設定
            if (!int.TryParse(ConfigurationManager.AppSettings["LicenseApiTimeout"], out timeoutSeconds))
            {
                timeoutSeconds = 30; // デフォルト値
            }

            // リトライ回数設定
            if (!int.TryParse(ConfigurationManager.AppSettings["LicenseApiRetryCount"], out maxRetryCount))
            {
                maxRetryCount = 3; // デフォルト値
            }

            // HttpClientの初期化
            httpClient = new HttpClient
            {
                BaseAddress = new Uri(this.baseUrl),
                Timeout = TimeSpan.FromSeconds(timeoutSeconds)
            };

            // プロキシ設定の適用
            ConfigureProxy();

            httpClient.DefaultRequestHeaders.Add("User-Agent", "PowerPointEfficiencyAddin/1.0");
            httpClient.DefaultRequestHeaders.Add("Accept", "application/json");

            logger.Debug($"LicenseApiClient initialized - URL: {this.baseUrl}, Timeout: {timeoutSeconds}s, Retry: {maxRetryCount}");
        }

        private void ConfigureProxy()
        {
            try
            {
                var useSystemProxy = ConfigurationManager.AppSettings["UseSystemProxy"];
                if (string.Equals(useSystemProxy, "true", StringComparison.OrdinalIgnoreCase))
                {
                    // システムプロキシを使用
                    var handler = new HttpClientHandler
                    {
                        UseProxy = true,
                        Proxy = WebRequest.GetSystemWebProxy()
                    };

                    // カスタムプロキシが指定されている場合
                    var proxyAddress = ConfigurationManager.AppSettings["ProxyAddress"];
                    var proxyPort = ConfigurationManager.AppSettings["ProxyPort"];

                    if (!string.IsNullOrEmpty(proxyAddress) && !string.IsNullOrEmpty(proxyPort))
                    {
                        handler.Proxy = new WebProxy($"{proxyAddress}:{proxyPort}");
                        logger.Info($"Using custom proxy: {proxyAddress}:{proxyPort}");
                    }

                    // HttpClientを再作成
                    httpClient?.Dispose();
                    httpClient = new HttpClient(handler)
                    {
                        BaseAddress = new Uri(this.baseUrl),
                        Timeout = TimeSpan.FromSeconds(timeoutSeconds)
                    };
                }
            }
            catch (Exception ex)
            {
                logger.Warn(ex, "Failed to configure proxy, using direct connection");
            }
        }

        /// <summary>
        /// ライセンスキーを検証
        /// </summary>
        public async Task<LicenseValidationResult> ValidateLicenseAsync(string licenseKey)
        {
            if (string.IsNullOrWhiteSpace(licenseKey))
            {
                return LicenseValidationResult.Invalid("ライセンスキーが指定されていません");
            }

            // 開発モードの場合はモック応答を返す
            if (LicenseManager.DevelopmentMode)
            {
                return CreateMockValidationResult(licenseKey);
            }

            int retryCount = 0;
            Exception lastException = null;

            while (retryCount < maxRetryCount)
            {
                try
                {
                    logger.Debug($"Validating license (attempt {retryCount + 1}/{MAX_RETRY_COUNT})");

                    // リクエストボディ作成
                    var request = new
                    {
                        license_key = licenseKey,
                        machine_id = GetMachineId(),
                        version = GetAddinVersion()
                    };

                    var json = JsonConvert.SerializeObject(request);
                    var content = new StringContent(json, Encoding.UTF8, "application/json");

                    // API呼び出し
                    var response = await httpClient.PostAsync("/api/license/validate", content);

                    if (response.IsSuccessStatusCode)
                    {
                        var responseJson = await response.Content.ReadAsStringAsync();
                        var result = ParseValidationResponse(responseJson);

                        logger.Info("License validation successful");
                        return result;
                    }
                    else if (response.StatusCode == System.Net.HttpStatusCode.Unauthorized)
                    {
                        logger.Warn("License validation failed: Unauthorized");
                        return LicenseValidationResult.Invalid("無効なライセンスキーです");
                    }
                    else if (response.StatusCode == System.Net.HttpStatusCode.Forbidden)
                    {
                        logger.Warn("License validation failed: Forbidden");
                        return LicenseValidationResult.Expired("ライセンスの有効期限が切れています");
                    }
                    else
                    {
                        logger.Warn($"License validation failed with status: {response.StatusCode}");
                        lastException = new HttpRequestException($"Server returned {response.StatusCode}");
                    }
                }
                catch (TaskCanceledException ex)
                {
                    logger.Warn($"License validation timeout (attempt {retryCount + 1})");
                    lastException = ex;
                }
                catch (HttpRequestException ex)
                {
                    logger.Warn($"Network error during license validation (attempt {retryCount + 1}): {ex.Message}");
                    lastException = ex;
                }
                catch (Exception ex)
                {
                    logger.Error(ex, "Unexpected error during license validation");
                    lastException = ex;
                    break; // 予期しないエラーの場合はリトライしない
                }

                retryCount++;

                // リトライ前に少し待機（指数バックオフ）
                if (retryCount < MAX_RETRY_COUNT)
                {
                    await Task.Delay(TimeSpan.FromSeconds(Math.Pow(2, retryCount)));
                }
            }

            // すべてのリトライが失敗
            logger.Error($"License validation failed after {retryCount} attempts");
            return LicenseValidationResult.NetworkError();
        }

        /// <summary>
        /// ライセンス検証と更新チェックを同時実行（新規追加メソッド）
        /// </summary>
        public async Task<LicenseValidationWithUpdateResponse> ValidateLicenseWithUpdate(string licenseKey)
        {
            if (string.IsNullOrWhiteSpace(licenseKey))
            {
                return new LicenseValidationWithUpdateResponse
                {
                    LicenseResult = LicenseValidationResult.Invalid("ライセンスキーが指定されていません"),
                    UpdateInfo = null
                };
            }

            // 開発モードの場合
            if (LicenseManager.DevelopmentMode)
            {
                return new LicenseValidationWithUpdateResponse
                {
                    LicenseResult = CreateMockValidationResult(licenseKey),
                    UpdateInfo = new UpdateInfo
                    {
                        Version = "99.99.99", // 開発モードでは更新なし
                        ReleaseDate = DateTime.Now
                    }
                };
            }

            try
            {
                logger.Debug("Validating license with update check");

                // リクエストボディ作成
                var request = new
                {
                    license_key = licenseKey,
                    machine_id = GetMachineId(),
                    version = GetAddinVersion(),
                    include_update = true  // 更新情報を含めるフラグ
                };

                var json = JsonConvert.SerializeObject(request);
                var content = new StringContent(json, Encoding.UTF8, "application/json");

                var response = await httpClient.PostAsync("/api/license/validate", content);

                if (response.IsSuccessStatusCode)
                {
                    var responseJson = await response.Content.ReadAsStringAsync();

                    // レスポンス解析
                    dynamic parsedResponse = JsonConvert.DeserializeObject(responseJson);

                    // ライセンス情報を解析
                    var licenseResult = ParseValidationResponse(responseJson);

                    // 更新情報を解析（存在する場合）
                    UpdateInfo updateInfo = null;
                    if (parsedResponse.update_info != null)
                    {
                        updateInfo = new UpdateInfo
                        {
                            Version = parsedResponse.update_info.version?.ToString(),
                            ReleaseDate = parsedResponse.update_info.release_date != null
                                ? DateTime.Parse(parsedResponse.update_info.release_date.ToString())
                                : DateTime.Now,
                            DownloadUrl = parsedResponse.update_info.download_url?.ToString(),
                            Checksum = parsedResponse.update_info.checksum?.ToString(),
                            FileSize = parsedResponse.update_info.file_size ?? 0,
                            IsCritical = parsedResponse.update_info.is_critical ?? false,
                            ReleaseNotes = parsedResponse.update_info.release_notes?.ToString(),
                            MinimumVersion = parsedResponse.update_info.minimum_version?.ToString()
                        };
                    }

                    return new LicenseValidationWithUpdateResponse
                    {
                        LicenseResult = licenseResult,
                        UpdateInfo = updateInfo
                    };
                }
                else
                {
                    // エラーレスポンス処理
                    if (response.StatusCode == System.Net.HttpStatusCode.Unauthorized)
                    {
                        return new LicenseValidationWithUpdateResponse
                        {
                            LicenseResult = LicenseValidationResult.Invalid("無効なライセンスキーです"),
                            UpdateInfo = null
                        };
                    }
                    else if (response.StatusCode == System.Net.HttpStatusCode.Forbidden)
                    {
                        return new LicenseValidationWithUpdateResponse
                        {
                            LicenseResult = LicenseValidationResult.Expired("ライセンスの有効期限が切れています"),
                            UpdateInfo = null
                        };
                    }
                    else
                    {
                        logger.Warn($"License validation failed with status: {response.StatusCode}");
                        return new LicenseValidationWithUpdateResponse
                        {
                            LicenseResult = LicenseValidationResult.NetworkError(),
                            UpdateInfo = null
                        };
                    }
                }
            }
            catch (TaskCanceledException)
            {
                logger.Warn("License validation timeout");
                return new LicenseValidationWithUpdateResponse
                {
                    LicenseResult = LicenseValidationResult.NetworkError(),
                    UpdateInfo = null
                };
            }
            catch (HttpRequestException ex)
            {
                logger.Warn($"Network error during license validation: {ex.Message}");
                return new LicenseValidationWithUpdateResponse
                {
                    LicenseResult = LicenseValidationResult.NetworkError(),
                    UpdateInfo = null
                };
            }
            catch (Exception ex)
            {
                logger.Error(ex, "Unexpected error during license validation with update check");

                // フォールバック：既存のValidateLicenseAsyncを使用
                var fallbackResult = await ValidateLicenseAsync(licenseKey);
                return new LicenseValidationWithUpdateResponse
                {
                    LicenseResult = fallbackResult,
                    UpdateInfo = null
                };
            }
        }

        /// <summary>
        /// ハートビート送信（使用状況記録）
        /// </summary>
        public async Task<bool> SendHeartbeatAsync(string licenseKey)
        {
            if (LicenseManager.DevelopmentMode)
            {
                return true;
            }

            try
            {
                var request = new
                {
                    license_key = licenseKey,
                    machine_id = GetMachineId(),
                    timestamp = DateTime.UtcNow
                };

                var json = JsonConvert.SerializeObject(request);
                var content = new StringContent(json, Encoding.UTF8, "application/json");

                var response = await httpClient.PostAsync("/api/license/heartbeat", content);
                return response.IsSuccessStatusCode;
            }
            catch (Exception ex)
            {
                logger.Debug(ex, "Heartbeat failed (non-critical)");
                return false;
            }
        }

        #region Private Methods

        /// <summary>
        /// API応答を解析
        /// </summary>
        private LicenseValidationResult ParseValidationResponse(string json)
        {
            try
            {
                dynamic response = JsonConvert.DeserializeObject(json);

                bool isValid = response.valid ?? false;

                if (isValid)
                {
                    return LicenseValidationResult.Success(
                        "ライセンスが確認されました",
                        userId: response.user_id?.ToString(),
                        planType: response.plan_type?.ToString(),
                        expiryDate: response.end_date != null ?
                            DateTime.Parse(response.end_date.ToString()) : (DateTime?)null
                    );
                }
                else
                {
                    string reason = response.reason?.ToString() ?? "Unknown";

                    switch (reason.ToLower())
                    {
                        case "expired":
                            return LicenseValidationResult.Expired("ライセンスの有効期限が切れています");
                        case "suspended":
                            return LicenseValidationResult.Invalid("ライセンスが一時停止されています");
                        default:
                            return LicenseValidationResult.Invalid($"ライセンスが無効です: {reason}");
                    }
                }
            }
            catch (Exception ex)
            {
                logger.Error(ex, "Failed to parse validation response");
                return LicenseValidationResult.Error("サーバー応答の解析に失敗しました");
            }
        }

        /// <summary>
        /// 開発モード用のモック応答を作成
        /// </summary>
        private LicenseValidationResult CreateMockValidationResult(string licenseKey)
        {
            logger.Debug("Creating mock validation result for development mode");

            // テスト用の特殊なライセンスキー
            if (licenseKey.StartsWith("EXPIRED"))
            {
                return LicenseValidationResult.Expired("テスト: 期限切れライセンス");
            }
            else if (licenseKey.StartsWith("INVALID"))
            {
                return LicenseValidationResult.Invalid("テスト: 無効なライセンス");
            }
            else
            {
                return LicenseValidationResult.Success(
                    "テスト: 有効なライセンス",
                    userId: "test-user",
                    planType: "Premium",
                    expiryDate: DateTime.Now.AddDays(30)
                );
            }
        }

        /// <summary>
        /// APIベースURLを取得
        /// </summary>
        private string GetApiBaseUrl()
        {
            var configUrl = ConfigurationManager.AppSettings["LicenseApiUrl"];
            if (!string.IsNullOrEmpty(configUrl))
            {
                logger.Info($"Using API URL from configuration: {configUrl}");
                return configUrl;
            }

            logger.Warn($"API URL not found in configuration, using default: {DEFAULT_BASE_URL}");
            return DEFAULT_BASE_URL;
        }

        /// <summary>
        /// マシンIDを取得
        /// </summary>
        private string GetMachineId()
        {
            try
            {
                return $"{Environment.MachineName}-{Environment.UserName}";
            }
            catch
            {
                return "Unknown";
            }
        }

        /// <summary>
        /// アドインバージョンを取得
        /// </summary>
        private string GetAddinVersion()
        {
            try
            {
                var assembly = System.Reflection.Assembly.GetExecutingAssembly();
                var version = assembly.GetName().Version;
                return version.ToString();
            }
            catch
            {
                return "1.0.0.0";
            }
        }

        #endregion

        #region IDisposable

        public void Dispose()
        {
            if (!disposed)
            {
                httpClient?.Dispose();
                disposed = true;
                logger.Debug("LicenseApiClient disposed");
            }
        }

        #endregion
    }
}