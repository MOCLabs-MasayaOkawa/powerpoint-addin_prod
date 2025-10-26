using Newtonsoft.Json;
using NLog;
using PowerPointEfficiencyAddin.Models.Licensing;
using PowerPointEfficiencyAddin.Services.Security;
using System;
using System.Diagnostics;
using System.IO;
using System.Net.Http;
using System.Reflection;
using System.Security.Cryptography;
using System.Threading.Tasks;

namespace PowerPointEfficiencyAddin.Services.Licensing
{
    /// <summary>
    /// 自動更新管理サービス（MVP版）
    /// </summary>
    public sealed class UpdateService : IDisposable
    {
        private static readonly Logger logger = LogManager.GetCurrentClassLogger();
        private static UpdateService instance;
        private static readonly object lockObject = new object();

        private readonly RegistryManager registryManager;
        private readonly HttpClient httpClient;
        private UpdateInfo pendingUpdate;
        private string downloadedFilePath;
        private bool disposed = false;

        // 定数
        private const string UPDATE_REGISTRY_KEY = "Updates";
        private const string PENDING_UPDATE_KEY = "PendingUpdate";
        private const string UPDATE_CHECK_KEY = "LastUpdateCheck";
        private const string AUTO_UPDATE_KEY = "AutoUpdateEnabled";
        private const string UPDATE_CHANNEL_KEY = "UpdateChannel";

        // 開発モードフラグ（LicenseManagerと共有）
        public bool DevelopmentMode => LicenseManager.DevelopmentMode;

        // 現在のバージョン
        public string CurrentVersion { get; }

        /// <summary>
        /// シングルトンインスタンス
        /// </summary>
        public static UpdateService Instance
        {
            get
            {
                if (instance == null)
                {
                    lock (lockObject)
                    {
                        if (instance == null)
                        {
                            instance = new UpdateService();
                        }
                    }
                }
                return instance;
            }
        }

        /// <summary>
        /// プライベートコンストラクタ
        /// </summary>
        private UpdateService()
        {
            try
            {
                logger.Info("Initializing UpdateService");

                registryManager = new RegistryManager();
                httpClient = CreateHttpClient();

                // 現在のバージョンを取得
                CurrentVersion = Assembly.GetExecutingAssembly()
                    .GetName().Version.ToString(3); // Major.Minor.Build

                // 更新用一時フォルダの作成
                var tempPath = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                    "PowerPointEfficiencyAddin", "Updates");

                if (!Directory.Exists(tempPath))
                {
                    Directory.CreateDirectory(tempPath);
                }

                logger.Info($"UpdateService initialized. Current version: {CurrentVersion}");
            }
            catch (Exception ex)
            {
                logger.Error(ex, "Failed to initialize UpdateService");
                throw;
            }
        }

        #region 公開メソッド

        /// <summary>
        /// ライセンス検証と同時に更新をチェック（LicenseManagerから呼び出し）
        /// </summary>
        public async Task<UpdateCheckResult> CheckForUpdateAsync(UpdateInfo updateInfo)
        {
            try
            {
                // 開発モードでは更新チェックをスキップ
                if (DevelopmentMode)
                {
                    logger.Debug("Development mode - skipping update check");
                    return new UpdateCheckResult
                    {
                        UpdateAvailable = false,
                        ErrorMessage = "Development mode"
                    };
                }

                // 自動更新が無効な場合
                if (!IsAutoUpdateEnabled())
                {
                    logger.Info("Auto-update is disabled by policy");
                    return new UpdateCheckResult { UpdateAvailable = false };
                }

                // バージョン比較
                if (updateInfo != null && updateInfo.IsNewerThan(CurrentVersion))
                {
                    if (updateInfo.CanUpdateFrom(CurrentVersion))
                    {
                        logger.Info($"Update available: {updateInfo.Version}");
                        pendingUpdate = updateInfo;

                        // レジストリに保存
                        SavePendingUpdate(updateInfo);

                        return new UpdateCheckResult
                        {
                            UpdateAvailable = true,
                            UpdateInfo = updateInfo
                        };
                    }
                    else
                    {
                        logger.Warn($"Cannot update from {CurrentVersion} to {updateInfo.Version} directly");
                        return new UpdateCheckResult
                        {
                            UpdateAvailable = false,
                            ErrorMessage = "直接更新できないバージョンです。手動でのインストールが必要です。"
                        };
                    }
                }

                return new UpdateCheckResult { UpdateAvailable = false };
            }
            catch (Exception ex)
            {
                logger.Error(ex, "Failed to check for updates");
                return new UpdateCheckResult
                {
                    UpdateAvailable = false,
                    ErrorMessage = ex.Message
                };
            }
        }

        /// <summary>
        /// バックグラウンドで更新をダウンロード
        /// </summary>
        public async Task<bool> DownloadUpdateAsync(UpdateInfo updateInfo)
        {
            if (updateInfo == null)
                return false;

            try
            {
                logger.Info($"Starting download of version {updateInfo.Version}");

                // ダウンロード先パス
                var tempPath = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                    "PowerPointEfficiencyAddin", "Updates",
                    $"update_{updateInfo.Version}.msi");

                // 既にダウンロード済みかチェック
                if (File.Exists(tempPath) && VerifyChecksum(tempPath, updateInfo.Checksum))
                {
                    logger.Info("Update already downloaded and verified");
                    downloadedFilePath = tempPath;
                    return true;
                }

                // ダウンロード実行
                using (var response = await httpClient.GetAsync(updateInfo.DownloadUrl))
                {
                    response.EnsureSuccessStatusCode();

                    using (var fileStream = File.Create(tempPath))
                    {
                        await response.Content.CopyToAsync(fileStream);
                    }
                }

                // チェックサム検証
                if (!VerifyChecksum(tempPath, updateInfo.Checksum))
                {
                    logger.Error("Downloaded file checksum verification failed");
                    File.Delete(tempPath);
                    return false;
                }

                logger.Info($"Update downloaded successfully to {tempPath}");
                downloadedFilePath = tempPath;

                // ダウンロード完了をレジストリに記録
                registryManager.SaveSecureString("DownloadedFile", tempPath);

                return true;
            }
            catch (Exception ex)
            {
                logger.Error(ex, "Failed to download update");
                return false;
            }
        }

        /// <summary>
        /// 更新を適用（PowerPoint終了時に呼び出し）
        /// </summary>
        public bool ApplyPendingUpdate()
        {
            try
            {
                if (string.IsNullOrEmpty(downloadedFilePath) || !File.Exists(downloadedFilePath))
                {
                    logger.Warn("No downloaded update file found");
                    return false;
                }

                logger.Info($"Applying update from {downloadedFilePath}");

                // MSIをサイレントモードで実行（再起動なし）
                var startInfo = new ProcessStartInfo
                {
                    FileName = "msiexec.exe",
                    Arguments = $"/i \"{downloadedFilePath}\" /quiet /norestart REBOOT=ReallySuppress",
                    UseShellExecute = false,
                    CreateNoWindow = true
                };

                var process = Process.Start(startInfo);

                // 非同期で実行（PowerPointの終了をブロックしない）
                Task.Run(() =>
                {
                    process.WaitForExit(60000); // 最大1分待機

                    if (process.ExitCode == 0)
                    {
                        logger.Info("Update installation completed successfully");

                        // クリーンアップ
                        ClearPendingUpdate();

                        try
                        {
                            File.Delete(downloadedFilePath);
                        }
                        catch { /* ファイル削除失敗は無視 */ }
                    }
                    else
                    {
                        logger.Error($"Update installation failed with exit code: {process.ExitCode}");
                    }
                });

                return true;
            }
            catch (Exception ex)
            {
                logger.Error(ex, "Failed to apply update");
                return false;
            }
        }

        /// <summary>
        /// 保留中の更新があるかチェック
        /// </summary>
        public bool HasPendingUpdate()
        {
            return pendingUpdate != null || LoadPendingUpdate() != null;
        }

        /// <summary>
        /// 保留中の更新情報を取得
        /// </summary>
        public UpdateInfo GetPendingUpdate()
        {
            return pendingUpdate ?? LoadPendingUpdate();
        }

        #endregion

        #region プライベートメソッド

        /// <summary>
        /// HttpClientの作成（プロキシ対応）
        /// </summary>
        private HttpClient CreateHttpClient()
        {
            var handler = new HttpClientHandler();

            // プロキシ設定をレジストリから読み込み
            var proxyUrl = registryManager.LoadSecureString("ProxyUrl");
            if (!string.IsNullOrEmpty(proxyUrl))
            {
                handler.Proxy = new System.Net.WebProxy(proxyUrl)
                {
                    UseDefaultCredentials = true // Windows認証
                };
            }

            return new HttpClient(handler)
            {
                Timeout = TimeSpan.FromMinutes(5) // ダウンロード用に長めに設定
            };
        }

        /// <summary>
        /// チェックサム検証
        /// </summary>
        private bool VerifyChecksum(string filePath, string expectedChecksum)
        {
            if (string.IsNullOrEmpty(expectedChecksum))
                return true; // チェックサムが提供されていない場合はスキップ

            try
            {
                using (var sha256 = SHA256.Create())
                using (var stream = File.OpenRead(filePath))
                {
                    var hash = sha256.ComputeHash(stream);
                    var hashString = BitConverter.ToString(hash).Replace("-", "").ToUpperInvariant();

                    var result = hashString == expectedChecksum.ToUpperInvariant();

                    if (!result)
                    {
                        logger.Error($"Checksum mismatch. Expected: {expectedChecksum}, Actual: {hashString}");
                    }

                    return result;
                }
            }
            catch (Exception ex)
            {
                logger.Error(ex, "Failed to verify checksum");
                return false;
            }
        }

        /// <summary>
        /// 自動更新が有効かチェック
        /// </summary>
        private bool IsAutoUpdateEnabled()
        {
            // グループポリシーチェック（企業環境）
            // LoadBoolは内部でポリシーキーを優先的にチェックしているため、1回の呼び出しで十分
            var enabled = registryManager.LoadBool("AutoUpdateEnabled");
            return enabled ?? true; // nullの場合はデフォルト有効
        }

        /// <summary>
        /// 保留中の更新情報を保存
        /// </summary>
        private void SavePendingUpdate(UpdateInfo updateInfo)
        {
            try
            {
                var json = JsonConvert.SerializeObject(updateInfo);
                registryManager.SaveSecureString("PendingUpdate", json);
            }
            catch (Exception ex)
            {
                logger.Error(ex, "Failed to save pending update");
            }
        }

        /// <summary>
        /// 保留中の更新情報を読み込み
        /// </summary>
        private UpdateInfo LoadPendingUpdate()
        {
            var json = registryManager.LoadSecureString("PendingUpdate");
            if (!string.IsNullOrEmpty(json))
                return JsonConvert.DeserializeObject<UpdateInfo>(json);
            return null;
        }

        /// <summary>
        /// 保留中の更新をクリア
        /// </summary>
        private void ClearPendingUpdate()
        {
            try
            {
                pendingUpdate = null;
                registryManager.DeleteValue("PendingUpdate");
                registryManager.DeleteValue("DownloadedFile");
            }
            catch (Exception ex)
            {
                logger.Error(ex, "Failed to clear pending update");
            }
        }

        #endregion

        #region IDisposable

        public void Dispose()
        {
            if (!disposed)
            {
                try
                {
                    httpClient?.Dispose();
                    instance = null;
                    disposed = true;
                    logger.Info("UpdateService disposed");
                }
                catch (Exception ex)
                {
                    logger.Error(ex, "Error disposing UpdateService");
                }
            }
        }

        #endregion
    }
}