using System;
using System.Security.Cryptography;
using System.Text;
using Microsoft.Win32;
using NLog;

namespace PowerPointEfficiencyAddin.Services.Security
{
    /// <summary>
    /// Windows レジストリへの安全なアクセスを提供するマネージャー
    /// ライセンス情報の保存・読み込みを暗号化して管理
    /// </summary>
    public sealed class RegistryManager : IDisposable
    {
        private static readonly Logger logger = LogManager.GetCurrentClassLogger();

        // レジストリキーのパス
        private const string REGISTRY_KEY_PATH = @"SOFTWARE\PowerPointEfficiencyAddin";
        private const string LICENSE_SUBKEY = "License";

        // 暗号化用の固定キー（本番環境では環境変数や別の安全な場所から取得すべき）
        private readonly byte[] encryptionKey;
        private readonly byte[] encryptionIV;

        private bool disposed = false;

        /// <summary>
        /// コンストラクタ
        /// </summary>
        public RegistryManager()
        {
            try
            {
                // マシン固有の情報を使用して暗号化キーを生成
                string machineIdentifier = GetMachineIdentifier();
                using (var sha256 = SHA256.Create())
                {
                    var hash = sha256.ComputeHash(Encoding.UTF8.GetBytes(machineIdentifier + "PPTAddin2025"));
                    encryptionKey = new byte[32];
                    Array.Copy(hash, encryptionKey, 32);

                    // IV用に別のハッシュを生成
                    var ivHash = sha256.ComputeHash(Encoding.UTF8.GetBytes(machineIdentifier + "IV"));
                    encryptionIV = new byte[16];
                    Array.Copy(ivHash, encryptionIV, 16);
                }

                logger.Debug("RegistryManager initialized with machine-specific encryption");
            }
            catch (Exception ex)
            {
                logger.Error(ex, "Failed to initialize RegistryManager");
                throw;
            }
        }

        /// <summary>
        /// 文字列値を暗号化してレジストリに保存
        /// </summary>
        public bool SaveSecureString(string valueName, string value)
        {
            if (string.IsNullOrEmpty(valueName))
            {
                logger.Warn("SaveSecureString called with empty valueName");
                return false;
            }

            RegistryKey key = null;
            try
            {
                // レジストリキーを開く（なければ作成）
                key = Registry.CurrentUser.CreateSubKey($@"{REGISTRY_KEY_PATH}\{LICENSE_SUBKEY}");

                if (key == null)
                {
                    logger.Error("Failed to create or open registry key");
                    return false;
                }

                // 値を暗号化
                string encryptedValue = string.IsNullOrEmpty(value)
                    ? string.Empty
                    : EncryptString(value);

                // レジストリに保存
                key.SetValue(valueName, encryptedValue, RegistryValueKind.String);

                logger.Debug($"Saved secure value '{valueName}' to registry");
                return true;
            }
            catch (UnauthorizedAccessException ex)
            {
                logger.Error(ex, "Unauthorized access to registry");
                return false;
            }
            catch (Exception ex)
            {
                logger.Error(ex, $"Failed to save secure string '{valueName}'");
                return false;
            }
            finally
            {
                key?.Close();
                key?.Dispose();
            }
        }

        /// <summary>
        /// サブキー配下の文字列値を保存（階層対応）
        /// </summary>
        public bool SaveSecureString(string subKeyPath, string valueName, string value)
        {
            if (string.IsNullOrEmpty(subKeyPath) || string.IsNullOrEmpty(valueName))
            {
                logger.Warn("SaveSecureString called with empty parameters");
                return false;
            }

            // スラッシュで分割されている場合の処理
            if (subKeyPath.Contains("\\"))
            {
                var lastSlashIndex = subKeyPath.LastIndexOf('\\');
                var keyPath = subKeyPath.Substring(0, lastSlashIndex);
                var actualValueName = subKeyPath.Substring(lastSlashIndex + 1);

                RegistryKey key = null;
                try
                {
                    key = Registry.CurrentUser.CreateSubKey($@"{REGISTRY_KEY_PATH}\{keyPath}");

                    if (key == null)
                    {
                        logger.Error("Failed to create or open registry subkey");
                        return false;
                    }

                    // 値を暗号化
                    string encryptedValue = string.IsNullOrEmpty(value)
                        ? string.Empty
                        : EncryptString(value);

                    key.SetValue(actualValueName, encryptedValue, RegistryValueKind.String);

                    logger.Debug($"Secure string saved to registry subkey: {keyPath}\\{actualValueName}");
                    return true;
                }
                catch (Exception ex)
                {
                    logger.Error(ex, $"Failed to save secure string to subkey: {subKeyPath}");
                    return false;
                }
                finally
                {
                    key?.Close();
                    key?.Dispose();
                }
            }
            else
            {
                // 既存のメソッドにフォールバック
                return SaveSecureString(valueName, value);
            }
        }

        /// <summary>
        /// レジストリから暗号化された文字列を読み込んで復号化
        /// </summary>
        public string LoadSecureString(string valueName)
        {
            if (string.IsNullOrEmpty(valueName))
            {
                logger.Warn("LoadSecureString called with empty valueName");
                return null;
            }

            RegistryKey key = null;
            try
            {
                // レジストリキーを開く（読み取り専用）
                key = Registry.CurrentUser.OpenSubKey($@"{REGISTRY_KEY_PATH}\{LICENSE_SUBKEY}", false);

                if (key == null)
                {
                    logger.Debug("Registry key does not exist");
                    return null;
                }

                // 値を読み込み
                var encryptedValue = key.GetValue(valueName) as string;

                if (string.IsNullOrEmpty(encryptedValue))
                {
                    logger.Debug($"Value '{valueName}' not found or empty");
                    return null;
                }

                // 復号化して返す
                return DecryptString(encryptedValue);
            }
            catch (Exception ex)
            {
                logger.Error(ex, $"Failed to load secure string '{valueName}'");
                return null;
            }
            finally
            {
                key?.Close();
                key?.Dispose();
            }
        }

        /// <summary>
        /// サブキー配下の文字列値を読み込み（階層対応）
        /// </summary>
        public string LoadSecureStringFromPath(string subKeyPath)
        {
            if (string.IsNullOrEmpty(subKeyPath))
            {
                logger.Warn("LoadSecureString called with empty subKeyPath");
                return null;
            }

            // スラッシュで分割されている場合の処理
            if (subKeyPath.Contains("\\"))
            {
                var lastSlashIndex = subKeyPath.LastIndexOf('\\');
                var keyPath = subKeyPath.Substring(0, lastSlashIndex);
                var actualValueName = subKeyPath.Substring(lastSlashIndex + 1);

                RegistryKey key = null;
                try
                {
                    key = Registry.CurrentUser.OpenSubKey($@"{REGISTRY_KEY_PATH}\{keyPath}", false);

                    if (key == null)
                    {
                        return null;
                    }

                    var encryptedValue = key.GetValue(actualValueName) as string;
                    if (string.IsNullOrEmpty(encryptedValue))
                    {
                        return null;
                    }

                    return DecryptString(encryptedValue);
                }
                catch (Exception ex)
                {
                    logger.Error(ex, $"Failed to load secure string from subkey: {subKeyPath}");
                    return null;
                }
                finally
                {
                    key?.Close();
                    key?.Dispose();
                }
            }
            else
            {
                // 既存のメソッドにフォールバック
                return LoadSecureString(subKeyPath);
            }
        }

        /// <summary>
        /// DateTime値を保存
        /// </summary>
        public bool SaveDateTime(string valueName, DateTime value)
        {
            try
            {
                // ISO 8601形式で保存
                string dateString = value.ToString("O");
                return SaveSecureString(valueName, dateString);
            }
            catch (Exception ex)
            {
                logger.Error(ex, $"Failed to save DateTime '{valueName}'");
                return false;
            }
        }

        /// <summary>
        /// DateTime値を読み込み
        /// </summary>
        public DateTime? LoadDateTime(string valueName)
        {
            try
            {
                string dateString = LoadSecureString(valueName);
                if (string.IsNullOrEmpty(dateString))
                {
                    return null;
                }

                if (DateTime.TryParse(dateString, out DateTime result))
                {
                    return result;
                }

                logger.Warn($"Failed to parse DateTime from '{valueName}'");
                return null;
            }
            catch (Exception ex)
            {
                logger.Error(ex, $"Failed to load DateTime '{valueName}'");
                return null;
            }
        }

        /// <summary>
        /// 特定の値を削除
        /// </summary>
        public bool DeleteValue(string valueName)
        {
            if (string.IsNullOrEmpty(valueName))
            {
                return false;
            }

            RegistryKey key = null;
            try
            {
                key = Registry.CurrentUser.OpenSubKey($@"{REGISTRY_KEY_PATH}\{LICENSE_SUBKEY}", true);

                if (key == null)
                {
                    return true; // キーが存在しない場合は成功とみなす
                }

                key.DeleteValue(valueName, false);
                logger.Debug($"Deleted value '{valueName}' from registry");
                return true;
            }
            catch (Exception ex)
            {
                logger.Error(ex, $"Failed to delete value '{valueName}'");
                return false;
            }
            finally
            {
                key?.Close();
                key?.Dispose();
            }
        }

        /// <summary>
        /// すべてのライセンス関連データをクリア
        /// </summary>
        public bool ClearAllLicenseData()
        {
            try
            {
                Registry.CurrentUser.DeleteSubKeyTree($@"{REGISTRY_KEY_PATH}\{LICENSE_SUBKEY}", false);
                logger.Info("Cleared all license data from registry");
                return true;
            }
            catch (Exception ex)
            {
                logger.Error(ex, "Failed to clear license data");
                return false;
            }
        }

        /// <summary>
        /// レジストリキーの存在確認
        /// </summary>
        public bool HasLicenseData()
        {
            RegistryKey key = null;
            try
            {
                key = Registry.CurrentUser.OpenSubKey($@"{REGISTRY_KEY_PATH}\{LICENSE_SUBKEY}", false);
                if (key == null)
                {
                    return false;
                }

                // ライセンスキーが存在するかチェック
                var licenseKey = key.GetValue("LicenseKey");
                return licenseKey != null;
            }
            catch (Exception ex)
            {
                logger.Error(ex, "Failed to check license data existence");
                return false;
            }
            finally
            {
                key?.Close();
                key?.Dispose();
            }
        }

        /// <summary>
        /// bool値を保存
        /// </summary>
        public bool SaveBool(string valueName, bool value)
        {
            if (string.IsNullOrEmpty(valueName))
            {
                logger.Warn("SaveBool called with empty valueName");
                return false;
            }

            RegistryKey key = null;
            try
            {
                // レジストリキーを開く（なければ作成）
                key = Registry.CurrentUser.CreateSubKey($@"{REGISTRY_KEY_PATH}");

                if (key == null)
                {
                    logger.Error("Failed to create or open registry key");
                    return false;
                }

                // DWORD値として保存
                key.SetValue(valueName, value ? 1 : 0, RegistryValueKind.DWord);

                logger.Debug($"Bool value saved to registry: {valueName} = {value}");
                return true;
            }
            catch (Exception ex)
            {
                logger.Error(ex, $"Failed to save bool value: {valueName}");
                return false;
            }
            finally
            {
                key?.Close();
                key?.Dispose();
            }
        }

        /// <summary>
        /// bool値を読み込み
        /// </summary>
        public bool? LoadBool(string valueName)
        {
            if (string.IsNullOrEmpty(valueName))
            {
                logger.Warn("LoadBool called with empty valueName");
                return null;
            }

            RegistryKey key = null;
            try
            {
                // まずポリシーキーをチェック（企業設定）
                key = Registry.LocalMachine.OpenSubKey($@"SOFTWARE\Policies\PowerPointEfficiencyAddin", false);
                if (key != null)
                {
                    var policyValue = key.GetValue(valueName);
                    if (policyValue != null)
                    {
                        key.Close();
                        return Convert.ToInt32(policyValue) != 0;
                    }
                    key.Close();
                }

                // ユーザー設定をチェック
                key = Registry.CurrentUser.OpenSubKey(REGISTRY_KEY_PATH, false);

                if (key == null)
                {
                    return null;
                }

                var value = key.GetValue(valueName);
                if (value == null)
                {
                    return null;
                }

                return Convert.ToInt32(value) != 0;
            }
            catch (Exception ex)
            {
                logger.Error(ex, $"Failed to load bool value: {valueName}");
                return null;
            }
            finally
            {
                key?.Close();
                key?.Dispose();
            }
        }



        #region Private Methods

        /// <summary>
        /// 文字列を暗号化
        /// </summary>
        private string EncryptString(string plainText)
        {
            if (string.IsNullOrEmpty(plainText))
            {
                return string.Empty;
            }

            try
            {
                using (var aes = Aes.Create())
                {
                    aes.Key = encryptionKey;
                    aes.IV = encryptionIV;
                    aes.Mode = CipherMode.CBC;
                    aes.Padding = PaddingMode.PKCS7;

                    using (var encryptor = aes.CreateEncryptor())
                    {
                        byte[] plainBytes = Encoding.UTF8.GetBytes(plainText);
                        byte[] encryptedBytes = encryptor.TransformFinalBlock(plainBytes, 0, plainBytes.Length);

                        // Base64エンコードして返す
                        return Convert.ToBase64String(encryptedBytes);
                    }
                }
            }
            catch (Exception ex)
            {
                logger.Error(ex, "Encryption failed");
                throw new CryptographicException("暗号化に失敗しました", ex);
            }
        }

        /// <summary>
        /// 文字列を復号化
        /// </summary>
        private string DecryptString(string encryptedText)
        {
            if (string.IsNullOrEmpty(encryptedText))
            {
                return string.Empty;
            }

            try
            {
                using (var aes = Aes.Create())
                {
                    aes.Key = encryptionKey;
                    aes.IV = encryptionIV;
                    aes.Mode = CipherMode.CBC;
                    aes.Padding = PaddingMode.PKCS7;

                    using (var decryptor = aes.CreateDecryptor())
                    {
                        byte[] encryptedBytes = Convert.FromBase64String(encryptedText);
                        byte[] decryptedBytes = decryptor.TransformFinalBlock(encryptedBytes, 0, encryptedBytes.Length);

                        return Encoding.UTF8.GetString(decryptedBytes);
                    }
                }
            }
            catch (Exception ex)
            {
                logger.Error(ex, "Decryption failed");
                throw new CryptographicException("復号化に失敗しました", ex);
            }
        }

        /// <summary>
        /// マシン固有の識別子を取得
        /// </summary>
        private string GetMachineIdentifier()
        {
            try
            {
                // マシン名とユーザー名を組み合わせて一意性を確保
                string machineName = Environment.MachineName;
                string userName = Environment.UserName;
                string processorId = GetProcessorId();

                return $"{machineName}-{userName}-{processorId}";
            }
            catch (Exception ex)
            {
                logger.Warn(ex, "Failed to get machine identifier, using fallback");
                // フォールバック値
                return "PPT-Addin-Default-2025";
            }
        }

        /// <summary>
        /// プロセッサIDを取得（可能な場合）
        /// </summary>
        private string GetProcessorId()
        {
            try
            {
                // WMIを使用せずに環境変数から取得
                string processorArchitecture = Environment.GetEnvironmentVariable("PROCESSOR_ARCHITECTURE");
                string processorLevel = Environment.GetEnvironmentVariable("PROCESSOR_LEVEL");
                string processorRevision = Environment.GetEnvironmentVariable("PROCESSOR_REVISION");

                return $"{processorArchitecture}-{processorLevel}-{processorRevision}";
            }
            catch
            {
                return "Unknown";
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
                    // 暗号化キーをクリア
                    if (encryptionKey != null)
                    {
                        Array.Clear(encryptionKey, 0, encryptionKey.Length);
                    }

                    if (encryptionIV != null)
                    {
                        Array.Clear(encryptionIV, 0, encryptionIV.Length);
                    }

                    disposed = true;
                    logger.Debug("RegistryManager disposed");
                }
                catch (Exception ex)
                {
                    logger.Error(ex, "Error disposing RegistryManager");
                }
            }
        }

        #endregion
    }
}