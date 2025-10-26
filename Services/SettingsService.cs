using System;
using System.Drawing;
using Microsoft.Win32;
using PowerPointEfficiencyAddin.Models;
using PowerPointEfficiencyAddin.Utils;
using NLog;

namespace PowerPointEfficiencyAddin.Services
{
    /// <summary>
    /// 設定の永続化サービス
    /// レジストリを使用してアプリケーション単位で設定を保存・読み込み
    /// </summary>
    public class SettingsService : IDisposable
    {
        #region フィールド・定数

        private static readonly Logger logger = LogManager.GetCurrentClassLogger();

        // レジストリキーパス（HKEY_CURRENT_USER配下）
        private const string REGISTRY_BASE_PATH = @"Software\PowerPointEfficiencyAddin";
        private const string SHAPE_STYLE_SUBKEY = "ShapeStyle";

        // 設定値キー名（色設定のみ）
        private const string KEY_ENABLE_STYLING = "EnableStyling";
        private const string KEY_FILL_COLOR = "FillColor";
        private const string KEY_LINE_COLOR = "LineColor";
        private const string KEY_FONT_COLOR = "FontColor";
        // 【削除】フォント名、サイズ、太字、斜体のキーは不要

        #endregion

        #region シングルトンパターン

        private static readonly Lazy<SettingsService> _instance = new Lazy<SettingsService>(() => new SettingsService());

        /// <summary>
        /// シングルトンインスタンス
        /// </summary>
        public static SettingsService Instance => _instance.Value;

        /// <summary>
        /// プライベートコンストラクタ
        /// </summary>
        private SettingsService()
        {
            logger.Debug("SettingsService instance created");
        }

        #endregion

        #region パブリックメソッド

        /// <summary>
        /// 図形スタイル設定を保存（色設定のみ）
        /// </summary>
        /// <param name="settings">保存する設定</param>
        /// <returns>保存成功時true</returns>
        public bool SaveShapeStyleSettings(ShapeStyleSettings settings)
        {
            if (settings == null)
            {
                logger.Warn("SaveShapeStyleSettings called with null settings");
                return false;
            }

            try
            {
                logger.Info("Saving shape style settings to registry");

                using (var key = CreateOrOpenRegistryKey())
                {
                    if (key == null)
                    {
                        logger.Error("Failed to create/open registry key for saving");
                        return false;
                    }

                    // 色設定のみを保存
                    key.SetValue(KEY_ENABLE_STYLING, settings.EnableStyling, RegistryValueKind.DWord);
                    key.SetValue(KEY_FILL_COLOR, ColorToArgb(settings.FillColor), RegistryValueKind.DWord);
                    key.SetValue(KEY_LINE_COLOR, ColorToArgb(settings.LineColor), RegistryValueKind.DWord);
                    key.SetValue(KEY_FONT_COLOR, ColorToArgb(settings.FontColor), RegistryValueKind.DWord);

                    logger.Info($"Shape style settings saved successfully: {settings}");
                    return true;
                }
            }
            catch (Exception ex)
            {
                logger.Error(ex, "Failed to save shape style settings");
                return false;
            }
        }

        /// <summary>
        /// 図形スタイル設定を読み込み（色設定のみ）
        /// </summary>
        /// <returns>読み込まれた設定（失敗時はデフォルト設定）</returns>
        public ShapeStyleSettings LoadShapeStyleSettings()
        {
            try
            {
                logger.Debug("Loading shape style settings from registry");

                using (var key = OpenRegistryKeyReadOnly())
                {
                    if (key == null)
                    {
                        logger.Info("Registry key not found, returning default settings");
                        return new ShapeStyleSettings();
                    }

                    var settings = new ShapeStyleSettings();

                    // 色設定のみを読み込み（存在しない場合はデフォルト値を使用）
                    settings.EnableStyling = GetRegistryBool(key, KEY_ENABLE_STYLING, false);
                    settings.FillColor = ArgbToColor(GetRegistryInt(key, KEY_FILL_COLOR, ColorToArgb(settings.FillColor)));
                    settings.LineColor = ArgbToColor(GetRegistryInt(key, KEY_LINE_COLOR, ColorToArgb(settings.LineColor)));
                    settings.FontColor = ArgbToColor(GetRegistryInt(key, KEY_FONT_COLOR, ColorToArgb(settings.FontColor)));

                    logger.Info($"Shape style settings loaded successfully: {settings}");
                    return settings;
                }
            }
            catch (Exception ex)
            {
                logger.Error(ex, "Failed to load shape style settings, returning defaults");
                return new ShapeStyleSettings();
            }
        }

        /// <summary>
        /// 図形スタイル設定をリセット（削除）
        /// </summary>
        /// <returns>リセット成功時true</returns>
        public bool ResetShapeStyleSettings()
        {
            try
            {
                logger.Info("Resetting shape style settings");

                using (var baseKey = Registry.CurrentUser.OpenSubKey(REGISTRY_BASE_PATH, true))
                {
                    if (baseKey != null)
                    {
                        baseKey.DeleteSubKeyTree(SHAPE_STYLE_SUBKEY, false);
                        logger.Info("Shape style settings reset successfully");
                        return true;
                    }
                }

                logger.Info("No settings to reset (registry key not found)");
                return true;
            }
            catch (Exception ex)
            {
                logger.Error(ex, "Failed to reset shape style settings");
                return false;
            }
        }

        #endregion

        #region プライベートヘルパーメソッド

        /// <summary>
        /// レジストリキーを作成または開く（書き込み用）
        /// </summary>
        /// <returns>レジストリキー（失敗時null）</returns>
        private RegistryKey CreateOrOpenRegistryKey()
        {
            try
            {
                var baseKey = Registry.CurrentUser.CreateSubKey(REGISTRY_BASE_PATH);
                if (baseKey == null) return null;

                var subKey = baseKey.CreateSubKey(SHAPE_STYLE_SUBKEY);
                baseKey.Dispose();
                return subKey;
            }
            catch (Exception ex)
            {
                logger.Error(ex, "Failed to create/open registry key for writing");
                return null;
            }
        }

        /// <summary>
        /// レジストリキーを開く（読み込み専用）
        /// </summary>
        /// <returns>レジストリキー（失敗時null）</returns>
        private RegistryKey OpenRegistryKeyReadOnly()
        {
            try
            {
                return Registry.CurrentUser.OpenSubKey($@"{REGISTRY_BASE_PATH}\{SHAPE_STYLE_SUBKEY}", false);
            }
            catch (Exception ex)
            {
                logger.Error(ex, "Failed to open registry key for reading");
                return null;
            }
        }

        /// <summary>
        /// レジストリからbool値を取得
        /// </summary>
        /// <param name="key">レジストリキー</param>
        /// <param name="valueName">値名</param>
        /// <param name="defaultValue">デフォルト値</param>
        /// <returns>取得された値</returns>
        private bool GetRegistryBool(RegistryKey key, string valueName, bool defaultValue)
        {
            try
            {
                var value = key.GetValue(valueName);
                if (value is int intValue)
                {
                    return intValue != 0;
                }
            }
            catch (Exception ex)
            {
                logger.Warn(ex, $"Failed to get registry bool value: {valueName}");
            }
            return defaultValue;
        }

        /// <summary>
        /// レジストリからint値を取得
        /// </summary>
        /// <param name="key">レジストリキー</param>
        /// <param name="valueName">値名</param>
        /// <param name="defaultValue">デフォルト値</param>
        /// <returns>取得された値</returns>
        private int GetRegistryInt(RegistryKey key, string valueName, int defaultValue)
        {
            try
            {
                var value = key.GetValue(valueName);
                if (value is int intValue)
                {
                    return intValue;
                }
            }
            catch (Exception ex)
            {
                logger.Warn(ex, $"Failed to get registry int value: {valueName}");
            }
            return defaultValue;
        }

        /// <summary>
        /// ColorをARGB整数値に変換
        /// </summary>
        /// <param name="color">色</param>
        /// <returns>ARGB値</returns>
        private int ColorToArgb(Color color)
        {
            return color.ToArgb();
        }

        /// <summary>
        /// ARGB整数値をColorに変換
        /// </summary>
        /// <param name="argb">ARGB値</param>
        /// <returns>色</returns>
        private Color ArgbToColor(int argb)
        {
            try
            {
                return Color.FromArgb(argb);
            }
            catch
            {
                // 無効なARGB値の場合はデフォルト色を返す
                return Color.Black;
            }
        }

        #endregion

        #region IDisposable実装（将来の拡張用）

        /// <summary>
        /// リソース解放フラグ
        /// </summary>
        private bool _disposed = false;

        /// <summary>
        /// リソースを解放
        /// </summary>
        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        /// <summary>
        /// リソース解放の実装
        /// </summary>
        /// <param name="disposing">マネージドリソースも解放するかどうか</param>
        protected virtual void Dispose(bool disposing)
        {
            if (!_disposed)
            {
                if (disposing)
                {
                    // マネージドリソースの解放（現在は特になし）
                }

                _disposed = true;
                logger.Debug("SettingsService disposed");
            }
        }

        #endregion
    }
}