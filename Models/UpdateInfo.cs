using System;
using System.Collections.Generic;

namespace PowerPointEfficiencyAddin.Models.Licensing
{
    /// <summary>
    /// 更新情報モデル（MVP版）
    /// </summary>
    public class UpdateInfo
    {
        /// <summary>
        /// バージョン番号（例: "1.2.0"）
        /// </summary>
        public string Version { get; set; }

        /// <summary>
        /// リリース日時
        /// </summary>
        public DateTime ReleaseDate { get; set; }

        /// <summary>
        /// ダウンロードURL（署名付きURL想定）
        /// </summary>
        public string DownloadUrl { get; set; }

        /// <summary>
        /// ファイルのSHA256チェックサム
        /// </summary>
        public string Checksum { get; set; }

        /// <summary>
        /// ファイルサイズ（バイト）
        /// </summary>
        public long FileSize { get; set; }

        /// <summary>
        /// 重要更新かどうか（セキュリティ修正等）
        /// </summary>
        public bool IsCritical { get; set; }

        /// <summary>
        /// 更新内容の説明
        /// </summary>
        public string ReleaseNotes { get; set; }

        /// <summary>
        /// 最小必要バージョン（これ以下からは直接更新不可）
        /// </summary>
        public string MinimumVersion { get; set; }

        /// <summary>
        /// 現在のバージョンと比較
        /// </summary>
        public bool IsNewerThan(string currentVersion)
        {
            try
            {
                var current = new Version(currentVersion);
                var update = new Version(Version);
                return update > current;
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// 更新可能かチェック
        /// </summary>
        public bool CanUpdateFrom(string currentVersion)
        {
            if (!IsNewerThan(currentVersion))
                return false;

            if (string.IsNullOrEmpty(MinimumVersion))
                return true;

            try
            {
                var current = new Version(currentVersion);
                var minimum = new Version(MinimumVersion);
                return current >= minimum;
            }
            catch
            {
                return true; // エラー時は更新を許可
            }
        }
    }

    /// <summary>
    /// 更新チェック結果
    /// </summary>
    public class UpdateCheckResult
    {
        /// <summary>
        /// 更新が利用可能か
        /// </summary>
        public bool UpdateAvailable { get; set; }

        /// <summary>
        /// 更新情報
        /// </summary>
        public UpdateInfo UpdateInfo { get; set; }

        /// <summary>
        /// エラーメッセージ
        /// </summary>
        public string ErrorMessage { get; set; }

        /// <summary>
        /// チェック成功したか
        /// </summary>
        public bool Success => string.IsNullOrEmpty(ErrorMessage);
    }
}