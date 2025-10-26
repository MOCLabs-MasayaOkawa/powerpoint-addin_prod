namespace PowerPointEfficiencyAddin.Models.Licensing
{
    /// <summary>
    /// 機能定義
    /// </summary>
    public class FeatureDefinition
    {
        /// <summary>機能ID（一意）</summary>
        public string FeatureId { get; set; }

        /// <summary>機能名</summary>
        public string DisplayName { get; set; }

        /// <summary>カテゴリ</summary>
        public FunctionCategory Category { get; set; }

        /// <summary>必要な最小アクセスレベル</summary>
        public FeatureAccessLevel RequiredLevel { get; set; }

        /// <summary>説明</summary>
        public string Description { get; set; }

        /// <summary>ソート順</summary>
        public int Order { get; set; }

        /// <summary>有効/無効</summary>
        public bool IsEnabled { get; set; } = true;
    }
}