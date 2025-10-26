namespace PowerPointEfficiencyAddin.Models
{
    /// <summary>
    /// 間隔調整方法
    /// </summary>
    public enum SpacingAdjustmentMethod
    {
        /// <summary>オブジェクトのサイズを変更</summary>
        ResizeObjects,
        /// <summary>オブジェクトを移動</summary>
        MoveObjects
    }

    /// <summary>
    /// 間隔調整設定
    /// </summary>
    public class SpacingSettings
    {
        public float Spacing { get; set; }
        public SpacingAdjustmentMethod AdjustmentMethod { get; set; }
    }
}