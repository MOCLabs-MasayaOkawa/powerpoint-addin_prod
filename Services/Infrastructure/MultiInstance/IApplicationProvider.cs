using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PowerPointEfficiencyAddin.Services.Infrastructure.MultiInstance
{
    /// <summary>
    /// PowerPointアプリケーション提供インターフェース
    /// 複数インスタンス対応のDI基盤
    /// </summary>
    public interface IApplicationProvider
    {
        /// <summary>
        /// 現在アクティブなPowerPointアプリケーションを取得
        /// </summary>
        PowerPoint.Application GetCurrentApplication();

        /// <summary>
        /// アプリケーション提供者の有効性チェック
        /// </summary>
        bool IsValid();
    }
}