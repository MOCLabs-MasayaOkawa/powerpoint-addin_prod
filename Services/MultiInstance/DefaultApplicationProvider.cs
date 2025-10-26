using NLog;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PowerPointEfficiencyAddin.Services
{
    /// <summary>
    /// デフォルトアプリケーション提供者（従来動作）
    /// </summary>
    public class DefaultApplicationProvider : IApplicationProvider
    {
        private static readonly Logger logger = LogManager.GetCurrentClassLogger();

        public PowerPoint.Application GetCurrentApplication()
        {
            return Globals.ThisAddIn.Application;
        }

        public bool IsValid()
        {
            try
            {
                var app = GetCurrentApplication();
                if (app != null)
                {
                    var _ = app.Version; // 有効性チェック
                    return true;
                }
            }
            catch (System.Exception ex)
            {
                logger.Debug(ex, "Default application provider validation failed");
            }

            return false;
        }
    }
}