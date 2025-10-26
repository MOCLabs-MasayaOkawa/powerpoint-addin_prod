using NLog;
//using PowerPointEfficiencyAddin.Services.MultiInstance.MultiInstance;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PowerPointEfficiencyAddin.Services
{
    /// <summary>
    /// 複数インスタンス対応アプリケーション提供者
    /// </summary>
    public class MultiInstanceApplicationProvider : IApplicationProvider
    {
        private static readonly Logger logger = LogManager.GetCurrentClassLogger();
        private readonly ApplicationContextManager contextManager;

        public MultiInstanceApplicationProvider(ApplicationContextManager contextManager)
        {
            this.contextManager = contextManager ?? throw new System.ArgumentNullException(nameof(contextManager));
        }

        public PowerPoint.Application GetCurrentApplication()
        {
            return contextManager.CurrentApplication;
        }

        public bool IsValid()
        {
            try
            {
                return contextManager != null && GetCurrentApplication() != null;
            }
            catch (System.Exception ex)
            {
                logger.Debug(ex, "Multi-instance application provider validation failed");
                return false;
            }
        }
    }
}