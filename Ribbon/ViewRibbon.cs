using System;
using System.IO;
using System.Reflection;
using Microsoft.Office.Core;
using PowerPointEfficiencyAddin.Utils;
using NLog;

namespace PowerPointEfficiencyAddin.Ribbon
{
    /// <summary>
    /// PowerPoint効率化アドインの最小リボンUI（View切り替えのみ）
    /// </summary>
    [System.Runtime.InteropServices.ComVisible(true)]
    public class ViewRibbon : IRibbonExtensibility
    {
        private IRibbonUI ribbon;
        private static readonly Logger logger = LogManager.GetCurrentClassLogger();

        public ViewRibbon()
        {
            logger.Info("ViewRibbon constructor called");
            logger.Info($"ViewRibbon instance created at {DateTime.Now:HH:mm:ss.fff}");
        }

        #region IRibbonExtensibility Members

        public string GetCustomUI(string ribbonID)
        {
            try
            {
                logger.Info($"GetCustomUI called with ribbonID: {ribbonID}");
                logger.Info($"Current thread: {System.Threading.Thread.CurrentThread.ManagedThreadId}");

                var customUI = GetResourceText("PowerPointEfficiencyAddin.Ribbon.ViewRibbon.xml");

                if (string.IsNullOrEmpty(customUI))
                {
                    logger.Error("ViewRibbon.xml content is null or empty");
                    return string.Empty;
                }

                logger.Info($"ViewRibbon.xml loaded successfully, length: {customUI.Length}");
                logger.Debug($"XML content preview: {customUI.Substring(0, Math.Min(200, customUI.Length))}...");

                return customUI;
            }
            catch (Exception ex)
            {
                logger.Error(ex, "Failed to load ViewRibbon.xml");
                return string.Empty;
            }
        }

        #endregion

        #region Ribbon Callbacks

        /// <summary>
        /// リボンロード時のコールバック
        /// </summary>
        /// <param name="ribbonUI">リボンUI</param>
        public void Ribbon_Load(IRibbonUI ribbonUI)
        {
            try
            {
                logger.Info("*** RIBBON_LOAD CALLED - THIS IS CRITICAL ***");
                logger.Info($"Ribbon_Load called at {DateTime.Now:HH:mm:ss.fff}");
                logger.Info($"RibbonUI parameter: {ribbonUI?.GetType().Name ?? "NULL"}");

                this.ribbon = ribbonUI;

                // リボンが正常に読み込まれたことを確認
                if (ribbonUI != null)
                {
                    logger.Info("*** RibbonUI object is valid, ViewRibbon is ACTIVE ***");
                    logger.Info("*** PowerPoint efficiency panel should now be visible in View tab ***");
                }
                else
                {
                    logger.Error("*** RibbonUI object is NULL - CRITICAL ERROR ***");
                }
            }
            catch (Exception ex)
            {
                logger.Error(ex, "*** CRITICAL ERROR during ViewRibbon load ***");
            }
        }

        /// <summary>
        /// 効率化ペイン表示ボタンクリック
        /// </summary>
        /// <param name="control">コントロール</param>
        public void ShowEfficiencyPane_Click(IRibbonControl control)
        {
            logger.Info("*** ShowEfficiencyPane_Click invoked ***");

            try
            {
                var addIn = Globals.ThisAddIn;
                if (addIn != null)
                {
                    addIn.ShowPanel();
                    logger.Info("ShowPanel called successfully");
                }
                else
                {
                    logger.Error("ThisAddIn is null");
                }
            }
            catch (Exception ex)
            {
                logger.Error(ex, "Error in ShowEfficiencyPane_Click");
                System.Windows.Forms.MessageBox.Show(
                    $"効率化ペインの表示に失敗しました: {ex.Message}",
                    "エラー",
                    System.Windows.Forms.MessageBoxButtons.OK,
                    System.Windows.Forms.MessageBoxIcon.Error
                );
            }
        }

        /// <summary>
        /// 効率化ペイン非表示ボタンクリック
        /// </summary>
        /// <param name="control">コントロール</param>
        public void HideEfficiencyPane_Click(IRibbonControl control)
        {
            logger.Info("*** HideEfficiencyPane_Click invoked ***");

            try
            {
                var addIn = Globals.ThisAddIn;
                if (addIn != null)
                {
                    addIn.HidePanel();
                    logger.Info("HidePanel called successfully");
                }
                else
                {
                    logger.Error("ThisAddIn is null");
                }
            }
            catch (Exception ex)
            {
                logger.Error(ex, "Error in HideEfficiencyPane_Click");
                System.Windows.Forms.MessageBox.Show(
                    $"効率化ペインの非表示に失敗しました: {ex.Message}",
                    "エラー",
                    System.Windows.Forms.MessageBoxButtons.OK,
                    System.Windows.Forms.MessageBoxIcon.Error
                );
            }
        }

        #endregion

        #region Helper Methods

        /// <summary>
        /// リソースからテキストを取得します
        /// </summary>
        /// <param name="resourceName">リソース名</param>
        /// <returns>リソーステキスト</returns>
        private static string GetResourceText(string resourceName)
        {
            try
            {
                Assembly asm = Assembly.GetExecutingAssembly();
                string[] resourceNames = asm.GetManifestResourceNames();

                logger.Debug($"Looking for resource: {resourceName}");
                logger.Debug($"Available resources: {string.Join(", ", resourceNames)}");

                for (int i = 0; i < resourceNames.Length; ++i)
                {
                    if (string.Compare(resourceName, resourceNames[i], StringComparison.OrdinalIgnoreCase) == 0)
                    {
                        logger.Debug($"Found matching resource: {resourceNames[i]}");

                        using (StreamReader resourceReader = new StreamReader(asm.GetManifestResourceStream(resourceNames[i])))
                        {
                            if (resourceReader != null)
                            {
                                var content = resourceReader.ReadToEnd();
                                logger.Info($"Resource content loaded successfully, length: {content.Length}");
                                return content;
                            }
                        }
                    }
                }

                logger.Error($"Resource not found: {resourceName}");
                return null;
            }
            catch (Exception ex)
            {
                logger.Error(ex, $"Failed to load resource: {resourceName}");
                return null;
            }
        }

        #endregion
    }
}