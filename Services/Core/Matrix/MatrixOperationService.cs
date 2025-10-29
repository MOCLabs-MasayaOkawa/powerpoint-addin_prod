using ImageMagick;
using Microsoft.Office.Core;
using NLog;
using PowerPointEfficiencyAddin.Models;
using PowerPointEfficiencyAddin.Services.Infrastructure.MultiInstance;
using PowerPointEfficiencyAddin.Services.Infrastructure.Settings;
using PowerPointEfficiencyAddin.Services.UI.Dialogs;
using PowerPointEfficiencyAddin.UI;
using PowerPointEfficiencyAddin.Utils;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PowerPointEfficiencyAddin.Services.Core.Matrix
{
    /// <summary>
    /// マトリクス操作（Excel貼り付け、行列追加、最適化など）を提供するサービスクラス
    /// </summary>
    public class MatrixOperationService
    {
        private static readonly Logger logger = LogManager.GetCurrentClassLogger();
        private readonly IApplicationProvider applicationProvider;
        private readonly PowerToolServiceHelper helper;

        // DI対応コンストラクタ
        public MatrixOperationService(IApplicationProvider applicationProvider)
        {
            this.applicationProvider = applicationProvider ?? throw new ArgumentNullException(nameof(applicationProvider));
            logger.Debug("MatrixOperationService initialized with DI application provider");
            helper = new PowerToolServiceHelper(applicationProvider);
        }

