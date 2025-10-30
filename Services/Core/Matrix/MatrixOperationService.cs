using Microsoft.Office.Core;
using NLog;
using PowerPointEfficiencyAddin.Models;
using PowerPointEfficiencyAddin.Services.Core.PowerTool;
using PowerPointEfficiencyAddin.Services.Infrastructure.MultiInstance;
using PowerPointEfficiencyAddin.Utils;
using System;
using System.Collections.Generic;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PowerPointEfficiencyAddin.Services.Core.Matrix
{
    /// <summary>
    /// マトリクス操作のファサードクラス（各専門サービスに処理を委譲）
    /// </summary>
    public class MatrixOperationService
    {
        private static readonly Logger logger = LogManager.GetCurrentClassLogger();
        private readonly IApplicationProvider applicationProvider;
        
        // 各専門サービス
        private readonly MatrixExcelService excelService;
        private readonly MatrixOptimizationService optimizationService;
        private readonly MatrixStructureService structureService;
        private readonly MatrixAlignmentService alignmentService;

        // DI対応コンストラクタ
        public MatrixOperationService(IApplicationProvider applicationProvider)
        {
            this.applicationProvider = applicationProvider ?? throw new ArgumentNullException(nameof(applicationProvider));
            logger.Debug("MatrixOperationService (Facade) initialized");
            
            // 各専門サービスを初期化
            excelService = new MatrixExcelService(applicationProvider);
            optimizationService = new MatrixOptimizationService(applicationProvider);
            structureService = new MatrixStructureService(applicationProvider);
            alignmentService = new MatrixAlignmentService(applicationProvider);
            
            logger.Debug("All matrix sub-services initialized");
        }

        #region Excel連携機能（MatrixExcelServiceへ委譲）

        /// <summary>
        /// ExcelデータをPowerPointに貼り付け
        /// </summary>
        public void ExcelToPptx()
        {
            excelService.ExcelToPptx();
        }

        #endregion

        #region 最適化機能（MatrixOptimizationServiceへ委譲）

        /// <summary>
        /// マトリクス行高さ最適化
        /// </summary>
        public void OptimizeMatrixRowHeights()
        {
            optimizationService.OptimizeMatrixRowHeights();
        }

        /// <summary>
        /// 表完全最適化
        /// </summary>
        public void OptimizeTableComplete()
        {
            optimizationService.OptimizeTableComplete();
        }

        /// <summary>
        /// 列幅統一
        /// </summary>
        public void EqualizeColumnWidths()
        {
            optimizationService.EqualizeColumnWidths();
        }

        /// <summary>
        /// 行高統一
        /// </summary>
        public void EqualizeRowHeights()
        {
            optimizationService.EqualizeRowHeights();
        }

        #endregion

        #region 構造変更機能（MatrixStructureServiceへ委譲）

        /// <summary>
        /// マトリクス行間区切り線追加
        /// </summary>
        public void AddMatrixRowSeparators()
        {
            structureService.AddMatrixRowSeparators();
        }

        /// <summary>
        /// 見出し行付与
        /// </summary>
        public void AddHeaderRowToMatrix()
        {
            structureService.AddHeaderRowToMatrix();
        }

        /// <summary>
        /// 行追加
        /// </summary>
        public void AddMatrixRow()
        {
            structureService.AddMatrixRow();
        }

        /// <summary>
        /// 列追加
        /// </summary>
        public void AddMatrixColumn()
        {
            structureService.AddMatrixColumn();
        }

        #endregion

        #region 配置・整列機能（MatrixAlignmentServiceへ委譲）

        /// <summary>
        /// 図形をセル中央に整列
        /// </summary>
        public void AlignShapesToCells()
        {
            alignmentService.AlignShapesToCells();
        }

        /// <summary>
        /// セルマージン設定
        /// </summary>
        public void SetCellMargins()
        {
            alignmentService.SetCellMargins();
        }

        /// <summary>
        /// Matrix Tuner（マトリクス調整ダイアログ）
        /// </summary>
        public void MatrixTuner()
        {
            alignmentService.MatrixTuner();
        }

        #endregion
    }
}
