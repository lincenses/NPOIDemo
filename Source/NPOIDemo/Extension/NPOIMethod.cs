using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace NPOIDemo.Extension
{
    public class NPOIMethod
    {
        #region 获取单个格的值
        /// <summary>
        /// 获取单元格的值。
        /// </summary>
        /// <param name="iCell">NPOI.SS.UserModel.ICell对象。</param>
        /// <returns>单元格的值。</returns>
        public static object GetCellValue(NPOI.SS.UserModel.ICell iCell)
        {
            if (iCell == null)
            { return null; }
            else
            {
                if (iCell.CellType == NPOI.SS.UserModel.CellType.Blank)
                { return ""; }
                else if (iCell.CellType == NPOI.SS.UserModel.CellType.Boolean)
                { return iCell.BooleanCellValue; }
                else if (iCell.CellType == NPOI.SS.UserModel.CellType.Error)
                { return iCell.ErrorCellValue.ToString(System.Globalization.CultureInfo.InvariantCulture); }
                else if (iCell.CellType == NPOI.SS.UserModel.CellType.Formula)
                {
                    if (iCell.CachedFormulaResultType == NPOI.SS.UserModel.CellType.Error)
                    { return ""; }
                    else
                    { return iCell.NumericCellValue; }
                }
                else if (iCell.CellType == NPOI.SS.UserModel.CellType.Numeric)
                {
                    if (NPOI.SS.UserModel.DateUtil.IsCellDateFormatted(iCell))
                    { return iCell.DateCellValue; }
                    else
                    { return iCell.NumericCellValue; }
                }
                else if (iCell.CellType == NPOI.SS.UserModel.CellType.String)
                { return iCell.StringCellValue; }
                else
                { return iCell.ToString(); }
            }
        }
        #endregion

        #region 获取单元格样式
        public static NPOI.SS.UserModel.ICellStyle GetCellStyle(NPOI.SS.UserModel.IWorkbook workbook, short backColorIndex, short fontColorIndex, System.Drawing.Font font, NPOI.SS.UserModel.HorizontalAlignment horizontalAlignment = NPOI.SS.UserModel.HorizontalAlignment.General, NPOI.SS.UserModel.VerticalAlignment verticalAlignment = NPOI.SS.UserModel.VerticalAlignment.None, NPOI.SS.UserModel.BorderStyle borderLeft = NPOI.SS.UserModel.BorderStyle.None, NPOI.SS.UserModel.BorderStyle borderTop = NPOI.SS.UserModel.BorderStyle.None, NPOI.SS.UserModel.BorderStyle borderRight = NPOI.SS.UserModel.BorderStyle.None, NPOI.SS.UserModel.BorderStyle borderBottom = NPOI.SS.UserModel.BorderStyle.None)
        {
            // 获取样式
            NPOI.SS.UserModel.ICellStyle cellStyle = workbook.CreateCellStyle();
            // 设置单元格样式
            if (backColorIndex > 0)
            {
                cellStyle.FillForegroundColor = backColorIndex;
                cellStyle.FillPattern = NPOI.SS.UserModel.FillPattern.SolidForeground;
            }
            cellStyle.Alignment = horizontalAlignment;
            cellStyle.VerticalAlignment = verticalAlignment;
            cellStyle.BorderLeft = borderLeft;
            cellStyle.BorderTop = borderTop;
            cellStyle.BorderRight = borderRight;
            cellStyle.BorderBottom = borderBottom;
            // 设置字体样式
            if (font != null)
            {
                NPOI.SS.UserModel.IFont cellFont = workbook.CreateFont();
                if (fontColorIndex > 0)
                { cellFont.Color = fontColorIndex; }
                cellFont.FontName = font.Name;
                cellFont.FontHeightInPoints = (short)font.Size;
                cellFont.Boldweight = (short)(font.Bold ? NPOI.SS.UserModel.FontBoldWeight.Bold : NPOI.SS.UserModel.FontBoldWeight.Normal);
                cellStyle.SetFont(cellFont);
            }
            return cellStyle;
        }
        #endregion

        #region 设置单元格的值
        public static void SetCellValue(NPOI.SS.UserModel.ICell iCell, object cellValue)
        {
            if (iCell == null)
            { return; }
            if (cellValue != null)
            {
                if (cellValue.GetType() == typeof(DBNull))
                {
                    iCell.SetCellType(NPOI.SS.UserModel.CellType.Blank);
                }
                else if (cellValue.GetType() == typeof(bool))
                {
                    iCell.SetCellValue((bool)cellValue);
                }
                else if (cellValue.GetType() == typeof(string))
                {
                    if (cellValue.ToString() == "#NUM!")
                    { iCell.SetCellType(NPOI.SS.UserModel.CellType.Blank); }
                    else
                    { iCell.SetCellValue((string)cellValue); }
                }
                else if (cellValue.GetType() == typeof(DateTime))
                {
                    NPOI.SS.UserModel.ICellStyle iCellStyle = iCell.Sheet.Workbook.CreateCellStyle();
                    NPOI.SS.UserModel.IDataFormat iDataFormat = iCell.Sheet.Workbook.CreateDataFormat();
                    iCellStyle.DataFormat = iDataFormat.GetFormat("yyyy/MM/dd HH:mm:ss");
                    iCell.CellStyle = iCellStyle;
                    iCell.SetCellValue((DateTime)cellValue);
                }
                else if (cellValue.GetType() == typeof(int) || cellValue.GetType() == typeof(double) || cellValue.GetType() == typeof(float))
                {
                    iCell.SetCellValue((double)cellValue);
                }
                else
                {
                    iCell.SetCellValue(cellValue.ToString());
                }
                
            }
        }
        #endregion

        #region 获取Sheet的最大数据列数
        public static int GetUsedColumnCount(NPOI.SS.UserModel.ISheet iSheet)
        {
            int usedColumnCount = 0;
            if (iSheet != null)
            {
                for (int rowIndex = 0; rowIndex < iSheet.LastRowNum + 1; rowIndex++)
                {
                    NPOI.SS.UserModel.IRow iRow = iSheet.GetRow(rowIndex);
                    if (iRow != null && iRow.LastCellNum > usedColumnCount)
                    { usedColumnCount = iRow.LastCellNum; }
                }
            }
            return usedColumnCount;
        }
        #endregion

        #region 从Sheet中获取DataTable
        /// <summary>
        /// 从NPOI.SS.UserModel.ISheet中获取数据表。
        /// </summary>
        /// <param name="iSheet">NPOI.SS.UserModel.ISheet对象。</param>
        /// <param name="firstRowIsColumnHead">是否将首行作为标题行。</param>
        /// <param name="maxColumnCount">最大列数，如果为0则不限制；如果大于0，则限制为指定的列数。</param>
        /// <param name="startColumnIndex">起始列的索引，从0开始。</param>
        /// <param name="startRowIndex">起始行的索引，从0开始。</param>
        /// <returns>System.Data.DataTable对象。</returns>
        public static System.Data.DataTable GetDataTable(NPOI.SS.UserModel.ISheet iSheet, bool firstRowIsColumnHead, int maxColumnCount, int startColumnIndex, int startRowIndex)
        {
            // 如果为空直接返回空。
            if (iSheet == null) { return null; }
            // 获取最大列数和最大行数。
            int usedColumnCount = Extension.NPOIMethod.GetUsedColumnCount(iSheet);
            int usedRowCount = iSheet.LastRowNum + 1;
            // 初始化DataTable。
            System.Data.DataTable dataTable = new System.Data.DataTable(iSheet.SheetName);

            #region 设置列标题
            // 声明标题行。
            NPOI.SS.UserModel.IRow headRow = null;
            // 如果首行为标题行，则获取首行为标题行。
            if (firstRowIsColumnHead)
            { headRow = iSheet.GetRow(0); }
            // 初始化DataTable最大列数为设定的最大列数。
            int maxDataTableColumnCount = maxColumnCount;
            // 如果未设定最大列数。
            if (maxColumnCount <= 0)
            {
                // 如果标题行为空，DataTable的最大列数为Sheet的最大列数-起始列的索引。
                if (headRow == null)
                { maxDataTableColumnCount = usedColumnCount - startColumnIndex; }
                else // 否则，DataTable的最大列数为标题行的列数-起始列的索引。
                { maxDataTableColumnCount = headRow.LastCellNum - startColumnIndex; }
            }
            // 创建DataTable数据列。
            for (int dataTableColumnIndex = 0; dataTableColumnIndex < maxDataTableColumnCount; dataTableColumnIndex++)
            {
                // 如果标题行为空，则添加默认数据列。
                if (headRow == null)
                { dataTable.Columns.Add("", typeof(object)); }
                else // 如果标题行不为空。
                {
                    // 从设定的起始列开始获取标题行的单元格。
                    NPOI.SS.UserModel.ICell iCell = headRow.GetCell(startColumnIndex + dataTableColumnIndex);
                    // 如果单元格为空，则添加默认数据列。
                    if (iCell == null)
                    {
                        dataTable.Columns.Add("", typeof(object));
                    }
                    else // 如果不为空，则添加名称为单元格值的数据列。
                    { dataTable.Columns.Add(iCell.ToString(), typeof(object)); }
                }
            }
            #endregion

            #region 读取数据
            // 如果起始行索引小于0，且首行最为标题行，则起始行的索引设置为1。
            if (startRowIndex <= 0 && firstRowIsColumnHead)
            { startRowIndex = 1; }
            // 循环读取Sheet中所有的数据行。
            for (int dataTableRowIndex = startRowIndex; dataTableRowIndex < usedRowCount; dataTableRowIndex++)
            {
                // 初始化DataTable数据行，并添加到DataTable。
                System.Data.DataRow dataRow = dataTable.NewRow();
                dataTable.Rows.Add(dataRow);
                // 初始化Sheet中的数据行。
                NPOI.SS.UserModel.IRow iRow = iSheet.GetRow(dataTableRowIndex);
                // 如果Sheet中的数据行不为空。
                if (iRow != null)
                {
                    // 初始化Sheet中的单元格。
                    NPOI.SS.UserModel.ICell iCell = null;
                    // 循环DataRow中的每一列。
                    for (int dataTableColumnIndex = 0; dataTableColumnIndex < dataTable.Columns.Count; dataTableColumnIndex++)
                    {
                        // 获取Sheet中的单元格，从设定的起始列开始
                        iCell = iRow.GetCell(startColumnIndex + dataTableColumnIndex);
                        // 如果Sheet中的单元格不为空，则将单元格的值赋给DataTable数据行对应的列。
                        if (iCell != null)
                        { dataRow[dataTableColumnIndex] = Extension.NPOIMethod.GetCellValue(iCell); }
                    }
                }
            }
            #endregion
            return dataTable;
        }
        #endregion

        #region 从文件中获取工作簿
        public static NPOI.SS.UserModel.IWorkbook GetWorkbookFromExcelFile(string fileName)
        {
            System.IO.FileInfo fileInfo = new System.IO.FileInfo(fileName);
            if (!fileInfo.Exists)
            { return null; }
            System.IO.FileStream fileStream = new System.IO.FileStream(fileInfo.FullName, System.IO.FileMode.Open, System.IO.FileAccess.Read);
            NPOI.SS.UserModel.IWorkbook iWorkbook = null;
            iWorkbook = NPOI.SS.UserModel.WorkbookFactory.Create(fileStream);
            //if (fileInfo.Extension == ".xls")
            //{ iWorkbook = new NPOI.HSSF.UserModel.HSSFWorkbook(fileStream); }
            //else if (fileInfo.Extension == ".xlsx")
            //{ iWorkbook = new NPOI.XSSF.UserModel.XSSFWorkbook(fileStream); }
            fileStream.Close();
            return iWorkbook;
        }
        #endregion

        #region 将工作簿保存到文件
        public static void SaveToFile(string fileName, NPOI.SS.UserModel.IWorkbook iWorkbook)
        {
            System.IO.FileInfo fileInfo = new System.IO.FileInfo(fileName);
            System.IO.FileStream fileStream = new System.IO.FileStream(fileInfo.FullName, System.IO.FileMode.Create, System.IO.FileAccess.ReadWrite);
            iWorkbook.Write(fileStream);
            fileStream.Close();
            iWorkbook.Close();
        }
        #endregion

        #region 从文件中获取DataSet
        public static System.Data.DataSet GetDataSet(string fileName, bool firstRowIsColumnHead, bool ignoreHiddenSheet)
        {
            NPOI.SS.UserModel.IWorkbook iWorkbook = Extension.NPOIMethod.GetWorkbookFromExcelFile(fileName);
            if (iWorkbook == null)
            { return null; }
            System.Data.DataSet dataSet = new System.Data.DataSet();
            for (int iSheetIndex = 0; iSheetIndex < iWorkbook.NumberOfSheets; iSheetIndex++)
            {
                bool isHiddenSheet = iWorkbook.IsSheetHidden(iSheetIndex) || iWorkbook.IsSheetHidden(iSheetIndex);
                if (ignoreHiddenSheet && isHiddenSheet)
                { continue; }
                dataSet.Tables.Add(GetDataTable(iWorkbook.GetSheetAt(iSheetIndex), firstRowIsColumnHead, 0, 0, 0));
            }
            return dataSet;
        }
        #endregion

        #region 从文件中获取DataTable
        public static System.Data.DataTable GetDataTable(string fileName, string sheetName, bool firstRowIsColumnHead)
        {
            NPOI.SS.UserModel.IWorkbook iWorkbook = Extension.NPOIMethod.GetWorkbookFromExcelFile(fileName);
            if (iWorkbook == null)
            { return null; }
            NPOI.SS.UserModel.ISheet iSheet = iWorkbook.GetSheet(sheetName);
            return GetDataTable(iSheet, firstRowIsColumnHead, 0, 0, 0);
        }

        public static System.Data.DataTable GetDataTable(string fileName, string sheetName, bool firstRowIsColumnHead, int maxColumnCount, int startColumnIndex, int startRowIndex)
        {
            NPOI.SS.UserModel.IWorkbook iWorkbook = Extension.NPOIMethod.GetWorkbookFromExcelFile(fileName);
            if (iWorkbook == null)
            { return null; }
            NPOI.SS.UserModel.ISheet iSheet = iWorkbook.GetSheet(sheetName);
            return GetDataTable(iSheet, firstRowIsColumnHead, maxColumnCount, startColumnIndex, startRowIndex);
        }
        #endregion

        #region 将DataTable填充到Sheet
        public static void FullSheet(NPOI.SS.UserModel.ISheet iSheet, System.Data.DataTable dataTable, bool firstRowIsColumnHead)
        {
            if (iSheet == null)
            { throw new AggregateException("参数NPOI.SS.UserModel.ISheet对象为null。"); }
            NPOI.SS.UserModel.IRow iRow = iSheet.CreateRow(0);
            for (int dataTableColumnIndex = 0; dataTableColumnIndex < dataTable.Columns.Count; dataTableColumnIndex++)
            {
                NPOI.SS.UserModel.ICell iCell = iRow.CreateCell(dataTableColumnIndex);
                iCell.SetCellValue(dataTable.Columns[dataTableColumnIndex].ColumnName);
            }
            int startRowIndex = 0;
            if (firstRowIsColumnHead)
            { startRowIndex = 1; }
            for (int dataTableRowIndex = 0; dataTableRowIndex < dataTable.Rows.Count; dataTableRowIndex++)
            {
                iRow = iSheet.CreateRow(dataTableRowIndex + startRowIndex);
                for (int dataTableColumnIndex = 0; dataTableColumnIndex < dataTable.Columns.Count; dataTableColumnIndex++)
                {
                    NPOI.SS.UserModel.ICell iCell = iRow.CreateCell(dataTableColumnIndex);
                    SetCellValue(iCell, dataTable.Rows[dataTableRowIndex][dataTableColumnIndex]);
                }
            }
        }
        #endregion

        #region 将DataTable保存到文件
        public static void SaveToFile(string fileName, System.Data.DataSet dataSet, bool firstRowIsColumnHead)
        {
            System.IO.FileInfo fileInfo = new System.IO.FileInfo(fileName);
            NPOI.SS.UserModel.IWorkbook iWorkbook = null;
            if (fileInfo.Extension == ".xls")
            { iWorkbook = new NPOI.HSSF.UserModel.HSSFWorkbook(); }
            else if (fileInfo.Extension == ".xlsx")
            { iWorkbook = new NPOI.XSSF.UserModel.XSSFWorkbook(); }
            for (int dataTableIndex = 0; dataTableIndex < dataSet.Tables.Count; dataTableIndex++)
            {
                System.Data.DataTable dataTable = dataSet.Tables[dataTableIndex];
                NPOI.SS.UserModel.ISheet iSheet = iWorkbook.CreateSheet(dataTable.TableName);
                FullSheet(iSheet, dataTable, firstRowIsColumnHead);
            }
            SaveToFile(fileName, iWorkbook);
        }
        #endregion

    }
}
