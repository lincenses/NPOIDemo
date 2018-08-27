using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace NPOIDemo
{
    public partial class FormMain : Form
    {
        public FormMain()
        {
            InitializeComponent();
            InitializeEvent();

            label1.Text = "";
        }

        #region 初始化事件
        private void InitializeEvent()
        {
            toolStripButtonLoad.Click += ToolStripButtonLoad_Click;

        }
        #endregion

        #region Load
        private void ToolStripButtonLoad_Click(object sender, EventArgs e)
        {
            string fileName = "AAA.xlsx";

            OpenFileDialog fileDialog = new OpenFileDialog();
            fileDialog.Filter = "*.xlsx|*.xlsx|*.xls|*.xls";
            if (fileDialog.ShowDialog() == DialogResult.OK)
            { fileName = fileDialog.FileName; }

            DateTime startTime = DateTime.Now;

            //NPOI.SS.UserModel.IWorkbook iWorkbook = Extension.NPOIMethod.GetWorkbookFromExcelFile(fileName);
            //NPOI.SS.UserModel.ISheet iSheet = iWorkbook.GetSheetAt(iWorkbook.ActiveSheetIndex);
            //DataTable dataTable = GetDataTable(iSheet, true, 0, 0, 0, false);

            //DataSet dataSet = GetDataSet(fileName, true, true);

            DataTable dataTable = GetDataTable(fileName, "sheet1", true);

            string time = DateTime.Now.Subtract(startTime).ToString();

            dataGridViewMain.DataSource = dataTable;
            //List<string> tableNames = new List<string>();
            //for (int i = 0; i < dataSet.Tables.Count; i++)
            //{
            //    tableNames.Add(dataSet.Tables[i].TableName);
            //}
            //label1.Text = string.Join(",", tableNames.ToArray());

            MessageBox.Show(time);
        }
        #endregion

        #region 获取Sheet的最大列数
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
        /// 
        /// </summary>
        /// <param name="iSheet"></param>
        /// <param name="firstRowIsColumnHead"></param>
        /// <param name="maxColumnCount">最大列数，默认值为0。
        /// <para>如果小于或等于0，则读取所有列的数据。</para>
        /// <para>如果同时指定了首行作为标题行的话，最大列数为标题行最大列数。</para>
        /// <para>如果不将首行作为标题行，最大列数为整个Sheet的最大列数。</para>
        /// <para>如果大于0，则只读取指定列数的数据。</para>
        /// <para>如果同时指定首行作为标题行的话，最大列数为指定的最大列数，如果标题行列数小于指定的最大列数，则补足到指定的最大列数；</para>
        /// <para>如果标题行列数大于指定的最大列数，超出部分自动舍弃。</para>
        /// <para>如果不将首行作为标题行，则按照指定的最大列数读取数据，超出部分自动舍弃。</para>
        /// </param>
        /// <param name="startRowIndex"></param>
        /// <param name="startColumnIndex"></param>
        /// <param name="ignoreBlankRow"></param>
        /// <returns></returns>
        private static System.Data.DataTable GetDataTable(NPOI.SS.UserModel.ISheet iSheet, bool firstRowIsColumnHead, int maxColumnCount, int startRowIndex, int startColumnIndex)
        {
            // 如果为空直接返回空。
            if (iSheet == null) { return null; }
            int usedColumnCount = GetUsedColumnCount(iSheet);
            int usedRowCount = iSheet.LastRowNum + 1;
            // 初始化DataTable。
            System.Data.DataTable dataTable = new System.Data.DataTable(iSheet.SheetName);

            #region 设置列标题
            // 声明IRow。
            NPOI.SS.UserModel.IRow iHeadRow = null;
            // 如果首行为标题行，则获取首行，
            if (firstRowIsColumnHead)
            { iHeadRow = iSheet.GetRow(0); }
            else // 否则获取指定的起始行。
            { iHeadRow = iSheet.GetRow(startRowIndex); }
            // 初始化当前列数，用于判断是否超过最大列数。
            int currentDataTableColumnCount = 0;
            // 如果标题行不为空。
            if (iHeadRow != null)
            {
                // 循环读取一行里的所有单元格，从指定的列的索引开始。
                for (int cellIndex = startColumnIndex; cellIndex < iHeadRow.LastCellNum; cellIndex++)
                {
                    // 如果设置了最大列数，并且当前列数大于或等于最大列数则跳出循环。
                    if (maxColumnCount > 0 && currentDataTableColumnCount >= maxColumnCount)
                    { break; }
                    // 获取单元格。
                    NPOI.SS.UserModel.ICell iCell = iHeadRow.GetCell(cellIndex);
                    // 如果单元格不为空且首行为标题行的话，添加一个以单元格的值作为列标题的列，
                    if (iCell != null && firstRowIsColumnHead)
                    { dataTable.Columns.Add(iCell.ToString()); }
                    else // 否则添加一个默认列标题的列。
                    { dataTable.Columns.Add(); }
                    // 当前列数累加1。
                    currentDataTableColumnCount++;
                }
            }
            // 如果设定的最大列数大于实际列数，则添加默认列达到指定的最大列数。
            for (int i = 0; i < maxColumnCount - currentDataTableColumnCount; i++)
            { dataTable.Columns.Add(); }
            #endregion

            #region 读取数据
            // 如果起始行索引为0，且首行最为标题行，则起始行的索引设置为1。
            if (startRowIndex == 0 && firstRowIsColumnHead)
            { startRowIndex = 1; }
            // 初始化空白行列表。
            List<System.Data.DataRow> blankDataRowList = new List<System.Data.DataRow>();
            // 循环获取Sheet中的所有行。
            for (int rowIndex = startRowIndex; rowIndex < iSheet.LastRowNum + 1; rowIndex++)
            {
                // 初始化DataRow。
                System.Data.DataRow dataRow = dataTable.NewRow();
                // 读取Sheet中的数据行。
                NPOI.SS.UserModel.IRow iRow = iSheet.GetRow(rowIndex);
                // 判断是否为空，如果为空，则按照空行处理，添加到空行列表。
                if (iRow == null)
                { blankDataRowList.Add(dataRow); }
                else
                {
                    // 初始化Cell的值的列表。
                    List<object> cellValues = iRow.Select(x => Extension.NPOIMethod.GetCellValue(x)).ToList();
                    // 判断是否是空行，如果是空行，则添加到空行列表。
                    if (string.IsNullOrWhiteSpace(string.Join("", cellValues.ToArray())))
                    { blankDataRowList.Add(dataRow); }
                    else // 如果不是空行。
                    {
                        // 将空行列表中的空行添加到DataTable。
                        for (int blankDataRowIndex = 0; blankDataRowIndex < blankDataRowList.Count; blankDataRowIndex++)
                        { dataTable.Rows.Add(blankDataRowList[blankDataRowIndex]); }
                        // 清空空行列表。
                        blankDataRowList.Clear();
                        // 如果不是首行为列标题行，则需要自动添加列。
                        if (!firstRowIsColumnHead)
                        {
                            // 初始化DataTable列数。
                            int dataTableColumnCount = dataTable.Columns.Count;
                            // 初始化最大列数。
                            int maxDataTableColumnCount = maxColumnCount;
                            // 如果最大列数设为0，则最大列数为数据行中单元格的个数，便于添加DataTable的列。
                            if (maxColumnCount == 0)
                            { maxDataTableColumnCount = cellValues.Count; }
                            // 如果DataTable的列数小于DataTable最大列数，则添加列。
                            for (int i = dataTableColumnCount; i < maxDataTableColumnCount; i++)
                            { dataTable.Columns.Add(); }
                        }
                        // 给DataRow赋值，循环次数为DataRow的Item数和单元格集合的个数的最小值。
                        for (int dataRowItemIndex = 0; dataRowItemIndex < Math.Min(dataRow.ItemArray.Length, cellValues.Count); dataRowItemIndex++)
                        { dataRow[dataRowItemIndex] = cellValues[dataRowItemIndex]; }
                        dataTable.Rows.Add(dataRow);
                    }
                }
            }
            #endregion

            return dataTable;
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
            
            if (iSheet == null)
            { return null; }
            else
            { return GetDataTable(iSheet, firstRowIsColumnHead, 0, 0, 0); }
        }
        #endregion


    }
}
