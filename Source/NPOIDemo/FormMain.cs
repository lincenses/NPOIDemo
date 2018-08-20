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

            DateTime startTime = DateTime.Now;

            //NPOI.SS.UserModel.IWorkbook iWorkbook = Extension.NPOIMethod.GetWorkbookFromExcelFile(fileName);
            //NPOI.SS.UserModel.ISheet iSheet = iWorkbook.GetSheetAt(iWorkbook.ActiveSheetIndex);
            //DataTable dataTable = GetDataTable(iSheet, true, 0, 0, 0, false);

            DataSet dataSet = GetDataSet(fileName, true, true);

            string time = DateTime.Now.Subtract(startTime).ToString();

            dataGridViewMain.DataSource = dataSet.Tables[0];

            MessageBox.Show(time);
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
            if (iSheet == null) { return null; }
            System.Data.DataTable dataTable = new System.Data.DataTable(iSheet.SheetName);
            NPOI.SS.UserModel.IRow iRow = null;
            if (firstRowIsColumnHead)
            { iRow = iSheet.GetRow(0); }
            else
            { iRow = iSheet.GetRow(startRowIndex); }
            int columnCount = 0;
            if (iRow != null)
            {
                for (int iColumnIndex = startColumnIndex; iColumnIndex < iRow.LastCellNum; iColumnIndex++)
                {
                    if (maxColumnCount > 0 && columnCount >= maxColumnCount)
                    { break; }
                    NPOI.SS.UserModel.ICell iCell = iRow.GetCell(iColumnIndex);
                    if (iCell != null && firstRowIsColumnHead)
                    { dataTable.Columns.Add(iCell.ToString()); }
                    else
                    { dataTable.Columns.Add(); }
                    columnCount++;
                }
            }
            for (int i = 0; i < maxColumnCount - columnCount; i++)
            {
                dataTable.Columns.Add();
            }
            if (startRowIndex == 0 && firstRowIsColumnHead)
            { startRowIndex = 1; }
            int lastRowNumber = iSheet.LastRowNum + 1;
            for (int iRowIndex = startRowIndex; iRowIndex < lastRowNumber; iRowIndex++)
            {
                iRow = iSheet.GetRow(iRowIndex);
                if (iRow != null)
                {
                    System.Data.DataRow dataRow = dataTable.NewRow();
                    dataTable.Rows.Add(dataRow);
                    int dataTableColumnIndex = 0;
                    for (int iColumnIndex = startColumnIndex; iColumnIndex < iRow.LastCellNum; iColumnIndex++)
                    {
                        if (dataTableColumnIndex >= dataTable.Columns.Count)
                        {
                            if (maxColumnCount == 0 && !firstRowIsColumnHead)
                            { dataTable.Columns.Add(); }
                            else
                            { break; }
                        }
                        dataRow[dataTableColumnIndex] = Extension.NPOIMethod.GetCellValue(iRow.GetCell(iColumnIndex));
                        dataTableColumnIndex++;
                    }
                }
            }
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



    }
}
