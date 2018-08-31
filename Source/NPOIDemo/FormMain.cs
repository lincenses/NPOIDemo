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
            {
                fileName = fileDialog.FileName;
                string time = "";
                bool firstRowIsColumnHead = true;
                bool ignoreHiddenSheet = true;
                DateTime startTime = DateTime.Now;

                startTime = DateTime.Now;
                DataSet dataSet = Extension.NPOIMethod.GetDataSet(fileName, firstRowIsColumnHead, ignoreHiddenSheet);
                time = DateTime.Now.Subtract(startTime).ToString();
                MessageBox.Show("读取时间：" + time);

                dataGridViewMain.DataSource = dataSet.Tables[0];

                startTime = DateTime.Now;
                Extension.NPOIMethod.SaveToFile("NewExcel.xlsx", dataSet, firstRowIsColumnHead);
                time = DateTime.Now.Subtract(startTime).ToString();
                MessageBox.Show("保存时间：" + time);

            }

            





        }
        #endregion



        

        

        

        ////else // 否则获取指定的起始行。
        ////{ iHeadRow = iSheet.GetRow(startRowIndex); }
        //// 初始化当前列数，用于判断是否超过最大列数。
        //int currentDataTableColumnCount = 0;
        //// 如果标题行不为空。
        //if (headRow != null)
        //{
        //    // 循环读取一行里的所有单元格，从指定的列的索引开始。
        //    for (int cellIndex = startColumnIndex; cellIndex < headRow.LastCellNum; cellIndex++)
        //    {
        //        // 如果设置了最大列数，并且当前列数大于或等于最大列数则跳出循环。
        //        if (maxColumnCount > 0 && currentDataTableColumnCount >= maxColumnCount)
        //        { break; }
        //        // 获取单元格。
        //        NPOI.SS.UserModel.ICell iCell = headRow.GetCell(cellIndex);
        //        // 如果单元格不为空且首行为标题行的话，添加一个以单元格的值作为列标题的列，
        //        if (iCell != null && firstRowIsColumnHead)
        //        { dataTable.Columns.Add(iCell.ToString()); }
        //        else // 否则添加一个默认列标题的列。
        //        { dataTable.Columns.Add(); }
        //        // 当前列数累加1。
        //        currentDataTableColumnCount++;
        //    }
        //}
        //// 如果设定的最大列数大于实际列数，则添加默认列达到指定的最大列数。
        //for (int i = 0; i < maxColumnCount - currentDataTableColumnCount; i++)
        //{ dataTable.Columns.Add(); }


        //// 初始化空白行列表。
        //List<System.Data.DataRow> blankDataRowList = new List<System.Data.DataRow>();
        //// 循环获取Sheet中的所有行。
        //for (int rowIndex = startRowIndex; rowIndex < iSheet.LastRowNum + 1; rowIndex++)
        //{
        //    // 初始化DataRow。
        //    System.Data.DataRow dataRow = dataTable.NewRow();
        //    // 读取Sheet中的数据行。
        //    NPOI.SS.UserModel.IRow iRow = iSheet.GetRow(rowIndex);
        //    // 判断是否为空，如果为空，则按照空行处理，添加到空行列表。
        //    if (iRow == null)
        //    { blankDataRowList.Add(dataRow); }
        //    else
        //    {
        //        // 初始化Cell的值的列表。
        //        List<object> cellValues = iRow.Select(x => Extension.NPOIMethod.GetCellValue(x)).ToList();
        //        // 判断是否是空行，如果是空行，则添加到空行列表。
        //        if (string.IsNullOrWhiteSpace(string.Join("", cellValues.ToArray())))
        //        { blankDataRowList.Add(dataRow); }
        //        else // 如果不是空行。
        //        {
        //            // 将空行列表中的空行添加到DataTable。
        //            for (int blankDataRowIndex = 0; blankDataRowIndex < blankDataRowList.Count; blankDataRowIndex++)
        //            { dataTable.Rows.Add(blankDataRowList[blankDataRowIndex]); }
        //            // 清空空行列表。
        //            blankDataRowList.Clear();
        //            // 如果不是首行为列标题行，则需要自动添加列。
        //            if (!firstRowIsColumnHead)
        //            {
        //                // 初始化DataTable列数。
        //                int dataTableColumnCount = dataTable.Columns.Count;
        //                // 初始化最大列数。
        //                maxDataTableColumnCount = maxColumnCount;
        //                // 如果最大列数设为0，则最大列数为数据行中单元格的个数，便于添加DataTable的列。
        //                if (maxColumnCount == 0)
        //                { maxDataTableColumnCount = cellValues.Count; }
        //                // 如果DataTable的列数小于DataTable最大列数，则添加列。
        //                for (int i = dataTableColumnCount; i < maxDataTableColumnCount; i++)
        //                { dataTable.Columns.Add(); }
        //            }
        //            // 给DataRow赋值，循环次数为DataRow的Item数和单元格集合的个数的最小值。
        //            for (int dataRowItemIndex = 0; dataRowItemIndex < Math.Min(dataRow.ItemArray.Length, cellValues.Count); dataRowItemIndex++)
        //            { dataRow[dataRowItemIndex] = cellValues[dataRowItemIndex]; }
        //            dataTable.Rows.Add(dataRow);
        //        }
        //    }
        //}
    }
}
