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
        public static dynamic GetCellValue(NPOI.SS.UserModel.ICell cell)
        {
            if (cell == null)
            { return null; }
            else
            {
                switch (cell.CellType)
                {
                    case NPOI.SS.UserModel.CellType.Blank:
                        return "";
                    case NPOI.SS.UserModel.CellType.Boolean:
                        return cell.BooleanCellValue;
                    case NPOI.SS.UserModel.CellType.Error:
                        return cell.ErrorCellValue.ToString(System.Globalization.CultureInfo.InvariantCulture);
                    case NPOI.SS.UserModel.CellType.Formula:
                        if (cell.CachedFormulaResultType == NPOI.SS.UserModel.CellType.Error)
                        { return "#NUM!"; }
                        else
                        { return cell.NumericCellValue; }
                    case NPOI.SS.UserModel.CellType.Numeric:
                        if (NPOI.SS.UserModel.DateUtil.IsCellDateFormatted(cell))
                        { return cell.DateCellValue; }
                        else
                        { return cell.NumericCellValue; }
                    case NPOI.SS.UserModel.CellType.String:
                        return cell.StringCellValue;
                    default:
                        return cell.ToString();
                }
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

        #region 从文件中获取工作簿
        public static NPOI.SS.UserModel.IWorkbook GetWorkbookFromExcelFile(string fileName)
        {
            System.IO.FileInfo fileInfo = new System.IO.FileInfo(fileName);
            if (!fileInfo.Exists)
            { return null; }
            System.IO.FileStream fileStream = new System.IO.FileStream(fileInfo.FullName, System.IO.FileMode.Open, System.IO.FileAccess.Read);
            NPOI.SS.UserModel.IWorkbook workbook = NPOI.SS.UserModel.WorkbookFactory.Create(fileStream);
            fileStream.Close();
            return workbook;
        }
        #endregion

        #region 将工作簿保存到文件
        public static void SaveToFile(NPOI.SS.UserModel.IWorkbook workbook, string fileName)
        {
            System.IO.FileInfo fileInfo = new System.IO.FileInfo(fileName);
            System.IO.FileStream fileStream = new System.IO.FileStream(fileInfo.FullName, System.IO.FileMode.OpenOrCreate, System.IO.FileAccess.ReadWrite);
            workbook.Write(fileStream);
            fileStream.Close();
            workbook.Close();
        }
        #endregion

        



    }
}
