using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;

namespace CreateExcel
{
	public partial class Form1 : Form
	{
		public Form1() {
			InitializeComponent();
		}

		private void btnCreate_Click(object sender, EventArgs e) {
			Microsoft.Office.Interop.Excel.Application theExcelApp = new Microsoft.Office.Interop.Excel.Application();
			Workbook theExcelBook = theExcelApp.Workbooks.Add(true);
			Worksheet theSheet = (Worksheet)theExcelBook.ActiveSheet;
			Range theCell;

			//整体设置
			theSheet.Name = "1余额明细";
			theSheet.Application.ActiveWindow.DisplayGridlines = false;
			theSheet.Cells.Font.Name = "楷体_GB2312";
			theSheet.Cells.Font.Size = 11;
			theSheet.Cells.RowHeight = 18;

			//表序号
			theCell = theSheet.Cells[1, 1];
			theCell.Value2 = "表一";
			theCell.Font.Name = "华文细黑";
			theCell.Font.Size = 12;
	
			//标题
			theCell = theSheet.Range[theSheet.Cells[2, 1], theSheet.Cells[2, 43]];
			theCell.Merge();
			theCell.Value2 = "2009年地方政府性债务明细表";
			theCell.Font.Name = "宋体";
			theCell.Font.Size = 20;
			theCell.HorizontalAlignment = XlHAlign.xlHAlignCenter;

			//填报单位 & 数额单位
			theCell = theSheet.Range[theSheet.Cells[3,1], theSheet.Cells[3,43]];
			theCell.RowHeight = 24.00;
			theCell.VerticalAlignment = XlVAlign.xlVAlignBottom;
			theSheet.Cells[3,1].Value2 = "填报单位：";
			theCell = theSheet.Range[theSheet.Cells[3, 42], theSheet.Cells[3, 43]];
			theCell.Merge();
			theCell.Value2 = "单位：万元";

			//表头
			theCell = theSheet.Range[theSheet.Cells[4, 1], theSheet.Cells[6, 43]];
			theCell.WrapText = true;
			theCell.HorizontalAlignment = XlHAlign.xlHAlignCenter;
			theCell.Borders.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);
			//行高
			theSheet.Range[theSheet.Cells[4, 1], theSheet.Cells[4, 43]].RowHeight = 24.00;
			theSheet.Range[theSheet.Cells[5, 1], theSheet.Cells[5, 43]].RowHeight = 37.50;
			theSheet.Range[theSheet.Cells[6, 1], theSheet.Cells[6, 43]].RowHeight = 69.75;
			//列宽
			((Range) theSheet.Cells[4, 1]).ColumnWidth = 46.50;
			theSheet.Range[theSheet.Cells[4, 2], theSheet.Cells[4, 43]].ColumnWidth = 4.75;
			//债务人
			theCell = theSheet.Range[theSheet.Cells[4,1], theSheet.Cells[6,1]];
			theCell.Merge();
			theCell.Value2 = "债务人";
			//年初数
			theCell = theSheet.Range[theSheet.Cells[4, 2], theSheet.Cells[4, 11]];
			theCell.Merge();
			theCell.Value2 = "年初数";
			//合计
			theCell = theSheet.Range[theSheet.Cells[5,2], theSheet.Cells[6,2]];
			theCell.Merge();
			theCell.Value2 = "合计";
			//财政性资金偿还
			theCell = theSheet.Range[theSheet.Cells[5, 3], theSheet.Cells[5, 7]];
			theCell.Merge();
			theCell.Value2 = "财政性资金偿还";
			//非财政性资金偿还
			theCell = theSheet.Range[theSheet.Cells[5, 8], theSheet.Cells[5, 11]];
			theCell.Merge();
			theCell.Value2 = "非财政性资金偿还";
			//细项
			theSheet.Cells[6, 3].Value2 = "小计";
			theSheet.Cells[6, 4].Value2 = "一般预算";
			theSheet.Cells[6, 5].Value2 = "基金预算";
			theSheet.Cells[6, 6].Value2 = "预算外";
			theSheet.Cells[6, 7].Value2 = "国有资本经营预算";
			theSheet.Cells[6, 8].Value2 = "小计";
			theSheet.Cells[6, 9].Value2 = "事业收入";
			theSheet.Cells[6, 10].Value2 = "经营收入";
			theSheet.Cells[6, 11].Value2 = "其他";

			//各部门数据
			for (int i = 0; i < 9; i++) {

			}

			// Save to file and close excel application
			theExcelApp.DisplayAlerts = false;
			string fileName = @"E:\Lab\VS2010\CreateExcel\Test.xls";
			theExcelBook.SaveCopyAs(fileName);
			theExcelApp.Workbooks.Close();
			theExcelApp.Quit();
			GC.Collect();

			MessageBox.Show("Success");
			System.Windows.Forms.Application.Exit();
		}
	}
}
