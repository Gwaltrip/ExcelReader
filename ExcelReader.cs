using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;

using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelReader
{
    class ExcelReader
    {
        private Excel.Application xlApp;
        private Excel.Workbook xlWBook;
        private Excel.Worksheet xlSheet;
        private Excel.Range range;


        public ExcelReader(string path)
        {
            xlApp = new Excel.Application();
            xlWBook = xlApp.Workbooks.Open(path);
            xlSheet = (Excel.Worksheet)xlWBook.Worksheets.get_Item(1);
            range = xlSheet.UsedRange;
        }

        public List<double> Range(int col)
        {
            List<double> list = new List<double>();
            int lastRow = xlSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;

            for (int i = 1; i <= lastRow; i++)
            {
                list.Add(Double.Parse((range.Cells[i, col] as Excel.Range).Value2.ToString()));
            }

            return list;
        }
    }
}
