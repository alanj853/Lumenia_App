using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace ConsoleApplication2
{
    class Location
    {
        private int row;
        private int column;
        int cellType = 0;

        public Location(int row, int column)
        {
            this.row = row;
            this.column = column;
        }

        public Location(int row, int column, int cellType)
        {
            this.row = row;
            this.column = column;
            this.cellType = cellType;
        }

        public int getCellType()
        {
            return this.cellType;
        }

        public int getRow()
        {
            return row;
        }

        public void setRow(int row)
        {
            this.row = row;
        }

        public int getColumn()
        {
            return column;
        }

        public String getAddress(){
            return "(" + row.ToString() + "," + column.ToString() + ")";
        }

        public String getExcelAddress()
        {
            const string letters = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";

            String value = "";
            int col = column - 1;

            if (col >= letters.Length)
                value += letters[col / letters.Length - 1];

            value += letters[col % letters.Length];

            return value + row.ToString();// +"  " + getAddress();
        }

        private String getExcelAddress(Excel.Worksheet sheet)
        {
            String address = "";

            Excel.Range r = sheet.Cells[row, column];
            address = r.get_Address(row, column, Excel.XlReferenceStyle.xlA1);

            address = address.Replace("$", "");

            return address;
        }

        public void setColumn(int column)
        {
            this.column = column;
        }
    }
}
