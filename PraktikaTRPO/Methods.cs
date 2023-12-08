using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using Guna.UI2.WinForms;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Excel;
using System.Security.Cryptography.X509Certificates;
using System.Runtime.CompilerServices;

namespace PraktikaTRPO
{
    internal class Methods
    {
        public static void SearchData(DataGridView dataGrid, Guna2TextBox searchBox)
        {
            for (int i = 0; i < dataGrid.RowCount; i++)
            {
                int count = 0;
                for (int j = 1; j < dataGrid.ColumnCount; j++)
                {
                    if (dataGrid[j, i].Value.ToString().IndexOf(searchBox.Text, StringComparison.OrdinalIgnoreCase) >= 0)
                        count++;
                }
                if (count > 0)
                    dataGrid.Rows[i].DefaultCellStyle.BackColor = Color.FromArgb(40,55,70);
                else
                    dataGrid.Rows[i].DefaultCellStyle.BackColor = Color.FromArgb(193, 199, 206);
            }
            if(searchBox.Text == "")
                for (int i = 0; i < dataGrid.RowCount; i++)
                    dataGrid.Rows[i].DefaultCellStyle.BackColor = Color.FromArgb(193, 199, 206);
        }
        public static void ExportExcel(DataGridView dataGrid,string headerName,string listname) 
        {            
            Excel.Application exApp = new Excel.Application();
            exApp.Workbooks.Add();
            Excel.Worksheet wsh = (Excel.Worksheet)exApp.ActiveSheet;         
            wsh.Cells.Range[wsh.Cells[1, 1], wsh.Cells[1, dataGrid.ColumnCount-1]].Merge();
            wsh.Columns.Style.HorizontalAlignment = Excel.XlVAlign.xlVAlignCenter;
            wsh.Cells[1, 1] = $"{listname}";
            wsh.Name = $"{listname}";
           
            for (int i = 1; i < dataGrid.ColumnCount; i++)
            {
                wsh.Cells[2, i] = dataGrid.Columns[i].HeaderText;
                wsh.Cells[2, i].Borders.Value = BorderStyle.FixedSingle;
            }
            for (int i = 0; i < dataGrid.RowCount; i++)
            {
                for (int j = 1; j < dataGrid.ColumnCount; j++)
                {
                    wsh.Cells[i + 3, j] = dataGrid[j, i].Value.ToString();
                    wsh.Cells[i + 3, j].Borders.Value = BorderStyle.FixedSingle;
                }
            }
            wsh.Cells.Range[wsh.Cells[dataGrid.RowCount+3,1], wsh.Cells[dataGrid.RowCount+3, dataGrid.ColumnCount - 1]].Merge();
            wsh.Columns.Style.HorizontalAlignment = Excel.XlVAlign.xlVAlignCenter;
            wsh.Cells[dataGrid.RowCount+3,1] = $"{headerName}";
            wsh.Cells.Range[wsh.Cells[dataGrid.RowCount + 4, 1], wsh.Cells[dataGrid.RowCount + 4, dataGrid.ColumnCount - 1]].Merge();
            wsh.Columns.Style.HorizontalAlignment = Excel.XlVAlign.xlVAlignCenter;
            wsh.Cells[dataGrid.RowCount + 4, 1] = $"Составил:";
            wsh.Columns.AutoFit();
            exApp.Visible = true;
        }
    }
}
