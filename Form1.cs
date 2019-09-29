using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using MySql.Data.MySqlClient;
using Excel = Microsoft.Office.Interop.Excel;
namespace test
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Button1_Click(object sender, EventArgs e)
        {
            Excel.Application excelApp = new Excel.Application();  // Создаём экземпляр нашего приложения
            
            Excel.Workbook workBook; // Создаём экземпляр рабочий книги Excel
        
            Excel.Worksheet workSheet; // Создаём экземпляр листа Excel

            workBook = excelApp.Workbooks.Add();
            workSheet = (Excel.Worksheet)workBook.Worksheets.get_Item(1);
            
            var j2 = 1;
            for (int j = 2; j2 <= 4; j++)
            {
                var strng = $"Число: {j}";
                
                workSheet.Cells[j - 1, j2] = strng;
                workSheet.Cells[j - 1, j2].Font.Size = 30; // Изменение размера шрифта.
                workSheet.Cells[j - 1, j2].Font.Color = Color.Green; // Изменение цвета текста.
                if (j == 5)
                {
                    j = 1;
                    j2++;
                }
                #region Border(рамка)
                Excel.Range rng = workSheet.Range[$"A{j}"];
                Excel.Borders border = rng.Borders;
                border.LineStyle = Excel.XlLineStyle.xlContinuous;
                rng = workSheet.Range[$"B{j}"];
                border = rng.Borders;
                border.LineStyle = Excel.XlLineStyle.xlContinuous;
                 rng = workSheet.Range[$"C{j}"];
                border = rng.Borders;
                border.LineStyle = Excel.XlLineStyle.xlContinuous;
                 rng = workSheet.Range[$"D{j}"];
                border = rng.Borders;
                border.LineStyle = Excel.XlLineStyle.xlContinuous;
                 rng = workSheet.Range[$"E{j}"];
                border = rng.Borders;
                border.LineStyle = Excel.XlLineStyle.xlContinuous;
                 rng = workSheet.Range[$"F{j}"];
                border = rng.Borders;
                border.LineStyle = Excel.XlLineStyle.xlContinuous;
                #endregion
            }



            /// Открываем созданный excel-файл
            excelApp.Visible = true;
            excelApp.UserControl = true;
        }
    }
}
