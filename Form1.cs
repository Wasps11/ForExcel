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
        string gobal;
       


        private void Button1_Click(object sender, EventArgs e)
        {
        
            #region Excel

            Excel.Application excelApp = new Excel.Application();  // Создаём экземпляр нашего приложения

            Excel.Workbook workBook; // Создаём экземпляр рабочий книги Excel

            Excel.Worksheet workSheet; // Создаём экземпляр листа Excel

            workBook = excelApp.Workbooks.Add();
            workSheet = (Excel.Worksheet)workBook.Worksheets.get_Item(1);
            workSheet.Cells.Font.Color = Color.Red;
            workSheet.Cells.Font.Name = "BroadWay"; // Изменение шрифта.
            workSheet.Cells.Font.Size = 10; // Изменение размера шрифта.
            var j2 = 1;
            for (int j = 2; j2 <= 4; j++)
            {
                // workSheet.Cells[j - 1, j2].Font.Color = Color.Green; // Изменение цвета текста.
                workSheet.Cells[2, "A"].Value2 = "test";
                if (j == 5)
                {
                    j = 1;
                    j2++;
                }
                #region Border(рамка)
                Excel.Range rng = workSheet.Range[$"A{j}" , $"D{j}"]; // Рамка от одного края до другого (то что ты просил =))
                Excel.Borders border = rng.Borders;
                border.LineStyle = Excel.XlLineStyle.xlContinuous;
                #endregion
            }


         
            /// Открываем созданный excel-файл
            excelApp.Visible = true;
            excelApp.UserControl = true;
            #endregion  
            #region BD
            #region Подключение
            string serverName1 = "127.0.0.1"; // Адрес сервера (для локальной базы пишите "localhost")
            string userName = "root"; // Имя пользователя
            string dbName = "test"; //Имя базы данных
            string port = "3306"; // Порт для подключения
            string password = ""; // Пароль для подключения
            string connStr = "server=" + serverName1 +
               ";user=" + userName +
               ";database=" + dbName +
               ";port=" + port +
               ";password=" + password + ";";

            string sql = "SELECT * FROM t_test"; // Строка запроса

            MySqlConnection connection = new MySqlConnection(connStr);
            MySqlCommand sqlCom = new MySqlCommand(sql, connection);
            connection.Open();
            #endregion
            sqlCom.ExecuteNonQuery();
            MySqlDataAdapter dataAdapter = new MySqlDataAdapter(sqlCom);
            DataTable dt = new DataTable();
            dataAdapter.Fill(dt);

            var myData = dt.Select();
            for (int i = 0; i < myData.Length; i++)
            {
                for (int j = 0; j < myData[i].ItemArray.Length; j++)
                {
                    var text = myData[i].ItemArray[j];
                    //textBox1.Text += text;
                    workSheet.Cells[$"А{j}"] = "test";
                    
                }
            }
            #endregion
        }
    }
}
