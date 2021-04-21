using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.SqlClient;
using Excel = Microsoft.Office.Interop.Excel;
namespace Курсовая
{
    public partial class Main : Form
    {
        public Main()
        {
            InitializeComponent();
        }
        public int actions;
        private void Main_Load(object sender, EventArgs e)
        {
            actions = 1;
        }
        public string conString = "Data Source=DESKTOP-IPB1D3P;Initial Catalog=library;Integrated Security=True";
        public void MyExecuteNonQuery(string SqlText)
        {
            SqlConnection cn;
            SqlCommand cmd;
            cn = new SqlConnection(conString);
            cn.Open(); 
            cmd = cn.CreateCommand(); 
            cmd.CommandText = SqlText; 
            cmd.ExecuteNonQuery(); 
            cn.Close(); 
        }

        private void добавитьГруппуToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (actions == 1) MessageBox.Show("У вас нет прав на выполнение данной операции");
            else
            {
                AddGroup f = new AddGroup();
                f.ShowDialog();
            }
        }

        private void редактироватьИнформациюToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (actions == 1) MessageBox.Show("У вас нет прав на выполнение данной операции");
            else
            {
                UpdateGroup f = new UpdateGroup();
                f.ShowDialog();
            }
        }

        private void удалитьГруппуToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (actions != 3) MessageBox.Show("У вас нет прав на выполнение данной операции");
            else
            {
                DeleteGroup f = new DeleteGroup();
                f.ShowDialog();
            }
        }

        private void добавитьКнигуToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (actions == 1) MessageBox.Show("У вас нет прав на выполнение данной операции");
            else
            {
                AddBook f = new AddBook();
                f.ShowDialog();
            }
        }

        private void удалитьКнигуToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (actions != 3) MessageBox.Show("У вас нет прав на выполнение данной операции");
            else
            {
                DeleteBook f = new DeleteBook();
                f.ShowDialog();
            }
        }

        private void добавитьФакультетToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (actions == 1) MessageBox.Show("У вас нет прав на выполнение данной операции");
            else
            {
                AddFaculty f = new AddFaculty();
                f.ShowDialog();
            }

        }

        private void удалитьФакультетToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (actions != 3) MessageBox.Show("У вас нет прав на выполнение данной операции");
            else
            {
                DeleteFaculty f = new DeleteFaculty();
                f.ShowDialog();
            }
        }

        private void добавитьФормуToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (actions == 1) MessageBox.Show("У вас нет прав на выполнение данной операции");
            else
            {
                AddForm f = new AddForm();
                f.ShowDialog();
            }
        }

        private void удалитьФормуToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (actions != 3) MessageBox.Show("У вас нет прав на выполнение данной операции");
            else
            {
                DeleteForm f = new DeleteForm();
                f.ShowDialog();
            }
        }

        private void добавитьТематикуToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (actions == 1) MessageBox.Show("У вас нет прав на выполнение данной операции");
            else
            {
                AddTheme f = new AddTheme();
                f.ShowDialog();
            }
        }

        private void удалитьТематикуToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (actions != 3) MessageBox.Show("У вас нет прав на выполнение данной операции");
            else
            {
                DeleteTheme f = new DeleteTheme();
                f.ShowDialog();
            }
        }

        private void добавитьВидИзданияToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (actions == 1) MessageBox.Show("У вас нет прав на выполнение данной операции");
            else
            {
                AddPublish f = new AddPublish();
                f.ShowDialog();
            }
        }

        private void удалитьВидИзданияToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (actions != 3) MessageBox.Show("У вас нет прав на выполнение данной операции");
            else
            {
                DeletePublish f = new DeletePublish();
                f.ShowDialog();
            }
        }

        private void вывестиСписокToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            string SqlText = "SELECT * FROM dbo.[Книга]";
            SqlDataAdapter da = new SqlDataAdapter(SqlText, conString);
            DataSet ds = new DataSet();
            da.Fill(ds, "dbo.[Книга]");
            dataGridView1.DataSource = ds.Tables["dbo.[Книга]"].DefaultView;
        }

        private void вывестиСписокToolStripMenuItem_Click(object sender, EventArgs e)
        {
            string SqlText = "SELECT * FROM dbo.[Факультет]";
            SqlDataAdapter da = new SqlDataAdapter(SqlText, conString);
            DataSet ds = new DataSet();
            da.Fill(ds, "dbo.[Факультет]");
            dataGridView1.DataSource = ds.Tables["dbo.[Факультет]"].DefaultView;

        }

        private void вывестиСписокToolStripMenuItem5_Click(object sender, EventArgs e)
        {
            string SqlText = "SELECT * FROM dbo.[Тематика]";
            SqlDataAdapter da = new SqlDataAdapter(SqlText, conString);
            DataSet ds = new DataSet();
            da.Fill(ds, "dbo.[Тематика]");
            dataGridView1.DataSource = ds.Tables["dbo.[Тематика]"].DefaultView;
        }

        private void вывестиСписокToolStripMenuItem4_Click(object sender, EventArgs e)
        {
            string SqlText = "SELECT * FROM dbo.[Обучение]";
            SqlDataAdapter da = new SqlDataAdapter(SqlText, conString);
            DataSet ds = new DataSet();
            da.Fill(ds, "dbo.[Обучение]");
            dataGridView1.DataSource = ds.Tables["dbo.[Обучение]"].DefaultView;
        }

        private void вывестиСписокToolStripMenuItem3_Click(object sender, EventArgs e)
        {
            string SqlText = "SELECT * FROM dbo.[Издание]";
            SqlDataAdapter da = new SqlDataAdapter(SqlText, conString);
            DataSet ds = new DataSet();
            da.Fill(ds, "dbo.[Издание]");
            dataGridView1.DataSource = ds.Tables["dbo.[Издание]"].DefaultView;
        }

        private void вывестиСписокToolStripMenuItem2_Click(object sender, EventArgs e)
        {
            string SqlText = "SELECT * FROM dbo.[Группа]";
            SqlDataAdapter da = new SqlDataAdapter(SqlText, conString);
            DataSet ds = new DataSet();
            da.Fill(ds, "dbo.[Группа]");
            dataGridView1.DataSource = ds.Tables["dbo.[Группа]"].DefaultView;
        }


        private void редактироватьИнформациюToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            if (actions == 1) MessageBox.Show("У вас нет прав на выполнение данной операции");
            else
            {
                UpdateBook f = new UpdateBook();
                f.ShowDialog();
            }
        }

        private void вывестиСписокToolStripMenuItem6_Click(object sender, EventArgs e)
        {
            string SqlText = "SELECT * FROM dbo.[Выдача]";
            SqlDataAdapter da = new SqlDataAdapter(SqlText, conString);
            DataSet ds = new DataSet();
            da.Fill(ds, "dbo.[Выдача]");
            dataGridView1.DataSource = ds.Tables["dbo.[Выдача]"].DefaultView;
        }

        private void добавитьЗаписьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (actions == 1) MessageBox.Show("У вас нет прав на выполнение данной операции");
            else
            {
                AddExtradition f = new AddExtradition();
                f.ShowDialog();
            }
        }

        private void удалитьЗаписьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (actions != 3) MessageBox.Show("У вас нет прав на выполнение данной операции");
            else
            {
                DeleteExtradition f = new DeleteExtradition();
                f.ShowDialog();
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string lo = textBox1.Text;
            string pa = textBox2.Text;
            string SqlText = "SELECT * FROM dbo.[User] WHERE dbo.[User].[login] = " + "'" + lo + "'";
            SqlText = SqlText + " AND dbo.[User].[password] = " + "'" + pa + "';";
            SqlDataAdapter da = new SqlDataAdapter(SqlText, conString);
            
            DataTable table = new DataTable();
            da.Fill(table);
            if (table.Rows.Count > 0)
            {
                switch (textBox1.Text)
                {
                    case "admin": { actions = 3; MessageBox.Show("Вы вошли как администратор!\n Вам доступны следующие действия:\n Просмотр информации\n Добавление информации\n Редактирование информации\n Удаление информации"); break; }
                    case "user": { actions = 2; MessageBox.Show("Вы вошли как сотрудник!\n Вам доступны следующие действия:\n Просмотр информации\n Добавление информации\n Редактирование информации"); break; }
                    case "guest": { actions = 1; MessageBox.Show("Вы вошли как гость!\n Вам доступен только просмотр данных.\n Пожалуйста, авторизируйтесь или свяжитесь с администратором!"); break; }
                }
            }
            else MessageBox.Show("Неправильное имя пользователя или пароль!");
            textBox1.Text = "";
            textBox2.Text = "";
        }

        private void button2_Click(object sender, EventArgs e)
        {
            string SqlText = "SELECT dbo.[Книга].[Название_книги], dbo.[Книга].[Автор], dbo.[Выдача].[Код_группы], " +
                "dbo.[Выдача].[Книг_выдано] FROM dbo.[Выдача], dbo.[Книга] WHERE dbo.[Книга].[Код_книги] = dbo.[Выдача].[Код_книги]";
            SqlDataAdapter da = new SqlDataAdapter(SqlText, conString);
            DataSet ds = new DataSet();
            da.Fill(ds, "[library]");
            dataGridView1.DataSource = ds.Tables["[library]"].DefaultView;
        }

        private void button3_Click(object sender, EventArgs e)
        {
            string SqlText = "SELECT  dbo.[Книга].[Название_книги], dbo.[Книга].[Автор], dbo.[Выдача].[Дата_выдачи] " +
               " from dbo.[Книга], dbo.[Выдача] where  (dbo.[Выдача].[Дата_выдачи] < '" + DateTime.Now.AddMonths(-2).ToString("dd-MM-yyyy") + "') and (dbo.[Выдача].[Код_книги] = dbo.[Книга].[Код_книги])"; 
            SqlDataAdapter da = new SqlDataAdapter(SqlText, conString);
            DataSet ds = new DataSet();
            da.Fill(ds, "[library]");
            dataGridView1.DataSource = ds.Tables["[library]"].DefaultView;
        }

        private void button4_Click(object sender, EventArgs e)
        {
            string SqlText = "SELECT dbo.[Книга].[Название_книги], dbo.[Книга].[Автор], dbo.[Книга].[Количество_экземпляров], dbo.[Выдача].[Книг_выдано], dbo.[Выдача].[Код_группы]" +
            "FROM dbo.[Книга], dbo.[Выдача] WHERE dbo.[Книга].[Код_книги] = dbo.[Выдача].[Код_книги] AND " +
            "(dbo.[Книга].[Количество_экземпляров] - dbo.[Выдача].[Книг_выдано] <= (dbo.[Книга].[Количество_экземпляров] / 2))";
            SqlDataAdapter da = new SqlDataAdapter(SqlText, conString);
            DataSet ds = new DataSet();
            da.Fill(ds, "[library]");
            dataGridView1.DataSource = ds.Tables["[library]"].DefaultView;
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            string fa = comboBox1.Text;
            string qu = "";
            switch (fa)
            {
                case "Инженерный": { qu = "ИФ"; break; }
                case "Педагогика и Психология": { qu = "ПиП"; break; }
                case "Славянские и Германские языки": { qu = "СиГЯ"; break; }
                case "Экономика и Право": { qu = "ФЭП"; break; }
            }
            string SqlText = "SELECT dbo.[Группа].[Код_группы] FROM dbo.[Группа] WHERE dbo.[Группа].[Код_факультета] = ";
            SqlText = SqlText + "'" + qu + "'";
            SqlDataAdapter da = new SqlDataAdapter(SqlText, conString);
            DataSet ds = new DataSet();
            da.Fill(ds, "[library]");
            dataGridView1.DataSource = ds.Tables["[library]"].DefaultView;
        }

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            string fa = comboBox2.Text;
            int qu = 0;
            switch (fa)
            {
                case "Программирование": { qu = 1; break; }
                case "Биология": { qu = 2; break; }
                case "Психология": { qu = 3; break; }
                case "Математика": { qu = 4; break; }
                case "Философия": { qu = 5; break; }
            }
            string SqlText = "SELECT distinct dbo.[Книга].[Название_Книги], dbo.[Книга].[Автор], dbo.[Книга].[Количество_экземпляров], dbo.[Книга].[Год_издания]" +
                " FROM dbo.[Книга], dbo.[Тематика] WHERE dbo.[Книга].[Код_тематики] = ";
            SqlText = SqlText + "'" + qu + "'";
            SqlDataAdapter da = new SqlDataAdapter(SqlText, conString);
            DataSet ds = new DataSet();
            da.Fill(ds, "[library]");
            dataGridView1.DataSource = ds.Tables["[library]"].DefaultView;
        }

        private void comboBox3_SelectedIndexChanged(object sender, EventArgs e)
        {
            string fa = comboBox3.Text;
            string qu = "";
            switch (fa)
            {
                case "Дневная": { qu = "Д"; break; }
                case "Заочная": { qu = "З"; break; }
                case "Дневная Сокращенная": { qu = "ДС"; break; }
                case "Заочная Сокращенная": { qu = "ЗС"; break; }
            }
            string SqlText = "SELECT distinct dbo.[Группа].[Код_группы], dbo.[Группа].[Количество_обучающихся] " + 
                " FROM dbo.[Группа] WHERE dbo.[Группа].[Код_формы_обучения] = ";
            SqlText = SqlText + "'" + qu + "'";
            SqlDataAdapter da = new SqlDataAdapter(SqlText, conString);
            DataSet ds = new DataSet();
            da.Fill(ds, "[library]");
            dataGridView1.DataSource = ds.Tables["[library]"].DefaultView;
        }

        private void button5_Click(object sender, EventArgs e)
        {
            
            string SqlText = "SELECT distinct dbo.[Группа].[Код_группы], dbo.[Группа].[Количество_обучающихся], dbo.[Факультет].[Факультет]," +
                "dbo.[Обучение].[Форма_обучения]" +
                " FROM dbo.[Группа], dbo.[Факультет], dbo.[Обучение]" +
                " WHERE dbo.[Группа].[Код_формы_обучения] = dbo.[Обучение].[Код_формы_обучения] AND dbo.[Группа].[Код_факультета] = dbo.[Факультет].[Код_факультета]";
            
            SqlDataAdapter da = new SqlDataAdapter(SqlText, conString);
            DataSet ds = new DataSet();
            da.Fill(ds, "[library]");
            dataGridView1.DataSource = ds.Tables["[library]"].DefaultView;
        }

        private void button6_Click(object sender, EventArgs e)
        {
            string SqlText = "SELECT distinct dbo.[Книга].[Название_книги], dbo.[Книга].[Автор], dbo.[Книга].[Год_издания]," +
                "dbo.[Издание].[Вид_издания], dbo.[Тематика].[Название_тематики]" +
                " FROM dbo.[Книга], dbo.[Издание], dbo.[Тематика]" +
                " WHERE dbo.[Книга].[Код_тематики] = dbo.[Тематика].[Код_тематики] AND dbo.[Книга].[Код_издания] = dbo.[Издание].[Код_издания]";

            SqlDataAdapter da = new SqlDataAdapter(SqlText, conString);
            DataSet ds = new DataSet();
            da.Fill(ds, "[library]");
            dataGridView1.DataSource = ds.Tables["[library]"].DefaultView;
        }

        private void button7_Click(object sender, EventArgs e)
        {
            saveFileDialog1.Filter = "xls files (*.xlsx)|*.xlsx|All files (*.*)|*.*";

            if (saveFileDialog1.ShowDialog() == DialogResult.OK)
            {
                string path = saveFileDialog1.FileName;

                Excel.Application excelapp = new Excel.Application();
                Excel.Workbook workbook = excelapp.Workbooks.Add();
                Excel.Worksheet worksheet = workbook.ActiveSheet;

                for (int i = 1; i < dataGridView1.ColumnCount + 1; i++)
                    worksheet.Rows[1].Columns[i] = dataGridView1.Columns[i - 1].HeaderText;

                for (int i = 1; i < dataGridView1.RowCount + 1; i++)
                    for (int j = 1; j < dataGridView1.ColumnCount + 1; j++)
                        worksheet.Rows[i + 1].Columns[j] = dataGridView1.Rows[i - 1].Cells[j - 1].Value;

                excelapp.Columns.AutoFit();
                excelapp.AlertBeforeOverwriting = false;
                workbook.SaveAs(path);
                excelapp.Quit();
            }
        }

        private void редактироватьЗаписьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            UpdateExtradition f = new UpdateExtradition();
            f.ShowDialog();
        }
    }
}
 