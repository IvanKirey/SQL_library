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
namespace Курсовая
{
    public partial class AddExtradition : Form
    {
        public AddExtradition()
        {
            InitializeComponent();
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
        private void button1_Click(object sender, EventArgs e)
        {
            string SqlText;
            {
                SqlText = "INSERT INTO dbo.[Выдача] ([Код_выдачи], [Код_группы], [Код_книги], [Книг_выдано], [Дата_выдачи], [Дата_возврата]) VALUES (";
                SqlText = SqlText + "\'" + textBox1.Text + "\', ";
                SqlText = SqlText + "\'" + textBox2.Text + "\', ";
                SqlText = SqlText + "\'" + textBox3.Text + "\', ";
                SqlText = SqlText + "\'" + textBox4.Text + "\', ";
                SqlText = SqlText + "\'" + textBox5.Text + "\', ";
                if (textBox6.Text == "") SqlText = SqlText + "NULL)"; else
                SqlText = SqlText + "\'" + textBox6.Text + "\')";
                MyExecuteNonQuery(SqlText);
            }
            textBox1.Text = "";
            textBox2.Text = "";
            textBox3.Text = "";
            textBox4.Text = "";
            textBox5.Text = "";
            textBox6.Text = "";   
        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
