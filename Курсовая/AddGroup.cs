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
    public partial class AddGroup : Form
    {
        public AddGroup()
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
                SqlText = "INSERT INTO dbo.[Группа] ([Код_группы], [Код_факультета], [Код_формы_обучения], [Количество_обучающихся]) VALUES (";
                SqlText = SqlText + "\'" + textBox1.Text + "\', ";
                SqlText = SqlText + "\'" + textBox2.Text + "\', ";
                SqlText = SqlText + "\'" + textBox3.Text + "\', ";
                SqlText = SqlText + "\'" + textBox4.Text + "\')";
                MyExecuteNonQuery(SqlText);
            }
            textBox1.Text = "";
            textBox2.Text = "";
            textBox3.Text = "";
            textBox4.Text = "";
        }
    }
}
