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
    public partial class AddPublish : Form
    {
        public AddPublish()
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
                SqlText = "INSERT INTO dbo.[Издание] ([Код_издания], [Вид_издания]) VALUES (";
                SqlText = SqlText + "\'" + textBox1.Text + "\', ";
                SqlText = SqlText + "\'" + textBox2.Text + "\')";
                MyExecuteNonQuery(SqlText);
            }
            textBox1.Text = "";
            textBox2.Text = "";
        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
