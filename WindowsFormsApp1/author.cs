using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.Sql;
using System.Data.OleDb;
using System.Data.SqlClient;

namespace WindowsFormsApp1
{
    public partial class author : Form
    {

        SqlConnection con = new SqlConnection("Data Source=(LocalDB)\\MSSQLLocalDB;AttachDbFilename=D:\\Univer\\1Diplom\\WindowsFormsApp1\\WindowsFormsApp1\\WindowsFormsApp1\\WindowsFormsApp1\\Database1.mdf;");


        public author()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            this.Hide();
            Form1 Form1 = new Form1();
            Form1.FormClosed += (s, args) => this.Close();
            Form1.Show();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            /* */
            con.Open();
            SqlCommand sqlcmd = new SqlCommand("insert into author values (@login)", con);
            sqlcmd.Parameters.AddWithValue("@login", user.Text);
            sqlcmd.ExecuteNonQuery();
            con.Close();
            MessageBox.Show("Автор добавлен");
        }
    }
}
