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
    public partial class login : Form
    {

        SqlConnection con = new SqlConnection("Data Source=(LocalDB)\\MSSQLLocalDB;AttachDbFilename=D:\\Univer\\1Diplom\\WindowsFormsApp1\\WindowsFormsApp1\\WindowsFormsApp1\\WindowsFormsApp1\\Database1.mdf;");


        public static string Login = "";
        public login()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            con.Open();
            SqlCommand cmd = new SqlCommand("select * from Users", con);
            SqlDataAdapter da = new SqlDataAdapter(cmd);
            bool successLogin = false;
            using (SqlDataReader read = cmd.ExecuteReader())
            {
                while (read.Read())
                {
                    if (read["user"].ToString() == user.Text && read["password"].ToString() == pass.Text)
                    {
                        Login = user.Text;
                        successLogin = true;
                        break;
                    }
                }
            }
            if (successLogin == false)
            {
                con.Close();
                MessageBox.Show("Invalid Login please check username and password");
            }
            else
            {
                con.Close();
                MessageBox.Show("Login sucess");
                this.Hide();
                Form3 Form3 = new Form3();
                Form3.FormClosed += (s, args) => this.Close();
                Form3.Show();
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            /* */
            con.Open();
            SqlCommand cmd = new SqlCommand("select * from Users", con);
            bool reg = false;
            using (SqlDataReader read = cmd.ExecuteReader())
            {
                while (read.Read())
                {
                    if (read["user"].ToString() == user.Text)
                    {
                        Login = user.Text;
                        reg = true;
                        break;
                    }
                }
            }
            if (reg == false)
            {
                SqlCommand sqlcmd = new SqlCommand("insert into Users values (@login, @password)", con);
                sqlcmd.Parameters.AddWithValue("@login", user.Text);
                sqlcmd.Parameters.AddWithValue("@password", pass.Text);
                sqlcmd.ExecuteNonQuery();
                con.Close();
                MessageBox.Show("Регистрация успешна");
            }
            else
            {
                con.Close();
                MessageBox.Show("Логин уже существует");
            }
        }
    }
}
