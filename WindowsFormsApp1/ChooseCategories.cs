using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace WindowsFormsApp1
{
    public partial class ChooseCategories : Form
    {
        public static string[] pageList = new string[] { };
        public static string One = "";
        public static string Two = "";
        public static string Three = "";
        public static string Four = "";
        public static string Five = "";
        public static string Six = "";
        public static string Seven = "";
        public static string Eight = "";
        public string[] massiv = new string[] { };
        public ChooseCategories()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (checkBox1.Checked == true || checkBox2.Checked == true || checkBox3.Checked == true || checkBox4.Checked == true || checkBox5.Checked == true || checkBox6.Checked == true || checkBox7.Checked == true || checkBox8.Checked == true)
            {
                if (checkBox1.Checked == true) { One = checkBox1.Text; massiv = massiv.Concat(new string[] { checkBox1.Text }).ToArray(); }
                if (checkBox2.Checked == true) { Two = checkBox2.Text; massiv = massiv.Concat(new string[] { checkBox2.Text }).ToArray(); }
                if (checkBox3.Checked == true) { Three = checkBox3.Text; massiv = massiv.Concat(new string[] { checkBox3.Text }).ToArray(); }
                if (checkBox4.Checked == true) { Four = checkBox4.Text; massiv = massiv.Concat(new string[] { checkBox4.Text }).ToArray(); }
                if (checkBox5.Checked == true) { Five = checkBox5.Text; massiv = massiv.Concat(new string[] { checkBox5.Text }).ToArray(); }
                if (checkBox6.Checked == true) { Six = checkBox6.Text; massiv = massiv.Concat(new string[] { checkBox6.Text }).ToArray(); }
                if (checkBox7.Checked == true) { Seven = checkBox7.Text; massiv = massiv.Concat(new string[] { checkBox7.Text }).ToArray(); }
                if (checkBox8.Checked == true) { Eight = checkBox8.Text; massiv = massiv.Concat(new string[] { checkBox8.Text }).ToArray(); }

                pageList = pageList.Concat(massiv).ToArray();

                this.Hide();
                CreateDoc CreateDoc = new CreateDoc();
                CreateDoc.FormClosed += (s, args) => this.Close();
                CreateDoc.ShowDialog();
            }
            else
            {
                MessageBox.Show("Отметьте хотя-бы одну категорию!", "Ошибка!", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.Hide();
            Form3 Form3 = new Form3();
            Form3.FormClosed += (s, args) => this.Close();
            Form3.Show();
        }
    }
}
