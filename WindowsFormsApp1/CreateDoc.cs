using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace WindowsFormsApp1
{
    public partial class CreateDoc : Form
    {
        public int currentPage = 0;
        public int currentTab = 0;
        public List<TextBox> textBoxes = new List<TextBox>();

        public List<int> countsPages = new List<int>();
        public CreateDoc()
        {
            InitializeComponent();
            dataGridView1.Visible = false;

            button1.Enabled = false;

            lblPageCount.Text = ChooseCategories.pageList.Length.ToString();
            label6.Text = ChooseCategories.pageList[currentPage];
        }

        private bool glas(char bukva)
        {
            bool isGlas = false;
            string glasnie = "ауоыиэяюёе";
            char[] s = glasnie.ToCharArray();
            for (int i = 0; i < glasnie.Length; i++)
            {
                if (s[i] == bukva)
                {
                    isGlas = true;
                }
            }

            return isGlas;
        }

        private string[] rodPadezh(string lastname, string firstname, string middlename)
        {
            bool isMan = true; // Пол. По-умолчанию, мужской
            string[] result =
            {
                lastname,firstname,middlename
            };

            // Если отчество заказнчивается на -а, то это женщина
            if (middlename.EndsWith("а"))
            {
                isMan = false;
            }

            if (isMan)
            // Если это мужчина
            {
                // ФАМИЛИЯ
                // Заканчивается на -й, меняем предыдущую букву и -й на -ого
                if (lastname.EndsWith("й"))
                {
                    lastname = lastname.Substring(0, lastname.Length - 2) + "ого";
                }
                // Заканчивается на согласную - добавляем "а"
                else if (!glas(lastname[lastname.Length - 1]))
                {
                    lastname = lastname + "а";
                }

                // ИМЯ
                if (!glas(firstname[firstname.Length - 1]))
                {
                    firstname = firstname + "а";
                }
                else if ((firstname.EndsWith("й")) || (firstname.EndsWith("ь")))
                {
                    firstname = firstname.Substring(0, firstname.Length - 1) + "я";
                }
                else if (firstname.EndsWith("я"))
                {
                    firstname = firstname.Substring(0, firstname.Length - 1) + "и";
                }
                else if (firstname.EndsWith("ка"))
                {
                    firstname = firstname.Substring(0, firstname.Length - 2) + "ки";
                }
                else if (firstname.EndsWith("а"))
                {
                    firstname = firstname.Substring(0, firstname.Length - 1) + "ы";
                }

                // ОТЧЕСТВО
                middlename = middlename + "а";
            }
            else
            // Если это женщина:
            {
                // ФАМИЛИЯ
                // если заканчивается на -а или -ая, то пишем - ой
                // если заканчивается на другую букву, то фамилия не склоняется
                if (lastname.EndsWith("а"))
                {
                    lastname = lastname.Substring(0, lastname.Length - 1) + "ой";
                }
                else if (lastname.EndsWith("ая"))
                {
                    lastname = lastname.Substring(0, lastname.Length - 2) + "ой";
                }

                // ИМЯ
                if (firstname.EndsWith("а"))
                {
                    firstname = firstname.Substring(0, firstname.Length - 1) + "ы";
                }
                else if (firstname.EndsWith("я"))
                {
                    firstname = firstname.Substring(0, firstname.Length - 1) + "и";
                }

                // ОТЧЕСТВО
                middlename = middlename.Substring(0, middlename.Length - 1) + "ы";
            }

            result[0] = lastname;
            result[1] = firstname;
            result[2] = middlename;

            return result;
        }

        private int textBoxCount(TabPage tp)
        {
            textBoxes.Clear();
            if (currentTab == 0)
            {
                TextBox txtbA = tp.Controls.Find("txtName", true).FirstOrDefault() as TextBox;
                TextBox txtbB = tp.Controls.Find("txtIzdatelstvo", true).FirstOrDefault() as TextBox;
                TextBox txtbC = tp.Controls.Find("txtPages", true).FirstOrDefault() as TextBox;
                TextBox txtbD = tp.Controls.Find("txtSoavtori", true).FirstOrDefault() as TextBox;
                if (txtbA != null) { textBoxes.Add(txtbA); }
                if (txtbB != null) { textBoxes.Add(txtbB); }
                if (txtbC != null) { textBoxes.Add(txtbC); }
                if (txtbD != null) { textBoxes.Add(txtbD); }
            }
            else
            {
                for (int i = 0; i < 30; i++)
                {
                    TextBox txtbA = tp.Controls.Find("txta" + (i + 1), true).FirstOrDefault() as TextBox;
                    TextBox txtbB = tp.Controls.Find("txtb" + (i + 1), true).FirstOrDefault() as TextBox;
                    TextBox txtbC = tp.Controls.Find("txtc" + (i + 1), true).FirstOrDefault() as TextBox;
                    TextBox txtbD = tp.Controls.Find("txtd" + (i + 1), true).FirstOrDefault() as TextBox;
                    if (txtbA != null) { textBoxes.Add(txtbA); }
                    if (txtbB != null) { textBoxes.Add(txtbB); }
                    if (txtbC != null) { textBoxes.Add(txtbC); }
                    if (txtbD != null) { textBoxes.Add(txtbD); }
                }
            }
            return textBoxes.Count;
        }

        public static void EnableTab(TabPage page, bool enable)
        {
            foreach (Control ctl in page.Controls) ctl.Enabled = enable;
        }

        private void txtKolvoRabot_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) /*&& (e.KeyChar != '-')*/)
            {
                e.Handled = true;
            }
        }

        private void txtPeriod_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) && (e.KeyChar != '-'))
            {
                e.Handled = true;
            }
        }

        private void txtKolvoRabot_TextChanged(object sender, EventArgs e)
        {
            if (!String.IsNullOrWhiteSpace(txtKolvoRabot.Text))
            {
                if (Convert.ToInt32(txtKolvoRabot.Text) <= 10 && Convert.ToInt32(txtKolvoRabot.Text) > tabControl1.TabPages.Count)
                {
                    for (int i = tabControl1.TabPages.Count; i < Convert.ToInt32(txtKolvoRabot.Text); i++)
                    {
                        tabControl1.TabPages.Add(new TabPage { Text = "Работа " + (i + 1), Name = "tabPage" + (i + 2) });
                        tabControl1.TabPages[i].BackColor = Color.White;

                        TextBox txt1 = new TextBox();
                        txt1.Name = "txta" + i;
                        txt1.Location = new Point(txtName.Location.X, txtName.Location.Y);
                        txt1.Size = new Size(620, 26);
                        txt1.Anchor = (AnchorStyles.Bottom | AnchorStyles.Left);
                        tabControl1.TabPages[i].Controls.Add(txt1);
                        txt1.BringToFront();

                        TextBox txt2 = new TextBox();
                        txt2.Name = "txtb" + i;
                        txt2.Location = new Point(txtIzdatelstvo.Location.X, txtIzdatelstvo.Location.Y);
                        txt2.Multiline = true;
                        txt2.Size = new Size(350, 70);
                        txt2.Anchor = (AnchorStyles.Bottom | AnchorStyles.Left);
                        tabControl1.TabPages[i].Controls.Add(txt2);
                        txt2.BringToFront();

                        TextBox txt3 = new TextBox();
                        txt3.Name = "txtc" + i;
                        txt3.Location = new Point(txtPages.Location.X, txtPages.Location.Y);
                        txt3.Size = new Size(350, 26);
                        txt3.Anchor = (AnchorStyles.Bottom | AnchorStyles.Left);
                        tabControl1.TabPages[i].Controls.Add(txt3);
                        txt3.BringToFront();

                        TextBox txt4 = new TextBox();
                        txt4.Name = "txtd" + i;
                        txt4.Location = new Point(txtSoavtori.Location.X, txtSoavtori.Location.Y);
                        txt4.Multiline = true;
                        txt4.Size = new Size(350, 80);
                        txt4.Anchor = (AnchorStyles.Bottom | AnchorStyles.Left);
                        tabControl1.TabPages[i].Controls.Add(txt4);
                        txt4.BringToFront();

                        Label lbl1 = new Label();
                        lbl1.Name = "lbla" + i;
                        lbl1.Text = "Наименование научного или методического труда:";
                        lbl1.Location = new Point(label2.Location.X, label2.Location.Y);
                        lbl1.Size = new Size(270, 20);
                        lbl1.Anchor = (AnchorStyles.Bottom | AnchorStyles.Left);
                        tabControl1.TabPages[i].Controls.Add(lbl1);
                        lbl1.BringToFront();

                        Label lbl2 = new Label();
                        lbl2.Name = "lblb" + i;
                        lbl2.Text = "Издательство, журнал (название, номер, год, страницы) или номер авторского свидетельства, патента:";
                        lbl2.Location = new Point(label3.Location.X, label3.Location.Y);
                        lbl2.Size = new Size(540, 20);
                        lbl2.Anchor = (AnchorStyles.Bottom | AnchorStyles.Left);
                        tabControl1.TabPages[i].Controls.Add(lbl2);
                        lbl2.BringToFront();

                        Label lbl3 = new Label();
                        lbl3.Name = "lblc" + i;
                        lbl3.Text = "Кол-во печатных листов или страниц:";
                        lbl3.Location = new Point(label4.Location.X, label4.Location.Y);
                        lbl3.Size = new Size(200, 20);
                        lbl3.Anchor = (AnchorStyles.Bottom | AnchorStyles.Left);
                        tabControl1.TabPages[i].Controls.Add(lbl3);
                        lbl3.BringToFront();

                        Label lbl4 = new Label();
                        lbl4.Name = "lbld" + i;
                        lbl4.Text = "Соавторы:";
                        lbl4.Location = new Point(label5.Location.X, label5.Location.Y);
                        lbl4.Size = new Size(60, 20);
                        lbl4.Anchor = (AnchorStyles.Bottom | AnchorStyles.Left);
                        tabControl1.TabPages[i].Controls.Add(lbl4);
                        lbl4.BringToFront();
                    }
                }
                else if (Convert.ToInt32(txtKolvoRabot.Text) <= 10 && Convert.ToInt32(txtKolvoRabot.Text) > 0 && Convert.ToInt32(txtKolvoRabot.Text) < tabControl1.TabPages.Count)
                {
                    for (int i = tabControl1.TabPages.Count - 1; i > Convert.ToInt32(txtKolvoRabot.Text) - 1; i--)
                    {
                        tabControl1.TabPages.RemoveByKey("tabPage" + (i + 2));
                    }
                }
            }
        }

        bool CheckPage()
        {
            bool isZapolneno = false;
            for (int i = 0; i < textBoxCount(tabControl1.TabPages[currentTab]); i++)
            {
                if (!String.IsNullOrWhiteSpace(textBoxes[i].Text))
                {
                    isZapolneno = true;
                    //textBoxes[i].BackColor = Color.White;
                }
                else
                {
                    isZapolneno = false;
                    //textBoxes[i].BackColor = Color.FromArgb(242, 150, 150);
                    break;
                }
            }
            return isZapolneno;
        }

        private void tabControl1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (tabControl1.TabPages.Count <= currentTab) { currentTab = tabControl1.TabPages.Count - 1; }

            if (tabControl1.SelectedIndex - currentTab == 1)
            {
                if (CheckPage() == true)
                {
                    //MessageBox.Show("Страница перевернута, индекс = " + tabControl1.SelectedIndex);
                    EnableTab(tabControl1.TabPages[currentTab], false);
                    currentTab = tabControl1.SelectedIndex;
                }
                else { tabControl1.SelectedIndex = currentTab; MessageBox.Show("Не все поля заполнены! Переход на следующую вкладку невозможен!", "Ошибка!", MessageBoxButtons.OK, MessageBoxIcon.Error); }
            }
            else if (tabControl1.SelectedIndex > currentTab)
            {
                tabControl1.SelectedIndex = currentTab;
                MessageBox.Show("Открыть можно только следующую вкладку! Страница осталась = " + (currentTab + 1), "Ошибка!", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            GC.Collect();
        }

        public bool EndOfPages = false;
        private void btnRigth_Click(object sender, EventArgs e)
        {
            tabControl1.SelectedIndex = (tabControl1.TabPages.Count - 1);
            if (tabControl1.SelectedIndex == (tabControl1.TabPages.Count - 1) && CheckPage() == true && EndOfPages == false)
            {
                if ((currentPage + 1) <= (ChooseCategories.pageList.Length - 1))
                {
                    //Занесение в ДатаГрид
                    ExportToDataGrid();

                    //Обновление текущей категории (перелистывание страницы)
                    currentPage++;
                    label6.Text = ChooseCategories.pageList[currentPage];

                    txtKolvoRabot.Text = "1";
                    currentTab = 0;
                    EnableTab(tabControl1.TabPages[0], true);
                    for (int i = 0; i < textBoxCount(tabControl1.TabPages[0]); i++)
                    {
                        textBoxes[i].Text = "";
                    }
                }
                else
                {
                    if (EndOfPages == false)
                    {
                        ExportToDataGrid();
                        EndOfPages = true;
                    }
                    txtKolvoRabot.Text = "1";
                    currentTab = 0;
                    EnableTab(tabControl1.TabPages[0], false);
                    txtKolvoRabot.Enabled = false;
                    button1.Enabled = true;
                    for (int i = 0; i < textBoxCount(tabControl1.TabPages[0]); i++)
                    {
                        textBoxes[i].Text = "";
                    }
                    MessageBox.Show("Это последняя страница!", "Информация", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            else
            {
                if (EndOfPages == false)
                {
                    MessageBox.Show("Не все работы заполнены", "Ошибка!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        void ExportToDataGrid()
        {
            //dataGridView1.RowCount += 1;
            int startIndex = dataGridView1.RowCount;
            int curTabPage = 1;
            //dataGridView1.Rows[startIndex].Cells[0].Value = label6.Text;
            for (int i = startIndex; i < tabControl1.TabPages.Count + startIndex; i++)
            {
                dataGridView1.RowCount += 1;
                dataGridView1.Rows[i].Cells[0].Value = (i + 1);
                if (i == startIndex)
                {
                    TextBox tbName = tabControl1.TabPages[0].Controls.Find("txtName", true).FirstOrDefault() as TextBox;
                    dataGridView1.Rows[i].Cells[1].Value = tbName.Text;
                    TextBox tbIzdatelstvo = tabControl1.TabPages[0].Controls.Find("txtIzdatelstvo", true).FirstOrDefault() as TextBox;
                    dataGridView1.Rows[i].Cells[2].Value = tbIzdatelstvo.Text;
                    TextBox tbPages = tabControl1.TabPages[0].Controls.Find("txtPages", true).FirstOrDefault() as TextBox;
                    dataGridView1.Rows[i].Cells[3].Value = tbPages.Text;
                    TextBox tbSoavtori = tabControl1.TabPages[0].Controls.Find("txtSoavtori", true).FirstOrDefault() as TextBox;
                    dataGridView1.Rows[i].Cells[4].Value = tbSoavtori.Text;
                }
                else
                {
                    int usefulPerem = textBoxCount(tabControl1.TabPages[curTabPage]);
                    dataGridView1.Rows[i].Cells[1].Value = textBoxes[0].Text;
                    dataGridView1.Rows[i].Cells[2].Value = textBoxes[1].Text;
                    dataGridView1.Rows[i].Cells[3].Value = textBoxes[2].Text;
                    dataGridView1.Rows[i].Cells[4].Value = textBoxes[3].Text;
                    curTabPage++;
                }
                dataGridView1.Rows[i].Cells[5].Value = label6.Text;
            }
            countsPages.Add(curTabPage);
            GC.Collect();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (!String.IsNullOrWhiteSpace(txtPeriod.Text) && !String.IsNullOrWhiteSpace(txtAuthor.Text) && txtAuthor.Text.Split(' ').Length == 3)
            {
                string[] FIO = rodPadezh(txtAuthor.Text.Split(' ')[0], txtAuthor.Text.Split(' ')[1], txtAuthor.Text.Split(' ')[2]);

                using (SaveFileDialog sfd = new SaveFileDialog() { FileName = "Список трудов " + FIO[0] + " " + FIO[1] + " " + FIO[2] + " " + DateTime.Now.ToShortDateString(), Filter = "DOC|*.doc" })
                {
                    if(sfd.ShowDialog() == DialogResult.OK)
                    {
                        string html = "";
                        html += "<html>";
                        html += @"<head><meta charset=""utf-8""></head>";
                        /*createtable();*/
                        //Table start.
                        html += "<body>";
                        html += @"<p style ='font-family: ""Times New Roman""'><center><font size=""3""><b>СПИСОК</b></font></center></p>";
                        html += @"<p style ='font-family: ""Times New Roman""'><center><font size=""3""><b>научных и научно-методических трудов за период " + txtPeriod.Text + " гг.</b></font></center></p>";
                        html += @"<p style ='font-family: ""Times New Roman""'><center><font size=""3""><b>PhD, доцента кафедры «IT-инжиниринг» </b></font></center></p>";
                        html += @"<p style ='font-family: ""Times New Roman""'><center><font size=""3""><b>НАО «Алматинский университет энергетики и связи имени Гумарбека Даукеева» </b></font></center></p>";
                        html += @"<p style ='font-family: ""Times New Roman""'><center><font size=""3""><b>" + FIO[0] + " " + FIO[1] + " " + FIO[2] + "</b></font></center></p>";
                        html += @"<p style ='font-family: ""Times New Roman""'><center>&nbsp</center></p>";
                        //html += "<table cellpadding='5' cellspacing='0' style='border: 1px solid #ccc;font-size: 9pt;font-family:arial'>";
                        html += @"<center><table border=""1"" cellspacing=""0"" cellpadding=""1"" width=""800""></center>";
                        //Adding HeaderRow.
                        html += "<tr>";

                        for (int i = 0; i < dataGridView1.ColumnCount - 1; i++)
                        {
                            html += @"<th style ='font-family: ""Times New Roman""'><center><font size=""2""><b>" + dataGridView1.Columns[i].HeaderText + "</b></font></center></th>";
                        }
                        html += "</tr>";

                        int startIndex = 0;
                        //Adding DataRow.
                        for (int i = 0; i < Convert.ToInt32(lblPageCount.Text); i++)
                        {
                            html += @"<tr><td colspan=""5"" style ='font-family: ""Times New Roman""'><center><font size=""2""><center>" + ChooseCategories.pageList[i] + "</center></font></center></td></tr>";

                            if (i == 0)
                            {
                                for (int f = 0; f < countsPages[i]; f++)
                                {
                                    html += "<tr>";
                                    for (int j = 0; j < dataGridView1.ColumnCount - 1; j++)
                                    {
                                        html += @"<td style ='font-family: ""Times New Roman""'><center><font size=""2"">" + dataGridView1.Rows[f].Cells[j].Value.ToString() + "</font></center></td>";
                                    }
                                    html += "</tr>";
                                }
                            }
                            else
                            {
                                startIndex += countsPages[i - 1];
                                for (int f = startIndex; f < startIndex + countsPages[i]; f++)
                                {
                                    html += "<tr>";
                                    for (int j = 0; j < dataGridView1.ColumnCount - 1; j++)
                                    {
                                        html += @"<td style ='font-family: ""Times New Roman""'><center><font size=""2"">" + dataGridView1.Rows[f].Cells[j].Value.ToString() + "</font></center></td>";
                                    }
                                    html += "</tr>";
                                }
                            }
                        }

                        //Table end.
                        html += "</center>";
                        html += "</table>";
                        html += "</center>";
                        html += "</body>";
                        html += "</html>";

                        //Save the HTML string as HTML File.
                        string htmlFilePath = @"C:\DataGridView.htm";
                        File.WriteAllText(htmlFilePath, html);
                        object oMissing = System.Reflection.Missing.Value;
                        //Convert the HTML File to Word document.
                        Microsoft.Office.Interop.Word._Application word = new Microsoft.Office.Interop.Word.Application();
                        Microsoft.Office.Interop.Word._Document wordDoc = word.Documents.Open(FileName: htmlFilePath, ReadOnly: false);
                        //wordDoc.SaveAs2(@"D:\temp2\DataGridView.doc");
                        wordDoc.SaveAs(FileName: sfd.FileName, FileFormat: Microsoft.Office.Interop.Word.WdSaveFormat.wdFormatRTF);
                        ((Microsoft.Office.Interop.Word._Document)wordDoc).Close();
                        ((Microsoft.Office.Interop.Word._Application)word).Quit();
                        Process.Start(sfd.FileName);
                        //Application.Exit();
                        //Delete the HTML File.
                        File.Delete(htmlFilePath);

                        //ЧТО ДЕЛАТЬ ПОСЛЕ ВЫВОДА ДОКА

                    }
                }
            }
            else
            {
                MessageBox.Show("поля Период или Автор не заполнены!", "Ошибка!", MessageBoxButtons.OK, MessageBoxIcon.Error);
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
