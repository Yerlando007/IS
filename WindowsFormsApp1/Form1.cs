using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Windows.Forms;
using System.Linq;
using System.Management;
using Microsoft.Office.Interop.Word;
using SautinSoft.Document;
using SautinSoft.Document.Drawing;
using System.Collections;

namespace WindowsFormsApp1
{
    public partial class Form1 : Form
    {

        /// ///////////////////
        List<string> category = new List<string>();
        List<int> countsPages = new List<int>();
        ///////////////////////
        /*public int currentPage = 0;
        public int currentTab = 0;
        public List<TextBox> textBoxes = new List<TextBox>();*/
        //////////////////////
        SqlConnection sqlcon = new SqlConnection("Data Source=(LocalDB)\\MSSQLLocalDB;AttachDbFilename=D:\\Univer\\1Diplom\\WindowsFormsApp1\\WindowsFormsApp1\\WindowsFormsApp1\\WindowsFormsApp1\\Database1.mdf;");
        SqlCommand sqlcmd = new SqlCommand();
        SqlDataAdapter da = new SqlDataAdapter();
        System.Data.DataTable dt = new System.Data.DataTable();
        public string path = "";
        public static string date = "";
        public int selectedIndex = 0;
        public int selectedID = 0;
        public bool IndexIsChanged = false;
        string formatForRedakt = "";
        string nameForRedakt = "";
        string RedaktID = "";

        ManagementEventWatcher stopWatch = new ManagementEventWatcher(new WqlEventQuery("SELECT * FROM Win32_ProcessStopTrace"));
        ManagementEventWatcher startWatch = new ManagementEventWatcher(new WqlEventQuery("SELECT * FROM Win32_ProcessStartTrace"));
        //private static readonly Random random = new Random();
        /*private static double RandomNumberBetween(double minValue, double maxValue)
        {
            var next = random.NextDouble();
            return minValue + (next * (maxValue - minValue));
        }*/
        public Form1()
        {
            InitializeComponent();
            stopWatch.EventArrived += stopWatch_EventArrived;
            startWatch.EventArrived += startWatch_EventArrived;
        }


        //СЛЕЖЕНИЕ ЗА ЗАКРЫТЫМИ ПРИЛОЖЕНИЯМИ НАЧАЛО
        void stopWatch_EventArrived(object sender, EventArrivedEventArgs e)
        {
            if (e.NewEvent.Properties["ProcessName"].Value.ToString() == "WINWORD.EXE" || e.NewEvent.Properties["ProcessName"].Value.ToString() == "EXCEL.EXE" || e.NewEvent.Properties["ProcessName"].Value.ToString() == "notepad.exe")
            {
                if(e.NewEvent.Properties["ProcessID"].Value.ToString() == RedaktID)
                {
                    //MessageBox.Show("ID Процесса = " + RedaktID + ". Оно равно " + e.NewEvent.Properties["ProcessID"].Value.ToString(), "Редактирование зафейлено!", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    byte[] FileBytes = null;
                    try
                    {
                        FileStream FS = new FileStream("c:\\temp\\RedaktFile" + formatForRedakt.Replace(" ", ""), System.IO.FileMode.Open, System.IO.FileAccess.Read);
                        BinaryReader BR = new BinaryReader(FS);
                        long allbytes = new FileInfo("c:\\temp\\RedaktFile" + formatForRedakt.Replace(" ", "")).Length;
                        FileBytes = BR.ReadBytes((Int32)allbytes);

                        FS.Close();
                        FS.Dispose();
                        BR.Close();
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Ошибка при чтении файла. " + ex.ToString());
                    }

                    string connectionString = "Data Source=(LocalDB)\\MSSQLLocalDB;AttachDbFilename=D:\\Univer\\1Diplom\\WindowsFormsApp1\\WindowsFormsApp1\\WindowsFormsApp1\\WindowsFormsApp1\\Database1.mdf;";
                    SqlConnection connection = new SqlConnection(connectionString);
                    SqlCommand sCommand = new SqlCommand();
                    sCommand.Connection = connection;
                    connection.Open();
                    sCommand.CommandType = CommandType.Text;
                    sCommand.CommandText = "UPDATE uploadfile SET fcontent = @FCONTENT WHERE Имя = @NAME";
                    sCommand.Parameters.AddWithValue("@FCONTENT", FileBytes);
                    sCommand.Parameters.AddWithValue("@NAME", nameForRedakt);
                    sCommand.ExecuteNonQuery();
                    connection.Close();

                    StopScanClosedApplications();

                    File.Delete("c:\\temp\\RedaktFile" + formatForRedakt.Replace(" ", ""));

                    MessageBox.Show("Редактирование файла прошло успешно.", "Редактирование завершено!", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
        }

        void startWatch_EventArrived(object sender, EventArrivedEventArgs e)
        {
            if (e.NewEvent.Properties["ProcessName"].Value.ToString() == "WINWORD.EXE" || e.NewEvent.Properties["ProcessName"].Value.ToString() == "EXCEL.EXE" || e.NewEvent.Properties["ProcessName"].Value.ToString() == "notepad.exe")
            {
                RedaktID = e.NewEvent.Properties["ProcessID"].Value.ToString();
                //MessageBox.Show("ID открытого процесса = " + RedaktID);
                startWatch.Stop();
            }
        }

        void ScanClosedApplications()
        {
            stopWatch.Start();
        }
        void StopScanClosedApplications()
        {
            stopWatch.Stop();
        }
        //СЛЕЖЕНИЕ ЗА ЗАКРЫТЫМИ ПРИЛОЖЕНИЯМИ КОНЕЦ

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            if(checkBox1.Checked) checkBox2.Checked = false;
            chart1.Series[0].ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Line;
        }

        private void checkBox2_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox2.Checked) checkBox1.Checked = false;
            chart1.Series[0].ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Column;
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            label4.Text = DateTime.Now.ToLongTimeString();
            label10.Text = DateTime.Now.ToLongDateString();
        }

        private void comboBox1_SelectedIndexChanged_1(object sender, EventArgs e)
        {
            if(chart1.Series.Count > 0)
            {
                chart1.Series[0].Points.Clear();
            }
            chart1.Series.Clear();
            dataGridView1.Columns.Clear();
            sqlcon.Open();
            string selectedState = comboBox1.SelectedItem.ToString();
            string selectedState2 = comboBox2.SelectedItem.ToString();
            if (login.Login == "admin" && comboBox1.Text == selectedState)
            {
                sqlcmd = new SqlCommand("select Имя, Дата, Принадлежность, Страниц, Соавтор, Автор, Категория, Год from uploadfile where Категория=@Категория and Автор=@Автор", sqlcon);
                sqlcmd.Parameters.AddWithValue("@Категория", selectedState);
                sqlcmd.Parameters.AddWithValue("@Автор", selectedState2);
            }
            else if (comboBox1.Text == selectedState)
            {
                sqlcmd = new SqlCommand("select Имя, Дата, Принадлежность, Страниц, Соавтор, Автор, Категория, Год from uploadfile where Аккаунт = @Аккаунт and Категория=@Категория and Автор=@Автор", sqlcon);
                sqlcmd.Parameters.AddWithValue("@Аккаунт", login.Login);
                sqlcmd.Parameters.AddWithValue("@Категория", selectedState);
                sqlcmd.Parameters.AddWithValue("@Автор", selectedState2);
            }
            /*sqlcmd = new SqlCommand("select Имя, Дата from uploadfile where Категория=@Категория", sqlcon);
            sqlcmd.Parameters.AddWithValue("@Категория", comboBox1.Text);*/
            da = new SqlDataAdapter(sqlcmd);
            dt = new System.Data.DataTable();
            da.Fill(dt);
            sqlcon.Close();
            if (dt.Rows.Count > 0)
            {
                DataGridViewLinkColumn links = new DataGridViewLinkColumn();
                links.UseColumnTextForLinkValue = true;
                links.HeaderText = "Скачать";
                links.DataPropertyName = "lnkColumn";
                links.ActiveLinkColor = System.Drawing.Color.White;
                links.LinkBehavior = LinkBehavior.SystemDefault;
                links.LinkColor = System.Drawing.Color.Blue;
                links.Text = "⏬";
                links.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                links.TrackVisitedState = true;
                links.VisitedLinkColor = System.Drawing.Color.YellowGreen;
                dataGridView1.Columns.Add(links);
                dataGridView1.DataSource = dt;
                dataGridView1.AutoResizeColumns();
            }

            //-----
            List<int> years = new List<int>();
            List<int> bookCountForYear = new List<int>();
            sqlcon.Open();
            SqlCommand cmd = new SqlCommand();
            if (login.Login == "admin" && comboBox1.Text == selectedState)
            {
                cmd = new SqlCommand("select Дата from uploadfile where Категория=@Категория", sqlcon);
                cmd.Parameters.AddWithValue("@Категория", selectedState);
            }
            else if (comboBox1.Text == selectedState)
            {
                cmd = new SqlCommand("select Дата from uploadfile where Аккаунт = @Аккаунт and Категория=@Категория", sqlcon);
                cmd.Parameters.AddWithValue("@Аккаунт", login.Login);
                cmd.Parameters.AddWithValue("@Категория", selectedState);
            }
            //SqlCommand cmd = new SqlCommand("select Дата from uploadfile where Аккаунт = '" + Form2.Login + "'" + " group by Дата", sqlcon);
            using (SqlDataReader read = cmd.ExecuteReader())
            {
                while (read.Read())
                {
                    if (!years.Contains(Convert.ToInt32(read["Дата"].ToString().Split(' ')[2])))
                    {
                        years.Add(Convert.ToInt32(read["Дата"].ToString().Split(' ')[2]));
                    }
                }
            }
            sqlcon.Close();
            for (int i = 0; i < years.Count; i++)
            {
                chart1.Series.Add(years[i].ToString());
                chart1.Series[i].Color = System.Drawing.Color.Blue;
            }

            for (int i = 0; i < years.Count; i++)
            {
                bookCountForYear.Add(0);
            }
            for (int i = 0; i < dataGridView1.RowCount; i++)
            {
                for (int j = 0; j < years.Count; j++)
                {
                    try
                    {
                        if (Convert.ToInt32(dataGridView1.Rows[i].Cells[2].Value.ToString().Split(' ')[2]) == years[j])
                        {
                            bookCountForYear[j] = bookCountForYear[j] + 1;
                        }
                    }
                    catch (Exception) { }
                }
            }

            if(chart1.Series.Count > 0)
            {
                chart1.Series[0].ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Column;
                chart1.Series[0].Color = System.Drawing.Color.Blue;
                for (int i = 0; i < years.Count; i++)
                {
                    double x = years[i];
                    double y = bookCountForYear[i];
                    chart1.Series[0].Points.AddXY(x, y);
                }
            }
            //--------
            if (checkBox4.Checked == true)
            {
                button4.Enabled = true;
            }
        }


        void selected()
        {
            if (chart1.Series.Count > 0)
            {
                chart1.Series[0].Points.Clear();
            }
            chart1.Series.Clear();
            dataGridView1.Columns.Clear();
            sqlcon.Open();
            string selectedState = comboBox1.SelectedItem.ToString();
            if (login.Login == "admin" && comboBox1.Text == selectedState)
            {
                sqlcmd = new SqlCommand("select Имя, Дата from uploadfile where Категория=@Категория", sqlcon);
                sqlcmd.Parameters.AddWithValue("@Категория", selectedState);
            }
            else if (comboBox1.Text == selectedState)
            {
                sqlcmd = new SqlCommand("select Имя, Дата from uploadfile where Аккаунт = @Аккаунт and Категория=@Категория", sqlcon);
                sqlcmd.Parameters.AddWithValue("@Аккаунт", login.Login);
                sqlcmd.Parameters.AddWithValue("@Категория", selectedState);
            }
            /*sqlcmd = new SqlCommand("select Имя, Дата from uploadfile where Категория=@Категория", sqlcon);
            sqlcmd.Parameters.AddWithValue("@Категория", comboBox1.Text);*/
            da = new SqlDataAdapter(sqlcmd);
            dt = new System.Data.DataTable();
            da.Fill(dt);
            sqlcon.Close();
            if (dt.Rows.Count > 0)
            {
                DataGridViewLinkColumn links = new DataGridViewLinkColumn();
                links.UseColumnTextForLinkValue = true;
                links.HeaderText = "Скачать";
                links.DataPropertyName = "lnkColumn";
                links.ActiveLinkColor = System.Drawing.Color.White;
                links.LinkBehavior = LinkBehavior.SystemDefault;
                links.LinkColor = System.Drawing.Color.Blue;
                links.Text = "⏬";
                links.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                links.TrackVisitedState = true;
                links.VisitedLinkColor = System.Drawing.Color.YellowGreen;
                dataGridView1.Columns.Add(links);
                dataGridView1.DataSource = dt;
                dataGridView1.AutoResizeColumns();
            }

            //-----
            List<int> years = new List<int>();
            List<int> bookCountForYear = new List<int>();
            sqlcon.Open();
            SqlCommand cmd = new SqlCommand();
            if (login.Login == "admin" && comboBox1.Text == selectedState)
            {
                cmd = new SqlCommand("select Дата from uploadfile where Категория=@Категория", sqlcon);
                cmd.Parameters.AddWithValue("@Категория", selectedState);
            }
            else if (comboBox1.Text == selectedState)
            {
                cmd = new SqlCommand("select Дата from uploadfile where Аккаунт = @Аккаунт and Категория=@Категория", sqlcon);
                cmd.Parameters.AddWithValue("@Аккаунт", login.Login);
                cmd.Parameters.AddWithValue("@Категория", selectedState);
            }
            //SqlCommand cmd = new SqlCommand("select Дата from uploadfile where Аккаунт = '" + Form2.Login + "'" + " group by Дата", sqlcon);
            using (SqlDataReader read = cmd.ExecuteReader())
            {
                while (read.Read())
                {
                    if (!years.Contains(Convert.ToInt32(read["Дата"].ToString().Split(' ')[2])))
                    {
                        years.Add(Convert.ToInt32(read["Дата"].ToString().Split(' ')[2]));
                    }
                }
            }
            sqlcon.Close();
            for (int i = 0; i < years.Count; i++)
            {
                chart1.Series.Add(years[i].ToString());
                chart1.Series[i].Color = System.Drawing.Color.Blue;
            }

            for (int i = 0; i < years.Count; i++)
            {
                bookCountForYear.Add(0);
            }
            for (int i = 0; i < dataGridView1.RowCount; i++)
            {
                for (int j = 0; j < years.Count; j++)
                {
                    try
                    {
                        if (Convert.ToInt32(dataGridView1.Rows[i].Cells[2].Value.ToString().Split(' ')[2]) == years[j])
                        {
                            bookCountForYear[j] = bookCountForYear[j] + 1;
                        }
                    }
                    catch (Exception) { }
                }
            }

            if (chart1.Series.Count > 0)
            {
                chart1.Series[0].ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Column;
                chart1.Series[0].Color = System.Drawing.Color.Blue;
            }
            for (int i = 0; i < years.Count; i++)
            {
                double x = years[i];
                double y = bookCountForYear[i];
                chart1.Series[0].Points.AddXY(x, y);
            }
            //--------


        }

        void LoadGrid()
        {
            dataGridView1.Columns.Clear();
            sqlcon.Open();
            if (login.Login == "admin")
            {
                sqlcmd = new SqlCommand("select Имя, Дата, Принадлежность, Страниц, Соавтор, Автор, Категория from uploadfile", sqlcon);
            }
            else
            {
                sqlcmd = new SqlCommand("select Имя, Дата, Принадлежность, Страниц, Соавтор, Автор, Категория from uploadfile where Аккаунт = '" + login.Login + "'", sqlcon);
            }         
            da = new SqlDataAdapter(sqlcmd);
            dt = new System.Data.DataTable();
            da.Fill(dt);
            sqlcon.Close();
            if (dt.Rows.Count > 0)
            {
                DataGridViewLinkColumn links = new DataGridViewLinkColumn();
                links.UseColumnTextForLinkValue = true;
                links.HeaderText = "Скачать";
                links.DataPropertyName = "lnkColumn";
                links.ActiveLinkColor = System.Drawing.Color.White;
                links.LinkBehavior = LinkBehavior.SystemDefault;
                links.LinkColor = System.Drawing.Color.Blue;
                links.Text = "⏬";
                links.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                links.TrackVisitedState = true;
                links.VisitedLinkColor = System.Drawing.Color.YellowGreen;
                dataGridView1.Columns.Add(links);
                dataGridView1.DataSource = dt;
                dataGridView1.AutoResizeColumns();
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
                button4.Enabled = false;
            if (textBox1.Text == "" || textBox2.Text == "" || textBox3.Text == "" || comboBox1.Text == "" || comboBox1.SelectedItem == null)
            {
                MessageBox.Show("Указаны не все данные");
            }
            else
            {
                OpenFileDialog fDialog = new OpenFileDialog();
                fDialog.Title = "Выберите файл";
                fDialog.Filter = "docx files (*.docx)|*.docx|doc files (*.doc)|*.doc";
                if (fDialog.ShowDialog() == DialogResult.OK)
                {
                    path = fDialog.FileName.ToString();
                    lbl.Text = fDialog.FileName.ToString();
                }

                if (path == "")
                {
                    MessageBox.Show("Файл не выбран!", "Ошибка!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                string filetype = path.Substring(Convert.ToInt32(path.LastIndexOf(".")), path.Length - (Convert.ToInt32(path.LastIndexOf("."))));
                string filename = path.Substring(Convert.ToInt32(path.LastIndexOf("\\")) + 1, path.Length - (Convert.ToInt32(path.LastIndexOf("\\")) + 1));

                /*if (filetype != ".docx" && filetype != ".xlsx" && filetype != ".pdf" && filetype != ".png" && filetype != ".jpg" && filetype != ".jpeg")
                {
                    MessageBox.Show("Upload Only Documents and Images");
                    return;
                }*/

                byte[] FileBytes = null;

                try
                {
                    // Open file to read using file path
                    FileStream FS = new FileStream(path, System.IO.FileMode.Open, System.IO.FileAccess.Read);

                    // Add filestream to binary reader
                    BinaryReader BR = new BinaryReader(FS);

                    // get total byte length of the file
                    long allbytes = new FileInfo(path).Length;

                    // read entire file into buffer
                    FileBytes = BR.ReadBytes((Int32)allbytes);
              
                    // close all instances
                    FS.Close();
                    FS.Dispose();
                    BR.Close();
                  
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error during File Read " + ex.ToString());
                }

                if (filetype == ".docx")
                {
                    // Update and count the number of words and pages in the file.

                    DocumentCore dc = DocumentCore.Load(path);

                    dc.CalculateStats();


                    // Show statistics.
                    //Console.WriteLine("Pages: {0}", dc.Document.Properties.BuiltIn[BuiltInDocumentProperty.Pages]);
                    //textBox3.Text = dc.Document.Properties.BuiltIn[BuiltInDocumentProperty.Pages];

                    if (nf.Text == "")
                    {
                        InsertFile(filename, FileBytes, label10.Text, login.Login, comboBox1.Text, textBox1.Text, dc.Document.Properties.BuiltIn[BuiltInDocumentProperty.Pages], textBox2.Text, comboBox2.Text, textBox3.Text);
                        /* LoadGrid();*/
                        selected();
                    }
                    else
                    {
                        InsertFile(nf.Text + filetype, FileBytes, label10.Text, login.Login, comboBox1.Text, textBox1.Text, dc.Document.Properties.BuiltIn[BuiltInDocumentProperty.Pages], textBox2.Text, comboBox2.Text, textBox3.Text);
                        /* LoadGrid();*/
                        selected();
                    }
                }
                else if (filetype == ".doc")
                {
                    ////////////////////



                    // Convert DOC file to DOCX file.
                    SautinSoft.UseOffice u = new SautinSoft.UseOffice();

                    string inpFile = Path.GetFullPath(path);
                    string outFile = Path.GetFullPath("D:\\temp\\temporary.docx");

                    // Prepare UseOffice .Net, loads MS Word in memory
                    int ret = u.InitWord();

                    // Return values:
                    // 0 - Loading successfully
                    // 1 - Can't load MS Word library in memory

                    if (ret == 1)
                    {
                        Console.WriteLine("Error! Can't load MS Word library in memory");
                        return;
                    }

                    // Perform the conversion.
                    ret = u.ConvertFile(inpFile, outFile, SautinSoft.UseOffice.eDirection.DOC_to_DOCX);

                    // Release MS Word from memory
                    u.CloseWord();



                    // 0 - Converting successfully
                    // 1 - Can't open input file. Check that you are using full local path to input file, URL and relative path are not supported
                    // 2 - Can't create output file. Please check that you have permissions to write by this path or probably this path already used by another application
                    // 3 - Converting failed, please contact with our Support Team
                    // 4 - MS Office isn't installed. The component requires that any of these versions of MS Office should be installed: 2000, XP, 2003, 2007, 2010, 2013, 2016 or 2019.
                    /*if (ret == 0)
                    {
                        // Open the result.
                        System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(outFile) { UseShellExecute = true });
                    }
                    else
                        Console.WriteLine("Error! Please contact with SautinSoft support: support@sautinsoft.com.");*/




                    ////////////////////


                    // Update and count the number of words and pages in the file.

                    DocumentCore dc = DocumentCore.Load("D:\\temp\\temporary.docx");

                    dc.CalculateStats();


                    File.Delete("D:\\temp\\temporary.docx");

                    // Show statistics.
                    //Console.WriteLine("Pages: {0}", dc.Document.Properties.BuiltIn[BuiltInDocumentProperty.Pages]);
                    //textBox3.Text = dc.Document.Properties.BuiltIn[BuiltInDocumentProperty.Pages];

                    if (nf.Text == "")
                    {
                        InsertFile(filename, FileBytes, label10.Text, login.Login, comboBox1.Text, textBox1.Text, dc.Document.Properties.BuiltIn[BuiltInDocumentProperty.Pages], textBox2.Text, comboBox2.Text, textBox3.Text);
                        /* LoadGrid();*/
                        selected();
                    }
                    else
                    {
                        InsertFile(nf.Text + filetype, FileBytes, label10.Text, login.Login, comboBox1.Text, textBox1.Text, dc.Document.Properties.BuiltIn[BuiltInDocumentProperty.Pages], textBox2.Text, comboBox2.Text, textBox3.Text);
                        /* LoadGrid();*/
                        selected();
                    }
                }
                /*this.Hide();
                Form1 Form1 = new Form1();
                Form1.FormClosed += (s, args) => this.Close();
                Form1.Show();*/
            }
        comboBox2.Enabled = true;
        comboBox1.Enabled = false;
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.ColumnIndex == 0)
            {
                //string id;
                string name;
                FileStream FS = null;
                byte[] dbbyte;

                name = dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex+1].Value.ToString();
                //id = Convert.ToString(e.RowIndex + 1);

                sqlcon.Open();
                sqlcmd = new SqlCommand("select * from uploadfile where Имя = @name ", sqlcon);
                sqlcmd.Parameters.AddWithValue("@name", name);
                da = new SqlDataAdapter(sqlcmd);
                dt = new System.Data.DataTable();
                da.Fill(dt);
                sqlcon.Close();

                if (dt.Rows.Count > 0 && dataGridView1.Rows[e.RowIndex].Cells[1].Value.ToString().Contains("."))
                {
                    if(Properties.Settings.Default.DownloadPath == "")
                    {
                        MessageBox.Show("Укажите путь для постоянного хранения файлов в следующем диалоговом окне...", "Путь к файлам.", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        ChangeDownloadPath();
                    }
                    else
                    {
                        try
                        {
                            string filetype = dataGridView1.Rows[e.RowIndex].Cells[1].Value.ToString().Substring(Convert.ToInt32(dataGridView1.Rows[e.RowIndex].Cells[1].Value.ToString().LastIndexOf(".")), dataGridView1.Rows[e.RowIndex].Cells[1].Value.ToString().Length - (Convert.ToInt32(dataGridView1.Rows[e.RowIndex].Cells[1].Value.ToString().LastIndexOf("."))));
                            string filename = dataGridView1.Rows[e.RowIndex].Cells[1].Value.ToString().Substring(Convert.ToInt32(dataGridView1.Rows[e.RowIndex].Cells[1].Value.ToString().LastIndexOf("\\")) + 1, dataGridView1.Rows[e.RowIndex].Cells[1].Value.ToString().Length - (Convert.ToInt32(dataGridView1.Rows[e.RowIndex].Cells[1].Value.ToString().LastIndexOf("\\")) + 1));
                            dbbyte = (byte[])dt.Rows[0]["fcontent"];
                            string filepath = Path.Combine(Properties.Settings.Default.DownloadPath, filename);

                            FS = new FileStream(filepath, FileMode.Create);
                            FS.Write(dbbyte, 0, dbbyte.Length);
                            FS.Close();

                            Process Proc = new Process();
                            Proc.StartInfo.FileName = filepath;
                            Proc.Start();
                        }
                        catch(Exception)
                        {
                            MessageBox.Show("Укажите путь для постоянного хранения файлов в следующем диалоговом окне...", "Путь к файлам.", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            ChangeDownloadPath();
                        }
                    }
                }
                sqlcon.Close();
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            ChangeDownloadPath();
        }

        void ChangeDownloadPath()
        {
            using (var dialog = new FolderBrowserDialog())
            {
                if (dialog.ShowDialog() == DialogResult.OK)
                {
                    Properties.Settings.Default.DownloadPath = dialog.SelectedPath;
                    Properties.Settings.Default.Save();
                }
            }
        }

        void InsertFile(object FN, object FB, object FD, object FA, object FK, object FP, object FS, object FC, object FO, object FG)
        {
            string sql = "SELECT * FROM uploadfile";
            SqlCommand sCom = new SqlCommand(sql, sqlcon);
            int max = 0;
            sqlcon.Open();
            using (SqlDataReader read = sCom.ExecuteReader())
            {
                while (read.Read())
                {
                    if (int.Parse(read["Id"].ToString()) > max)
                    {
                        max = int.Parse(read["Id"].ToString());
                    }
                }
            }
            max++;
            sqlcon.Close();

            /*if (textBox1.Text == "" && textBox2.Text == "" && textBox3.Text == "")
            {

            }*/
            sqlcon.Open();
            SqlCommand sqlcmd = new SqlCommand("insert into uploadfile(Id, Имя, fcontent, Дата, Аккаунт, Категория, Принадлежность, Страниц, Соавтор, Автор, Год) values (@Id, @FN, @FB, @FD, @FA, @FK, @FP, @FS, @FC, @FO, @FG)", sqlcon);
            sqlcmd.Parameters.AddWithValue("@Id", max);
            sqlcmd.Parameters.AddWithValue("@FN", FN);
            sqlcmd.Parameters.AddWithValue("@FB", FB);
            sqlcmd.Parameters.AddWithValue("@FD", FD);
            sqlcmd.Parameters.AddWithValue("@FA", FA);
            sqlcmd.Parameters.AddWithValue("@FK", FK);
            sqlcmd.Parameters.AddWithValue("@FP", FP);
            sqlcmd.Parameters.AddWithValue("@FS", FS);
            sqlcmd.Parameters.AddWithValue("@FC", FC);
            sqlcmd.Parameters.AddWithValue("@FO", FO);
            sqlcmd.Parameters.AddWithValue("@FG", FG);
            sqlcmd.ExecuteNonQuery();
            sqlcon.Close();
        }

        private void timer2_Tick(object sender, EventArgs e)
        {
            GC.Collect();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            comboBox1.Items.Clear();
            comboBox1.Items.Add("Учебные пособия, рекомендованные Республиканским учебно-методическим советом Министерства образования и науки Республики Казахстан");
            comboBox1.Items.Add("Публикации в материалах конференций, индексируемых в базах Web of Science, Scopus");
            comboBox1.Items.Add("Публикация в изданиях, имеющих ненулевой импакт-фактор в базе-данных информационной компании Томсон Рейтер (WebofScience, ThomsonRuters)");
            comboBox1.Items.Add("Публикация в изданиях, рекомендуемых ККСОН МОН РК (кроме материалов конференций");
            comboBox1.Items.Add("Публикация в базе данных Scopus, Pubmed, zbMath, MathScinet, Agris, Georef, Astrophysical journal");
            comboBox1.Items.Add("Учебники, рекомендованные Министерством образования и науки Республики Казахстан");
            comboBox1.Items.Add("Монография");
            comboBox1.Items.Add("Учебники и учебные пособия, рекомендованные ученым советом вуза");


            ToolTip t1 = new ToolTip();
            t1.SetToolTip(button2, "Загрузить файл в БД");


            label2.Text = "Добро пожаловать: " + login.Login;

            LoadGrid();
            timer1.Start();
            if (chart1.Series.Count > 0)
            {
                chart1.Series[0].Points.Clear();
            }
            chart1.Series.Clear();

            //-----
            List<int> years = new List<int>();
            List<int> bookCountForYear = new List<int>();
            sqlcon.Open();
            /*string selectedState = comboBox1.SelectedItem.ToString();*/
            /*SqlCommand cmd = new SqlCommand();*/
            /*if (Form2.Login == "admin" && comboBox1.Text == selectedState)
            {
                cmd = new SqlCommand("select Дата from uploadfile where Категория=@Категория", sqlcon);
                cmd.Parameters.AddWithValue("@Категория", selectedState);
            }
            else if (comboBox1.Text == selectedState)
            {
                cmd = new SqlCommand("select Дата from uploadfile where Аккаунт = @Аккаунт and Категория=@Категория", sqlcon);
                cmd.Parameters.AddWithValue("@Аккаунт", Form2.Login);
                cmd.Parameters.AddWithValue("@Категория", selectedState);
            }*/
            SqlCommand cmd = new SqlCommand("select Дата from uploadfile where Аккаунт = '" + login.Login + "'" + " group by Дата", sqlcon);
            using (SqlDataReader read = cmd.ExecuteReader())
            {
                while (read.Read())
                {
                    if (!years.Contains(Convert.ToInt32(read["Дата"].ToString().Split(' ')[2])))
                    {
                        years.Add(Convert.ToInt32(read["Дата"].ToString().Split(' ')[2]));
                    }
                }
            }
            sqlcon.Close();
            for (int i = 0; i < years.Count; i++)
            {
                chart1.Series.Add(years[i].ToString());
                chart1.Series[i].Color = System.Drawing.Color.Blue;
            }

            for (int i = 0; i < years.Count; i++)
            {
                bookCountForYear.Add(0);
            }
            for (int i = 0; i < dataGridView1.RowCount; i++)
            {
                for (int j = 0; j < years.Count; j++)
                {
                    try
                    {
                        if (Convert.ToInt32(dataGridView1.Rows[i].Cells[2].Value.ToString().Split(' ')[2]) == years[j])
                        {
                            bookCountForYear[j] = bookCountForYear[j] + 1;
                        }
                    }
                    catch (Exception) { }
                }
            }

            if (chart1.Series.Count > 0)
            { 
                chart1.Series[0].ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Column;
                chart1.Series[0].Color = System.Drawing.Color.Blue;
            }
            for (int i = 0; i < years.Count; i++)
            {
                double x = years[i];
                double y = bookCountForYear[i];
                chart1.Series[0].Points.AddXY(x, y);
            }
            //--------
            comboBox2.Items.Clear();
            /*SqlConnection sqlcon = new SqlConnection("Data Source=(LocalDB)\\MSSQLLocalDB;AttachDbFilename=D:\\Univer\\1Diplom\\WindowsFormsApp1\\WindowsFormsApp1\\WindowsFormsApp1\\WindowsFormsApp1\\Database1.mdf;");
        SqlCommand sqlcmd = new SqlCommand();
        SqlDataAdapter da = new SqlDataAdapter();
        System.Data.DataTable dt = new System.Data.DataTable();*/
            sqlcon.Open();
            sqlcmd = new SqlCommand("select authors from author", sqlcon);
            da = new SqlDataAdapter(sqlcmd);
            dt = new System.Data.DataTable();
            da.Fill(dt);
            foreach (DataRow dr in dt.Rows)
            {
                comboBox2.Items.Add(dr["authors"].ToString());
            }
            sqlcon.Close();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Form1_Load(this, EventArgs.Empty);
            comboBox2.Enabled = true;
            comboBox1.Enabled = false;
            button4.Enabled = false;
            //RefreshForm();
        }

        private void RefreshForm()
        {
            int memoryWidth = this.Size.Width;
            int memoryHeight = this.Size.Height;
            bool windowMaximized = false;
            if (this.WindowState == FormWindowState.Maximized)
            {
                windowMaximized = true;
            }
            this.WindowState = FormWindowState.Minimized;
            this.Show();
            if (windowMaximized == true)
            {
                this.WindowState = FormWindowState.Maximized;
            }
            else
            {
                this.WindowState = FormWindowState.Normal;
                this.Size = new System.Drawing.Size(memoryWidth, memoryHeight);
            }
        }


        private void checkBox3_CheckedChanged(object sender, EventArgs e)
        {
           /* if (checkBox3.Checked == true)
            {
                sqlcon.Open();
                if (Form2.Login == "admin" && checkBox3.Checked == true)
                {
                    sqlcmd = new SqlCommand("select Имя, Дата, Категория from uploadfile where Категория=@Категория", sqlcon);
                    sqlcmd.Parameters.AddWithValue("@Категория", y1);
                }
                else if (checkBox3.Checked == true)
                {
                    sqlcmd = new SqlCommand("select Имя, Дата, Категория from uploadfile where Аккаунт = @Аккаунт and Категория=@Категория", sqlcon);
                    sqlcmd.Parameters.AddWithValue("@Аккаунт", Form2.Login);
                    sqlcmd.Parameters.AddWithValue("@Категория", y1);
                }
                da = new SqlDataAdapter(sqlcmd);
                dt = new DataTable();
                da.Fill(dt);
                sqlcon.Close();
                if (dt.Rows.Count > 0)
                {
                    DataGridViewLinkColumn links = new DataGridViewLinkColumn();
                    links.UseColumnTextForLinkValue = true;
                    links.HeaderText = "Скачать";
                    links.DataPropertyName = "lnkColumn";
                    links.ActiveLinkColor = Color.White;
                    links.LinkBehavior = LinkBehavior.SystemDefault;
                    links.LinkColor = Color.Blue;
                    links.Text = "Нажать сюда";
                    links.TrackVisitedState = true;
                    links.VisitedLinkColor = Color.YellowGreen;
                    dataGridView1.Columns.Add(links);
                    dataGridView1.DataSource = dt;
                    dataGridView1.AutoResizeColumns();
                }
            }
            else
            {
                List<int> indexes = new List<int>();
                for (int i = 0; i < dataGridView1.RowCount; i++)
                {
                    if (dataGridView1.Rows[i].Cells["Категория"].Value.ToString().Contains(y1))
                    {
                        indexes.Add(i);
                    }
                }
                
                for(int i = indexes.Count - 1; i >= 0; i--)
                {
                    dataGridView1.Rows.RemoveAt(indexes[i]);
                }
            }*/
        }

        private void checkBox4_CheckedChanged(object sender, EventArgs e)
        {
           /* if (checkBox4.Checked == true)
            {
                sqlcon.Open();
                if (Form2.Login == "admin" && checkBox4.Checked == true)
                {
                    sqlcmd = new SqlCommand("select Имя, Дата, Категория from uploadfile where Категория=@Категория", sqlcon);
                    sqlcmd.Parameters.AddWithValue("@Категория", y2);
                }
                else if (checkBox4.Checked == true)
                {
                    sqlcmd = new SqlCommand("select Имя, Дата, Категория from uploadfile where Аккаунт = @Аккаунт and Категория=@Категория", sqlcon);
                    sqlcmd.Parameters.AddWithValue("@Аккаунт", Form2.Login);
                    sqlcmd.Parameters.AddWithValue("@Категория", y2);
                }
                da = new SqlDataAdapter(sqlcmd);
                dt = new DataTable();
                da.Fill(dt);
                sqlcon.Close();
                if (dt.Rows.Count > 0)
                {
                    DataGridViewLinkColumn links = new DataGridViewLinkColumn();
                    links.UseColumnTextForLinkValue = true;
                    links.HeaderText = "Скачать";
                    links.DataPropertyName = "lnkColumn";
                    links.ActiveLinkColor = Color.White;
                    links.LinkBehavior = LinkBehavior.SystemDefault;
                    links.LinkColor = Color.Blue;
                    links.Text = "Нажать сюда";
                    links.TrackVisitedState = true;
                    links.VisitedLinkColor = Color.YellowGreen;
                    dataGridView1.Columns.Add(links);
                    dataGridView1.DataSource = dt;
                    dataGridView1.AutoResizeColumns();
                }
                List<int> indexes = new List<int>();
                for (int i = 0; i < dataGridView1.RowCount; i++)
                {
                    if (dataGridView1.Rows[i].Cells["Категория"].Value.ToString().Contains(y2))
                    {
                        indexes.Add(i);
                    }
                }

                for (int i = 0; i < dataGridView1.RowCount; i++)
                {
                    dataGridView1.Rows.Add(indexes[i]); //(indexes[i]);
                }
            }
            else
            {
                List<int> indexes = new List<int>();
                for (int i = 0; i < dataGridView1.RowCount; i++)
                {
                    if (dataGridView1.Rows[i].Cells["Категория"].Value.ToString().Contains(y2))
                    {
                        indexes.Add(i);
                    }
                }

                for (int i = indexes.Count - 1; i >= 0; i--)
                {
                    dataGridView1.Rows.RemoveAt(indexes[i]);
                }
            }*/
        }

        private void button4_Click(object sender, EventArgs e)
        {          
            if (comboBox1.Text == "")
            {
                MessageBox.Show("Выберите категорию");
            }
            else
            {
                //for (int i = indexes.Count - 1; i >= 0; i--)
                if (dataGridView2.Rows.Count > 0)
                {
                    for (int i = 0; i < dataGridView1.Rows.Count; i++)
                    {
                        dataGridView2.Rows.Add();
                        int l = 1;
                        foreach (DataGridViewRow row in dataGridView2.Rows)
                        {
                            row.Cells[0].Value = l;
                            l++;
                        }
                        //dataGridView2.Rows[dataGridView2.Rows.Count - 1].Cells[0].Value = dataGridView1.Rows[i].Cells[0].Value.ToString();
                        dataGridView2.Rows[dataGridView2.Rows.Count - 1].Cells[1].Value = dataGridView1.Rows[i].Cells[1].Value.ToString();
                        dataGridView2.Rows[dataGridView2.Rows.Count - 1].Cells[2].Value = dataGridView1.Rows[i].Cells[3].Value.ToString();
                        dataGridView2.Rows[dataGridView2.Rows.Count - 1].Cells[3].Value = dataGridView1.Rows[i].Cells[4].Value.ToString();
                        dataGridView2.Rows[dataGridView2.Rows.Count - 1].Cells[4].Value = dataGridView1.Rows[i].Cells[5].Value.ToString();
                        dataGridView2.Rows[dataGridView2.Rows.Count - 1].Cells[5].Value = dataGridView1.Rows[i].Cells[7].Value.ToString();
                        dataGridView2.Rows[dataGridView2.Rows.Count - 1].Cells[6].Value = dataGridView1.Rows[i].Cells[8].Value.ToString();
                    }
                }
                else
                {
                    for (int i = 0; i < dataGridView1.Rows.Count; i++)
                    {
                        dataGridView2.Rows.Add();
                        int l = 1;
                        foreach (DataGridViewRow row in dataGridView2.Rows)
                        {
                            row.Cells[0].Value = l;
                            l++;
                        }
                        //dataGridView2.Rows[i].Cells[0].Value = dataGridView1.Rows[i].Cells[0].Value.ToString();
                        dataGridView2.Rows[i].Cells[1].Value = dataGridView1.Rows[i].Cells[1].Value.ToString();
                        dataGridView2.Rows[i].Cells[2].Value = dataGridView1.Rows[i].Cells[3].Value.ToString();
                        dataGridView2.Rows[i].Cells[3].Value = dataGridView1.Rows[i].Cells[4].Value.ToString();
                        dataGridView2.Rows[i].Cells[4].Value = dataGridView1.Rows[i].Cells[5].Value.ToString();
                        dataGridView2.Rows[i].Cells[5].Value = dataGridView1.Rows[i].Cells[7].Value.ToString();
                        dataGridView2.Rows[i].Cells[6].Value = dataGridView1.Rows[i].Cells[8].Value.ToString();
                    }
                }
                category.Add(Convert.ToString(comboBox1.SelectedItem));
                countsPages.Add(dataGridView1.RowCount);
                comboBox1.Items.Remove(comboBox1.SelectedItem);
                int num1 = int.Parse(lblPageCount.Text);
                int sum = num1 + 1;
                lblPageCount.Text = sum.ToString();
                vivod.Enabled = true;             
            }
        }

        private void button5_Click(object sender, EventArgs e)
        {
            DialogResult dialog = MessageBox.Show("Вы действительно хотите удалить выбранную запись?", "?", MessageBoxButtons.YesNo, MessageBoxIcon.Information);
            if(dialog == DialogResult.Yes)
            {
                if (IndexIsChanged == true)
                {
                    dataGridView2.Rows.RemoveAt(selectedIndex);
                }
                else
                {
                    MessageBox.Show("Вы ничего не выделили!", "Ошибка!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private void dataGridView2_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                selectedIndex = e.RowIndex;
                IndexIsChanged = true;
            }
            catch (Exception) { selectedIndex = 0; IndexIsChanged = true; }
        }

        private void button6_Click(object sender, EventArgs e)
        {
            if (dataGridView1.SelectedRows[0].Cells[0].Value == null)
            {
                return;
            }

            byte[] data = GetDataFromBase();
            if (toggle == true)
            {
                string fileformat = dataGridView1.SelectedRows[0].Cells[1].Value.ToString().Substring(Convert.ToInt32(dataGridView1.SelectedRows[0].Cells[1].Value.ToString().LastIndexOf(".")), dataGridView1.SelectedRows[0].Cells[1].Value.ToString().Length - (Convert.ToInt32(dataGridView1.SelectedRows[0].Cells[1].Value.ToString().LastIndexOf("."))));
                formatForRedakt = fileformat.Replace(" ", "");
                File.WriteAllBytes("c:\\temp\\RedaktFile" + fileformat.Replace(" ", ""), data);

                startWatch.Start();

                Process Proc = new Process();
                Proc.StartInfo.FileName = "c:\\temp\\RedaktFile" + fileformat.Replace(" ", "");
                Proc.Start();

                ScanClosedApplications();
            }
            else { MessageBox.Show("Произошла критическая ошибка!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error); }
            vivod.Enabled = true;
        }

        //ФУНКЦИЯ ВЫНЕСЕНИЯ ФАЙЛА ИЗ БД НАЧАЛО
        bool toggle = true;
        private static readonly string CONNECT_STRING = "Data Source = (LocalDB)\\MSSQLLocalDB;AttachDbFilename=D:\\Univer\\1Diplom\\WindowsFormsApp1\\WindowsFormsApp1\\WindowsFormsApp1\\WindowsFormsApp1\\Database1.mdf;";
        public byte[] GetDataFromBase()
        {
            string selectCMD = "SELECT fcontent FROM uploadfile WHERE Имя = '" + nameForRedakt + "'";
            
            using (SqlConnection cnn = new SqlConnection(CONNECT_STRING))
            {
                cnn.Open();
                using (SqlCommand cmd = new SqlCommand(selectCMD, cnn))
                {
                    cmd.CommandType = CommandType.Text;
                    cmd.CommandText = "SELECT fcontent FROM uploadfile WHERE Имя = @NAME";
                    cmd.Parameters.AddWithValue("@NAME", nameForRedakt);
                    using (SqlDataReader dr = cmd.ExecuteReader())
                    {
                        try
                        {
                            dr.Read();
                            byte[] data = (byte[])dr["fcontent"];
                            toggle = true;
                            return data;
                        }
                        catch (Exception)
                        {
                            toggle = false;
                            return null;
                        }
                    }
                }
            }
        }
        //ФУНКЦИЯ ВЫНЕСЕНИЯ ФАЙЛА ИЗ БД КОНЕЦ

        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            nameForRedakt = dataGridView1.Rows[e.RowIndex].Cells[1].Value.ToString();
            try
            {
                selectedID = e.RowIndex;
                IndexIsChanged = true;
            }
            catch (Exception) { selectedID = 0; IndexIsChanged = true; }
        }

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            comboBox1.Enabled = true;
            comboBox2.Enabled = false;
        }

        private void button7_Click(object sender, EventArgs e)
        {
            this.Hide();
            Form3 Form3 = new Form3();
            Form3.FormClosed += (s, args) => this.Close();
            Form3.Show();
        }

        private void manual_Click(object sender, EventArgs e)
        {
            this.Hide();
            ChooseCategories ChooseCategories = new ChooseCategories();
            ChooseCategories.FormClosed += (s, args) => this.Close();
            ChooseCategories.Show();
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


        private void vivod_Click(object sender, EventArgs e)
        {
            if (textBox6.Text == "")
            {
                MessageBox.Show("Заполните должность");
            }
            else
            {
                string[] FIO = rodPadezh(comboBox2.Text.Split(' ')[0], comboBox2.Text.Split(' ')[1], comboBox2.Text.Split(' ')[2]);

                var MaxID = dataGridView2.Rows.Cast<DataGridViewRow>().Max(r => Convert.ToInt32(r.Cells[6].Value));
                var MinID = dataGridView2.Rows.Cast<DataGridViewRow>().Min(r => Convert.ToInt32(r.Cells[6].Value));


                using (SaveFileDialog sfd = new SaveFileDialog() { FileName = "Список трудов " + FIO[0] + " " + FIO[1] + " " + FIO[2] + " " + DateTime.Now.ToShortDateString(), Filter = "DOC|*.doc" })
                {
                    if (sfd.ShowDialog() == DialogResult.OK)
                    {
                        string html = "";
                        html += "<html>";
                        html += @"<head><meta charset=""utf-8""></head>";
                        //createtable();
                        //Table start.
                        html += "<body>";
                        html += @"<p style ='font-family: ""Times New Roman""'><center><font size=""3""><b>СПИСОК</b></font></center></p>";
                        html += @"<p style ='font-family: ""Times New Roman""'><center><font size=""3""><b>научных и научно-методических трудов за период " + MinID + " - " + MaxID + " гг.</b></font></center></p>";
                        html += @"<p style ='font-family: ""Times New Roman""'><center><font size=""3""><b>" + textBox6.Text + "</b></font></center></p>";
                        html += @"<p style ='font-family: ""Times New Roman""'><center><font size=""3""><b>НАО «Алматинский университет энергетики и связи имени Гумарбека Даукеева» </b></font></center></p>";
                        html += @"<p style ='font-family: ""Times New Roman""'><center><font size=""3""><b>" + FIO[0] + " " + FIO[1] + " " + FIO[2] + "</b></font></center></p>";
                        html += @"<p style ='font-family: ""Times New Roman""'><center>&nbsp</center></p>";
                        //html += "<table cellpadding='5' cellspacing='0' style='border: 1px solid #ccc;font-size: 9pt;font-family:arial'>";
                        html += @"<center><table border=""1"" cellspacing=""0"" cellpadding=""1"" width=""800""></center>";
                        //Adding HeaderRow.
                        html += "<tr>";

                        for (int i = 0; i < dataGridView2.ColumnCount - 2; i++)
                        {
                            html += @"<th style ='font-family: ""Times New Roman""'><center><font size=""2""><b>" + dataGridView2.Columns[i].HeaderText + "</b></font></center></th>";
                        }
                        html += "</tr>";

                        int startIndex = 0;
                        //Adding DataRow.
                        for (int i = 0; i < Convert.ToInt32(lblPageCount.Text); i++)
                        {
                            html += @"<tr><td colspan=""5"" style ='font-family: ""Times New Roman""'><center><font size=""2""><center><b>" + category[i] + "</b></center></font></center></td></tr>";

                            if (i == 0)
                            {
                                for (int f = 0; f < countsPages[i]; f++)
                                {
                                    html += "<tr>";
                                    for (int j = 0; j < dataGridView2.ColumnCount - 2; j++)
                                    {
                                        html += @"<td style ='font-family: ""Times New Roman""'><center><font size=""2"">" + dataGridView2.Rows[f].Cells[j].Value.ToString() + "</font></center></td>";
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
                                    for (int j = 0; j < dataGridView2.ColumnCount - 2; j++)
                                    {
                                        html += @"<td style ='font-family: ""Times New Roman""'><center><font size=""2"">" + dataGridView2.Rows[f].Cells[j].Value.ToString() + "</font></center></td>";
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
           
        }

        private void button8_Click(object sender, EventArgs e)
        {
            var MaxID = dataGridView2.Rows.Cast<DataGridViewRow>().Max(r => Convert.ToInt32(r.Cells[6].Value));
            var MinID = dataGridView2.Rows.Cast<DataGridViewRow>().Min(r => Convert.ToInt32(r.Cells[6].Value));
        }

        private void button8_Click_1(object sender, EventArgs e)
        {
            this.Hide();
            author author = new author();
            author.FormClosed += (s, args) => this.Close();
            author.Show();
        }

        private void checkBox3_CheckedChanged_1(object sender, EventArgs e)
        {
            if (checkBox3.Checked) checkBox4.Checked = false;
            nf.Enabled = true;
            button2.Enabled = true;
            button3.Enabled = true;
            textBox1.Enabled = true;
            textBox3.Enabled = true;
            textBox2.Enabled = true;
            button3.Enabled = true;
            dataGridView2.Enabled = false;
            button4.Enabled = false;
            vivod.Enabled = false;
            textBox6.Enabled =false;
        }

        private void checkBox4_CheckedChanged_1(object sender, EventArgs e)
        {
            if (checkBox4.Checked) checkBox3.Checked = false;
            dataGridView2.Enabled = true;
            button4.Enabled = true;
            vivod.Enabled = true;
            textBox6.Enabled = true;
            textBox1.Enabled = false;
            textBox3.Enabled = false;
            textBox2.Enabled = false;
            button3.Enabled = false;
            button2.Enabled = false;
            button3.Enabled = false;
        }


        /*private void button7_Click(object sender, EventArgs e)
        {
            string filePath2 = @"D:\temp\sss2.docx";

            DocumentCore dc = DocumentCore.Load(filePath2);

            // Update and count the number of words and pages in the file.
            dc.CalculateStats();

            // Show statistics.
            //Console.WriteLine("Pages: {0}", dc.Document.Properties.BuiltIn[BuiltInDocumentProperty.Pages]);
            textBox3.Text = dc.Document.Properties.BuiltIn[BuiltInDocumentProperty.Pages];
        }*/

        /*private void button7_Click(object sender, EventArgs e)
        /
            //Table start.
            string html = "<table cellpadding='5' cellspacing='0' style='border: 1px solid #ccc;font-size: 9pt;font-family:arial'>";

            //Adding HeaderRow.
            html += "<tr>";
            foreach (DataGridViewColumn column in dataGridView3.Columns)
            {
                html += "<th style='background-color: #B8DBFD;border: 1px solid #ccc'>" + column.HeaderText + "</th>";
            }
            html += "</tr>";

            //Adding DataRow.
            foreach (DataGridViewRow row in dataGridView3.Rows)
            {
                html += "<tr>";
                foreach (DataGridViewCell cell in row.Cells)
                {
                    html += "<td style='width:120px;border: 1px solid #ccc'>" + cell.Value.ToString() + "</td>";
                }
                html += "</tr>";
            }

            //Table end.
            html += "</table>";

            //Save the HTML string as HTML File.
            string htmlFilePath = @"D:\temp\DataGridView.htm";
            File.WriteAllText(htmlFilePath, html);

            //Convert the HTML File to Word document.
            _Application word = new Microsoft.Office.Interop.Word.Application();
            _Document wordDoc = word.Documents.Open(FileName: htmlFilePath, ReadOnly: false);
            wordDoc.SaveAs(FileName: @"D:\temp\DataGridView.doc", FileFormat: WdSaveFormat.wdFormatRTF);
            ((_Document)wordDoc).Close();
            ((_Application)word).Quit();

            //Delete the HTML File.
            File.Delete(htmlFilePath);
        }*/
        /*
for (int i = 0; i < dataGridView1.Rows.Count; i++)
{
if (dataGridView1["Категория", i].Value != null && (dataGridView1["Категория", i].Value.ToString().Trim() == y1))
{
DataGridViewRow dgvDelRow = dataGridView1.Rows[i];
dataGridView1.Rows.Remove(dgvDelRow);
}
}
dataGridView1.Refresh();*/
    }
           /* {
                List<DataGridViewRow> RowsToDelete = new List<DataGridViewRow>();
                foreach (DataGridViewRow row in dataGridView1.Rows)
                    if (row.Cells[2].Value != null &&
                         row.Cells[2].Value.ToString() == y1) RowsToDelete.Add(row);
                foreach (DataGridViewRow row in RowsToDelete) dataGridView1.Rows.Remove(row);
                RowsToDelete.Clear();
            }*/
            //int rowIndex = dataGridView1.CurrentCell.RowIndex;
            // dataGridView1.Rows.RemoveAt(rowIndex);
            /* for (int i = 0; i < dataGridView1.RowCount; i++)
             {
                 DataGridViewRow dgvDelRow = dataGridView1.Rows[i];
                 dataGridView1.Rows.Remove(dgvDelRow);
             }
             */
                /*for (int i = 0; i < dataGridView1.Rows.Count; i++)
                {
                    if (/*dataGridView1["Категория", i].Value != null && (dataGridView1["Категория", i].Value.ToString().Trim() == y1))
                {
                    DataGridViewRow dgvDelRow = dataGridView1.Rows[i];
                    dataGridView1.Rows.Remove(dgvDelRow);
                }
            }
           // dataGridView1.Refresh();
        }*/
}
        /*sqlcon.Open();
SqlCommand sqlcmd = new SqlCommand("DBCC CHECKIDENT ('[uploadfile]', RESEED, 0)", sqlcon);
sqlcmd.ExecuteNonQuery();
sqlcon.Close();*/
