namespace WindowsFormsApp1
{
    partial class CreateDoc
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.Column6 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Column4 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Column3 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Column2 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Column1 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.dataGridView1 = new System.Windows.Forms.DataGridView();
            this.Column5 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.tabControl1 = new System.Windows.Forms.TabControl();
            this.tabPage1 = new System.Windows.Forms.TabPage();
            this.txtName = new System.Windows.Forms.TextBox();
            this.txtIzdatelstvo = new System.Windows.Forms.TextBox();
            this.txtPages = new System.Windows.Forms.TextBox();
            this.txtSoavtori = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.label8 = new System.Windows.Forms.Label();
            this.txtPeriod = new System.Windows.Forms.TextBox();
            this.button1 = new System.Windows.Forms.Button();
            this.label7 = new System.Windows.Forms.Label();
            this.txtKolvoRabot = new System.Windows.Forms.TextBox();
            this.label9 = new System.Windows.Forms.Label();
            this.txtAuthor = new System.Windows.Forms.TextBox();
            this.label6 = new System.Windows.Forms.Label();
            this.lblPageCount = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.btnRigth = new System.Windows.Forms.Button();
            this.button2 = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            this.tabControl1.SuspendLayout();
            this.tabPage1.SuspendLayout();
            this.SuspendLayout();
            // 
            // Column6
            // 
            this.Column6.HeaderText = "Категория";
            this.Column6.MinimumWidth = 8;
            this.Column6.Name = "Column6";
            this.Column6.Width = 150;
            // 
            // Column4
            // 
            this.Column4.HeaderText = "Кол-во страниц";
            this.Column4.MinimumWidth = 8;
            this.Column4.Name = "Column4";
            this.Column4.Width = 150;
            // 
            // Column3
            // 
            this.Column3.HeaderText = "Издательство, журнал или номер патента";
            this.Column3.MinimumWidth = 8;
            this.Column3.Name = "Column3";
            this.Column3.Width = 150;
            // 
            // Column2
            // 
            this.Column2.HeaderText = "Наименование";
            this.Column2.MinimumWidth = 8;
            this.Column2.Name = "Column2";
            this.Column2.Width = 150;
            // 
            // Column1
            // 
            this.Column1.HeaderText = "№";
            this.Column1.MinimumWidth = 8;
            this.Column1.Name = "Column1";
            this.Column1.Width = 150;
            // 
            // dataGridView1
            // 
            this.dataGridView1.AllowUserToAddRows = false;
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView1.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.Column1,
            this.Column2,
            this.Column3,
            this.Column4,
            this.Column5,
            this.Column6});
            this.dataGridView1.Location = new System.Drawing.Point(1292, 186);
            this.dataGridView1.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.RowHeadersWidth = 62;
            this.dataGridView1.RowTemplate.Height = 28;
            this.dataGridView1.Size = new System.Drawing.Size(1025, 398);
            this.dataGridView1.TabIndex = 49;
            // 
            // Column5
            // 
            this.Column5.HeaderText = "Соавторы";
            this.Column5.MinimumWidth = 8;
            this.Column5.Name = "Column5";
            this.Column5.Width = 150;
            // 
            // tabControl1
            // 
            this.tabControl1.Controls.Add(this.tabPage1);
            this.tabControl1.Location = new System.Drawing.Point(9, 163);
            this.tabControl1.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.tabControl1.Name = "tabControl1";
            this.tabControl1.SelectedIndex = 0;
            this.tabControl1.Size = new System.Drawing.Size(1231, 424);
            this.tabControl1.TabIndex = 46;
            this.tabControl1.SelectedIndexChanged += new System.EventHandler(this.tabControl1_SelectedIndexChanged);
            // 
            // tabPage1
            // 
            this.tabPage1.Controls.Add(this.txtName);
            this.tabPage1.Controls.Add(this.txtIzdatelstvo);
            this.tabPage1.Controls.Add(this.txtPages);
            this.tabPage1.Controls.Add(this.txtSoavtori);
            this.tabPage1.Controls.Add(this.label2);
            this.tabPage1.Controls.Add(this.label3);
            this.tabPage1.Controls.Add(this.label4);
            this.tabPage1.Controls.Add(this.label5);
            this.tabPage1.Location = new System.Drawing.Point(4, 25);
            this.tabPage1.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.tabPage1.Name = "tabPage1";
            this.tabPage1.Padding = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.tabPage1.Size = new System.Drawing.Size(1223, 395);
            this.tabPage1.TabIndex = 0;
            this.tabPage1.Text = "Работа 1";
            this.tabPage1.UseVisualStyleBackColor = true;
            // 
            // txtName
            // 
            this.txtName.Location = new System.Drawing.Point(371, 59);
            this.txtName.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.txtName.Name = "txtName";
            this.txtName.Size = new System.Drawing.Size(835, 22);
            this.txtName.TabIndex = 15;
            // 
            // txtIzdatelstvo
            // 
            this.txtIzdatelstvo.Location = new System.Drawing.Point(734, 102);
            this.txtIzdatelstvo.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.txtIzdatelstvo.Multiline = true;
            this.txtIzdatelstvo.Name = "txtIzdatelstvo";
            this.txtIzdatelstvo.Size = new System.Drawing.Size(471, 85);
            this.txtIzdatelstvo.TabIndex = 18;
            // 
            // txtPages
            // 
            this.txtPages.Location = new System.Drawing.Point(449, 205);
            this.txtPages.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.txtPages.Name = "txtPages";
            this.txtPages.Size = new System.Drawing.Size(455, 22);
            this.txtPages.TabIndex = 20;
            // 
            // txtSoavtori
            // 
            this.txtSoavtori.Location = new System.Drawing.Point(449, 246);
            this.txtSoavtori.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.txtSoavtori.Multiline = true;
            this.txtSoavtori.Name = "txtSoavtori";
            this.txtSoavtori.Size = new System.Drawing.Size(455, 85);
            this.txtSoavtori.TabIndex = 22;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(10, 62);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(347, 17);
            this.label2.TabIndex = 16;
            this.label2.Text = "Наименование научного или методического труда:";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(10, 135);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(699, 17);
            this.label3.TabIndex = 17;
            this.label3.Text = "Издательство, журнал (название, номер, год, страницы) или номер авторского свидет" +
    "ельства, патента:";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(180, 207);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(259, 17);
            this.label4.TabIndex = 19;
            this.label4.Text = "Кол-во печатных листов или страниц:";
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(364, 283);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(77, 17);
            this.label5.TabIndex = 21;
            this.label5.Text = "Соавторы:";
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Location = new System.Drawing.Point(193, 101);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(107, 17);
            this.label8.TabIndex = 44;
            this.label8.Text = "Период (годы):";
            // 
            // txtPeriod
            // 
            this.txtPeriod.Location = new System.Drawing.Point(308, 98);
            this.txtPeriod.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.txtPeriod.Name = "txtPeriod";
            this.txtPeriod.Size = new System.Drawing.Size(160, 22);
            this.txtPeriod.TabIndex = 45;
            this.txtPeriod.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtPeriod_KeyPress);
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(230, 591);
            this.button1.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(215, 50);
            this.button1.TabIndex = 37;
            this.button1.Text = "Вывести в документ";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Location = new System.Drawing.Point(473, 101);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(51, 17);
            this.label7.TabIndex = 42;
            this.label7.Text = "Автор:";
            // 
            // txtKolvoRabot
            // 
            this.txtKolvoRabot.Location = new System.Drawing.Point(747, 134);
            this.txtKolvoRabot.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.txtKolvoRabot.Name = "txtKolvoRabot";
            this.txtKolvoRabot.Size = new System.Drawing.Size(41, 22);
            this.txtKolvoRabot.TabIndex = 48;
            this.txtKolvoRabot.Text = "1";
            this.txtKolvoRabot.TextChanged += new System.EventHandler(this.txtKolvoRabot_TextChanged);
            this.txtKolvoRabot.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtKolvoRabot_KeyPress);
            // 
            // label9
            // 
            this.label9.AutoSize = true;
            this.label9.Location = new System.Drawing.Point(468, 137);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(267, 17);
            this.label9.TabIndex = 47;
            this.label9.Text = "Количество работ в данной категории:";
            // 
            // txtAuthor
            // 
            this.txtAuthor.Location = new System.Drawing.Point(532, 98);
            this.txtAuthor.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.txtAuthor.Name = "txtAuthor";
            this.txtAuthor.Size = new System.Drawing.Size(499, 22);
            this.txtAuthor.TabIndex = 43;
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(26, 65);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(85, 17);
            this.label6.TabIndex = 41;
            this.label6.Text = "PAGENAME";
            // 
            // lblPageCount
            // 
            this.lblPageCount.AutoSize = true;
            this.lblPageCount.Location = new System.Drawing.Point(136, 10);
            this.lblPageCount.Name = "lblPageCount";
            this.lblPageCount.Size = new System.Drawing.Size(12, 17);
            this.lblPageCount.TabIndex = 40;
            this.lblPageCount.Text = " ";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(23, 10);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(111, 17);
            this.label1.TabIndex = 39;
            this.label1.Text = "Число страниц:";
            // 
            // btnRigth
            // 
            this.btnRigth.Location = new System.Drawing.Point(9, 591);
            this.btnRigth.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.btnRigth.Name = "btnRigth";
            this.btnRigth.Size = new System.Drawing.Size(215, 50);
            this.btnRigth.TabIndex = 38;
            this.btnRigth.Text = "----------->";
            this.btnRigth.UseVisualStyleBackColor = true;
            this.btnRigth.Click += new System.EventHandler(this.btnRigth_Click);
            // 
            // button2
            // 
            this.button2.Location = new System.Drawing.Point(451, 591);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(215, 50);
            this.button2.TabIndex = 50;
            this.button2.Text = "На главную";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.button2_Click);
            // 
            // CreateDoc
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1251, 652);
            this.Controls.Add(this.button2);
            this.Controls.Add(this.dataGridView1);
            this.Controls.Add(this.tabControl1);
            this.Controls.Add(this.label8);
            this.Controls.Add(this.txtPeriod);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.label7);
            this.Controls.Add(this.txtKolvoRabot);
            this.Controls.Add(this.label9);
            this.Controls.Add(this.txtAuthor);
            this.Controls.Add(this.label6);
            this.Controls.Add(this.lblPageCount);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.btnRigth);
            this.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.MaximumSize = new System.Drawing.Size(1269, 706);
            this.Name = "CreateDoc";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "CreateDoc";
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            this.tabControl1.ResumeLayout(false);
            this.tabPage1.ResumeLayout(false);
            this.tabPage1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.DataGridViewTextBoxColumn Column6;
        private System.Windows.Forms.DataGridViewTextBoxColumn Column4;
        private System.Windows.Forms.DataGridViewTextBoxColumn Column3;
        private System.Windows.Forms.DataGridViewTextBoxColumn Column2;
        private System.Windows.Forms.DataGridViewTextBoxColumn Column1;
        private System.Windows.Forms.DataGridView dataGridView1;
        private System.Windows.Forms.DataGridViewTextBoxColumn Column5;
        private System.Windows.Forms.TabControl tabControl1;
        private System.Windows.Forms.TabPage tabPage1;
        private System.Windows.Forms.TextBox txtName;
        private System.Windows.Forms.TextBox txtIzdatelstvo;
        private System.Windows.Forms.TextBox txtPages;
        private System.Windows.Forms.TextBox txtSoavtori;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.TextBox txtPeriod;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.TextBox txtKolvoRabot;
        private System.Windows.Forms.Label label9;
        private System.Windows.Forms.TextBox txtAuthor;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.Label lblPageCount;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button btnRigth;
        private System.Windows.Forms.Button button2;
    }
}