namespace WindowsFormsApp1
{
    partial class author
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
            this.button3 = new System.Windows.Forms.Button();
            this.user = new System.Windows.Forms.TextBox();
            this.button1 = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // button3
            // 
            this.button3.Location = new System.Drawing.Point(12, 40);
            this.button3.Name = "button3";
            this.button3.Size = new System.Drawing.Size(290, 23);
            this.button3.TabIndex = 8;
            this.button3.Text = "Добавить автора";
            this.button3.UseVisualStyleBackColor = true;
            this.button3.Click += new System.EventHandler(this.button3_Click);
            // 
            // user
            // 
            this.user.Location = new System.Drawing.Point(12, 12);
            this.user.Name = "user";
            this.user.Size = new System.Drawing.Size(558, 22);
            this.user.TabIndex = 6;
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(308, 40);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(262, 23);
            this.button1.TabIndex = 5;
            this.button1.Text = "Вернуться";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // author
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(587, 76);
            this.Controls.Add(this.button3);
            this.Controls.Add(this.user);
            this.Controls.Add(this.button1);
            this.Name = "author";
            this.Text = "author";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button button3;
        private System.Windows.Forms.TextBox user;
        private System.Windows.Forms.Button button1;
    }
}