namespace OtomatikSiparisGirisi
{
    partial class SiparisGirisi
    {
        /// <summary>
        ///  Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        ///  Clean up any resources being used.
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
        ///  Required method for Designer support - do not modify
        ///  the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            button1 = new Button();
            dataGridView1 = new DataGridView();
            button2 = new Button();
            textBox1 = new TextBox();
            label1 = new Label();
            button4 = new Button();
            dateTimePicker1 = new DateTimePicker();
            label3 = new Label();
            label4 = new Label();
            textBox3 = new TextBox();
            label6 = new Label();
            textBox4 = new TextBox();
            textBox2 = new TextBox();
            label2 = new Label();
            label7 = new Label();
            textBox5 = new TextBox();
            label5 = new Label();
            textBox6 = new TextBox();
            button5 = new Button();
            ((System.ComponentModel.ISupportInitialize)dataGridView1).BeginInit();
            SuspendLayout();
            // 
            // button1
            // 
            button1.Location = new Point(148, 46);
            button1.Name = "button1";
            button1.Size = new Size(237, 45);
            button1.TabIndex = 0;
            button1.Text = "Exel Tablosunu Aktar";
            button1.UseVisualStyleBackColor = true;
            button1.Click += button1_Click;
            // 
            // dataGridView1
            // 
            dataGridView1.AllowUserToAddRows = false;
            dataGridView1.BackgroundColor = Color.White;
            dataGridView1.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            dataGridView1.Location = new Point(-1, 242);
            dataGridView1.Name = "dataGridView1";
            dataGridView1.RowHeadersWidth = 51;
            dataGridView1.RowTemplate.Height = 25;
            dataGridView1.Size = new Size(1900, 731);
            dataGridView1.TabIndex = 1;
            // 
            // button2
            // 
            button2.Location = new Point(1655, 28);
            button2.Name = "button2";
            button2.Size = new Size(257, 56);
            button2.TabIndex = 2;
            button2.Text = "Veritabanına Siparişleri İşle";
            button2.UseVisualStyleBackColor = true;
            button2.Click += button2_Click;
            // 
            // textBox1
            // 
            textBox1.Location = new Point(148, 15);
            textBox1.Name = "textBox1";
            textBox1.Size = new Size(237, 29);
            textBox1.TabIndex = 3;
            textBox1.Text = "sonuc";
            // 
            // label1
            // 
            label1.AutoSize = true;
            label1.Location = new Point(13, 15);
            label1.Name = "label1";
            label1.Size = new Size(129, 23);
            label1.TabIndex = 4;
            label1.Text = "Exel Sayfa Adı:";
            // 
            // button4
            // 
            button4.Location = new Point(634, 180);
            button4.Name = "button4";
            button4.Size = new Size(277, 56);
            button4.TabIndex = 8;
            button4.Text = "Barkoda Göre Stok Kodu";
            button4.UseVisualStyleBackColor = true;
            button4.Click += button4_Click;
            // 
            // dateTimePicker1
            // 
            dateTimePicker1.Location = new Point(648, 15);
            dateTimePicker1.Name = "dateTimePicker1";
            dateTimePicker1.Size = new Size(200, 29);
            dateTimePicker1.TabIndex = 9;
            // 
            // label3
            // 
            label3.AutoSize = true;
            label3.Location = new Point(422, 18);
            label3.Name = "label3";
            label3.Size = new Size(217, 23);
            label3.TabIndex = 12;
            label3.Text = "Sipariş Güncelleme Tarihi:";
            // 
            // label4
            // 
            label4.AutoSize = true;
            label4.Location = new Point(470, 55);
            label4.Name = "label4";
            label4.Size = new Size(130, 23);
            label4.TabIndex = 14;
            label4.Text = "Sip_Satici_kod:\r\n";
            // 
            // textBox3
            // 
            textBox3.Location = new Point(606, 52);
            textBox3.Name = "textBox3";
            textBox3.Size = new Size(51, 29);
            textBox3.TabIndex = 13;
            // 
            // label6
            // 
            label6.AutoSize = true;
            label6.Location = new Point(663, 55);
            label6.Name = "label6";
            label6.Size = new Size(131, 23);
            label6.TabIndex = 18;
            label6.Text = "Alt Döviz kuru:";
            // 
            // textBox4
            // 
            textBox4.Location = new Point(800, 55);
            textBox4.Name = "textBox4";
            textBox4.Size = new Size(48, 29);
            textBox4.TabIndex = 17;
            // 
            // textBox2
            // 
            textBox2.Location = new Point(1101, 12);
            textBox2.Name = "textBox2";
            textBox2.Size = new Size(53, 29);
            textBox2.TabIndex = 19;
            // 
            // label2
            // 
            label2.AutoSize = true;
            label2.Location = new Point(875, 15);
            label2.Name = "label2";
            label2.Size = new Size(220, 23);
            label2.TabIndex = 20;
            label2.Text = "Sip_Oluşturan_kulancı_No:";
            // 
            // label7
            // 
            label7.AutoSize = true;
            label7.Location = new Point(857, 52);
            label7.Name = "label7";
            label7.Size = new Size(238, 23);
            label7.TabIndex = 22;
            label7.Text = "Sip_Güncelleyen_kulancı_No:";
            // 
            // textBox5
            // 
            textBox5.Location = new Point(1101, 49);
            textBox5.Name = "textBox5";
            textBox5.Size = new Size(53, 29);
            textBox5.TabIndex = 21;
            // 
            // label5
            // 
            label5.AutoSize = true;
            label5.Location = new Point(13, 114);
            label5.Name = "label5";
            label5.Size = new Size(129, 23);
            label5.TabIndex = 26;
            label5.Text = "Exel Sayfa Adı:";
            // 
            // textBox6
            // 
            textBox6.Location = new Point(148, 111);
            textBox6.Name = "textBox6";
            textBox6.Size = new Size(237, 29);
            textBox6.TabIndex = 25;
            textBox6.Text = "FİRMA TANIMLAYICI KOD";
            // 
            // button5
            // 
            button5.Location = new Point(148, 142);
            button5.Name = "button5";
            button5.Size = new Size(237, 45);
            button5.TabIndex = 24;
            button5.Text = "Exel Tablosunu Aktar";
            button5.UseVisualStyleBackColor = true;
            button5.Click += button5_Click;
            // 
            // SiparisGirisi
            // 
            AutoScaleDimensions = new SizeF(10F, 21F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(1924, 985);
            Controls.Add(label5);
            Controls.Add(textBox6);
            Controls.Add(button5);
            Controls.Add(label7);
            Controls.Add(textBox5);
            Controls.Add(label2);
            Controls.Add(textBox2);
            Controls.Add(label6);
            Controls.Add(textBox4);
            Controls.Add(label4);
            Controls.Add(textBox3);
            Controls.Add(label3);
            Controls.Add(dateTimePicker1);
            Controls.Add(button4);
            Controls.Add(label1);
            Controls.Add(textBox1);
            Controls.Add(button2);
            Controls.Add(dataGridView1);
            Controls.Add(button1);
            Font = new Font("Segoe UI", 9.75F, FontStyle.Bold, GraphicsUnit.Point);
            Name = "SiparisGirisi";
            StartPosition = FormStartPosition.CenterScreen;
            Text = "Form1";
            WindowState = FormWindowState.Maximized;
            Load += SiparisGirisi_Load;
            ((System.ComponentModel.ISupportInitialize)dataGridView1).EndInit();
            ResumeLayout(false);
            PerformLayout();
        }

        #endregion

        private Button button1;
        private DataGridView dataGridView1;
        private Button button2;
        private TextBox textBox1;
        private Label label1;
        private Button button4;
        private DateTimePicker dateTimePicker1;
        private Label label3;
        private Label label4;
        private TextBox textBox3;
        private Label label6;
        private TextBox textBox4;
        private TextBox textBox2;
        private Label label2;
        private Label label7;
        private TextBox textBox5;
        private Label label5;
        private TextBox textBox6;
        private Button button5;
    }
}