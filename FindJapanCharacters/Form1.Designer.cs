namespace FindJapanCharacters
{
    partial class Form1
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
            button2 = new Button();
            richTextBox1 = new RichTextBox();
            label1 = new Label();
            label2 = new Label();
            button3 = new Button();
            button4 = new Button();
            SuspendLayout();
            // 
            // button1
            // 
            button1.Location = new Point(48, 31);
            button1.Name = "button1";
            button1.Size = new Size(188, 36);
            button1.TabIndex = 0;
            button1.Text = "Chin mời Tâm yêu chọn file <3";
            button1.UseVisualStyleBackColor = true;
            button1.Click += button1_Click;
            // 
            // button2
            // 
            button2.Location = new Point(339, 461);
            button2.Name = "button2";
            button2.Size = new Size(144, 30);
            button2.TabIndex = 1;
            button2.Text = "Tìm kí tự tiếng Nhật";
            button2.UseVisualStyleBackColor = true;
            button2.Click += button2_Click;
            // 
            // richTextBox1
            // 
            richTextBox1.Location = new Point(269, 82);
            richTextBox1.Name = "richTextBox1";
            richTextBox1.Size = new Size(842, 355);
            richTextBox1.TabIndex = 2;
            richTextBox1.Text = "";
            richTextBox1.TextChanged += richTextBox1_TextChanged;
            // 
            // label1
            // 
            label1.AutoSize = true;
            label1.Location = new Point(48, 505);
            label1.Name = "label1";
            label1.Size = new Size(203, 15);
            label1.TabIndex = 3;
            label1.Text = "Chúc Tâm yêu luôn luôn vui vẻ nhó!!!";
            label1.Click += label1_Click_1;
            // 
            // label2
            // 
            label2.AutoSize = true;
            label2.Location = new Point(48, 529);
            label2.Name = "label2";
            label2.Size = new Size(260, 15);
            label2.TabIndex = 4;
            label2.Text = "Chúc Tâm yêu luôn luôn xinh đẹp, mạnh khoẻ!!!";
            // 
            // button3
            // 
            button3.Location = new Point(601, 461);
            button3.Name = "button3";
            button3.Size = new Size(140, 30);
            button3.TabIndex = 5;
            button3.Text = "Tìm khác tiếng Nhật";
            button3.UseVisualStyleBackColor = true;
            button3.Click += button3_Click;
            // 
            // button4
            // 
            button4.Location = new Point(856, 461);
            button4.Name = "button4";
            button4.Size = new Size(116, 30);
            button4.TabIndex = 6;
            button4.Text = "Tìm tiếng Việt";
            button4.UseVisualStyleBackColor = true;
            button4.Click += button4_Click;
            // 
            // Form1
            // 
            AutoScaleDimensions = new SizeF(7F, 15F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(1256, 597);
            Controls.Add(button4);
            Controls.Add(button3);
            Controls.Add(label2);
            Controls.Add(label1);
            Controls.Add(richTextBox1);
            Controls.Add(button2);
            Controls.Add(button1);
            Name = "Form1";
            Text = "Form1";
            ResumeLayout(false);
            PerformLayout();
        }

        #endregion

        private Button button1;
        private Button button2;
        private RichTextBox richTextBox1;
        private string _excelPath;
        private Label label1;
        private Label label2;
        private Button button3;
        private Button button4;
    }
}
