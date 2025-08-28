namespace MysqlToExcelWord
{
    partial class FormStart
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
            label1 = new Label();
            button1 = new Button();
            button2 = new Button();
            button3 = new Button();
            SuspendLayout();
            // 
            // label1
            // 
            label1.AutoSize = true;
            label1.Font = new Font("Microsoft YaHei UI", 40F);
            label1.Location = new Point(275, 202);
            label1.Name = "label1";
            label1.Size = new Size(573, 88);
            label1.TabIndex = 63;
            label1.Text = "标书文档生成系统";
            // 
            // button1
            // 
            button1.Font = new Font("Microsoft YaHei UI", 20F);
            button1.Location = new Point(213, 440);
            button1.Name = "button1";
            button1.Size = new Size(180, 124);
            button1.TabIndex = 64;
            button1.Text = "低压电缆";
            button1.UseVisualStyleBackColor = true;
            button1.Click += button1_Click;
            // 
            // button2
            // 
            button2.Font = new Font("Microsoft YaHei UI", 20F);
            button2.Location = new Point(469, 439);
            button2.Name = "button2";
            button2.Size = new Size(180, 125);
            button2.TabIndex = 65;
            button2.Text = "中压电缆";
            button2.UseVisualStyleBackColor = true;
            button2.Click += button2_Click;
            // 
            // button3
            // 
            button3.Font = new Font("Microsoft YaHei UI", 20F);
            button3.Location = new Point(715, 440);
            button3.Name = "button3";
            button3.Size = new Size(180, 124);
            button3.TabIndex = 66;
            button3.Text = "高压电缆";
            button3.UseVisualStyleBackColor = true;
            button3.Click += button3_Click;
            // 
            // FormStart
            // 
            AutoScaleDimensions = new SizeF(13F, 30F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(1184, 880);
            Controls.Add(button3);
            Controls.Add(button2);
            Controls.Add(button1);
            Controls.Add(label1);
            Name = "FormStart";
            Text = "FormStart";
            ResumeLayout(false);
            PerformLayout();
        }

        #endregion

        private Label label1;
        private Button button1;
        private Button button2;
        private Button button3;
    }
}