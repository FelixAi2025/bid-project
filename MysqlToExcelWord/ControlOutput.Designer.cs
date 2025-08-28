namespace MysqlToExcelWord
{
    partial class ControlOutput
    {
        /// <summary> 
        /// 必需的设计器变量。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary> 
        /// 清理所有正在使用的资源。
        /// </summary>
        /// <param name="disposing">如果应释放托管资源，为 true；否则为 false。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region 组件设计器生成的代码

        /// <summary> 
        /// 设计器支持所需的方法 - 不要修改
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InitializeComponent()
        {
            button4 = new Button();
            button3 = new Button();
            button2 = new Button();
            textBox1 = new TextBox();
            button1 = new Button();
            //button5 = new Button();
            button6 = new Button();
            button7 = new Button();
            SuspendLayout();
            // 
            // button4
            // 
            button4.Location = new Point(515, 209);
            button4.Name = "button4";
            button4.Size = new Size(159, 63);
            button4.TabIndex = 69;
            button4.Text = "Word替代";
            button4.UseVisualStyleBackColor = true;
            button4.Click += button4_Click;
            // 
            // button3
            // 
            button3.Location = new Point(393, 209);
            button3.Name = "button3";
            button3.Size = new Size(116, 64);
            button3.TabIndex = 68;
            button3.Text = "写Word";
            button3.UseVisualStyleBackColor = true;
            button3.Click += button3_Click;
            // 
            // button2
            // 
            button2.Location = new Point(680, 169);
            button2.Name = "button2";
            button2.Size = new Size(123, 103);
            button2.TabIndex = 67;
            button2.Text = "取消";
            button2.UseVisualStyleBackColor = true;
            button2.Click += button2_Click;
            // 
            // textBox1
            // 
            textBox1.Location = new Point(276, 169);
            textBox1.Name = "textBox1";
            textBox1.Size = new Size(398, 35);
            textBox1.TabIndex = 66;
            // 
            // button1
            // 
            button1.Location = new Point(276, 209);
            button1.Name = "button1";
            button1.Size = new Size(111, 64);
            button1.TabIndex = 65;
            button1.Text = "写Excel";
            button1.UseVisualStyleBackColor = true;
            button1.Click += button1_Click;
            // 
            // button5
            // 
            //button5.Location = new Point(103, 62);
            //button5.Name = "button5";
            //button5.Size = new Size(170, 56);
            //button5.TabIndex = 70;
            //button5.Text = "button5";
            //button5.UseVisualStyleBackColor = true;
            //button5.Click += button5_Click;
            // 
            // button6
            // 
            button6.Location = new Point(1012, 79);
            button6.Name = "button6";
            button6.Size = new Size(123, 56);
            button6.TabIndex = 168;
            button6.Text = "返回首页";
            button6.UseVisualStyleBackColor = true;
            button6.Click += button6_Click;
            // 
            // button7
            // 
            button7.Location = new Point(1012, 17);
            button7.Name = "button7";
            button7.Size = new Size(123, 56);
            button7.TabIndex = 167;
            button7.Text = "退出";
            button7.UseVisualStyleBackColor = true;
            button7.Click += button7_Click;
            // 
            // ControlOutput
            // 
            AutoScaleDimensions = new SizeF(13F, 30F);
            AutoScaleMode = AutoScaleMode.Font;
            Controls.Add(button6);
            Controls.Add(button7);
            //Controls.Add(button5);
            Controls.Add(button4);
            Controls.Add(button3);
            Controls.Add(button2);
            Controls.Add(textBox1);
            Controls.Add(button1);
            Name = "ControlOutput";
            Size = new Size(1285, 518);
            Load += ControlOutput_Load;
            ResumeLayout(false);
            PerformLayout();
        }

        #endregion

        //  private Button button6;
        // private Button button5;
        internal Button button4;
        internal Button button3;
        internal Button button2;
        internal TextBox textBox1;
        internal Button button1;
        //private Button button5;
        private Button button6;
        private Button button7;
    }
}
