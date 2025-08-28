//using System;
//using System.Collections.Generic;
//using System.ComponentModel;
//using System.Data;
//using System.Drawing;
//using System.Linq;
//using System.Text;
//using System.Threading.Tasks;
//using System.Windows.Forms;
//using MysqlToExcelWord;
//using NPOI.OpenXmlFormats.Wordprocessing;
//using NPOI.SS.UserModel;
//using NPOI.SS.Util;
//using NPOI.WP.UserModel;
//using NPOI.XSSF.UserModel;  // 对应xlsx格式
//using NPOI.XWPF.UserModel;
//using System.Text.RegularExpressions;
//using ICell = NPOI.SS.UserModel.ICell;
//using MatchRegex = System.Text.RegularExpressions.Match;
//using MySqlCommand = MySqlConnector.MySqlCommand;
//using MySqlConnection = MySqlConnector.MySqlConnection;
//using static MysqlToExcelWord.FormLV;
//using System.Runtime.CompilerServices;

using System.Diagnostics;
using HorizontalAlignmentForm = System.Windows.Forms.HorizontalAlignment;
using static MysqlToExcelWord.FormStart;


namespace MysqlToExcelWord
{
    public partial class ControlOutput : UserControl
    {
        public event EventHandler EventCancel;
        public event EventHandler EventTestClose;
        public event EventHandler EventCloseAll; // EventHandler<TEventArgs> 默认泛型委托  定义监听    
        public event EventHandler EventCloseType; // EventHandler<TEventArgs> 默认泛型委托  定义监听    
        public ControlOutput()
        {
            InitializeComponent();
            textBox1.ForeColor = Color.SteelBlue;
            textBox1.TextAlign = HorizontalAlignmentForm.Center;
        }

        private void ControlOutput_Load(object sender, EventArgs e)
        {
            button1.Enabled = button2.Enabled = button3.Enabled = button4.Enabled = false;

        }

        private void button1_Click(object sender, EventArgs e)//写Excel
        {
            //MySql数据查询
            //  getMysqlData();
            Trace.WriteLine("写Excel");
            //写Excel
            CreateExcel createExcel1 = new();
            createExcel1.ToCreateExcel();
            createExcel1 = null;
        }


        private void button3_Click(object sender, EventArgs e) //写Word
        {
            //inPutTemplatePath = @"D:\1Development\2bid\";
            //MySql数据查询
            //   getMysqlData();
            Trace.WriteLine("写Word=");
            CreateWord generator = new();
            //generator.TableDataInputNew();
            //generator.CreateWordTechKeyValuePair(generator.TableDataInputNew());
            generator.CreateWordTechKeyValuePair();
            generator.CreateWordQuoteKeyValuePair();
        }

        private void button4_Click(object sender, EventArgs e)//替代Word
        {
            //MySql数据查询
            //   getMysqlData();
            Trace.WriteLine("替代Word");
            var generator1 = new CreateWord();
            generator1.ReplaceWordTech(@"D:\1Development\2bid\");
            generator1.ReplaceWordQuote(@"D:\1Development\2bid\");
        }

        private void button2_Click(object sender, EventArgs e)  //取消原选项
        {
            Trace.WriteLine("取消原选项");
            EventCancel?.Invoke(this, EventArgs.Empty);
            button1.Enabled = button2.Enabled = button3.Enabled = button4.Enabled = false;
            textBox1.Text = "";

        }


        private void button7_Click(object sender, EventArgs e)    //退出
        {
            //EventCloseAll?.Invoke(this,true);
            EventCloseAll?.Invoke(this, EventArgs.Empty);
            Trace.WriteLine("ControlType 关闭界面");
        }

        private void button6_Click(object sender, EventArgs e)   // 返回首页

        {
            EventCloseType?.Invoke(this, EventArgs.Empty);
        }






    }
}
