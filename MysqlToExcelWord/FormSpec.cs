//using System.Threading.Tasks;
//using System.Windows.Forms;
//using static NPOI.HSSF.Util.HSSFColor;
//using static System.ComponentModel.Design.ObjectSelectorEditor;
//using NPOI.SS.Formula.Functions;
//using System;
//using System.Collections.Generic;
//using System.ComponentModel;
//using System.Diagnostics.Metrics;
//using System.Drawing;
//using System.Linq;

using MySql.Data.MySqlClient;
using System.Data;
using System.Diagnostics;
using System.Text;



namespace MysqlToExcelWord
{
    public partial class FormSpec : Form
    {
        public event EventHandler<string> EventCableSpec; // EventHandler<TEventArgs> 默认泛型委托    
        string tableNameFromButton;
        string mysqlSpec;
       // string  SpecFromButton;
        string type_spec;
      //  bool specSelected;
        List<Button> buttonList;
        public string SpecFromButton { get; private set; }
        public string TableNameFromButton
        {
            set { tableNameFromButton = value; }
        }

        public string Type_specFromMysql
        {
            // set { type_spec = value; }
            get { return type_spec; }
        }

        public FormSpec()
        {
            InitializeComponent();            
        }



        private void specForm_Load(object sender, EventArgs e)
        {
              buttonList = this.Controls.OfType<Button>()
                                .Where(b => b.Name.StartsWith("button"))
                                 //不要orderby,                               //排序从右至左，从上至下排
                                 .OrderBy(b => b.TabIndex) //按 button33.TabIndex = 32;
                              //  .OrderBy(b => b.Name)// 实际乱了
                                //下面两句，从左至右，从上至下
                               //.OrderBy(b => b.Left)
                               //.ThenBy(b => b.Top)                              
                                .ToList();
            /* */
            foreach (var button in buttonList)
               {
                   button.Visible = false;  
               }

            //_ = GetMysqlSpec(); //异步连接
             GetMysqlSpec(); //异步连接

            foreach (var button in buttonList)
            {
                button.Click += Button_Click;
            }
            //this.FormClosing += SpecForm_FormClosing1;     
           
        }

        private void Button_Click(object sender, EventArgs e)
        {
            // 获取被点击的按钮
            Button clickedButton = sender as Button;

            if (clickedButton != null)
            {
                SpecFromButton = clickedButton.Text;
                // _ =GetMysqlType_Spec(); //异步连接
                GetMysqlType_Spec(); //异步连接
                Trace.WriteLine("specForm 按钮动作");


                // 存储被点击按钮的文本
                // SpecFromButton = clickedButton.Text;
                /*  //EventCableSpec.Invoke(this, clickedButton.Text);
                 // EventCableSpec.Invoke(this, SpecFromButton);
                  EventCableSpec.Invoke(this, type_spec);

                  */
               // specSelected = true;
                this.Close();
            }
        }
        /*
        public  string CableSpec {
            get { return SpecFromButton; }
         }
        */


        void GetMysqlSpec()
        // public async Task GetMysqlSpec() //异步连接
        {  

            var connectionString = "server=localhost;user=root;password=8888;database=cableLv";

            try
            {
                using var mysqlConnection = new MySqlConnection(connectionString);
                //await mysqlConnection.OpenAsync(); //异步连接
                mysqlConnection.Open(); 
                Trace.WriteLine("\n GetMysqlSpec成功连接到MySQL数据库");
         
                string queryString = $"SELECT spec FROM {tableNameFromButton}";
                Trace.WriteLine($"tableNameFromButton: {tableNameFromButton}");
                using var mysqlCommand = new MySqlCommand(queryString, mysqlConnection);

                // using var mysqlReader = await mysqlCommand.ExecuteReaderAsync(); //异步连接
                using var mysqlReader = mysqlCommand.ExecuteReader();
                Trace.WriteLine("\nspec查询结果:");
                Trace.WriteLine("spec");//\t 长空格tab
                //Trace.WriteLine(new string('-', 50));//打印50个横杠
                int buttonCounter = 1;
                while (mysqlReader.Read())
                //while (await mysqlReader.ReadAsync()) //异步连接
                {
                    mysqlSpec = mysqlReader.GetString("spec");//每个循环，得到spec列的一个数据
                    buttonList[buttonCounter-1].Text = $"{mysqlSpec}";
                    buttonList[buttonCounter - 1].Visible = true; 
                    buttonCounter++;
                    Trace.WriteLine($"来自数据库的spec：{mysqlSpec}");
                }

            Trace.WriteLine("spec查询完成");
            }
            catch (Exception ex)
            {
                Trace.WriteLine($"spec查询发生错误: {ex.Message}");
            }

        }

        void GetMysqlType_Spec()
        //public async Task GetMysqlType_Spec()
        {  //异步连接，用async Task

            var connectionString = "server=localhost;user=root;password=8888;database=cableLv";           
            try
            {
                // 1. 创建数据库连接
                using var mysqlConnection = new MySqlConnection(connectionString);
                //await mysqlConnection.OpenAsync(); //异步连接
                 mysqlConnection.Open();

                Trace.WriteLine("\nGetMysqlType_Spec 成功连接到MySQL数据库");

                // 2. 定义查询 - 这里我们只选择特定的列和行
                //string queryString = "SELECT type_spec FROM Z_YJVx2  WHERE spec = '3×10＋1×6'";
                string queryString = $"SELECT type_spec FROM {tableNameFromButton} WHERE spec = @spec";
               
                Trace.WriteLine($"数据表tableNameFromButton 为: {tableNameFromButton}");
                // 3. 创建命令对象并添加参数
                using var mysqlCommand = new MySqlCommand(queryString, mysqlConnection);
                mysqlCommand.Parameters.AddWithValue("@spec", SpecFromButton);
                // 4. 执行查询并读取结果
                //using var reader = await mysqlCommand.ExecuteReaderAsync();  //异步连接
                using var reader = mysqlCommand.ExecuteReader();
                Trace.WriteLine($"specFromButton 值: {SpecFromButton}");

                Trace.WriteLine("type_spec查询结果:");
                //Trace.WriteLine(new string('-', 50));//打印50个横杠

                // 5. 处理结果集
                
                if (reader.Read())
                {
                    type_spec = reader.GetString("type_spec");
                    Trace.WriteLine($"找到匹配的type_spec: {type_spec}");
                }
                else
                {
                    Trace.WriteLine($"无匹配的type_spec");
                }
                Trace.WriteLine("type_spec查询完成");
                EventCableSpec?.Invoke(this, type_spec);
            }
            catch (Exception ex)
            {
                Trace.WriteLine($"type_spec发生错误: {ex.Message}");
            }           

        }



    }
}
