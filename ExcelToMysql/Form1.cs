using MySql.Data.MySqlClient;
using NPOI.HSSF.UserModel;  // 对应xls格式
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;  // 对应xlsx格式
using System.Diagnostics;
using System.Web;



// ExcelToMySQL完善，多行写
namespace ExcelToMysql ////NPOI_MySQL_DataTable_Excel文件夹 , MySql=NPOI=Excel  和 MySqlFromCsharp文件夹
{
    public partial class Form1 : Form
    {
        string connectionString = "server=localhost;user=root;database=cableLv ;port=3306;password=8888;";
        public Form1()
        {
            InitializeComponent();
        }
        private void Form1_Load(object sender, EventArgs e)
        {
            CheckZxTable();
        }
        private void mysqlInputButton_Click(object sender, EventArgs e)//阻燃型
        {
            /*
            CREATE DATABASE bid
              CHARACTER SET utf8mb4
                COLLATE utf8mb4_unicode_ci;

            CREATE TABLE z_yjvx2 ( 				
            steel_width DECIMAL(6,1) COMMENT '钢带宽度',				
            sheath_thick DECIMAL(6,1) COMMENT '护套厚度',				
            diameter DECIMAL(6,1) COMMENT '电缆直径',				
            spec  VARCHAR(20)  PRIMARY KEY COMMENT '规格',     				
            conductor DECIMAL(8,1) NOT NULL COMMENT '铜',				
            mica_tape DECIMAL(6,1) COMMENT '云母带',				
            silane_insulation DECIMAL(6,1) COMMENT '二步法硅烷交联绝缘料',				
            UV_Insulation DECIMAL(6,1) COMMENT '紫外光绝缘料',				
            PP_buffer DECIMAL(6,1) COMMENT 'PP填充绳',				
            rockwool_buffer DECIMAL(6,1) COMMENT  '阻燃岩棉填充绳',				
            PP_tape DECIMAL(6,1) COMMENT 'PP带',				
            nonwoven  DECIMAL(6,1) COMMENT '无纺布',				
            inner_sheath  DECIMAL(6,1) COMMENT '内护套料',				
            armour DECIMAL(6,1) COMMENT '铠装',				
            outer_sheath  DECIMAL(6,1) COMMENT '外护套料',				
            weight DECIMAL(6,1) COMMENT '质量'				
         ) COMMENT= '阻燃型电缆yjv62，yjv22数据';				


            */

            //Trace.WriteLine("\n--------------------------------打印1");
            // Trace.WriteLine("\n--------------------------------打印2");  
            /*  
                ExcelToMySQL ExcelToMySQL9 = new ExcelToMySQL("server=localhost;user=root;database=cableLv;port=3306;password=8888;" //database=d1
                                          , "zXX"  //MySQL Table name
                                          , @"D:\1Development\2bid\定额\阻燃型\Z-VV报价定额.xlsx"                 //"D:\1Development\Temp\学生.xlsx"
                                          , "z"           // sheetName
                                          , 5       // header row number show in excel. The first row is 1
                                          , 2 // start column number shows in excel. The first column number show A in excel calls 1 here.
                                          );
                ExcelToMySQL9.TheExcelToMySql();
                ExcelToMySQL9 = null;
             */

            /*           
            if (Z_VV_CheckBox.Checked)
            {
                ExcelToMySQL ExcelToMySQL9 = new ExcelToMySQL("server=localhost;user=root;database=cableLv;port=3306;password=8888;"  //database=d1
                                                          , "z_vv"  //MySQL Table name
                                                          , @"D:\1Development\2bid\定额\阻燃型\Z-VV报价定额.xlsx"                 //"D:\1Development\Temp\学生.xlsx"
                                                          , "z"           // sheetName
                                                          , 5        // header row number show in excel. The first row is 1
                                                          , 2
                                                          );
                ExcelToMySQL9.TheExcelToMySql();
                ExcelToMySQL9 = null;
                


                ExcelToMySQL ExcelToMySQL10 = new ExcelToMySQL("server=localhost;user=root;database=cableLv;port=3306;password=8888;"  //database=d1
                                                          , "za_vv"  //MySQL Table name
                                                          , @"D:\1Development\2bid\定额\阻燃型\Z-VV报价定额.xlsx"                 //"D:\1Development\Temp\学生.xlsx"
                                                          , "za"           // sheetName
                                                          , 5        // header row number show in excel. The first row is 1 
                                                          , 2  // start column number shows in excel. The first column number show A in excel calls 1 here.
                                                          );
                ExcelToMySQL10.TheExcelToMySql();
                ExcelToMySQL10 = null;
                 

               

                ExcelToMySQL ExcelToMySQL11 = new ExcelToMySQL("server=localhost;user=root;database=cableLv;port=3306;password=8888;"  //database=d1
                                                          , "zb_vv"  //MySQL Table name
                                                          , @"D:\1Development\2bid\定额\阻燃型\Z-VV报价定额.xlsx"                 //"D:\1Development\Temp\学生.xlsx"
                                                          , "zb"           // sheetName
                                                          , 5        // header row number show in excel. The first row is 1 
                                                          , 2  // start column number shows in excel. The first column number show A in excel calls 1 here.
                                                          );
                ExcelToMySQL11.TheExcelToMySql();
                ExcelToMySQL11 = null;
              
                ExcelToMySQL ExcelToMySQL12 = new ExcelToMySQL("server=localhost;user=root;database=cableLv;port=3306;password=8888;"  //database=d1
                                                      , "zc_vv"  //MySQL Table name
                                                      , @"D:\1Development\2bid\定额\阻燃型\Z-VV报价定额.xlsx"                 //"D:\1Development\Temp\学生.xlsx"
                                                      , "zc"           // sheetName
                                                      , 5        // header row number show in excel. The first row is 1 
                                                      , 2  // start column number shows in excel. The first column number show A in excel calls 1 here.
                                                      );
                ExcelToMySQL12.TheExcelToMySql();
                ExcelToMySQL12 = null;
            }*/

            ///*
            if (Z_YJV_CheckBox.Checked)
            {

                /*
                
                ExcelToMySQL ExcelToMySQL10 = new ExcelToMySQL(connectionString   // "server =localhost;user=root;database=cableLv;port=3306;password=8888;"
                            , "area"  //MySQL Table name
                            , @"D:\1Development\2bid\数据.xlsx"              //"D:\1Development\Temp\学生.xlsx"
                            , "area"           // sheetName
                            , 2        // header row number show in excel. The first row is 1 
                            , 1  // start column number shows in excel. The first column number show A in excel calls 1 here.
                            );
                   ExcelToMySQL10.TheExcelToMySql();
                   ExcelToMySQL10 = null;

                  
                ExcelToMySQL ExcelToMySQL5 = new ExcelToMySQL(connectionString   // "server =localhost;user=root;database=cableLv;port=3306;password=8888;"
                       , "z_yjv_thick"  //MySQL Table name
                       , @"D:\1Development\2bid\定额\阻燃型\Z-YJV报价定额.xlsx"                 //"D:\1Development\Temp\学生.xlsx"
                       , "z_yjv_thick"           // sheetName
                       , 5        // header row number show in excel. The first row is 1 
                       , 1  // start column number shows in excel. The first column number show A in excel calls 1 here.
                       );
              ExcelToMySQL5.TheExcelToMySql();
              ExcelToMySQL5 = null;

                */


                ExcelToMySQL ExcelToMySQL6 = new ExcelToMySQL(connectionString   // "server =localhost;user=root;database=cableLv;port=3306;password=8888;"
                          , "z_yjv_m"  //MySQL Table name
                          , @"D:\1Development\2bid\定额\阻燃型\Z-YJV报价定额.xlsx"                 //"D:\1Development\Temp\学生.xlsx"
                          , "z_yjv_m"           // sheetName
                          , 2        // header row number show in excel. The first row is 1 
                          , 1  // start column number shows in excel. The first column number show A in excel calls 1 here.
                          );
                ExcelToMySQL6.TheExcelToMySql();
                ExcelToMySQL6 = null;

                /*
                ExcelToMySQL ExcelToMySQL1 = new ExcelToMySQL(connectionString   // "server =localhost;user=root;database=cableLv;port=3306;password=8888;"
                                                                       , "z_yjvx2"  //MySQL Table name
                                                                       , @"D:\1Development\2bid\定额\阻燃型\Z-YJV报价定额.xlsx" //@"D:\1Development\2bid\定额\阻燃型\Z-YJV报价定额.xlsx"                //"D:\1Development\Temp\学生.xlsx"
                                                                       , "z_yjvx2"           // sheetName
                                                                       , 5        // header row number show in excel. The first row is 1 
                                                                       , 1  // start column number shows in excel. The first column number show A in excel calls 1 here.
                                                                       );
              ExcelToMySQL1.TheExcelToMySql();
              ExcelToMySQL1 = null;
                
              ExcelToMySQL ExcelToMySQL6 = new ExcelToMySQL(connectionString   // "server =localhost;user=root;database=cableLv;port=3306;password=8888;"
                                        , "z_yjvx2_m"  //MySQL Table name
                                        , @"D:\1Development\2bid\定额\阻燃型\Z-YJV报价定额.xlsx"                 //"D:\1Development\Temp\学生.xlsx"
                                        , "z_yjvx2_m"           // sheetName
                                        , 5        // header row number show in excel. The first row is 1 
                                         , 1  // start column number shows in excel. The first column number show A in excel calls 1 here.
                                        );
              ExcelToMySQL6.TheExcelToMySql();
              ExcelToMySQL6 = null;
                
                ExcelToMySQL ExcelToMySQL2 = new ExcelToMySQL(connectionString   // "server =localhost;user=root;database=cableLv;port=3306;password=8888;"
                                                       , "za_yjv"  //MySQL Table name
                                                       , @"D:\1Development\2bid\定额\阻燃型\Z-YJV报价定额.xlsx"                 //"D:\1Development\Temp\学生.xlsx"
                                                       , "za_yjv"           // sheetName
                                                       , 5        // header row number show in excel. The first row is 1 
                                                        , 1  // start column number shows in excel. The first column number show A in excel calls 1 here.
                                                       );
              ExcelToMySQL2.TheExcelToMySql();
              ExcelToMySQL2 = null;


                
              ExcelToMySQL ExcelToMySQL7 = new ExcelToMySQL(connectionString   // "server =localhost;user=root;database=cableLv;port=3306;password=8888;"
                                       , "za_yjv_m"  //MySQL Table name
                                       , @"D:\1Development\2bid\定额\阻燃型\Z-YJV报价定额.xlsx"                 //"D:\1Development\Temp\学生.xlsx"
                                       , "za_yjv_m"           // sheetName
                                       , 5        // header row number show in excel. The first row is 1 
                                       , 1  // start column number shows in excel. The first column number show A in excel calls 1 here.
                                       );
              ExcelToMySQL7.TheExcelToMySql();
              ExcelToMySQL7 = null;

                

              ExcelToMySQL ExcelToMySQL3 = new ExcelToMySQL(connectionString   // "server =localhost;user=root;database=cableLv;port=3306;password=8888;"
                                                       , "zb_yjv"  //MySQL Table name
                                                       , @"D:\1Development\2bid\定额\阻燃型\Z-YJV报价定额.xlsx"                 //"D:\1Development\Temp\学生.xlsx"
                                                       , "zb_yjv"           // sheetName
                                                       , 5        // header row number show in excel. The first row is 1 
                                                        , 1  // start column number shows in excel. The first column number show A in excel calls 1 here.
                                                       );
              ExcelToMySQL3.TheExcelToMySql();
              ExcelToMySQL3 = null;

                /*
              ExcelToMySQL ExcelToMySQL8 = new ExcelToMySQL(connectionString   // "server =localhost;user=root;database=cableLv;port=3306;password=8888;"
                                       , "zb_yjv_m"  //MySQL Table name
                                       , @"D:\1Development\2bid\定额\阻燃型\Z-YJV报价定额.xlsx"                 //"D:\1Development\Temp\学生.xlsx"
                                       , "zb_yjv_m"           // sheetName
                                       , 5        // header row number show in excel. The first row is 1 
                                        , 1  // start column number shows in excel. The first column number show A in excel calls 1 here.
                                       );
              ExcelToMySQL8.TheExcelToMySql();
              ExcelToMySQL8 = null;
            

              ExcelToMySQL ExcelToMySQL4 = new ExcelToMySQL(connectionString   // "server =localhost;user=root;database=cableLv;port=3306;password=8888;"
                                                       , "zc_yjv"  //MySQL Table name
                                                       , @"D:\1Development\2bid\定额\阻燃型\Z-YJV报价定额.xlsx"                 //"D:\1Development\Temp\学生.xlsx"
                                                       , "zc_yjv"           // sheetName
                                                       , 5        // header row number show in excel. The first row is 1 
                                                        , 1  // start column number shows in excel. The first column number show A in excel calls 1 here.
                                                       );
              ExcelToMySQL4.TheExcelToMySql();
              ExcelToMySQL4 = null;
                //



               

              ExcelToMySQL ExcelToMySQL9 = new ExcelToMySQL(connectionString   // "server =localhost;user=root;database=cableLv;port=3306;password=8888;"
                                       , "zc_yjv_m"  //MySQL Table name
                                       , @"D:\1Development\2bid\定额\阻燃型\Z-YJV报价定额.xlsx"                 //"D:\1Development\Temp\学生.xlsx"
                                       , "zc_yjv_m"           // sheetName
                                       , 5        // header row number show in excel. The first row is 1 
                                        , 1  // start column number shows in excel. The first column number show A in excel calls 1 here.
                                       );
              ExcelToMySQL9.TheExcelToMySql();
              ExcelToMySQL9 = null;
                /**/

            }

            /* 
            if (Z_YJY_CheckBox.Checked)
            {
                ExcelToMySQL ExcelToMySQL5 = new ExcelToMySQL("server=localhost;user=root;database=cableLv;port=3306;password=8888;"  //database=d1
                                                          , "z_yjy"  //MySQL Table name
                                                          , @"D:\1Development\2bid\定额\阻燃型\Z-YJY报价定额.xlsx"                 //"D:\1Development\Temp\学生.xlsx"
                                                          , "z"           // sheetName
                                                          , 5        // header row number show in excel. The first row is 1 
                                                           , 2  // start column number shows in excel. The first column number show A in excel calls 1 here.
                                                          );
                ExcelToMySQL5.TheExcelToMySql();
                ExcelToMySQL5 = null;

                ExcelToMySQL ExcelToMySQL6 = new ExcelToMySQL("server=localhost;user=root;database=cableLv;port=3306;password=8888;"  //database=d1
                                                          , "za_yjy"  //MySQL Table name
                                                          , @"D:\1Development\2bid\定额\阻燃型\Z-YJY报价定额.xlsx"                 //"D:\1Development\Temp\学生.xlsx"
                                                          , "za"           // sheetName
                                                          , 5        // header row number show in excel. The first row is 1 
                                                           , 2  // start column number shows in excel. The first column number show A in excel calls 1 here.
                                                          );
                ExcelToMySQL6.TheExcelToMySql();
                ExcelToMySQL6 = null;

                ExcelToMySQL ExcelToMySQL7 = new ExcelToMySQL("server=localhost;user=root;database=cableLv;port=3306;password=8888;"  //database=d1
                                                          , "zb_yjy"  //MySQL Table name
                                                          , @"D:\1Development\2bid\定额\阻燃型\Z-YJY报价定额.xlsx"                 //"D:\1Development\Temp\学生.xlsx"
                                                          , "zb"           // sheetName
                                                          , 5        // header row number show in excel. The first row is 1 
                                                           , 2  // start column number shows in excel. The first column number show A in excel calls 1 here.
                                                          );
                ExcelToMySQL7.TheExcelToMySql();
                ExcelToMySQL7 = null;

                ExcelToMySQL ExcelToMySQL8 = new ExcelToMySQL("server=localhost;user=root;database=cableLv;port=3306;password=8888;"  //database=d1
                                                      , "zc_yjy"  //MySQL Table name
                                                      , @"D:\1Development\2bid\定额\阻燃型\Z-YJY报价定额.xlsx"                 //"D:\1Development\Temp\学生.xlsx"
                                                      , "zc"           // sheetName
                                                      , 5        // header row number show in excel. The first row is 1
                                                      , 2  // start column number shows in excel. The first column number show A in excel calls 1 here.
                                                      );
                ExcelToMySQL8.TheExcelToMySql();
                ExcelToMySQL8 = null;

            }
            */

            //CheckZxTable();
        }


        class ExcelToMySQL
        {
            string filePath;
            string connectionString;
            string mysqlTableName, sheetName;
            string specCellString , withoutSpaces;
            int primaryKeyPos;
            int excelHeaderRow, excelBeginningColumn;
            bool hasHeader;
            private List<string> indexListString;
            private List<string> columnListString;
            public ExcelToMySQL(string connectionString, string mysqlTableName, string filePath, string sheetName, int excelHeaderRow, int excelBeginningColumn, bool hasHeader = true)
            {
                this.filePath = filePath;
                this.hasHeader = hasHeader;
                this.connectionString = connectionString;
                this.mysqlTableName = mysqlTableName;
                this.sheetName = sheetName;
                this.excelHeaderRow = excelHeaderRow;
                this.excelBeginningColumn = excelBeginningColumn;
                indexListString = new List<string>();
                columnListString = new List<string>();
            }

            public void TheExcelToMySql()
            {
                // 注册编码提供程序（解决 GB2312 等编码问题）程序兼容性问题
                //using System.Text;
                //Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
                System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

                using (FileStream stream = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                {
                    IWorkbook workbook;

                    // 根据文件扩展名创建适当的workbook
                    if (Path.GetExtension(filePath).ToLower() == ".xlsx")
                        workbook = new XSSFWorkbook(stream);
                    else
                        workbook = new HSSFWorkbook(stream);

                    //ISheet sheet = workbook.GetSheetAt(0);
                    ISheet sheet = workbook.GetSheet(sheetName);
                    // IRow iHeaderRow = sheet.GetRow(0);
                    IRow iHeaderRow = sheet.GetRow(excelHeaderRow - 1);// 这儿headerRow从excel的0行开始计行数，而Excel表格从1开始计行数
                    int lastColumnCount = iHeaderRow.LastCellNum;
                    Trace.WriteLine($"列数为： {lastColumnCount}");
                    // 添加数据代数list

                    //for (int i = excelBeginningColumn - 1; i < excelBeginningColumn - 1 + lastColumnCount; i++)
                    for (int i = excelBeginningColumn - 1; i < lastColumnCount; i++)
                    {
                        // indexListString.Add($"@p{i}");//{@p0,@p1,@p2,@p3,@p4,@p5}
                        string cellValue = iHeaderRow.GetCell(i).ToString();                        
                        indexListString.Add($"@p{i - (excelBeginningColumn - 1)}");//{@p0,@p1,@p2,@p3,@p4,@p5}
                        columnListString.Add(cellValue);
                        if (cellValue == "spec")
                        {
                            primaryKeyPos = i;
                            Trace.WriteLine($"primaryKeyPos: {i}: =========================================");
                        }
                    }
                    // 2 添加数据行                   
                    //int dataStartRow = hasHeader ? 1 : 0;// 确定开始读取数据的行号
                    int dataStartRow = excelHeaderRow;// 既然iHeaderRow 是excelHeaderRow-1，数据行第一行就是excelHeaderRow-1+1= excelHeaderRow

                    //转换 数据代数list成 数据代数字符串
                    string paramNumString = string.Join(", ", indexListString);//"@p0,@p1,@p2,@p3,@p4,@p5"
                    string columnLongString1 = string.Join(", ", columnListString);
                    using (MySqlConnection mySqlConnection = new MySqlConnection(connectionString))
                    {
                        mySqlConnection.Open();
                        // 添加行数据
                        for (int i = dataStartRow; i <= sheet.LastRowNum; i++)
                        {
                            IRow iRow = sheet.GetRow(i);
                            //判断句不起作用  if (row == null) continue; // 跳过空行
                            List<object> valuesListObject = new List<object>();
                            //判断句不起作用  if (row == null) goto theEnd; //保留空行   
                            //  Trace.WriteLine("");//保留空行
                            //string sqlCommandString = $"INSERT INTO {mysqlTableName} ({columnLongString}) VALUES ({paramNumString})";
                            string sqlCommandString = $"INSERT INTO {mysqlTableName} ({columnLongString1}) VALUES ({paramNumString})";
                            using (MySqlCommand mySqlCommand1 = new MySqlCommand(sqlCommandString, mySqlConnection))
                            {
                                //for (int j = 0; j < lastColumnCount; j++)
                                // for (int j = excelBeginningColumn - 1; j < excelBeginningColumn - 1 + lastColumnCount; j++)
                                for (int j = excelBeginningColumn - 1; j < lastColumnCount; j++)
                                {
                                    //Trace.WriteLine($"j前={j}==========================================");
                                    ICell iCell = iRow.GetCell(j);
                                    if (iCell == null || iCell.CellType == CellType.Blank)
                                    {
                                        // if (j == 0) return;// excel 见空单元格退出。primary key 不能为空 
                                        if (j == excelBeginningColumn - 1) return;//  excel 见空单元格退出。primary key 不能为空 
                                        valuesListObject.Add(DBNull.Value);
                                        Trace.WriteLine($"{iHeaderRow.GetCell(j).ToString()}: DBNull (null)");
                                    }
                                    else if (j == primaryKeyPos)
                                    {
                                         specCellString = iCell.StringCellValue;
                                        //valuesListObject.Add(specCellString);
                                         withoutSpaces = specCellString.Replace(" ", string.Empty);
                                         valuesListObject.Add(withoutSpaces);
                                        Trace.WriteLine($"{iHeaderRow.GetCell(j).ToString()}: {valuesListObject[j - (excelBeginningColumn - 1)]} ({iCell.CellType})");
                                    }
                                    else
                                    {
                                        // 根据单元格类型获取值
                                        switch (iCell.CellType)
                                        {
                                            case CellType.String:
                                                valuesListObject.Add(iCell.StringCellValue);
                                                break;
                                            case CellType.Numeric:
                                                if (DateUtil.IsCellDateFormatted(iCell))
                                                {
                                                    valuesListObject.Add(iCell.DateCellValue);
                                                }
                                                else
                                                {
                                                    //Trace.WriteLine($"j={j}==========================================");
                                                    valuesListObject.Add(iCell.NumericCellValue);
                                                }
                                                break;
                                            case CellType.Boolean:
                                                valuesListObject.Add(iCell.BooleanCellValue);
                                                break;
                                            case CellType.Formula:
                                                // 对于公式单元格，获取计算后的值
                                                //  try
                                                //  {
                                                switch (iCell.CachedFormulaResultType)
                                                {
                                                    case CellType.String:
                                                        valuesListObject.Add(iCell.StringCellValue);
                                                        break;
                                                    case CellType.Numeric:
                                                        valuesListObject.Add(iCell.NumericCellValue);
                                                        break;
                                                    case CellType.Boolean:
                                                        valuesListObject.Add(iCell.BooleanCellValue);
                                                        break;
                                                    case CellType.Error:
                                                        valuesListObject.Add(DBNull.Value);
                                                        Trace.WriteLine($"excel单元格[{i},{j}]内, 公式结果错误============================");
                                                        MessageBox.Show($"excel单元格[{i},{j}]内, 公式结果错误");
                                                        break;
                                                        /*   default:
                                                              dataRow[j] = "=" + iCell.ToString();//公式错误，把公式按字符串输入
                                                              break;*/
                                                }
                                                /*   }
                                                   catch (Exception ex)
                                                   {
                                                       Trace.WriteLine("单元格错误：{0}", ex.Message);
                                                       Trace.WriteLine("---单元格 [列：{0}] [行{1}] 错误：{2}", iHeaderRow.GetCell(j).ToString(), j.ToString(), ex.Message);
                                                       dataRow[j] = "=" + iCell.CellFormula;
                                                   }*/
                                                break;
                                                /*  default:
                                                      dataRow[j] = iCell.ToString();//string.Empty;
                                                      break; */

                                        }
                                        // Trace.WriteLine($"{iHeaderRow.GetCell(j).ToString()}: {valuesListObject[j]} 测试============后");
                                        Trace.WriteLine($"{iHeaderRow.GetCell(j).ToString()}: {valuesListObject[j - (excelBeginningColumn - 1)]} ({iCell.CellType})");
                                    }
                                    mySqlCommand1.Parameters.AddWithValue($"@p{j - (excelBeginningColumn - 1)}", valuesListObject[j - (excelBeginningColumn - 1)] ?? DBNull.Value);
                                }
                                int rowsAffected = mySqlCommand1.ExecuteNonQuery();
                                Trace.WriteLine($"插入成功，影响行数: {rowsAffected}");
                                Trace.WriteLine("");
                            }
                        }
                    }
                }

            }
        }

        private void CheckZxTable()
        {
            using (MySqlConnection connection = new MySqlConnection(connectionString))
            {
                try
                {
                    connection.Open();
                    using (MySqlCommand command = new MySqlCommand("SELECT 1 FROM z_vv LIMIT 1", connection)) // 只检查是否存在任意一行
                    {
                        object result = command.ExecuteScalar();
                        if (result == null)
                        {
                            label1.Text = "空";
                            label1.BackColor = Color.Empty;
                            label1.ForeColor = Color.Gray;
                            Trace.WriteLine("z_vv表为空");
                        }
                        else
                        {
                            label1.Text = "实";
                            label1.BackColor = Color.Black;
                            label1.ForeColor = Color.White;
                            Trace.WriteLine("z_vv表不为空");
                        }
                    }
                    using (MySqlCommand command = new MySqlCommand("SELECT 1 FROM za_vv LIMIT 1", connection)) // 只检查是否存在任意一行
                    {
                        object result = command.ExecuteScalar();
                        if (result == null)
                        {
                            label2.Text = "空";
                            label2.BackColor = Color.Empty;
                            label2.ForeColor = Color.Gray;
                            Trace.WriteLine("za_vv表为空");
                        }
                        else
                        {
                            label2.Text = "实";
                            label2.BackColor = Color.Black;
                            label2.ForeColor = Color.White;
                            Trace.WriteLine("za_vv表不为空");
                        }
                    }
                    using (MySqlCommand command = new MySqlCommand("SELECT 1 FROM zb_vv LIMIT 1", connection)) // 只检查是否存在任意一行
                    {
                        object result = command.ExecuteScalar();
                        if (result == null)
                        {
                            label3.Text = "空";
                            label3.BackColor = Color.Empty;
                            label3.ForeColor = Color.Gray;
                            Trace.WriteLine("zb_vv表为空");
                        }
                        else
                        {
                            label3.Text = "实";
                            label3.BackColor = Color.Black;
                            label3.ForeColor = Color.White;
                            Trace.WriteLine("zb_vv表不为空");
                        }
                    }
                    using (MySqlCommand command = new MySqlCommand("SELECT 1 FROM zc_vv LIMIT 1", connection)) // 只检查是否存在任意一行
                    {
                        object result = command.ExecuteScalar();
                        if (result == null)
                        {
                            label4.Text = "空";
                            label4.BackColor = Color.Empty;
                            label4.ForeColor = Color.Gray;
                            Trace.WriteLine("zc_vv表为空");
                        }
                        else
                        {
                            label4.Text = "实";
                            label4.BackColor = Color.Black;
                            label4.ForeColor = Color.White;
                            Trace.WriteLine("zc_vv表不为空");
                        }
                    }

                    using (MySqlCommand command = new MySqlCommand("SELECT 1 FROM z_yjv LIMIT 1", connection)) // 只检查是否存在任意一行
                    {
                        object result = command.ExecuteScalar();
                        if (result == null)
                        {
                            label5.Text = "空";
                            label5.BackColor = Color.Empty;
                            label5.ForeColor = Color.Gray;
                            Trace.WriteLine("z_yjv表为空");
                        }
                        else
                        {
                            label5.Text = "实";
                            label5.BackColor = Color.Black;
                            label5.ForeColor = Color.White;
                            Trace.WriteLine("z_yjv表不为空");
                        }
                    }
                    using (MySqlCommand command = new MySqlCommand("SELECT 1 FROM za_yjv LIMIT 1", connection)) // 只检查是否存在任意一行
                    {
                        object result = command.ExecuteScalar();
                        if (result == null)
                        {
                            label6.Text = "空";
                            label6.BackColor = Color.Empty;
                            label6.ForeColor = Color.Gray;
                            Trace.WriteLine("za_yjv表为空");
                        }
                        else
                        {
                            label6.Text = "实";
                            label6.BackColor = Color.Black;
                            label6.ForeColor = Color.White;
                            Trace.WriteLine("za_yjv表不为空");
                        }
                    }
                    using (MySqlCommand command = new MySqlCommand("SELECT 1 FROM zb_yjv LIMIT 1", connection)) // 只检查是否存在任意一行
                    {
                        object result = command.ExecuteScalar();
                        if (result == null)
                        {
                            label7.Text = "空";
                            label7.BackColor = Color.Empty;
                            label7.ForeColor = Color.Gray;
                            Trace.WriteLine("zb_yjv表为空");
                        }
                        else
                        {
                            label7.Text = "实";
                            label7.BackColor = Color.Black;
                            label7.ForeColor = Color.White;
                            Trace.WriteLine("zb_yjv表不为空");
                        }
                    }
                    using (MySqlCommand command = new MySqlCommand("SELECT 1 FROM zc_yjv LIMIT 1", connection)) // 只检查是否存在任意一行
                    {
                        object result = command.ExecuteScalar();
                        if (result == null)
                        {
                            label8.Text = "空";
                            label8.BackColor = Color.Empty;
                            label8.ForeColor = Color.Gray;
                            Trace.WriteLine("zc_yjv表为空");
                        }
                        else
                        {
                            label8.Text = "实";
                            label8.BackColor = Color.Black;
                            label8.ForeColor = Color.White;
                            Trace.WriteLine("zc_yjv表不为空");
                        }
                    }

                    using (MySqlCommand command = new MySqlCommand("SELECT 1 FROM z_yjy LIMIT 1", connection)) // 只检查是否存在任意一行
                    {
                        object result = command.ExecuteScalar();
                        if (result == null)
                        {
                            label9.Text = "空";
                            label9.BackColor = Color.Empty;
                            label9.ForeColor = Color.Gray;
                            Trace.WriteLine("z_yjy 表为空");
                        }
                        else
                        {
                            label9.Text = "实";
                            label9.BackColor = Color.Black;
                            label9.ForeColor = Color.White;
                            Trace.WriteLine("z_yjy 表不为空");
                        }
                    }
                    using (MySqlCommand command = new MySqlCommand("SELECT 1 FROM za_yjy LIMIT 1", connection)) // 只检查是否存在任意一行
                    {
                        object result = command.ExecuteScalar();
                        if (result == null)
                        {
                            label10.Text = "空";
                            label10.BackColor = Color.Empty;
                            label10.ForeColor = Color.Gray;
                            Trace.WriteLine("za_yjy 表为空");
                        }
                        else
                        {
                            label10.Text = "实";
                            label10.BackColor = Color.Black;
                            label10.ForeColor = Color.White;
                            Trace.WriteLine("za_yjy 表不为空");
                        }
                    }
                    using (MySqlCommand command = new MySqlCommand("SELECT 1 FROM zb_yjy LIMIT 1", connection)) // 只检查是否存在任意一行
                    {
                        object result = command.ExecuteScalar();
                        if (result == null)
                        {
                            label11.Text = "空";
                            label11.BackColor = Color.Empty;
                            label11.ForeColor = Color.Gray;
                            Trace.WriteLine("zb_yjy 表为空");
                        }
                        else
                        {
                            label11.Text = "实";
                            label11.BackColor = Color.Black;
                            label11.ForeColor = Color.White;
                            Trace.WriteLine("zb_yjy 表不为空");
                        }
                    }
                    using (MySqlCommand command = new MySqlCommand("SELECT 1 FROM zc_yjy  LIMIT 1", connection)) // 只检查是否存在任意一行
                    {
                        object result = command.ExecuteScalar();
                        if (result == null)
                        {
                            label12.Text = "空";
                            label12.BackColor = Color.Empty;
                            label12.ForeColor = Color.Gray;
                            Trace.WriteLine("zc_yjy 表为空");
                        }
                        else
                        {
                            label12.Text = "实";
                            label12.BackColor = Color.Black;
                            label12.ForeColor = Color.White;
                            Trace.WriteLine("zc_yjy 表不为空");
                        }
                    }
                }
                catch (Exception ex)
                {
                    Trace.WriteLine($"发生错误: {ex.Message}");
                }
            }

        }



    }
}
