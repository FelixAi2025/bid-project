using MySql.Data.MySqlClient;
using NPOI.HSSF.UserModel;  // ��Ӧxls��ʽ
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;  // ��Ӧxlsx��ʽ
using System.Diagnostics;
using System.Web;



// ExcelToMySQL���ƣ�����д
namespace ExcelToMysql ////NPOI_MySQL_DataTable_Excel�ļ��� , MySql=NPOI=Excel  �� MySqlFromCsharp�ļ���
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
        private void mysqlInputButton_Click(object sender, EventArgs e)//��ȼ��
        {
            /*
            CREATE DATABASE bid
              CHARACTER SET utf8mb4
                COLLATE utf8mb4_unicode_ci;

            CREATE TABLE z_yjvx2 ( 				
            steel_width DECIMAL(6,1) COMMENT '�ִ����',				
            sheath_thick DECIMAL(6,1) COMMENT '���׺��',				
            diameter DECIMAL(6,1) COMMENT '����ֱ��',				
            spec  VARCHAR(20)  PRIMARY KEY COMMENT '���',     				
            conductor DECIMAL(8,1) NOT NULL COMMENT 'ͭ',				
            mica_tape DECIMAL(6,1) COMMENT '��ĸ��',				
            silane_insulation DECIMAL(6,1) COMMENT '���������齻����Ե��',				
            UV_Insulation DECIMAL(6,1) COMMENT '������Ե��',				
            PP_buffer DECIMAL(6,1) COMMENT 'PP�����',				
            rockwool_buffer DECIMAL(6,1) COMMENT  '��ȼ���������',				
            PP_tape DECIMAL(6,1) COMMENT 'PP��',				
            nonwoven  DECIMAL(6,1) COMMENT '�޷Ĳ�',				
            inner_sheath  DECIMAL(6,1) COMMENT '�ڻ�����',				
            armour DECIMAL(6,1) COMMENT '��װ',				
            outer_sheath  DECIMAL(6,1) COMMENT '�⻤����',				
            weight DECIMAL(6,1) COMMENT '����'				
         ) COMMENT= '��ȼ�͵���yjv62��yjv22����';				


            */

            //Trace.WriteLine("\n--------------------------------��ӡ1");
            // Trace.WriteLine("\n--------------------------------��ӡ2");  
            /*  
                ExcelToMySQL ExcelToMySQL9 = new ExcelToMySQL("server=localhost;user=root;database=cableLv;port=3306;password=8888;" //database=d1
                                          , "zXX"  //MySQL Table name
                                          , @"D:\1Development\2bid\����\��ȼ��\Z-VV���۶���.xlsx"                 //"D:\1Development\Temp\ѧ��.xlsx"
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
                                                          , @"D:\1Development\2bid\����\��ȼ��\Z-VV���۶���.xlsx"                 //"D:\1Development\Temp\ѧ��.xlsx"
                                                          , "z"           // sheetName
                                                          , 5        // header row number show in excel. The first row is 1
                                                          , 2
                                                          );
                ExcelToMySQL9.TheExcelToMySql();
                ExcelToMySQL9 = null;
                


                ExcelToMySQL ExcelToMySQL10 = new ExcelToMySQL("server=localhost;user=root;database=cableLv;port=3306;password=8888;"  //database=d1
                                                          , "za_vv"  //MySQL Table name
                                                          , @"D:\1Development\2bid\����\��ȼ��\Z-VV���۶���.xlsx"                 //"D:\1Development\Temp\ѧ��.xlsx"
                                                          , "za"           // sheetName
                                                          , 5        // header row number show in excel. The first row is 1 
                                                          , 2  // start column number shows in excel. The first column number show A in excel calls 1 here.
                                                          );
                ExcelToMySQL10.TheExcelToMySql();
                ExcelToMySQL10 = null;
                 

               

                ExcelToMySQL ExcelToMySQL11 = new ExcelToMySQL("server=localhost;user=root;database=cableLv;port=3306;password=8888;"  //database=d1
                                                          , "zb_vv"  //MySQL Table name
                                                          , @"D:\1Development\2bid\����\��ȼ��\Z-VV���۶���.xlsx"                 //"D:\1Development\Temp\ѧ��.xlsx"
                                                          , "zb"           // sheetName
                                                          , 5        // header row number show in excel. The first row is 1 
                                                          , 2  // start column number shows in excel. The first column number show A in excel calls 1 here.
                                                          );
                ExcelToMySQL11.TheExcelToMySql();
                ExcelToMySQL11 = null;
              
                ExcelToMySQL ExcelToMySQL12 = new ExcelToMySQL("server=localhost;user=root;database=cableLv;port=3306;password=8888;"  //database=d1
                                                      , "zc_vv"  //MySQL Table name
                                                      , @"D:\1Development\2bid\����\��ȼ��\Z-VV���۶���.xlsx"                 //"D:\1Development\Temp\ѧ��.xlsx"
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
                            , @"D:\1Development\2bid\����.xlsx"              //"D:\1Development\Temp\ѧ��.xlsx"
                            , "area"           // sheetName
                            , 2        // header row number show in excel. The first row is 1 
                            , 1  // start column number shows in excel. The first column number show A in excel calls 1 here.
                            );
                   ExcelToMySQL10.TheExcelToMySql();
                   ExcelToMySQL10 = null;

                  
                ExcelToMySQL ExcelToMySQL5 = new ExcelToMySQL(connectionString   // "server =localhost;user=root;database=cableLv;port=3306;password=8888;"
                       , "z_yjv_thick"  //MySQL Table name
                       , @"D:\1Development\2bid\����\��ȼ��\Z-YJV���۶���.xlsx"                 //"D:\1Development\Temp\ѧ��.xlsx"
                       , "z_yjv_thick"           // sheetName
                       , 5        // header row number show in excel. The first row is 1 
                       , 1  // start column number shows in excel. The first column number show A in excel calls 1 here.
                       );
              ExcelToMySQL5.TheExcelToMySql();
              ExcelToMySQL5 = null;

                */


                ExcelToMySQL ExcelToMySQL6 = new ExcelToMySQL(connectionString   // "server =localhost;user=root;database=cableLv;port=3306;password=8888;"
                          , "z_yjv_m"  //MySQL Table name
                          , @"D:\1Development\2bid\����\��ȼ��\Z-YJV���۶���.xlsx"                 //"D:\1Development\Temp\ѧ��.xlsx"
                          , "z_yjv_m"           // sheetName
                          , 2        // header row number show in excel. The first row is 1 
                          , 1  // start column number shows in excel. The first column number show A in excel calls 1 here.
                          );
                ExcelToMySQL6.TheExcelToMySql();
                ExcelToMySQL6 = null;

                /*
                ExcelToMySQL ExcelToMySQL1 = new ExcelToMySQL(connectionString   // "server =localhost;user=root;database=cableLv;port=3306;password=8888;"
                                                                       , "z_yjvx2"  //MySQL Table name
                                                                       , @"D:\1Development\2bid\����\��ȼ��\Z-YJV���۶���.xlsx" //@"D:\1Development\2bid\����\��ȼ��\Z-YJV���۶���.xlsx"                //"D:\1Development\Temp\ѧ��.xlsx"
                                                                       , "z_yjvx2"           // sheetName
                                                                       , 5        // header row number show in excel. The first row is 1 
                                                                       , 1  // start column number shows in excel. The first column number show A in excel calls 1 here.
                                                                       );
              ExcelToMySQL1.TheExcelToMySql();
              ExcelToMySQL1 = null;
                
              ExcelToMySQL ExcelToMySQL6 = new ExcelToMySQL(connectionString   // "server =localhost;user=root;database=cableLv;port=3306;password=8888;"
                                        , "z_yjvx2_m"  //MySQL Table name
                                        , @"D:\1Development\2bid\����\��ȼ��\Z-YJV���۶���.xlsx"                 //"D:\1Development\Temp\ѧ��.xlsx"
                                        , "z_yjvx2_m"           // sheetName
                                        , 5        // header row number show in excel. The first row is 1 
                                         , 1  // start column number shows in excel. The first column number show A in excel calls 1 here.
                                        );
              ExcelToMySQL6.TheExcelToMySql();
              ExcelToMySQL6 = null;
                
                ExcelToMySQL ExcelToMySQL2 = new ExcelToMySQL(connectionString   // "server =localhost;user=root;database=cableLv;port=3306;password=8888;"
                                                       , "za_yjv"  //MySQL Table name
                                                       , @"D:\1Development\2bid\����\��ȼ��\Z-YJV���۶���.xlsx"                 //"D:\1Development\Temp\ѧ��.xlsx"
                                                       , "za_yjv"           // sheetName
                                                       , 5        // header row number show in excel. The first row is 1 
                                                        , 1  // start column number shows in excel. The first column number show A in excel calls 1 here.
                                                       );
              ExcelToMySQL2.TheExcelToMySql();
              ExcelToMySQL2 = null;


                
              ExcelToMySQL ExcelToMySQL7 = new ExcelToMySQL(connectionString   // "server =localhost;user=root;database=cableLv;port=3306;password=8888;"
                                       , "za_yjv_m"  //MySQL Table name
                                       , @"D:\1Development\2bid\����\��ȼ��\Z-YJV���۶���.xlsx"                 //"D:\1Development\Temp\ѧ��.xlsx"
                                       , "za_yjv_m"           // sheetName
                                       , 5        // header row number show in excel. The first row is 1 
                                       , 1  // start column number shows in excel. The first column number show A in excel calls 1 here.
                                       );
              ExcelToMySQL7.TheExcelToMySql();
              ExcelToMySQL7 = null;

                

              ExcelToMySQL ExcelToMySQL3 = new ExcelToMySQL(connectionString   // "server =localhost;user=root;database=cableLv;port=3306;password=8888;"
                                                       , "zb_yjv"  //MySQL Table name
                                                       , @"D:\1Development\2bid\����\��ȼ��\Z-YJV���۶���.xlsx"                 //"D:\1Development\Temp\ѧ��.xlsx"
                                                       , "zb_yjv"           // sheetName
                                                       , 5        // header row number show in excel. The first row is 1 
                                                        , 1  // start column number shows in excel. The first column number show A in excel calls 1 here.
                                                       );
              ExcelToMySQL3.TheExcelToMySql();
              ExcelToMySQL3 = null;

                /*
              ExcelToMySQL ExcelToMySQL8 = new ExcelToMySQL(connectionString   // "server =localhost;user=root;database=cableLv;port=3306;password=8888;"
                                       , "zb_yjv_m"  //MySQL Table name
                                       , @"D:\1Development\2bid\����\��ȼ��\Z-YJV���۶���.xlsx"                 //"D:\1Development\Temp\ѧ��.xlsx"
                                       , "zb_yjv_m"           // sheetName
                                       , 5        // header row number show in excel. The first row is 1 
                                        , 1  // start column number shows in excel. The first column number show A in excel calls 1 here.
                                       );
              ExcelToMySQL8.TheExcelToMySql();
              ExcelToMySQL8 = null;
            

              ExcelToMySQL ExcelToMySQL4 = new ExcelToMySQL(connectionString   // "server =localhost;user=root;database=cableLv;port=3306;password=8888;"
                                                       , "zc_yjv"  //MySQL Table name
                                                       , @"D:\1Development\2bid\����\��ȼ��\Z-YJV���۶���.xlsx"                 //"D:\1Development\Temp\ѧ��.xlsx"
                                                       , "zc_yjv"           // sheetName
                                                       , 5        // header row number show in excel. The first row is 1 
                                                        , 1  // start column number shows in excel. The first column number show A in excel calls 1 here.
                                                       );
              ExcelToMySQL4.TheExcelToMySql();
              ExcelToMySQL4 = null;
                //



               

              ExcelToMySQL ExcelToMySQL9 = new ExcelToMySQL(connectionString   // "server =localhost;user=root;database=cableLv;port=3306;password=8888;"
                                       , "zc_yjv_m"  //MySQL Table name
                                       , @"D:\1Development\2bid\����\��ȼ��\Z-YJV���۶���.xlsx"                 //"D:\1Development\Temp\ѧ��.xlsx"
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
                                                          , @"D:\1Development\2bid\����\��ȼ��\Z-YJY���۶���.xlsx"                 //"D:\1Development\Temp\ѧ��.xlsx"
                                                          , "z"           // sheetName
                                                          , 5        // header row number show in excel. The first row is 1 
                                                           , 2  // start column number shows in excel. The first column number show A in excel calls 1 here.
                                                          );
                ExcelToMySQL5.TheExcelToMySql();
                ExcelToMySQL5 = null;

                ExcelToMySQL ExcelToMySQL6 = new ExcelToMySQL("server=localhost;user=root;database=cableLv;port=3306;password=8888;"  //database=d1
                                                          , "za_yjy"  //MySQL Table name
                                                          , @"D:\1Development\2bid\����\��ȼ��\Z-YJY���۶���.xlsx"                 //"D:\1Development\Temp\ѧ��.xlsx"
                                                          , "za"           // sheetName
                                                          , 5        // header row number show in excel. The first row is 1 
                                                           , 2  // start column number shows in excel. The first column number show A in excel calls 1 here.
                                                          );
                ExcelToMySQL6.TheExcelToMySql();
                ExcelToMySQL6 = null;

                ExcelToMySQL ExcelToMySQL7 = new ExcelToMySQL("server=localhost;user=root;database=cableLv;port=3306;password=8888;"  //database=d1
                                                          , "zb_yjy"  //MySQL Table name
                                                          , @"D:\1Development\2bid\����\��ȼ��\Z-YJY���۶���.xlsx"                 //"D:\1Development\Temp\ѧ��.xlsx"
                                                          , "zb"           // sheetName
                                                          , 5        // header row number show in excel. The first row is 1 
                                                           , 2  // start column number shows in excel. The first column number show A in excel calls 1 here.
                                                          );
                ExcelToMySQL7.TheExcelToMySql();
                ExcelToMySQL7 = null;

                ExcelToMySQL ExcelToMySQL8 = new ExcelToMySQL("server=localhost;user=root;database=cableLv;port=3306;password=8888;"  //database=d1
                                                      , "zc_yjy"  //MySQL Table name
                                                      , @"D:\1Development\2bid\����\��ȼ��\Z-YJY���۶���.xlsx"                 //"D:\1Development\Temp\ѧ��.xlsx"
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
                // ע������ṩ���򣨽�� GB2312 �ȱ������⣩�������������
                //using System.Text;
                //Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
                System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

                using (FileStream stream = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                {
                    IWorkbook workbook;

                    // �����ļ���չ�������ʵ���workbook
                    if (Path.GetExtension(filePath).ToLower() == ".xlsx")
                        workbook = new XSSFWorkbook(stream);
                    else
                        workbook = new HSSFWorkbook(stream);

                    //ISheet sheet = workbook.GetSheetAt(0);
                    ISheet sheet = workbook.GetSheet(sheetName);
                    // IRow iHeaderRow = sheet.GetRow(0);
                    IRow iHeaderRow = sheet.GetRow(excelHeaderRow - 1);// ���headerRow��excel��0�п�ʼ����������Excel����1��ʼ������
                    int lastColumnCount = iHeaderRow.LastCellNum;
                    Trace.WriteLine($"����Ϊ�� {lastColumnCount}");
                    // ������ݴ���list

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
                    // 2 ���������                   
                    //int dataStartRow = hasHeader ? 1 : 0;// ȷ����ʼ��ȡ���ݵ��к�
                    int dataStartRow = excelHeaderRow;// ��ȻiHeaderRow ��excelHeaderRow-1�������е�һ�о���excelHeaderRow-1+1= excelHeaderRow

                    //ת�� ���ݴ���list�� ���ݴ����ַ���
                    string paramNumString = string.Join(", ", indexListString);//"@p0,@p1,@p2,@p3,@p4,@p5"
                    string columnLongString1 = string.Join(", ", columnListString);
                    using (MySqlConnection mySqlConnection = new MySqlConnection(connectionString))
                    {
                        mySqlConnection.Open();
                        // ���������
                        for (int i = dataStartRow; i <= sheet.LastRowNum; i++)
                        {
                            IRow iRow = sheet.GetRow(i);
                            //�жϾ䲻������  if (row == null) continue; // ��������
                            List<object> valuesListObject = new List<object>();
                            //�жϾ䲻������  if (row == null) goto theEnd; //��������   
                            //  Trace.WriteLine("");//��������
                            //string sqlCommandString = $"INSERT INTO {mysqlTableName} ({columnLongString}) VALUES ({paramNumString})";
                            string sqlCommandString = $"INSERT INTO {mysqlTableName} ({columnLongString1}) VALUES ({paramNumString})";
                            using (MySqlCommand mySqlCommand1 = new MySqlCommand(sqlCommandString, mySqlConnection))
                            {
                                //for (int j = 0; j < lastColumnCount; j++)
                                // for (int j = excelBeginningColumn - 1; j < excelBeginningColumn - 1 + lastColumnCount; j++)
                                for (int j = excelBeginningColumn - 1; j < lastColumnCount; j++)
                                {
                                    //Trace.WriteLine($"jǰ={j}==========================================");
                                    ICell iCell = iRow.GetCell(j);
                                    if (iCell == null || iCell.CellType == CellType.Blank)
                                    {
                                        // if (j == 0) return;// excel ���յ�Ԫ���˳���primary key ����Ϊ�� 
                                        if (j == excelBeginningColumn - 1) return;//  excel ���յ�Ԫ���˳���primary key ����Ϊ�� 
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
                                        // ���ݵ�Ԫ�����ͻ�ȡֵ
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
                                                // ���ڹ�ʽ��Ԫ�񣬻�ȡ������ֵ
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
                                                        Trace.WriteLine($"excel��Ԫ��[{i},{j}]��, ��ʽ�������============================");
                                                        MessageBox.Show($"excel��Ԫ��[{i},{j}]��, ��ʽ�������");
                                                        break;
                                                        /*   default:
                                                              dataRow[j] = "=" + iCell.ToString();//��ʽ���󣬰ѹ�ʽ���ַ�������
                                                              break;*/
                                                }
                                                /*   }
                                                   catch (Exception ex)
                                                   {
                                                       Trace.WriteLine("��Ԫ�����{0}", ex.Message);
                                                       Trace.WriteLine("---��Ԫ�� [�У�{0}] [��{1}] ����{2}", iHeaderRow.GetCell(j).ToString(), j.ToString(), ex.Message);
                                                       dataRow[j] = "=" + iCell.CellFormula;
                                                   }*/
                                                break;
                                                /*  default:
                                                      dataRow[j] = iCell.ToString();//string.Empty;
                                                      break; */

                                        }
                                        // Trace.WriteLine($"{iHeaderRow.GetCell(j).ToString()}: {valuesListObject[j]} ����============��");
                                        Trace.WriteLine($"{iHeaderRow.GetCell(j).ToString()}: {valuesListObject[j - (excelBeginningColumn - 1)]} ({iCell.CellType})");
                                    }
                                    mySqlCommand1.Parameters.AddWithValue($"@p{j - (excelBeginningColumn - 1)}", valuesListObject[j - (excelBeginningColumn - 1)] ?? DBNull.Value);
                                }
                                int rowsAffected = mySqlCommand1.ExecuteNonQuery();
                                Trace.WriteLine($"����ɹ���Ӱ������: {rowsAffected}");
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
                    using (MySqlCommand command = new MySqlCommand("SELECT 1 FROM z_vv LIMIT 1", connection)) // ֻ����Ƿ��������һ��
                    {
                        object result = command.ExecuteScalar();
                        if (result == null)
                        {
                            label1.Text = "��";
                            label1.BackColor = Color.Empty;
                            label1.ForeColor = Color.Gray;
                            Trace.WriteLine("z_vv��Ϊ��");
                        }
                        else
                        {
                            label1.Text = "ʵ";
                            label1.BackColor = Color.Black;
                            label1.ForeColor = Color.White;
                            Trace.WriteLine("z_vv��Ϊ��");
                        }
                    }
                    using (MySqlCommand command = new MySqlCommand("SELECT 1 FROM za_vv LIMIT 1", connection)) // ֻ����Ƿ��������һ��
                    {
                        object result = command.ExecuteScalar();
                        if (result == null)
                        {
                            label2.Text = "��";
                            label2.BackColor = Color.Empty;
                            label2.ForeColor = Color.Gray;
                            Trace.WriteLine("za_vv��Ϊ��");
                        }
                        else
                        {
                            label2.Text = "ʵ";
                            label2.BackColor = Color.Black;
                            label2.ForeColor = Color.White;
                            Trace.WriteLine("za_vv��Ϊ��");
                        }
                    }
                    using (MySqlCommand command = new MySqlCommand("SELECT 1 FROM zb_vv LIMIT 1", connection)) // ֻ����Ƿ��������һ��
                    {
                        object result = command.ExecuteScalar();
                        if (result == null)
                        {
                            label3.Text = "��";
                            label3.BackColor = Color.Empty;
                            label3.ForeColor = Color.Gray;
                            Trace.WriteLine("zb_vv��Ϊ��");
                        }
                        else
                        {
                            label3.Text = "ʵ";
                            label3.BackColor = Color.Black;
                            label3.ForeColor = Color.White;
                            Trace.WriteLine("zb_vv��Ϊ��");
                        }
                    }
                    using (MySqlCommand command = new MySqlCommand("SELECT 1 FROM zc_vv LIMIT 1", connection)) // ֻ����Ƿ��������һ��
                    {
                        object result = command.ExecuteScalar();
                        if (result == null)
                        {
                            label4.Text = "��";
                            label4.BackColor = Color.Empty;
                            label4.ForeColor = Color.Gray;
                            Trace.WriteLine("zc_vv��Ϊ��");
                        }
                        else
                        {
                            label4.Text = "ʵ";
                            label4.BackColor = Color.Black;
                            label4.ForeColor = Color.White;
                            Trace.WriteLine("zc_vv��Ϊ��");
                        }
                    }

                    using (MySqlCommand command = new MySqlCommand("SELECT 1 FROM z_yjv LIMIT 1", connection)) // ֻ����Ƿ��������һ��
                    {
                        object result = command.ExecuteScalar();
                        if (result == null)
                        {
                            label5.Text = "��";
                            label5.BackColor = Color.Empty;
                            label5.ForeColor = Color.Gray;
                            Trace.WriteLine("z_yjv��Ϊ��");
                        }
                        else
                        {
                            label5.Text = "ʵ";
                            label5.BackColor = Color.Black;
                            label5.ForeColor = Color.White;
                            Trace.WriteLine("z_yjv��Ϊ��");
                        }
                    }
                    using (MySqlCommand command = new MySqlCommand("SELECT 1 FROM za_yjv LIMIT 1", connection)) // ֻ����Ƿ��������һ��
                    {
                        object result = command.ExecuteScalar();
                        if (result == null)
                        {
                            label6.Text = "��";
                            label6.BackColor = Color.Empty;
                            label6.ForeColor = Color.Gray;
                            Trace.WriteLine("za_yjv��Ϊ��");
                        }
                        else
                        {
                            label6.Text = "ʵ";
                            label6.BackColor = Color.Black;
                            label6.ForeColor = Color.White;
                            Trace.WriteLine("za_yjv��Ϊ��");
                        }
                    }
                    using (MySqlCommand command = new MySqlCommand("SELECT 1 FROM zb_yjv LIMIT 1", connection)) // ֻ����Ƿ��������һ��
                    {
                        object result = command.ExecuteScalar();
                        if (result == null)
                        {
                            label7.Text = "��";
                            label7.BackColor = Color.Empty;
                            label7.ForeColor = Color.Gray;
                            Trace.WriteLine("zb_yjv��Ϊ��");
                        }
                        else
                        {
                            label7.Text = "ʵ";
                            label7.BackColor = Color.Black;
                            label7.ForeColor = Color.White;
                            Trace.WriteLine("zb_yjv��Ϊ��");
                        }
                    }
                    using (MySqlCommand command = new MySqlCommand("SELECT 1 FROM zc_yjv LIMIT 1", connection)) // ֻ����Ƿ��������һ��
                    {
                        object result = command.ExecuteScalar();
                        if (result == null)
                        {
                            label8.Text = "��";
                            label8.BackColor = Color.Empty;
                            label8.ForeColor = Color.Gray;
                            Trace.WriteLine("zc_yjv��Ϊ��");
                        }
                        else
                        {
                            label8.Text = "ʵ";
                            label8.BackColor = Color.Black;
                            label8.ForeColor = Color.White;
                            Trace.WriteLine("zc_yjv��Ϊ��");
                        }
                    }

                    using (MySqlCommand command = new MySqlCommand("SELECT 1 FROM z_yjy LIMIT 1", connection)) // ֻ����Ƿ��������һ��
                    {
                        object result = command.ExecuteScalar();
                        if (result == null)
                        {
                            label9.Text = "��";
                            label9.BackColor = Color.Empty;
                            label9.ForeColor = Color.Gray;
                            Trace.WriteLine("z_yjy ��Ϊ��");
                        }
                        else
                        {
                            label9.Text = "ʵ";
                            label9.BackColor = Color.Black;
                            label9.ForeColor = Color.White;
                            Trace.WriteLine("z_yjy ��Ϊ��");
                        }
                    }
                    using (MySqlCommand command = new MySqlCommand("SELECT 1 FROM za_yjy LIMIT 1", connection)) // ֻ����Ƿ��������һ��
                    {
                        object result = command.ExecuteScalar();
                        if (result == null)
                        {
                            label10.Text = "��";
                            label10.BackColor = Color.Empty;
                            label10.ForeColor = Color.Gray;
                            Trace.WriteLine("za_yjy ��Ϊ��");
                        }
                        else
                        {
                            label10.Text = "ʵ";
                            label10.BackColor = Color.Black;
                            label10.ForeColor = Color.White;
                            Trace.WriteLine("za_yjy ��Ϊ��");
                        }
                    }
                    using (MySqlCommand command = new MySqlCommand("SELECT 1 FROM zb_yjy LIMIT 1", connection)) // ֻ����Ƿ��������һ��
                    {
                        object result = command.ExecuteScalar();
                        if (result == null)
                        {
                            label11.Text = "��";
                            label11.BackColor = Color.Empty;
                            label11.ForeColor = Color.Gray;
                            Trace.WriteLine("zb_yjy ��Ϊ��");
                        }
                        else
                        {
                            label11.Text = "ʵ";
                            label11.BackColor = Color.Black;
                            label11.ForeColor = Color.White;
                            Trace.WriteLine("zb_yjy ��Ϊ��");
                        }
                    }
                    using (MySqlCommand command = new MySqlCommand("SELECT 1 FROM zc_yjy  LIMIT 1", connection)) // ֻ����Ƿ��������һ��
                    {
                        object result = command.ExecuteScalar();
                        if (result == null)
                        {
                            label12.Text = "��";
                            label12.BackColor = Color.Empty;
                            label12.ForeColor = Color.Gray;
                            Trace.WriteLine("zc_yjy ��Ϊ��");
                        }
                        else
                        {
                            label12.Text = "ʵ";
                            label12.BackColor = Color.Black;
                            label12.ForeColor = Color.White;
                            Trace.WriteLine("zc_yjy ��Ϊ��");
                        }
                    }
                }
                catch (Exception ex)
                {
                    Trace.WriteLine($"��������: {ex.Message}");
                }
            }

        }



    }
}
