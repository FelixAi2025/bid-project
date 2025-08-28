//using System;
//using System.Collections.Generic;
//using System.ComponentModel;
//using System.Data;
//using System.Drawing;
//using System.Linq;
//using System.Text;
//using System.Threading.Tasks;
//using System.Windows.Forms;
using HorizontalAlignmentForm = System.Windows.Forms.HorizontalAlignment;
using NPOI.OpenXmlFormats.Wordprocessing;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.WP.UserModel;
using NPOI.XSSF.UserModel;  // 对应xlsx格式
using NPOI.XWPF.UserModel;
using System.Diagnostics;
using System.Text.RegularExpressions;
using ICell = NPOI.SS.UserModel.ICell;
using MatchRegex = System.Text.RegularExpressions.Match;
using MySqlCommand = MySqlConnector.MySqlCommand;
using MySqlConnection = MySqlConnector.MySqlConnection;


namespace MysqlToExcelWord
{
    public partial class FormStart : Form
    {
        internal static string voltageLevel, armourString, coreString;
        internal static string inPutTemplateFullName, inPutTemplatePath;
        internal static string insulationMaterialSelected, bufferMaterialSelected, tapeMaterialSelected, armourMaterialSelected, outer_sheathMaterialSelected;
        internal static double insulationWeightSelected, bufferWeightSelected, tapeWeightSelected;
        internal static bool isDoubleCable, isMultiCore, isArmoured;
        internal static int numChecked;//选中的型号,导入数据时，无需批量清零,InspectCheckBox()中清零
        internal static string flameRedartant, halogenFree, smokeFree, outer_sheathMaterialFront;
        internal static string theType, type_spec, spec, voltage, type_specInFileName;//导入数据时，无需批量清零,取消时清零,打开界面时清零
        internal static string tableForThick;
        internal static string tableNameFromButton, typeInMaterialTable; //导入数据时，无需批量清零,取消时清零,打开界面时清零 tableNameFromButton同库的table名称 如 Z_YJVx2,     typeInMaterialTable如Z-YJVx2
        internal static string conductorMaterial, mica_tapeMaterial, insulation1Material, insulation2Material, buffer1Material, buffer2Material, tape1Material, tape2Material, inner_sheathMaterial, armour1Material, armour2Material, outer_sheath1Material, outer_sheath2Material, production_name, file_code, source;
        internal static object conductorWeight, mica_tapeWeight, insulation1Weight, insulation2Weight, buffer1Weight, buffer2Weight, tape1Weight, tape2Weight, inner_sheathWeight, armourWeight, outer_sheathWeight, cableWeight, current40;
        internal static object inner_thick, steel_thick, sheathThick, insulationThick_1, insulationThick_2, steel_width;//,arm_sheathThick;
        internal static object cableDiameter, conductDiameter_1, conductDiameter_2;
        internal static object resistant20_1, resistant20_2, resistant90_1, resistant90_2;
        internal static object pieces_1, pieces_2;
        internal static List<string> cross_sectional_area = [];
        internal static List<string> areaConductor = [];
        //private internal static List<string> headerListString = [];
        internal static List<CheckBox> checkBoxList; // checkBox数据,导入数据时，无需批量清零
        internal static string[] faultMarkArray = { "?", "？", "错误" };// 导入数据时，无需批量清零
        public FormStart()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e) //低压界面
        {
            button1.Enabled = button2.Enabled = button3.Enabled = false;
            button1.Visible = button2.Visible = button3.Visible = false;
            label1.Visible = false;
            //FormLV formLv = new();
            ////formLv.ShowDialog();//这儿阻塞等待,这样打开，无法监听事件.formLv.Show();则不会阻塞。

            //formLv.Show();
            //formLv.EventCloseAll += (sender, e) => { Trace.WriteLine("formStart  退出"); this.Close(); };
            //formLv.FormClosed += (sender, e) => { Trace.WriteLine("formStart 返回首页"); button1.Enabled = button2.Enabled = button3.Enabled = true; };// 自带X关闭Form1
            //                                                                                                                                       //   formLv.FormClosed += (s, args) => { button1.Enabled = button2.Enabled = button3.Enabled = true; }; 
            //                                                                                                                                       //   formLv.Show();

            CnLv cnLv = new();
            //formLv.ShowDialog();//这儿阻塞等待,这样打开，无法监听事件.formLv.Show();则不会阻塞。

            // cnLv.Show();//
            Controls.Add(cnLv);
            cnLv.controlOutput.EventCloseAll += (sender, e) => { Trace.WriteLine("formStart  退出"); this.Close(); };
            cnLv.controlOutput.EventCloseType += (sender, e) => {
                Trace.WriteLine("formStart 返回首页");
                Controls.Remove(cnLv);
                cnLv.Dispose();
                cnLv = null;
                button1.Enabled = button2.Enabled = button3.Enabled = true;
                button1.Visible = button2.Visible = button3.Visible = true;
                label1.Visible = true;
            };// 自带X关闭Form1

            cnLv.controlOutput.EventTestClose += (sender, e) => { Trace.WriteLine("formStart  退出"); this.Close(); };
        }

        private void button2_Click(object sender, EventArgs e) //中压界面
        {

            button1.Enabled  = button3.Enabled = false;
        }

        private void button3_Click(object sender, EventArgs e)//高压界面
        {
            button1.Enabled = button2.Enabled  = false;

        }

        public static void getMysqlData()
        {
            cleanData();
            //调用数据
            preTreatData();

            getMysqlQuote();
            getMysqlThick();
            getMysqlMaterialAll();
            getMysqlArea();

            treatData();

            void cleanData()
            {

                voltageLevel = armourString = coreString = "";
                inPutTemplateFullName = inPutTemplatePath = "";
                insulationMaterialSelected = bufferMaterialSelected = tapeMaterialSelected = armourMaterialSelected = outer_sheathMaterialSelected = "";
                insulationWeightSelected = bufferWeightSelected = tapeWeightSelected = 0.0;
                isDoubleCable = isMultiCore = isArmoured = false;
                flameRedartant = halogenFree = smokeFree = outer_sheathMaterialFront = "";
                tableForThick = "";
                conductorMaterial = mica_tapeMaterial = insulation1Material = insulation2Material = buffer1Material = buffer2Material = tape1Material = tape2Material = inner_sheathMaterial = armour1Material = armour2Material = outer_sheath1Material = outer_sheath2Material = production_name = file_code = source = "";
                conductorWeight = mica_tapeWeight = insulation1Weight = insulation2Weight = buffer1Weight = buffer2Weight = tape1Weight = tape2Weight = inner_sheathWeight = armourWeight = outer_sheathWeight = cableWeight = current40 = "?清零状态";// null;
                inner_thick = steel_thick = sheathThick = insulationThick_1 = insulationThick_2 = steel_width = "?清零状态";//null;  //  =arm_sheathThick=null;
                cableDiameter = conductDiameter_1 = conductDiameter_2 = "?清零状态";//null;
                resistant20_1 = resistant20_2 = resistant90_1 = resistant90_2 = "?清零状态";//null;
                pieces_1 = pieces_2 = "?清零状态";// null;
                cross_sectional_area.Clear();
                areaConductor.Clear();
                // headerListString.Clear();

            }
            //方法
            void preTreatData()
            {
                string[] findString = { "＋", "+" };
                isDoubleCable = findString.Any(phrase => spec.Contains(phrase, StringComparison.OrdinalIgnoreCase));

                MatchCollection matchCollectionFront = Regex.Matches(spec, @"(×\d+\.*\d*)|(X\d+\.*\d*)|(x\d+\.*\d*)");


                int i = 0;
                foreach (MatchRegex match1 in matchCollectionFront)
                {
                    Trace.WriteLine($"match1 {i}：{match1.Value}");
                    MatchRegex theMatch = Regex.Match(match1.Value, @"\d+\.*\d*");
                    Trace.WriteLine($"theMatch{i}：{theMatch.Value}");
                    cross_sectional_area.Add(theMatch.Value + "平方");
                    Trace.WriteLine($"cross_sectional_area{i}：{cross_sectional_area[i]}");
                    areaConductor.Add($"对应{theMatch.Value} mm²截面");
                    Trace.WriteLine($"areaConductor{i}：{areaConductor[i++]}");
                }
                if (cross_sectional_area.Count == 1) cross_sectional_area.Add("? mm²平方");//20250825
                if (areaConductor.Count == 1) areaConductor.Add("? mm²截面");//20250825
            }

            void getMysqlQuote() //定额
                // public async Task getMysqlQuote()//异步连接
            {
                try
                {
                    using var mysqlConnection = new MySqlConnection("server=localhost;user=root;password=8888;database=cableLv");
                  //  await mysqlConnection.OpenAsync(); //异步连接
                    mysqlConnection.Open();

                    Trace.WriteLine($"\n成功连接到MySQL数据库, 基础表格名称：{tableNameFromButton}");
                    string queryString = $"SELECT    COALESCE(steel_width,'?未提供') As  steel_width, " +
                          $"COALESCE(sheath_Thick,'?未提供') As  sheath_Thick," +  //
                          $"COALESCE(diameter,'?未提供') As  diameter," +
                          $"COALESCE(conductor,'?未提供') As  conductor," +
                          $"COALESCE(mica_tape,'?未提供') As  mica_tape, " +
                          $"COALESCE(insulation1,'?未提供') As  insulation1, " +
                          $"COALESCE(insulation2,'?未提供') As  insulation2, " +
                          $"COALESCE(buffer1,'?未提供') As  buffer1, " +
                          $"COALESCE(buffer2,'?未提供') As  buffer2, " +
                          $"COALESCE(tape1,'?未提供') As  tape1, " +
                          $"COALESCE(tape2,'?未提供') As  tape2, " +
                          $"COALESCE(inner_sheath,'?未提供') As  inner_sheath," +
                          $"COALESCE(armour,'?未提供') As  armour, " +
                          $"COALESCE(outer_sheath,'?未提供') As  outer_sheath, " +
                          $"COALESCE(weight,'?未提供') As  weight," +
                          $"COALESCE(current,'?未提供')As  current " +
                          $"FROM {tableNameFromButton} WHERE spec = @spec";
                    /*
                       $"COALESCE(current,'?未提供') As  current" +
                    $"FROM {tableNameFromButton} WHERE spec = @spec";
                    */
                   
                    using var mysqlCommand = new MySqlCommand(queryString, mysqlConnection);   
                    mysqlCommand.Parameters.AddWithValue("@spec", spec);

                    // using var mysqlReader = await mysqlCommand.ExecuteReaderAsync(); //异步连接
                    using var mysqlReader = mysqlCommand.ExecuteReader();
                    while (mysqlReader.Read())
                    //while (await mysqlReader.ReadAsync()) //异步连接
                    {
                        steel_width = mysqlReader["steel_width"]; Trace.Write($"steel_width: {steel_width}   ");
                        sheathThick = mysqlReader["sheath_Thick"]; Trace.Write($"sheathThick: {sheathThick}    ");
                        cableDiameter = mysqlReader["diameter"]; Trace.Write($"cableDiameter: {cableDiameter}  ");
                        conductorWeight = mysqlReader["conductor"]; Trace.Write($"conductorWeight: {conductorWeight}  ");
                        mica_tapeWeight = mysqlReader["mica_tape"]; Trace.Write($"mica_tapeWeight: {mica_tapeWeight}");
                        // silane_insulationWeight = mysqlReader["silane_insulation"]; Trace.Write($"silane_insulationWeight: {silane_insulationWeight}");
                        insulation1Weight = mysqlReader["insulation1"]; Trace.Write($"insulation1Weight: {insulation1Weight}");
                        //UV_insulationWeight = mysqlReader["UV_insulation"]; Trace.Write($"UV_insulationWeight: {UV_insulationWeight}");
                        insulation2Weight = mysqlReader["insulation2"]; Trace.Write($"insulation2Weight: {insulation2Weight}");
                        // PP_bufferWeight = mysqlReader["PP_buffer"]; Trace.Write($"PP_bufferWeight: {PP_bufferWeight}");
                        buffer1Weight = mysqlReader["buffer1"]; Trace.Write($"buffer1Weight: {buffer1Weight}");
                        // rockwool_bufferWeight = mysqlReader["rockwool_buffer"]; Trace.Write($"rockwool_bufferWeight: {rockwool_bufferWeight}");
                        buffer2Weight = mysqlReader["buffer2"]; Trace.Write($"buffer2Weight: {buffer2Weight}");
                        tape1Weight = mysqlReader["tape1"]; Trace.Write($"tape1Weight: {tape1Weight}");
                        tape2Weight = mysqlReader["tape2"]; Trace.Write($"tape2Weight: {tape2Weight}");
                        inner_sheathWeight = mysqlReader["inner_sheath"]; Trace.Write($"inner_sheathWeight: {inner_sheathWeight}");
                        armourWeight = mysqlReader["armour"]; Trace.Write($"armourWeight: {armourWeight}");
                        outer_sheathWeight = mysqlReader["outer_sheath"]; Trace.Write($"outer_sheathWeight: {outer_sheathWeight}");
                        cableWeight = mysqlReader["weight"]; Trace.Write($"cableWeight: {cableWeight}");
                        current40 = mysqlReader["current"]; Trace.Write($"current40: {current40} ");
                        /* if (double.TryParse(mysqlReader["armour"].ToString(), out double armourWeight)) Trace.Write($"armourWeight: {armourWeight} "); else Trace.Write($"armourWeight: lost ");
                         if (double.TryParse(mysqlReader["inner_sheath"].ToString(), out double inner_sheathWeight)) Trace.Write($"armourWeight: {inner_sheathWeight} "); else Trace.Write($"inner_sheathWeight: lost ");
                         if (double.TryParse(mysqlReader["weight"].ToString(), out double cableWeight)) Trace.Write($"cableWeight: {cableWeight} "); else Trace.Write($"cableWeight: lost ");
                         if (double.TryParse(mysqlReader["steel_width"].ToString(), out double steel_width)) Trace.Write($"steel_width: {steel_width} "); else Trace.Write($"steel_width: lost ");
                         if (double.TryParse(mysqlReader["diameter"].ToString(), out double cableDiameter)) Trace.Write($"cableDiameter: {cableDiameter} "); else Trace.Write($"cableDiameter: lost ");
                         if (double.TryParse(mysqlReader["sheath_Thick"].ToString(), out double sheathThick)) Trace.Write($"sheathThick: {sheathThick} "); else Trace.Write($"sheathThick: lost ");
                        */
                    }
                    Trace.WriteLine($"steel_width: {steel_width}   sheathThick: {sheathThick}   cableDiameter: {cableDiameter}  conductorWeight: {conductorWeight}\n" +
                                    $"mica_tapeWeight: {mica_tapeWeight}   insulation1Weight: {insulation1Weight}   insulation2Weight: {insulation2Weight}  buffer1Weight: {buffer1Weight}\n" +
                                    $"buffer2Weight: {buffer2Weight}   tape1Weight: {tape1Weight}   tape2Weight: {tape2Weight}  inner_sheathWeight: {inner_sheathWeight}  armourWeight: {armourWeight}" +
                                    $"outer_sheathWeight: {outer_sheathWeight}   cableWeight: {cableWeight}   current40: {current40}");
                    Trace.WriteLine($"数据表{tableNameFromButton}查询完成");
                }
                catch (Exception ex)
                {
                    Trace.WriteLine($"getMysqlQuote 对 数据表{tableNameFromButton}查询发生错误: {ex.Message}");
                }
            }
            void getMysqlThick() // 厚度
            {
                //tableForThick = Regex.Match(type_spec, @"[^\d\s]+").Value.Replace(@"\s-", "_") + "_thick";
                tableForThick = Regex.Replace(Regex.Match(type_spec, @"[^\d\s]+").Value, @"([aAbBCc]-)|-", "_") + "_thick";
                try
                {
                    using var mysqlConnectionString = new MySqlConnection("server=localhost;user=root;password=8888;database=cableLv");
                    mysqlConnectionString.Open();

                    Trace.WriteLine($"\n成功连接到MySQL数据库, 厚度表格名称：{tableForThick}");
                    string queryString = $"SELECT COALESCE(inner_thick,'?未提供') As  inner_thick, " +
                                         $"COALESCE(steel_thick,'?未提供')As  steel_thick  " +
                                         $"FROM {tableForThick} WHERE spec = @spec";
                    using var mysqlCommand = new MySqlCommand(queryString, mysqlConnectionString);
                    mysqlCommand.Parameters.AddWithValue("@spec", spec);

                    using var mysqlReader = mysqlCommand.ExecuteReader();
                    while (mysqlReader.Read())
                    {
                        inner_thick = mysqlReader["inner_thick"];
                        steel_thick = mysqlReader["steel_thick"];
                        // if (double.TryParse(mysqlReader["inner_thick"].ToString(), out double inner_thick)) Trace.Write($"inner_thick: {inner_thick} "); else Trace.Write($"inner_thick: lost ");
                        // if (double.TryParse(mysqlReader["steel_thick"].ToString(), out double steel_thick)) Trace.Write($"steel_thick: {steel_thick} "); else Trace.Write($"steel_thick: lost ");
                    }
                    Trace.WriteLine($"inner_thick: {inner_thick}  steel_thick: {steel_thick}");
                    Trace.WriteLine($"数据表{tableForThick}查询完成");
                }
                catch (Exception ex)
                {
                    Trace.WriteLine($"getMysqlThick 对 数据表{tableForThick}查询发生错误: {ex.Message}");
                }
            }
            void getMysqlMaterialAll() // 材料统一新表格
            {
                // string tableMeterial = tableNameFromButton + "";
                string tableMeterial = Regex.Replace(Regex.Match(type_spec, @"[^\d\s]+").Value, @"([aAbBCc]-)|-", "_") + "_m";
                Trace.WriteLine($"tableMeterial: {tableMeterial}");
                try
                {
                    using var mysqlConnectionString = new MySqlConnection("server=localhost;user=root;password=8888;database=cableLv");
                    mysqlConnectionString.Open();

                    Trace.WriteLine($"\n成功连接到MySQL数据库,材料表格名称：{tableMeterial}");
                    string queryString = $"SELECT COALESCE(conductor,'?未提供') As  conductor," +
                         $"COALESCE(mica_tape,'?未提供') As  mica_tape," +
                         $"COALESCE(insulation1,'?未提供') As  insulation1," +
                         $"COALESCE(insulation2,'?未提供') As  insulation2," +
                         $"COALESCE(buffer1,'?未提供') As  buffer1," +
                         $"COALESCE(buffer2,'?未提供') As  buffer2," +
                         $"COALESCE(tape1,'?未提供') As  tape1," +
                         $"COALESCE(tape2,'?未提供') As  tape2," +
                         $"COALESCE(inner_sheath,'?未提供') As  inner_sheath," +
                         $"COALESCE(armour1,'?未提供') As  armour1," +
                         $"COALESCE(armour2,'?未提供') As  armour2," +
                         $"COALESCE(outer_sheath1,'?未提供') As  outer_sheath1,  " +
                         $"COALESCE(outer_sheath2,'?未提供') As  outer_sheath2,  " +
                         $"COALESCE(production_name,'?未提供') As  production_name,  " +
                         $"COALESCE(file_code,'?未提供') As  file_code,  " +
                         $"COALESCE(source,'?未提供') As  source  " +
                         $"FROM {tableMeterial} WHERE type = @typeInMaterialTable";

                    using var mysqlCommand = new MySqlCommand(queryString, mysqlConnectionString);
                    //mysqlCommand.Parameters.AddWithValue("@type", theType);
                    mysqlCommand.Parameters.AddWithValue("@typeInMaterialTable", typeInMaterialTable);
                    Trace.WriteLine($"typeInMaterialTable：{typeInMaterialTable}");
                    using var mysqlReader = mysqlCommand.ExecuteReader();

                    while (mysqlReader.Read())
                    {
                        conductorMaterial = mysqlReader.GetString("conductor");
                        mica_tapeMaterial = mysqlReader.GetString("mica_tape");
                        insulation1Material = mysqlReader.GetString("insulation1");
                        insulation2Material = mysqlReader.GetString("insulation2");
                        buffer1Material = mysqlReader.GetString("buffer1");
                        buffer2Material = mysqlReader.GetString("buffer2");
                        tape1Material = mysqlReader.GetString("tape1");
                        tape2Material = mysqlReader.GetString("tape2");
                        inner_sheathMaterial = mysqlReader.GetString("inner_sheath");
                        armour1Material = mysqlReader.GetString("armour1");
                        armour2Material = mysqlReader.GetString("armour2");
                        outer_sheath1Material = mysqlReader.GetString("outer_sheath1");
                        outer_sheath2Material = mysqlReader.GetString("outer_sheath2");
                        production_name = mysqlReader.GetString("production_name");
                        file_code = mysqlReader.GetString("file_code");
                        source = mysqlReader.GetString("source");
                        Trace.WriteLine("新材料");
                    }
                    Trace.WriteLine($"conductorMaterial: {conductorMaterial}   mica_tapeMaterial: {mica_tapeMaterial}   insulation1Material: {insulation1Material}  insulation2Material: {insulation2Material}\n" +
                        $"buffer1Material: {buffer1Material}   buffer2Material: {buffer2Material}   tape1Material: {tape1Material}  tape2Material: {tape2Material}\n" +
                        $"inner_sheathMaterial: {inner_sheathMaterial}   armour1Material: {armour1Material}   armour2Material: {armour2Material}  outer_sheath1Material: {outer_sheath1Material}  outer_sheath2Material: {outer_sheath2Material}" +
                        $"production_name: {production_name}   file_code: {file_code}   source: {source}");
                    Trace.WriteLine($"数据表{tableMeterial}查询完成");
                }
                catch (Exception ex)
                {
                    Trace.WriteLine($"getMysqlMaterial 对 数据表{tableMeterial}查询发生错误: {ex.Message}");
                }
            }
            void getMysqlArea() // 面积相关
            {
                try
                {
                    using var mysqlConnectionString = new MySqlConnection("server=localhost;user=root;password=8888;database=cableLv");
                    mysqlConnectionString.Open();

                    Trace.WriteLine("\n成功连接到MySQL数据库, 面积表格名称：area");
                    //string queryString1 = $"SELECT  copper_pieces FROM area WHERE spec = @spec";
                    string queryString1 = $"SELECT COALESCE(insulation_thickness,'?未提供') As  insulation_thickness," +
                             $"COALESCE(copper_diameter,'?未提供') As  copper_diameter," +
                             $"COALESCE(copper_pieces,'?未提供') As  copper_pieces," +
                             $"COALESCE(copTem20resist,'?未提供') As  copTem20resist,  " +
                             $"COALESCE(copTem90resist,'?未提供') As  copTem90resist  " +
                             $"FROM area WHERE spec = @spec";

                    using var mysqlCommand1 = new MySqlCommand(queryString1, mysqlConnectionString);
                    mysqlCommand1.Parameters.AddWithValue("@spec", cross_sectional_area[0]);//"4平方"
                    Trace.WriteLine($"cross_sectional_area :   {cross_sectional_area[0]}");

                    using var mysqlReader = mysqlCommand1.ExecuteReader();
                    while (mysqlReader.Read())
                    {
                        insulationThick_1 = mysqlReader["insulation_thickness"];
                        conductDiameter_1 = mysqlReader["copper_diameter"];
                        pieces_1 = mysqlReader["copper_pieces"];
                        resistant20_1 = mysqlReader["copTem20resist"];
                        resistant90_1 = mysqlReader["copTem90resist"];
                        // if (double.TryParse(mysqlReader["insulation_thickness"].ToString(), out double insulationThick_1)) Trace.Write($"insulationThick_1: {insulationThick_1} "); else Trace.Write($"insulationThick_1: lost ");
                        // if (double.TryParse(mysqlReader["copper_diameter"].ToString(), out double conductDiameter_1)) Trace.Write($"conductDiameter_1: {conductDiameter_1} "); else Trace.Write($"conductDiameter_1: lost ");
                        //if (int.TryParse(mysqlReader["copper_pieces"].ToString(), out int pieces_1)) Trace.Write($"pieces_1: {pieces_1} "); else { Trace.Write($"pieces_1: lost "); pieces_1 = mysqlReader["copper_pieces"].ToString(); }
                        //if (int.TryParse(mysqlReader["copper_pieces"].ToString(), out int pieces_1)) Trace.Write($"pieces_1: {pieces_1}"); else { pieces_1 = "?未提供";  Trace.Write($"pieces_1: ?未提供----{pieces_1} ");  }

                        //     if (double.TryParse(mysqlReader["copTem20resist"].ToString(), out double resistant20_1)) Trace.Write($"resistant20_1: {resistant20_1} "); else Trace.Write($"resistant20_1: lost ");
                        //  if (double.TryParse(mysqlReader["copTem90resist"].ToString(), out double resistant90_1)) Trace.Write($"resistant90_1: {resistant90_1} "); else Trace.Write($"resistant90_1: lost ");
                    }
                    Trace.WriteLine($"insulationThick_1: {insulationThick_1}  conductDiameter_1: {conductDiameter_1}   pieces_1: {pieces_1}  resistant20_1: {resistant20_1}  resistant90_1: {resistant90_1}");
                    Trace.WriteLine($"数据表area查询 1 完成");
                }
                catch (Exception ex)
                {
                    Trace.WriteLine($"getMysqlArea对 数据表 area 查询 1 发生错误: {ex.Message}");
                }

                if (isDoubleCable)
                {
                    try
                    {
                        using var mysqlConnectionString2 = new MySqlConnection("server=localhost;user=root;password=8888;database=cableLv");
                        mysqlConnectionString2.Open();

                        Trace.WriteLine("\n成功连接到MySQL数据库, 面积表格名称：area");
                        string queryString2 = $"SELECT COALESCE(insulation_thickness,'?未提供') As  insulation_thickness," +
                             $"COALESCE(copper_diameter,'?未提供') As  copper_diameter," +
                             $"COALESCE(copper_pieces,'?未提供') As  copper_pieces," +
                             $"COALESCE(copTem20resist,'?未提供') As  copTem20resist,  " +
                             $"COALESCE(copTem90resist,'?未提供') As  copTem90resist  " +
                             $"FROM area WHERE spec = @spec";

                        using var mysqlCommand2 = new MySqlCommand(queryString2, mysqlConnectionString2);
                        mysqlCommand2.Parameters.AddWithValue("@spec", cross_sectional_area[1]);
                        Trace.WriteLine($"cross_sectional_area :   {cross_sectional_area[1]}");
                        using var mysqlReader2 = mysqlCommand2.ExecuteReader();
                        while (mysqlReader2.Read())
                        {
                            insulationThick_2 = mysqlReader2["insulation_thickness"]; Trace.Write($"insulationThick_2: {insulationThick_2}  ");
                            conductDiameter_2 = mysqlReader2["copper_diameter"]; Trace.Write($"conductDiameter_2: {conductDiameter_2}   ");
                            pieces_2 = mysqlReader2["copper_pieces"]; Trace.Write($"pieces_2: {pieces_2}    ");
                            resistant20_2 = mysqlReader2["copTem20resist"]; Trace.Write($"resistant20_2: {resistant20_2}    ");
                            resistant90_2 = mysqlReader2["copTem90resist"]; Trace.Write($"resistant90_2: {resistant90_2}     ");
                            /*   if (double.TryParse(mysqlReader2["insulation_thickness"].ToString(), out double insulationThick_2)) Trace.Write($"insulationThick_2: {insulationThick_2} "); else Trace.Write($"insulationThick_2: lost ");
                               if (double.TryParse(mysqlReader2["copper_diameter"].ToString(), out double conductDiameter_2)) Trace.Write($"conductDiameter_2: {conductDiameter_2} "); else Trace.Write($"conductDiameter_2: lost ");
                               if (int.TryParse(mysqlReader2["copper_pieces"].ToString(), out int pieces_222)) { pieces_2 = pieces_222;  Trace.Write($"pieces_2: {pieces_2} "); } else Trace.Write($"pieces_2: lost ");
                               if (double.TryParse(mysqlReader2["copTem20resist"].ToString(), out double resistant20_2)) Trace.Write($"resistant20_2: {resistant20_2} "); else Trace.Write($"resistant20_2: lost ");
                               if (double.TryParse(mysqlReader2["copTem90resist"].ToString(), out double resistant90_2)) Trace.Write($"resistant90_2: {resistant90_2} "); else Trace.Write($"resistant90_2: lost ");
                           */
                        }
                        Trace.WriteLine($"insulationThick_2: {insulationThick_2}  conductDiameter_2: {conductDiameter_2}   pieces_2: {pieces_2}  resistant20_2: {resistant20_2}  resistant90_2: {resistant90_2}");
                        Trace.WriteLine($"\n 数据表area查询 2 完成");
                    }
                    catch (Exception ex)
                    {
                        Trace.WriteLine($"getMysqlArea对 数据表 area 查询 2 发生错误: {ex.Message}");
                    }
                }
                else
                {
                    insulationThick_2 = 0.0;
                    conductDiameter_2 = 0.0;
                    pieces_2 = 0;
                    resistant20_2 = 0.0;
                    resistant90_2 = 0.0;
                    Trace.WriteLine($"insulationThick_2: {insulationThick_2}  conductDiameter_2: {conductDiameter_2}   pieces_2: {pieces_2}  resistant20_2: {resistant20_2}  resistant90_2: {resistant90_2}");
                }

            }

            void treatData()
            {
                double insulation1WeightDouble = Convert.ToDouble(insulation1Weight);
                double insulation2WeightDouble = Convert.ToDouble(insulation2Weight);

                if (insulation1WeightDouble > 1 && insulation2WeightDouble > 1) { Trace.WriteLine("问题 insulationWeight 同时有读数"); }
                else if (insulation1WeightDouble > 1) { insulationMaterialSelected = insulation1Material; insulationWeightSelected = insulation1WeightDouble; }
                else if (insulation2WeightDouble > 1) { insulationMaterialSelected = insulation2Material; insulationWeightSelected = insulation2WeightDouble; }
                else { Trace.WriteLine("问题 insulationWeight 都为0"); }

                double buffer1WeightDouble = Convert.ToDouble(buffer1Weight);
                double buffer2WeightDouble = Convert.ToDouble(buffer2Weight);
                if (buffer1WeightDouble > 1 && buffer2WeightDouble > 1) { Trace.WriteLine("问题 bufferWeight 同时有读数"); }
                else if (buffer1WeightDouble > 1) { bufferMaterialSelected = buffer1Material; bufferWeightSelected = buffer1WeightDouble; Trace.WriteLine($"bufferMaterialSelected: {bufferMaterialSelected}"); }
                else if (buffer2WeightDouble > 1) { bufferMaterialSelected = buffer2Material; bufferWeightSelected = buffer2WeightDouble; Trace.WriteLine($"bufferMaterialSelected: {bufferMaterialSelected}"); }
                else { bufferMaterialSelected = "无"; Trace.WriteLine("问题 bufferWeight 都为0"); }

                double tape1WeightDouble = Convert.ToDouble(tape1Weight);
                double tape2WeightDouble = Convert.ToDouble(tape2Weight);
                if (tape1WeightDouble > 1 && tape2WeightDouble > 1) { Trace.WriteLine("问题 tapeWeight 同时有读数"); }
                else if (tape1WeightDouble > 1) { tapeMaterialSelected = tape1Material; tapeWeightSelected = tape1WeightDouble; }
                else if (tape2WeightDouble > 1) { tapeMaterialSelected = tape2Material; tapeWeightSelected = tape2WeightDouble; }
                else { tapeMaterialSelected = "无"; Trace.WriteLine("问题 tapeWeight 都为0"); }

                armourMaterialSelected = (isDoubleCable || isMultiCore) ? armour2Material : armour1Material;
                outer_sheathMaterialSelected = (isDoubleCable || isMultiCore) ? outer_sheath2Material : outer_sheath1Material;
                Trace.WriteLine($"outer_sheath2Material: {outer_sheath2Material}");
                Trace.WriteLine($"outer_sheathMaterialSelected: {outer_sheathMaterialSelected}");




                //forhead = Regex.Match(theType, @"(\w+)\-").Groups[1].Value;
                string forhead = Regex.Match(theType, @"(\w+)\-").Groups[1].Value;
                flameRedartant = Regex.Replace(forhead, @"w|d|W|D", "");
                if (flameRedartant == "") flameRedartant = "不适用";
                Trace.WriteLine($"flameRedartant: {flameRedartant}   ");

                halogenFree = (Regex.Match(forhead, @"(w|W)").Value);
                if (halogenFree == "") halogenFree = "不适用";
                Trace.WriteLine($"halogenFree: {halogenFree}");
                smokeFree = Regex.Match(forhead, @"(d|D)").Value;
                if (smokeFree == "") smokeFree = "不适用";
                Trace.WriteLine($"smokingFree: {smokeFree}");
                /*
                                MatchCollection matchCollection = Regex.Matches(spec, @"(\d×)|(\dX)|(\dx)");
                                foreach (MatchRegex match1 in matchCollection)
                                {
                                    MatchRegex theMatch = Regex.Match(match1.Value, @"\d");
                                    //   Trace.WriteLine($"数字：{theMatch.Value}");
                                    isMultiCore = (Convert.ToInt16(theMatch.Value) > 1);
                                    Trace.WriteLine($"theMatch.Value: {Convert.ToInt16(theMatch.Value)}");
                                    // specMini.Add(match1.Value);
                                }*/
                Match match = Regex.Match(spec, @"(\d)[×Xx]"); // Simplified regex with single digit group
                if (match.Success) isMultiCore = Convert.ToInt16(match.Groups[1].Value) > 1;

                outer_sheathMaterialFront = Regex.Match(outer_sheathMaterialSelected, @"(.+护套料)").Groups[1].Value;
                Trace.WriteLine($"outer_sheathMaterialFront: {outer_sheathMaterialFront}");

                /*     MatchCollection matchCollectionRear = Regex.Matches(spec, @"(×\d+)|(X\d+)|(x\d+)");
                     //int k = 0;
                     List<string> areaConductor = [];
                     foreach (MatchRegex match1 in matchCollectionRear)
                     {
                         MatchRegex theMatch = Regex.Match(match1.Value, @"\d+");
                         areaConductor.Add($"对应{theMatch.Value} mm²截面");
                         //Trace.WriteLine($"{areaConductor[k++]}");
                     }
                     */
                isArmoured = (Convert.ToDouble(armourWeight) >= 1);
                if (type_spec.Contains("0.6/1")) voltageLevel = "低压"; else voltageLevel = "高压？中压？低压？";
                if (isArmoured) armourString = "铠装"; else armourString = "无铠装";
                if (isDoubleCable) coreString = "双缆"; else if (isMultiCore) coreString = "多芯"; else coreString = "单芯";
            }

        }

        public class CreateExcel
        {
            #region field
            // 1 创建工作簿
            XSSFWorkbook workbook;
            // 2 创建工作表
            ISheet sheetQuote, sheetConstructionPara, sheetNonElectricPara, sheetMaterialConfiguration, sheetEnvironment;
            ICellStyle titleStyle, titleBlankBorderStyle;
            ICellStyle boolStyle, dateStyle, dateTimeStyle, numberStyle, warnNumStyle, warnStyle;
            ICellStyle stringCenterStyle, stringLeftStyle, stringBlueStyle;


            #endregion

            public CreateExcel()
            {
                workbook = new XSSFWorkbook();
                sheetQuote = workbook.CreateSheet("营销报价");
                sheetConstructionPara = workbook.CreateSheet("结构参数");
                sheetNonElectricPara = workbook.CreateSheet("非电气参数");
                sheetMaterialConfiguration = workbook.CreateSheet("材料配置");
                sheetEnvironment = workbook.CreateSheet("环境条件");
                titleStyle = CellStyleSet(4);//（1 背景无框 2 普通单元格 3 报警 4 标题）
                titleBlankBorderStyle = CellStyleSet(11);//（1 背景无框 2 普通单元格 3 报警 4 标题）
                boolStyle = CellStyleSet(7);//（1 背景无框 2 普通单元格 3 报警 4 标题,5 字体中间，7 boolStyle）
                dateStyle = CellStyleSet(8);//8 dateTime"yyyy-MM-dd"
                dateTimeStyle = CellStyleSet(9);//9 dateTime "yyyy-MM-dd HH:mm:ss")
                numberStyle = CellStyleSet(5);//5 number "#,##0.00"
                warnNumStyle = CellStyleSet(6);//6 红字报警 number "#,##0.00"
                warnStyle = CellStyleSet(3);//（1 背景无框 2 普通单元格 3 报警 4 标题）
                stringCenterStyle = CellStyleSet(2); //（1 背景无框 2 普通单元格 3 报警 4 标题,5 字体中间，7 boolStyle）
                stringLeftStyle = CellStyleSet(12); //（1 背景无框 2 普通单元格 3 报警 4 标题,5 字体中间，7 boolStyle）
                stringBlueStyle = CellStyleSet(10); //（1 背景无框 2 普通单元格 3 报警 4 标题,5 字体中间，7 boolStyle）
            }
            public void ToCreateExcel()
            {

                //A --------------------------Excel营销报价------------------------------------
                createSheetQuote();

                //B -------------------------- Excel结构参数表----------------------------------
                createSheetConstructionPara();

                //C --------------------------Excel非电气参数------------------------------------
                createSheetNonElectricPara();

                //D --------------------------Excel组件材料配置表------------------------------------
                createSheetMaterialconfiguration();

                //E --------------------------Excel使用环境条件表------------------------------------
                createSheetEnvironment();

                // 4. 保存文件（FileStream）含对话框铠装层
                saveExcel(workbook, $"报价定额及工艺参数{type_specInFileName} ");

            }

            void SetSheetDefault() //默认背景和文字颜色
            {    //默认列背景和文字颜色
                 // 已注释。读取或创建工作簿。类中已有，不需要重复
                #region
                /*    XSSFWorkbook workbook;
                 if (File.Exists(filePath))
                 {
                     using (FileStream fs = new FileStream(filePath, FileMode.Open, FileAccess.Read))
                     {
                         workbook = new XSSFWorkbook(fs);
                     }
                 }
                 else
                 {
                     workbook = new XSSFWorkbook();
                 }


                XSSFWorkbook workbook = new XSSFWorkbook(); ;
                XSSFSheet sheetConstructionPara = workbook.CreateSheet(sheetName) as XSSFSheet;
                 * */
                #endregion


                ICellStyle styleBackground = CellStyleSet(1);//（1 背景无框 2 普通单元格 3 报警 4 标题）

                // 设置列样式
                for (int i = 0; i < 70; i++)
                {
                    //sheetConstructionPara.SetDefaultColumnStyle(（列索引-最多16383，ICellStyle）
                    sheetConstructionPara.SetDefaultColumnStyle(i, styleBackground);
                }

            }

            ICellStyle CellStyleSet(short stypleSelector)//单元格颜色和边框设置 
            {
                ICellStyle style = workbook.CreateCellStyle();//默认格式为 常规 或General
                IFont fontInner = workbook.CreateFont(); //文字样式      

                Action styleDefault = () =>
                {
                    style.BorderTop = style.BorderBottom = style.BorderLeft = style.BorderRight = NPOI.SS.UserModel.BorderStyle.Thin;//边框 BorderStyle.Double
                                                                                                                                     //style.TopBorderColor = style.BottomBorderColor = style.LeftBorderColor = style.RightBorderColor = IndexedColors.Blue.Index;//边框颜色
                    style.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Center;
                };
                style.WrapText = true;  // 自动换行，文字超出单元格范围，行高自动增加。合并单元格，行高不会自动变化
                fontInner.Color = IndexedColors.Black.Index;//文字颜色  
                                                            //style.FillForegroundColor = IndexedColors.Black.Index; // 填充色 
                fontInner.FontName = "宋体";//"微软雅黑";//"Arial"
                fontInner.FontHeightInPoints = 9;
                //  style.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Center;
                style.VerticalAlignment = NPOI.SS.UserModel.VerticalAlignment.Center;
                /*
           XSSFCellStyle defaultStyle = (XSSFCellStyle)workbook.CreateCellStyle();
           XSSFFont font = workbook.CreateFont() as XSSFFont;
           font.Color = IndexedColors.White.Index;
           font.SetColor(new XSSFColor(new byte[] {255, 255, 255 }));
           defaultStyle.SetFont(font);

           defaultStyle.FillForegroundColor = IndexedColors.Black.Index;
           defaultStyle.SetFillForegroundColor(new XSSFColor(new byte[] { 0, 0, 0 }));

           // 设置填充模式（必须设置，否则颜色不会显示）
           defaultStyle.FillPattern = FillPattern.SolidForeground;
                */

                switch (stypleSelector)
                {
                    case 1: // 背景无框
                        break;
                    // break;
                    case 2: //普通单元格
                        styleDefault();
                        break;
                    case 3: // 报警
                        fontInner.Color = IndexedColors.Red.Index;
                        styleDefault();
                        break;
                    case 4: //标题
                        fontInner.FontName = "黑体";//"Arial"
                        fontInner.FontHeightInPoints = 11;
                        //fontInner.IsBold = true;// 粗体
                        //fontInner.Color = IndexedColors.Black.Index;
                        //     style.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Center;
                        //     style.VerticalAlignment = NPOI.SS.UserModel.VerticalAlignment.Center;
                        //fontInner.Color = IndexedColors.SkyBlue.Index;
                        //style.FillForegroundColor = IndexedColors.Black.Index; // IndexedColors.Black.Index;     
                        styleDefault();
                        break;
                    case 5: //number "#,##0.00"
                        fontInner.FontName = "微软雅黑";//"Arial"
                        styleDefault();
                        style.DataFormat = workbook.CreateDataFormat().GetFormat("#,##0.00");
                        break;
                    case 6: //红字报警 number "#,##0.00"
                        fontInner.Color = IndexedColors.Red.Index;
                        styleDefault();
                        style.DataFormat = workbook.CreateDataFormat().GetFormat("#,##0.00");
                        break;
                    case 7: // boolStyle
                        styleDefault();
                        style.DataFormat = workbook.CreateDataFormat().GetFormat("TRUE;FALSE");
                        break;
                    case 8: //dateTime"yyyy-MM-dd"
                        styleDefault();
                        style.DataFormat = workbook.CreateDataFormat().GetFormat("yyyy-MM-dd");
                        break;
                    case 9: //dateTime"yyyy-MM-dd HH:mm:ss")
                        styleDefault();
                        style.DataFormat = workbook.CreateDataFormat().GetFormat("yyyy-MM-dd HH:mm:ss");
                        break;
                    case 10: // string color blue                    
                        fontInner.Color = IndexedColors.Blue.Index;
                        styleDefault();
                        break;
                    case 11:  //无边框表头
                        fontInner.FontName = "黑体";//"Arial"
                        fontInner.FontHeightInPoints = 11;
                        style.BorderTop = style.BorderBottom = style.BorderLeft = style.BorderRight = NPOI.SS.UserModel.BorderStyle.Thin;//边框 BorderStyle.Double
                        style.TopBorderColor = style.BottomBorderColor = style.LeftBorderColor = style.RightBorderColor = IndexedColors.White.Index;//边框颜色
                        style.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Center;
                        break;
                    case 12:
                        style.BorderTop = style.BorderBottom = style.BorderLeft = style.BorderRight = NPOI.SS.UserModel.BorderStyle.Thin;//边框 BorderStyle.Double
                                                                                                                                         // style.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Left;
                        break;
                }
                // style.BorderTop = style.BorderBottom = style.BorderLeft = style.BorderRight = NPOI.SS.UserModel.BorderStyle.Thin;//边框 BorderStyle.Double
                // style.TopBorderColor = style.BottomBorderColor = style.LeftBorderColor = style.RightBorderColor = IndexedColors.Blue.Index;//边框颜色

                style.SetFont(fontInner);   //设定单元格字体                                 
                style.FillPattern = FillPattern.SolidForeground;//设定单元格填充

                return style;
            }

            void createSheetQuote()
            {
                feedDataByRowsForQuote();
                setColumnWidthQuote();
                mergeCellsQuote();

                void feedDataByRowsForQuote()
                {

                    int rowOutNum = 0;
                    var pairList = new List<KeyValuePair<int, object>>();
                    // 添加元素					
                    pairList.Add(new KeyValuePair<int, object>(0, "营销报价定额\r\n(kg/km)"));
                    pairList.Add(new KeyValuePair<int, object>(0, "——"));
                    pairList.Add(new KeyValuePair<int, object>(0, "产品型号"));
                    pairList.Add(new KeyValuePair<int, object>(1, theType));//待输入
                    pairList.Add(new KeyValuePair<int, object>(0, "——"));
                    pairList.Add(new KeyValuePair<int, object>(0, "电压等级"));
                    pairList.Add(new KeyValuePair<int, object>(01, voltage + " kV")); //"0.6/1kV"));//待输入
                    pairList.Add(new KeyValuePair<int, object>(0, "文件编号"));
                    pairList.Add(new KeyValuePair<int, object>(1, file_code)); //"Q/LN1 05 003-2024"));//待输入
                    createRowsTreatDataStyleQuote(pairList, rowOutNum++);
                    pairList.Clear();

                    pairList.Add(new KeyValuePair<int, object>(0, "——"));
                    pairList.Add(new KeyValuePair<int, object>(0, "——"));
                    pairList.Add(new KeyValuePair<int, object>(0, "产品名称"));
                    pairList.Add(new KeyValuePair<int, object>(1, production_name));  // "阻燃交联聚乙烯绝缘电力电缆"));//待输入
                    pairList.Add(new KeyValuePair<int, object>(0, "——"));
                    pairList.Add(new KeyValuePair<int, object>(0, "——"));
                    pairList.Add(new KeyValuePair<int, object>(0, "——"));
                    pairList.Add(new KeyValuePair<int, object>(0, "编制依据"));
                    pairList.Add(new KeyValuePair<int, object>(1, source));  // "GB/T 12706.1-2020"));//待输入
                    createRowsTreatDataStyleQuote(pairList, rowOutNum++);
                    pairList.Clear();

                    pairList.Add(new KeyValuePair<int, object>(0, "规格"));
                    pairList.Add(new KeyValuePair<int, object>(0, "导体"));
                    pairList.Add(new KeyValuePair<int, object>(0, "绕包"));
                    pairList.Add(new KeyValuePair<int, object>(0, "绝缘"));
                    pairList.Add(new KeyValuePair<int, object>(0, "成缆"));
                    pairList.Add(new KeyValuePair<int, object>(0, "——"));
                    pairList.Add(new KeyValuePair<int, object>(0, "内衬层"));
                    pairList.Add(new KeyValuePair<int, object>(0, "铠装"));
                    pairList.Add(new KeyValuePair<int, object>(0, "外护套"));
                    createRowsTreatDataStyleQuote(pairList, rowOutNum++);
                    pairList.Clear();

                    pairList.Add(new KeyValuePair<int, object>(0, "——"));
                    pairList.Add(new KeyValuePair<int, object>(1, conductorMaterial));  // "铜\r\n(kg/km)"));
                    pairList.Add(new KeyValuePair<int, object>(1, mica_tapeMaterial));  // "云母带"));//输入
                    pairList.Add(new KeyValuePair<int, object>(1, insulationMaterialSelected));  // "二步法硅烷交联绝缘料")); //输入    
                    pairList.Add(new KeyValuePair<int, object>(1, bufferMaterialSelected));  // "填充绳"));////输入
                    pairList.Add(new KeyValuePair<int, object>(1, tapeMaterialSelected));  // "三合一金云母带"));//输入
                    pairList.Add(new KeyValuePair<int, object>(1, inner_sheathMaterial));  // "H-90 PVC\r\n护套料"));//输入
                    pairList.Add(new KeyValuePair<int, object>(1, armourMaterialSelected));  // (isDoubleCable || isMultiCore) ? "多芯采用镀锌钢带" : "单芯采用不锈钢带"));  // 待输入(1, "单芯采用不锈钢带，\r\n多芯采用镀锌钢带"));//待输入
                    pairList.Add(new KeyValuePair<int, object>(1, outer_sheathMaterialSelected));  // "ZH-90 PVC护套料（氧指数大于等于36%）"));//输入
                    createRowsTreatDataStyleQuote(pairList, rowOutNum++);
                    pairList.Clear();

                    pairList.Add(new KeyValuePair<int, object>(1, spec));//待输入
                    pairList.Add(new KeyValuePair<int, object>(1, conductorWeight.ToString()));//待输入
                    pairList.Add(new KeyValuePair<int, object>(1, mica_tapeWeight));//待输入
                    pairList.Add(new KeyValuePair<int, object>(1, insulationWeightSelected)); //"insulation1Weight"));//待输入
                    pairList.Add(new KeyValuePair<int, object>(1, bufferWeightSelected));//待输入
                    pairList.Add(new KeyValuePair<int, object>(1, tapeWeightSelected));//待输入
                    pairList.Add(new KeyValuePair<int, object>(1, inner_sheathWeight));//待输入
                    pairList.Add(new KeyValuePair<int, object>(1, armourWeight));//待输入
                    pairList.Add(new KeyValuePair<int, object>(1, outer_sheathWeight));//待输入
                    createRowsTreatDataStyleQuote(pairList, rowOutNum++);
                    pairList.Clear();
                }

                void createRowsTreatDataStyleQuote(List<KeyValuePair<int, object>> inputDataList, int rowNum)
                {

                    // 3. 添加标题行
                    IRow iRow = sheetQuote.CreateRow(rowNum);
                    sheetQuote.AutoSizeRow(rowNum);
                    int colIdx = 0;
                    //for (int colIdx = 0; colIdx < 6; colIdx++)//方法2,3
                    foreach (var keyValuePair in inputDataList)//方法1
                    {
                        try
                        {
                            //ICell cell = iRow.GetCell(colIdx);//如果获得已有单元格，则这样写
                            ICell cell = iRow.CreateCell(colIdx);//IRow是地址引用，像指针，反过来赋值
                                                                 // cell.CellStyle = stringCenterStyle;
                                                                 //int styleInt = styleList[colIdx];//方法2
                            int styleInt = keyValuePair.Key;//方法1
                                                            // object value = inputDataList[colIdx];
                            object value = keyValuePair.Value;

                            // 根据数据类型应用样式
                            if (value == DBNull.Value || value == "")
                            {
                                // cell.SetCellValue("ut");// string.Empty;
                                // cell.CellStyle = nullStyle;// CreateNullStyle();
                                cell.SetCellValue("数据未提供");
                                cell.CellStyle = warnStyle;
                            }
                            else
                            {
                                switch (value)
                                {
                                    case string string1:
                                        cell.SetCellValue(Convert.ToString(value));
                                        // cell.CellStyle = stringStyle;
                                        break;
                                    case int int1:
                                        cell.SetCellValue(Convert.ToString(value));
                                        // cell.CellStyle = itemStyle;
                                        break;
                                    case DateTime dateTime1:
                                        cell.SetCellValue(Convert.ToDateTime(value));
                                        // cell.CellStyle = dateTimeStyle;
                                        break;
                                    case double double1:
                                        cell.SetCellValue(Convert.ToDouble(value));
                                        if (Convert.ToDouble(value) < 0)
                                        {
                                            // cell.CellStyle = warnNumStyle;//numberStyle; //
                                            break;
                                        }
                                        // cell.CellStyle = numberStyle;
                                        break;
                                    default:
                                        cell.SetCellValue(value.ToString());
                                        // cell.CellStyle = stringStyle; //文本自动换行
                                        break;
                                }
                                if (rowNum == 0 || rowNum == 1) iRow.Height = 16 * 20;// 14.4*2 *20;1/20个点为最小单位
                                if (value.ToString().Contains("营销报价定额", StringComparison.OrdinalIgnoreCase))
                                    cell.CellStyle = titleStyle;
                                else if (faultMarkArray.Any(term => value.ToString().Contains(term, StringComparison.OrdinalIgnoreCase)))
                                    cell.CellStyle = warnStyle;
                                else if (styleInt == 1)//方法2
                                    cell.CellStyle = stringBlueStyle;
                                /*  else if (value.ToString().Contains("电缆长期允许载流量", StringComparison.OrdinalIgnoreCase))
                                  {
                                      //合并单元格，文字超出单元格范围，行高不会自动变化
                                      iRow.Height = 576;// 14.4*2 *20;1/20个点为最小单位
                                      cell.CellStyle = stringCenterStyle;
                                  }*/
                                else
                                    cell.CellStyle = stringCenterStyle;

                            }
                            Trace.Write($"{value.ToString()}  ");//20250519 打印单元格数据                
                        }                             //                   }
                        catch (Exception excep1)
                        {
                            Trace.WriteLine(excep1.Message);
                            Trace.WriteLine($"问题在第{rowNum.ToString()}行  ");
                        }
                        ++colIdx;
                    }
                    Trace.WriteLine(""); ;
                }

                void setColumnWidthQuote()
                {
                    sheetQuote.SetColumnWidth(0, (19.11 + 0.78) * 256);
                    sheetQuote.SetColumnWidth(1, (7.56 + 0.78) * 256);
                    sheetQuote.SetColumnWidth(2, (11.44 + 0.78) * 256);
                    sheetQuote.SetColumnWidth(3, (12.98 + 0.78) * 256);
                    sheetQuote.SetColumnWidth(4, (9.44 + 0.78) * 256);
                    sheetQuote.SetColumnWidth(5, (11.44 + 0.78) * 256);
                    sheetQuote.SetColumnWidth(6, (12.56 + 0.78) * 256);
                    sheetQuote.SetColumnWidth(7, (13.11 + 0.78) * 256);
                    sheetQuote.SetColumnWidth(8, (23.56 + 0.78) * 256);
                }

                void mergeCellsQuote() //合并单元格
                {
                    CellRangeAddress cellMerge = new CellRangeAddress(0, 1, 0, 1);//（起始行，结束行，起始列，结束列）
                    sheetQuote.AddMergedRegion(cellMerge);
                    CellRangeAddress cellMerge2 = new CellRangeAddress(0, 0, 3, 4);//（起始行，结束行，起始列，结束列）
                    sheetQuote.AddMergedRegion(cellMerge2);
                    CellRangeAddress cellMerge1 = new CellRangeAddress(1, 1, 3, 6);//（起始行，结束行，起始列，结束列）         
                    sheetQuote.AddMergedRegion(cellMerge1);
                    CellRangeAddress cellMerge18 = new CellRangeAddress(2, 2, 4, 5);//（起始行，结束行，起始列，结束列）   
                    sheetQuote.AddMergedRegion(cellMerge18);
                    CellRangeAddress cellMerge3 = new CellRangeAddress(2, 3, 0, 0);//（起始行，结束行，起始列，结束列）
                    sheetQuote.AddMergedRegion(cellMerge3);

                    /*     CellRangeAddress cellMerge18 = new CellRangeAddress(32, 32, 2, 2);
                         sheetQuote.AddMergedRegion(cellMerge18);
                         CellRangeAddress cellMerge16 = new CellRangeAddress(32, 33, 3, 3);
                         sheetQuote.AddMergedRegion(cellMerge16);
                    */

                    //   CellRangeAddress cellMerge24 = new CellRangeAddress(35, 36, 2, 2);
                    //    sheetQuote.AddMergedRegion(cellMerge24);
                    //   CellRangeAddress cellMerge25 = new CellRangeAddress(35, 36, 3, 3);
                    //   sheetQuote.AddMergedRegion(cellMerge25);            
                }
            }


            void createSheetConstructionPara()
            {
                feedDataByRowsForConstructionPara();
                setColumnWidthConstructionPara();
                mergeCellsConstructionPara();

                //void feedDataByRowsForConstructionPara(bool isDoubleCable)
                void feedDataByRowsForConstructionPara()
                {

                    //待删除
                    #region
                    //forhead = Regex.Match(theType, @"(\w+)\-").Groups[1].Value;
                    //flameRedartant = Regex.Replace(forhead, @"w|d|W|D", "");
                    //if (flameRedartant == "") flameRedartant = "不适用";
                    //Trace.WriteLine($"flameRedartant: {flameRedartant}   ");

                    //halogenFree = (Regex.Match(forhead, @"(w|W)").Value);
                    //if (halogenFree == "") halogenFree = "不适用";
                    //Trace.WriteLine($"halogenFree: {halogenFree}");
                    //smokeFree = Regex.Match(forhead, @"(d|D)").Value;
                    //if (smokeFree == "") smokeFree = "不适用";
                    //Trace.WriteLine($"smokingFree: {smokeFree}");
                    ///*
                    //                MatchCollection matchCollection = Regex.Matches(spec, @"(\d×)|(\dX)|(\dx)");
                    //                foreach (MatchRegex match1 in matchCollection)
                    //                {
                    //                    MatchRegex theMatch = Regex.Match(match1.Value, @"\d");
                    //                    //   Trace.WriteLine($"数字：{theMatch.Value}");
                    //                    isMultiCore = (Convert.ToInt16(theMatch.Value) > 1);
                    //                    Trace.WriteLine($"theMatch.Value: {Convert.ToInt16(theMatch.Value)}");
                    //                    // specMini.Add(match1.Value);
                    //                }*/
                    //outer_sheathMaterialFront = Regex.Match(outer_sheathMaterialSelected, @"(.+)（").Groups[1].Value;
                    //Trace.WriteLine($"outer_sheathMaterialFront: {outer_sheathMaterialFront}");

                    ///*     MatchCollection matchCollectionRear = Regex.Matches(spec, @"(×\d+)|(X\d+)|(x\d+)");
                    //     //int k = 0;
                    //     List<string> areaConductor = [];
                    //     foreach (MatchRegex match1 in matchCollectionRear)
                    //     {
                    //         MatchRegex theMatch = Regex.Match(match1.Value, @"\d+");
                    //         areaConductor.Add($"对应{theMatch.Value} mm²截面");
                    //         //Trace.WriteLine($"{areaConductor[k++]}");
                    //     }

                    //     */
                    #endregion

                    int rowOutNum = 0;
                    var pairList = new List<KeyValuePair<int, object>>();
                    // 添加元素					
                    pairList.Add(new KeyValuePair<int, object>(0, "电缆结构技术参数"));
                    pairList.Add(new KeyValuePair<int, object>(0, "——"));
                    pairList.Add(new KeyValuePair<int, object>(0, "——"));
                    pairList.Add(new KeyValuePair<int, object>(0, "——"));
                    pairList.Add(new KeyValuePair<int, object>(0, "——"));
                    pairList.Add(new KeyValuePair<int, object>(0, "——"));
                    createRowsTreatDataStyleConstructionPara(pairList, rowOutNum++);
                    pairList.Clear();

                    pairList.Add(new KeyValuePair<int, object>(0, "电缆型号"));
                    pairList.Add(new KeyValuePair<int, object>(0, "——"));
                    pairList.Add(new KeyValuePair<int, object>(1, type_spec));
                    pairList.Add(new KeyValuePair<int, object>(0, "——"));
                    pairList.Add(new KeyValuePair<int, object>(0, "——"));
                    pairList.Add(new KeyValuePair<int, object>(0, "——"));
                    createRowsTreatDataStyleConstructionPara(pairList, rowOutNum++);
                    pairList.Clear();

                    pairList.Add(new KeyValuePair<int, object>(0, "项　　目"));
                    pairList.Add(new KeyValuePair<int, object>(0, "——"));
                    pairList.Add(new KeyValuePair<int, object>(0, "单位"));
                    pairList.Add(new KeyValuePair<int, object>(0, "标准参数值"));
                    pairList.Add(new KeyValuePair<int, object>(0, "投标人响应值"));
                    pairList.Add(new KeyValuePair<int, object>(0, "备注"));
                    createRowsTreatDataStyleConstructionPara(pairList, rowOutNum++);
                    pairList.Clear();


                    pairList.Add(new KeyValuePair<int, object>(0, "铜导体"));
                    pairList.Add(new KeyValuePair<int, object>(0, "材料"));
                    pairList.Add(new KeyValuePair<int, object>(0, "——"));
                    pairList.Add(new KeyValuePair<int, object>(1, conductorMaterial));  //type_spec.Contains('L', StringComparison.OrdinalIgnoreCase) ? "铝" : "铜")); // "铜"));//变色
                    pairList.Add(new KeyValuePair<int, object>(1, conductorMaterial));  //type_spec.Contains('L', StringComparison.OrdinalIgnoreCase) ? "铝" : "铜")); // "铜"));//变色
                    pairList.Add(new KeyValuePair<int, object>(0, "——"));
                    createRowsTreatDataStyleConstructionPara(pairList, rowOutNum++);
                    pairList.Clear();

                    pairList.Add(new KeyValuePair<int, object>(0, "——"));
                    pairList.Add(new KeyValuePair<int, object>(0, "芯数×标称截面"));
                    pairList.Add(new KeyValuePair<int, object>(0, "芯×mm²"));
                    pairList.Add(new KeyValuePair<int, object>(1, spec));// 待输入
                    pairList.Add(new KeyValuePair<int, object>(1, spec));// 待输入
                    pairList.Add(new KeyValuePair<int, object>(0, "——"));
                    createRowsTreatDataStyleConstructionPara(pairList, rowOutNum++);
                    pairList.Clear();

                    pairList.Add(new KeyValuePair<int, object>(0, "——"));
                    pairList.Add(new KeyValuePair<int, object>(0, "结构形式"));
                    pairList.Add(new KeyValuePair<int, object>(0, "芯×mm²"));
                    pairList.Add(new KeyValuePair<int, object>(0, "紧压圆形 / 实心导体"));
                    pairList.Add(new KeyValuePair<int, object>(0, "紧压圆形 / 实心导体"));
                    pairList.Add(new KeyValuePair<int, object>(0, "固定不变？"));
                    createRowsTreatDataStyleConstructionPara(pairList, rowOutNum++);
                    pairList.Clear();


                    pairList.Add(new KeyValuePair<int, object>(0, "——"));
                    pairList.Add(new KeyValuePair<int, object>(0, "最少单线根数"));
                    pairList.Add(new KeyValuePair<int, object>(0, "根"));
                    pairList.Add(new KeyValuePair<int, object>(1, pieces_1));// "?"));// 待输入  
                    pairList.Add(new KeyValuePair<int, object>(1, pieces_1));// "?"));// 待输入 pieces_1)); //
                    pairList.Add(new KeyValuePair<int, object>(1, areaConductor[0]));//"对应10mm²截面"));// 待输入
                    createRowsTreatDataStyleConstructionPara(pairList, rowOutNum++);
                    pairList.Clear();

                    if (isDoubleCable)
                    {
                        pairList.Add(new KeyValuePair<int, object>(0, "——"));
                        pairList.Add(new KeyValuePair<int, object>(0, "——"));
                        pairList.Add(new KeyValuePair<int, object>(0, "——"));
                        pairList.Add(new KeyValuePair<int, object>(1, pieces_2));//"3"));// 待输入
                        Trace.WriteLine($"pieces_2={pieces_2}");
                        pairList.Add(new KeyValuePair<int, object>(1, pieces_2));//"2"));// 待输入
                        pairList.Add(new KeyValuePair<int, object>(1, areaConductor[1])); //"对应6mm²截面"));// 待输入
                        createRowsTreatDataStyleConstructionPara(pairList, rowOutNum++);
                        pairList.Clear();
                        pairList.Clear();
                    }

                    pairList.Add(new KeyValuePair<int, object>(0, "——"));
                    pairList.Add(new KeyValuePair<int, object>(0, "导体外径（近似值）"));
                    pairList.Add(new KeyValuePair<int, object>(0, "mm"));
                    pairList.Add(new KeyValuePair<int, object>(0, "（投标人提供）"));
                    pairList.Add(new KeyValuePair<int, object>(1, conductDiameter_1));////4.1));  // 待输入
                    pairList.Add(new KeyValuePair<int, object>(1, areaConductor[0])); //"对应10mm²截面")); // 待输入
                    createRowsTreatDataStyleConstructionPara(pairList, rowOutNum++);
                    pairList.Clear();

                    if (isDoubleCable)
                    {
                        pairList.Add(new KeyValuePair<int, object>(0, "——"));
                        pairList.Add(new KeyValuePair<int, object>(0, "——"));
                        pairList.Add(new KeyValuePair<int, object>(0, "——"));
                        pairList.Add(new KeyValuePair<int, object>(0, "——"));
                        pairList.Add(new KeyValuePair<int, object>(1, conductDiameter_2));//2.2));  // 待输入
                        pairList.Add(new KeyValuePair<int, object>(1, areaConductor[1])); //"对应4mm²截面")); // 待输入
                        createRowsTreatDataStyleConstructionPara(pairList, rowOutNum++);
                        pairList.Clear();
                    }

                    pairList.Add(new KeyValuePair<int, object>(0, "——"));
                    pairList.Add(new KeyValuePair<int, object>(0, "紧压系数"));
                    pairList.Add(new KeyValuePair<int, object>(0, "——"));
                    pairList.Add(new KeyValuePair<int, object>(0, "≥0.9"));
                    pairList.Add(new KeyValuePair<int, object>(0, "≥0.9\r\n(对应紧压圆形导体结构)"));
                    pairList.Add(new KeyValuePair<int, object>(0, "——"));
                    createRowsTreatDataStyleConstructionPara(pairList, rowOutNum++);
                    pairList.Clear();

                    Trace.WriteLine($"insulationMaterialSelected: {insulationMaterialSelected}");
                    pairList.Add(new KeyValuePair<int, object>(0, "绝缘"));
                    pairList.Add(new KeyValuePair<int, object>(0, "材料"));
                    pairList.Add(new KeyValuePair<int, object>(0, "——"));
                    pairList.Add(new KeyValuePair<int, object>(1, insulationMaterialSelected)); //insulation1Material));// "XLPE"));// 待输入
                    pairList.Add(new KeyValuePair<int, object>(1, insulationMaterialSelected));//insulation1Material));// "XLPE"));// 待输入
                    pairList.Add(new KeyValuePair<int, object>(0, "——"));
                    createRowsTreatDataStyleConstructionPara(pairList, rowOutNum++);
                    pairList.Clear();

                    pairList.Add(new KeyValuePair<int, object>(0, "——"));
                    pairList.Add(new KeyValuePair<int, object>(0, "平均厚度不小于标称厚度 t"));
                    pairList.Add(new KeyValuePair<int, object>(0, "mm"));
                    pairList.Add(new KeyValuePair<int, object>(0, "（投标人提供）"));
                    pairList.Add(new KeyValuePair<int, object>(1, insulationThick_1));//"0.7"));// 待输入
                    pairList.Add(new KeyValuePair<int, object>(1, areaConductor[0])); //"对应10截面"));// 待输入
                    createRowsTreatDataStyleConstructionPara(pairList, rowOutNum++);
                    pairList.Clear();

                    if (isDoubleCable)
                    {
                        pairList.Add(new KeyValuePair<int, object>(0, "——"));
                        pairList.Add(new KeyValuePair<int, object>(0, "——"));
                        pairList.Add(new KeyValuePair<int, object>(0, "——"));
                        pairList.Add(new KeyValuePair<int, object>(0, "——"));
                        pairList.Add(new KeyValuePair<int, object>(1, insulationThick_2));//"0.7"));// 待输入
                        pairList.Add(new KeyValuePair<int, object>(1, areaConductor[1])); //"对应6截面"));// 待输入
                        createRowsTreatDataStyleConstructionPara(pairList, rowOutNum++);
                        pairList.Clear();
                    }

                    pairList.Add(new KeyValuePair<int, object>(0, "——"));
                    pairList.Add(new KeyValuePair<int, object>(0, "最薄点厚度不小于标称值"));
                    pairList.Add(new KeyValuePair<int, object>(0, "mm"));
                    pairList.Add(new KeyValuePair<int, object>(0, "90 % t"));
                    pairList.Add(new KeyValuePair<int, object>(0, "90 % t"));
                    pairList.Add(new KeyValuePair<int, object>(0, "——"));
                    createRowsTreatDataStyleConstructionPara(pairList, rowOutNum++);
                    pairList.Clear();

                    pairList.Add(new KeyValuePair<int, object>(0, "——"));
                    pairList.Add(new KeyValuePair<int, object>(0, "偏心度"));
                    pairList.Add(new KeyValuePair<int, object>(0, "%"));
                    pairList.Add(new KeyValuePair<int, object>(0, "10"));
                    pairList.Add(new KeyValuePair<int, object>(0, "≤10"));
                    pairList.Add(new KeyValuePair<int, object>(0, "——"));
                    createRowsTreatDataStyleConstructionPara(pairList, rowOutNum++);
                    pairList.Clear();

                    if (isDoubleCable|| isMultiCore)
                    {
                        pairList.Add(new KeyValuePair<int, object>(0, "填充层"));
                        pairList.Add(new KeyValuePair<int, object>(0, "填充材料"));
                        pairList.Add(new KeyValuePair<int, object>(0, "——"));
                        pairList.Add(new KeyValuePair<int, object>(0, "（投标人提供）"));
                        pairList.Add(new KeyValuePair<int, object>(1, bufferMaterialSelected)); //buffer1Material));//buffer1Material ?? "无"));// 待输入
                        pairList.Add(new KeyValuePair<int, object>(0, "——"));
                        createRowsTreatDataStyleConstructionPara(pairList, rowOutNum++);
                        pairList.Clear();
                    }


                    if (Convert.ToDouble(inner_sheathWeight) >= 1)
                    {
                        pairList.Add(new KeyValuePair<int, object>(0, "内衬层"));
                        pairList.Add(new KeyValuePair<int, object>(0, "材料"));
                        pairList.Add(new KeyValuePair<int, object>(0, "——"));
                        pairList.Add(new KeyValuePair<int, object>(0, "（投标人提供）"));
                        pairList.Add(new KeyValuePair<int, object>(1, inner_sheathMaterial));//"H - 90 PVC护套料"));  // 待输入 
                        pairList.Add(new KeyValuePair<int, object>(0, "——"));
                        createRowsTreatDataStyleConstructionPara(pairList, rowOutNum++);
                        pairList.Clear();

                        pairList.Add(new KeyValuePair<int, object>(0, "——"));
                        pairList.Add(new KeyValuePair<int, object>(0, "厚度\r\n（依据GB/T 12706.1假定外径对应选取）"));
                        pairList.Add(new KeyValuePair<int, object>(0, "mm"));
                        pairList.Add(new KeyValuePair<int, object>(0, "（投标人提供）"));
                        pairList.Add(new KeyValuePair<int, object>(1, inner_thick)); //"1.0 - 2.0"));  // 待输入
                        pairList.Add(new KeyValuePair<int, object>(0, "——"));
                        createRowsTreatDataStyleConstructionPara(pairList, rowOutNum++);
                        pairList.Clear();
                    }

                    if (isArmoured)
                    {
                        pairList.Add(new KeyValuePair<int, object>(0, "铠装层"));
                        pairList.Add(new KeyValuePair<int, object>(0, "材料"));
                        pairList.Add(new KeyValuePair<int, object>(0, "——"));
                        pairList.Add(new KeyValuePair<int, object>(0, "（投标人提供）"));
                        //  pairList.Add(new KeyValuePair<int, object>(1, "单芯采用不锈钢带，\r\n多芯采用镀锌钢带"));  // 待输入
                        pairList.Add(new KeyValuePair<int, object>(1, armourMaterialSelected)); //(isDoubleCable || isMultiCore) ? armour2Material : armour1Material)); //"多芯采用镀锌钢带" : "单芯采用不锈钢带"));  // 待输入
                        Trace.WriteLine($"isDoubleCable: {isDoubleCable}    isMultiCore:  {isMultiCore}");
                        pairList.Add(new KeyValuePair<int, object>(0, "与供货需求表一致"));
                        createRowsTreatDataStyleConstructionPara(pairList, rowOutNum++);
                        pairList.Clear();

                        pairList.Add(new KeyValuePair<int, object>(0, "——"));
                        pairList.Add(new KeyValuePair<int, object>(0, "钢带厚度/钢丝直径\r\n（依据GB/T 12706.1假定外径对应选取）"));
                        pairList.Add(new KeyValuePair<int, object>(0, "mm"));
                        pairList.Add(new KeyValuePair<int, object>(0, "（投标人提供）"));
                        pairList.Add(new KeyValuePair<int, object>(1, steel_thick)); //"0.2~0.5"));  // 待输入
                        pairList.Add(new KeyValuePair<int, object>(0, "——"));
                        createRowsTreatDataStyleConstructionPara(pairList, rowOutNum++);
                        pairList.Clear();


                        pairList.Add(new KeyValuePair<int, object>(0, "——"));
                        pairList.Add(new KeyValuePair<int, object>(0, "钢带层数"));
                        pairList.Add(new KeyValuePair<int, object>(0, "层"));
                        pairList.Add(new KeyValuePair<int, object>(1, 2)); // 待输入
                        pairList.Add(new KeyValuePair<int, object>(1, 2));  // 待输入
                        pairList.Add(new KeyValuePair<int, object>(0, "固定不变？"));
                        createRowsTreatDataStyleConstructionPara(pairList, rowOutNum++);
                        pairList.Clear();

                        pairList.Add(new KeyValuePair<int, object>(0, "——"));
                        pairList.Add(new KeyValuePair<int, object>(0, "钢带宽度"));
                        pairList.Add(new KeyValuePair<int, object>(0, "mm"));
                        pairList.Add(new KeyValuePair<int, object>(0, "（投标人提供）"));
                        pairList.Add(new KeyValuePair<int, object>(1, steel_width)); // 待输入
                        pairList.Add(new KeyValuePair<int, object>(0, "——"));
                        createRowsTreatDataStyleConstructionPara(pairList, rowOutNum++);
                        pairList.Clear();
                    }
                    pairList.Add(new KeyValuePair<int, object>(0, "外护套"));
                    pairList.Add(new KeyValuePair<int, object>(0, "材料"));
                    pairList.Add(new KeyValuePair<int, object>(0, "——"));
                    pairList.Add(new KeyValuePair<int, object>(0, "（投标人提供）"));
                    pairList.Add(new KeyValuePair<int, object>(1, outer_sheathMaterialFront));// outer_sheath1Material));//"ZH-90 PVC护套料")); // 待输入
                    pairList.Add(new KeyValuePair<int, object>(0, "——"));
                    createRowsTreatDataStyleConstructionPara(pairList, rowOutNum++);
                    pairList.Clear();

                    pairList.Add(new KeyValuePair<int, object>(0, "——"));
                    pairList.Add(new KeyValuePair<int, object>(0, "颜色"));
                    pairList.Add(new KeyValuePair<int, object>(0, "——"));
                    pairList.Add(new KeyValuePair<int, object>(0, "（投标人提供）"));
                    pairList.Add(new KeyValuePair<int, object>(1, "黑色"));  //高亮颜色
                    pairList.Add(new KeyValuePair<int, object>(1, "黑色？留空？")); //高亮颜色
                    createRowsTreatDataStyleConstructionPara(pairList, rowOutNum++);
                    pairList.Clear();


                    pairList.Add(new KeyValuePair<int, object>(0, "——"));
                    pairList.Add(new KeyValuePair<int, object>(1, isArmoured ? "标称厚度t(有铠装)" : "标称厚度t(无铠装)"));//(0, "标称厚度t（无铠装）"));
                    pairList.Add(new KeyValuePair<int, object>(0, "mm"));
                    pairList.Add(new KeyValuePair<int, object>(0, "（投标人提供）"));
                    pairList.Add(new KeyValuePair<int, object>(1, sheathThick)); //0.8)); // 待输入
                    pairList.Add(new KeyValuePair<int, object>(0, "——")); //"Z - YJV"));// 待输入
                    createRowsTreatDataStyleConstructionPara(pairList, rowOutNum++);
                    pairList.Clear();

                    string armourString = isArmoured ? "铠装80%" : "无铠装85%";
                    pairList.Add(new KeyValuePair<int, object>(0, "——"));
                    pairList.Add(new KeyValuePair<int, object>(0, "最薄点厚度不小于"));
                    pairList.Add(new KeyValuePair<int, object>(0, "mm"));
                    pairList.Add(new KeyValuePair<int, object>(1, armourString));// 待输入armourWeight
                    pairList.Add(new KeyValuePair<int, object>(1, armourString));// 待输入
                    pairList.Add(new KeyValuePair<int, object>(0, "——"));
                    createRowsTreatDataStyleConstructionPara(pairList, rowOutNum++);
                    pairList.Clear();

                    pairList.Add(new KeyValuePair<int, object>(0, " 电缆外径D："));
                    pairList.Add(new KeyValuePair<int, object>(0, "——"));
                    pairList.Add(new KeyValuePair<int, object>(0, "mm"));
                    pairList.Add(new KeyValuePair<int, object>(0, "（投标人提供）"));
                    pairList.Add(new KeyValuePair<int, object>(1, cableDiameter));// 19.74));// 待输入
                    pairList.Add(new KeyValuePair<int, object>(0, "——")); //(1, type_spec)); // 待输入
                    createRowsTreatDataStyleConstructionPara(pairList, rowOutNum++);
                    pairList.Clear();

                    pairList.Add(new KeyValuePair<int, object>(0, "20℃时铜导体最大直流电阻"));
                    pairList.Add(new KeyValuePair<int, object>(0, "——"));
                    pairList.Add(new KeyValuePair<int, object>(0, "Ω/km"));
                    pairList.Add(new KeyValuePair<int, object>(1, resistant20_1));//1.83));// 待输入
                    pairList.Add(new KeyValuePair<int, object>(1, resistant20_1));//1.83));// 待输入
                    pairList.Add(new KeyValuePair<int, object>(1, areaConductor[0])); //"错误对应10mm²截面"));// 待输入
                    createRowsTreatDataStyleConstructionPara(pairList, rowOutNum++);
                    pairList.Clear();

                    if (isDoubleCable)
                    {
                        pairList.Add(new KeyValuePair<int, object>(0, "——"));
                        pairList.Add(new KeyValuePair<int, object>(0, "——"));
                        pairList.Add(new KeyValuePair<int, object>(0, "——"));
                        pairList.Add(new KeyValuePair<int, object>(1, resistant20_2));//3.08));// 待输入
                        pairList.Add(new KeyValuePair<int, object>(1, resistant20_2));//3.08));// 待输入
                        pairList.Add(new KeyValuePair<int, object>(1, areaConductor[1])); //"对应6mm²截面"));// 待输入
                        createRowsTreatDataStyleConstructionPara(pairList, rowOutNum++);
                        pairList.Clear();
                    }

                    pairList.Add(new KeyValuePair<int, object>(0, "90℃时铜导体最大交流电阻"));
                    pairList.Add(new KeyValuePair<int, object>(0, "——"));
                    pairList.Add(new KeyValuePair<int, object>(0, "Ω/kμ"));
                    pairList.Add(new KeyValuePair<int, object>(0, "（投标人提供）"));
                    pairList.Add(new KeyValuePair<int, object>(1, resistant90_1));//2.3334));// 待输入
                    pairList.Add(new KeyValuePair<int, object>(1, areaConductor[0])); //"对应10mm²截面"));// 待输入
                    createRowsTreatDataStyleConstructionPara(pairList, rowOutNum++);
                    pairList.Clear();

                    if (isDoubleCable)
                    {
                        pairList.Add(new KeyValuePair<int, object>(0, "——"));
                        pairList.Add(new KeyValuePair<int, object>(0, "——"));
                        pairList.Add(new KeyValuePair<int, object>(0, "——"));
                        pairList.Add(new KeyValuePair<int, object>(0, "——"));
                        pairList.Add(new KeyValuePair<int, object>(1, resistant90_2));//3.9273));// 待输入
                        pairList.Add(new KeyValuePair<int, object>(1, areaConductor[1])); //"对应6mm²截面"));// 待输入
                        createRowsTreatDataStyleConstructionPara(pairList, rowOutNum++);
                        pairList.Clear();
                    }

                    pairList.Add(new KeyValuePair<int, object>(0, "电缆长期允许载流量\r\n（计算值，空气中40℃敷设）"));
                    pairList.Add(new KeyValuePair<int, object>(0, "——"));
                    pairList.Add(new KeyValuePair<int, object>(0, "A"));
                    pairList.Add(new KeyValuePair<int, object>(0, "（投标人提供）"));
                    pairList.Add(new KeyValuePair<int, object>(1, current40)); //272));// 待输入
                    pairList.Add(new KeyValuePair<int, object>(1, areaConductor[0])); //type_spec)); // 待输入
                    createRowsTreatDataStyleConstructionPara(pairList, rowOutNum++);
                    pairList.Clear();

                    pairList.Add(new KeyValuePair<int, object>(0, "出厂工频电压试验"));
                    pairList.Add(new KeyValuePair<int, object>(0, "——"));
                    pairList.Add(new KeyValuePair<int, object>(0, "kV/min"));
                    pairList.Add(new KeyValuePair<int, object>(0, "3.5 U0/5"));
                    pairList.Add(new KeyValuePair<int, object>(0, "3.5/5"));
                    pairList.Add(new KeyValuePair<int, object>(0, "——"));
                    createRowsTreatDataStyleConstructionPara(pairList, rowOutNum++);
                    pairList.Clear();

                    pairList.Add(new KeyValuePair<int, object>(0, "电缆盘尺寸"));
                    pairList.Add(new KeyValuePair<int, object>(0, "——"));
                    pairList.Add(new KeyValuePair<int, object>(0, "mm"));
                    pairList.Add(new KeyValuePair<int, object>(0, "（投标人提供）"));
                    pairList.Add(new KeyValuePair<int, object>(0, "根据订单长度选择"));
                    pairList.Add(new KeyValuePair<int, object>(0, "——"));
                    createRowsTreatDataStyleConstructionPara(pairList, rowOutNum++);
                    pairList.Clear();

                    pairList.Add(new KeyValuePair<int, object>(0, "电缆敷设时的最大牵引力"));
                    pairList.Add(new KeyValuePair<int, object>(0, "——"));
                    pairList.Add(new KeyValuePair<int, object>(0, "N/mm²"));
                    pairList.Add(new KeyValuePair<int, object>(1, "70"));// 待输入
                    pairList.Add(new KeyValuePair<int, object>(1, "70"));// 待输入
                    pairList.Add(new KeyValuePair<int, object>(1, "铜芯，牵引头?"));// 待输入
                    createRowsTreatDataStyleConstructionPara(pairList, rowOutNum++);
                    pairList.Clear();

                    pairList.Add(new KeyValuePair<int, object>(0, "电缆敷设时的最大侧压力"));
                    pairList.Add(new KeyValuePair<int, object>(0, "——"));
                    pairList.Add(new KeyValuePair<int, object>(0, "N/m"));
                    pairList.Add(new KeyValuePair<int, object>(0, "5000"));
                    pairList.Add(new KeyValuePair<int, object>(0, "5000"));
                    pairList.Add(new KeyValuePair<int, object>(0, "——"));
                    createRowsTreatDataStyleConstructionPara(pairList, rowOutNum++);
                    pairList.Clear();

                    pairList.Add(new KeyValuePair<int, object>(0, "电缆质量（近似值）"));
                    pairList.Add(new KeyValuePair<int, object>(0, "——"));
                    pairList.Add(new KeyValuePair<int, object>(0, "kg/m"));
                    pairList.Add(new KeyValuePair<int, object>(0, "（投标人提供）"));
                    pairList.Add(new KeyValuePair<int, object>(1, cableWeight));//(1, "2.7"));// 待输入  cableWeight
                    pairList.Add(new KeyValuePair<int, object>(0, "——")); //type_spec)); // 待输入
                    createRowsTreatDataStyleConstructionPara(pairList, rowOutNum++);
                    pairList.Clear();
                    Trace.WriteLine($"cableWeight 等于： {cableWeight} ");

                    pairList.Add(new KeyValuePair<int, object>(0, "电缆敷设时允许环境温度"));
                    pairList.Add(new KeyValuePair<int, object>(0, "——"));
                    pairList.Add(new KeyValuePair<int, object>(0, "℃"));
                    pairList.Add(new KeyValuePair<int, object>(0, "-5～＋40"));
                    pairList.Add(new KeyValuePair<int, object>(0, "-5～＋40"));
                    pairList.Add(new KeyValuePair<int, object>(0, "——"));
                    createRowsTreatDataStyleConstructionPara(pairList, rowOutNum++);
                    pairList.Clear();

                    pairList.Add(new KeyValuePair<int, object>(0, "电缆在正常使用条件下的寿命"));
                    pairList.Add(new KeyValuePair<int, object>(0, "——"));
                    pairList.Add(new KeyValuePair<int, object>(0, "年"));
                    pairList.Add(new KeyValuePair<int, object>(0, "≥30"));
                    pairList.Add(new KeyValuePair<int, object>(0, "≥30"));
                    pairList.Add(new KeyValuePair<int, object>(0, "——"));
                    createRowsTreatDataStyleConstructionPara(pairList, rowOutNum++);
                    pairList.Clear();


                    pairList.Add(new KeyValuePair<int, object>(0, "电缆阻燃级别"));
                    pairList.Add(new KeyValuePair<int, object>(0, "——"));
                    pairList.Add(new KeyValuePair<int, object>(0, "——"));
                    pairList.Add(new KeyValuePair<int, object>(0, "按供货需求表"));
                    pairList.Add(new KeyValuePair<int, object>(1, flameRedartant)); ////"ZC"));
                    pairList.Add(new KeyValuePair<int, object>(0, "——"));
                    createRowsTreatDataStyleConstructionPara(pairList, rowOutNum++);
                    pairList.Clear();

                    pairList.Add(new KeyValuePair<int, object>(0, "电缆的无卤性能"));
                    pairList.Add(new KeyValuePair<int, object>(0, "——"));
                    pairList.Add(new KeyValuePair<int, object>(0, "——"));
                    pairList.Add(new KeyValuePair<int, object>(0, "按供货需求表"));
                    pairList.Add(new KeyValuePair<int, object>(1, halogenFree));//"不适用"));// 待输入
                    pairList.Add(new KeyValuePair<int, object>(0, "——"));// 待输入
                    createRowsTreatDataStyleConstructionPara(pairList, rowOutNum++);
                    pairList.Clear();

                    pairList.Add(new KeyValuePair<int, object>(0, "电缆的低烟性能"));
                    pairList.Add(new KeyValuePair<int, object>(0, "——"));
                    pairList.Add(new KeyValuePair<int, object>(0, "——"));
                    pairList.Add(new KeyValuePair<int, object>(0, "按供货需求表"));
                    pairList.Add(new KeyValuePair<int, object>(1, smokeFree));// "不适用"));// 待输入
                    pairList.Add(new KeyValuePair<int, object>(0, "——"));// 待输入
                    createRowsTreatDataStyleConstructionPara(pairList, rowOutNum++);
                    pairList.Clear();
                }
                void feedDataByRows2()// 未使用
                {
                    /*
                      List<int> styleList = [];
                      List<Object> data1 = [];

                      data1.Add("电缆结构技术参数");
                      data1.Add("");
                      data1.Add("");
                      data1.Add("");
                      data1.Add("");
                      data1.Add("");
                      styleList.Add(0);
                      styleList.Add(0);
                      styleList.Add(0);
                      styleList.Add(0);
                      styleList.Add(0);
                      styleList.Add(0);
                      createRowsTreatDataStyleConstructionPara(data1, styleList, 0);
                      data1.Clear();
                      styleList.Clear();

                      data1.Add("电缆型号");
                      data1.Add("");
                      data1.Add("Z-YJV22 0.6/1 3×10＋1×6");//待输入
                      data1.Add("");
                      data1.Add("");
                      data1.Add("");
                      styleList.Add(0);
                      styleList.Add(0);
                      styleList.Add(1);
                      styleList.Add(0);
                      styleList.Add(0);
                      styleList.Add(0);
                      createRowsTreatDataStyleConstructionPara(data1, styleList, 1);
                      data1.Clear();            
                      styleList.Clear();


                      data1.Add("项　　目");
                      data1.Add("");
                      data1.Add("单位");
                      data1.Add("标准参数值");
                      data1.Add("投标人响应值");
                      data1.Add("备注");
                      styleList.Add(0);
                      styleList.Add(0);
                      styleList.Add(0);
                      styleList.Add(0);
                      styleList.Add(0);
                      styleList.Add(0);
                      createRowsTreatDataStyleConstructionPara(data1, styleList, 2);
                      data1.Clear();            styleList.Clear();

                      data1.Add("铜导体");
                      data1.Add("");
                      data1.Add("材料");
                      data1.Add("铜");//变色
                      data1.Add("铜");//变色
                      data1.Add("");
                      styleList.Add(0);
                      styleList.Add(0);
                      styleList.Add(0);
                      styleList.Add(0);
                      styleList.Add(1);
                      styleList.Add(1);
                      createRowsTreatDataStyleConstructionPara(data1, styleList, 3);
                      data1.Clear();            styleList.Clear();

                      data1.Add("");
                      data1.Add("芯数×标称截面");
                      data1.Add("芯×mm²");
                      data1.Add("3×10＋1×6");// 待输入
                      data1.Add("3×10＋1×6");// 待输入
                      data1.Add("");
                      styleList.Add(0);
                      styleList.Add(0);
                      styleList.Add(0);
                      styleList.Add(1);
                      styleList.Add(1);
                      styleList.Add(0);
                      createRowsTreatDataStyleConstructionPara(data1, styleList, 4);
                      data1.Clear();            styleList.Clear();


                      data1.Add("");
                      data1.Add("结构形式");
                      data1.Add("芯×mm²");
                      data1.Add("紧压圆形 / 实心导体");
                      data1.Add("紧压圆形 / 实心导体");
                      data1.Add("");
                      styleList.Add(0);
                      styleList.Add(0);
                      styleList.Add(0);
                      styleList.Add(1);
                      styleList.Add(1);
                      styleList.Add(0);
                      createRowsTreatDataStyleConstructionPara(data1, styleList, 5);
                      data1.Clear();            styleList.Clear();

                      data1.Add("");
                      data1.Add("最少单线根数");
                      data1.Add("根");
                      data1.Add("1");// 待输入
                      data1.Add("1");// 待输入
                      data1.Add("对应6mm²截面");// 待输入
                      styleList.Add(0);
                      styleList.Add(0);
                      styleList.Add(0);
                      styleList.Add(1);
                      styleList.Add(1);
                      styleList.Add(1);
                      createRowsTreatDataStyleConstructionPara(data1, styleList, 6);
                      data1.Clear();            styleList.Clear();

                      data1.Add("");
                      data1.Add("导体外径（近似值）");
                      data1.Add("mm");
                      data1.Add("（投标人提供）");
                      data1.Add(2.2);  // 待输入
                      data1.Add("对应4mm²截面"); // 待输入
                      styleList.Add(0);
                      styleList.Add(0);
                      styleList.Add(0);
                      styleList.Add(1);
                      styleList.Add(1);
                      styleList.Add(1);
                      createRowsTreatDataStyleConstructionPara(data1, styleList, 7);
                      data1.Clear();            styleList.Clear();

                      data1.Add("");
                      data1.Add("紧压系数");
                      data1.Add("");
                      data1.Add("≥0.9");
                      data1.Add("≥0.9（对应紧压圆形导体结构）");
                      data1.Add("");
                      styleList.Add(0);
                      styleList.Add(0);
                      styleList.Add(0);
                      styleList.Add(0);
                      styleList.Add(0);
                      styleList.Add(0);
                      createRowsTreatDataStyleConstructionPara(data1, styleList, 8);
                      data1.Clear();            styleList.Clear();

                      data1.Add("绝缘");
                      data1.Add("材料");
                      data1.Add("");
                      data1.Add("XLPE");// 待输入
                      data1.Add("XLPE");// 待输入
                      data1.Add("");
                      styleList.Add(0);
                      styleList.Add(0);
                      styleList.Add(0);
                      styleList.Add(1);
                      styleList.Add(1);
                      styleList.Add(0);
                      createRowsTreatDataStyleConstructionPara(data1, styleList, 9);
                      data1.Clear();            styleList.Clear();

                      data1.Add("");
                      data1.Add("平均厚度不小于标称厚度 t");
                      data1.Add("mm");
                      data1.Add("（投标人提供）");
                      data1.Add(0.7);// 待输入
                      data1.Add("对应6、10截面");// 待输入
                      styleList.Add(0);
                      styleList.Add(0);
                      styleList.Add(0);
                      styleList.Add(0);
                      styleList.Add(1);
                      styleList.Add(1);
                      createRowsTreatDataStyleConstructionPara(data1, styleList, 10);
                      data1.Clear();            styleList.Clear();

                      data1.Add("");
                      data1.Add("最薄点厚度不小于标称值");
                      data1.Add("mm");
                      data1.Add("90 % t");
                      data1.Add("90 % t");
                      data1.Add("");
                      styleList.Add(0);
                      styleList.Add(0);
                      styleList.Add(0);
                      styleList.Add(0);
                      styleList.Add(0);
                      styleList.Add(0);
                      createRowsTreatDataStyleConstructionPara(data1, styleList, 11);
                      data1.Clear();            styleList.Clear();

                      data1.Add("");
                      data1.Add("偏心度");
                      data1.Add("%");
                      data1.Add("10");
                      data1.Add("≤10");
                      data1.Add("");
                      styleList.Add(0);
                      styleList.Add(0);
                      styleList.Add(0);
                      styleList.Add(0);
                      styleList.Add(0);
                      styleList.Add(0);
                      createRowsTreatDataStyleConstructionPara(data1, styleList, 12);
                      data1.Clear();            styleList.Clear();

                      data1.Add("填充层");
                      data1.Add("填充材料");
                      data1.Add("");
                      data1.Add("（投标人提供）");
                      data1.Add("PP填充绳");// 待输入
                      data1.Add("");
                      styleList.Add(0);
                      styleList.Add(0);
                      styleList.Add(0);
                      styleList.Add(0);
                      styleList.Add(1);
                      styleList.Add(0);
                      createRowsTreatDataStyleConstructionPara(data1, styleList, 13);
                      data1.Clear();            styleList.Clear();

                      data1.Add("内衬层");
                      data1.Add("材料");
                      data1.Add("");
                      data1.Add("（投标人提供）");
                      data1.Add("H - 90 PVC护套料");  // 待输入
                      data1.Add("");
                      styleList.Add(0);
                      styleList.Add(0);
                      styleList.Add(0);
                      styleList.Add(0);
                      styleList.Add(1);
                      styleList.Add(0);
                      createRowsTreatDataStyleConstructionPara(data1, styleList, 14);
                      data1.Clear();            styleList.Clear();

                      data1.Add("");
                      data1.Add("厚度\r\n（依据GB/T 12706.1假定外径对应选取）");
                      data1.Add("mm");
                      data1.Add("（投标人提供）");
                      data1.Add("1.0 - 2.0");  // 待输入
                      data1.Add("");
                      styleList.Add(0);
                      styleList.Add(0);
                      styleList.Add(0);
                      styleList.Add(0);
                      styleList.Add(1);
                      styleList.Add(0);
                      createRowsTreatDataStyleConstructionPara(data1, styleList, 15);
                      data1.Clear();            styleList.Clear();


                      data1.Add("铠装层");
                      data1.Add("材料");
                      data1.Add("");
                      data1.Add("（投标人提供）");
                      data1.Add("单芯采用不锈钢带，\r\n多芯采用镀锌钢带");  // 待输入
                      data1.Add("与供货需求表一致");
                      styleList.Add(0);
                      styleList.Add(0);
                      styleList.Add(0);
                      styleList.Add(0);
                      styleList.Add(1);
                      styleList.Add(0);
                      createRowsTreatDataStyleConstructionPara(data1, styleList, 16);
                      data1.Clear();            styleList.Clear();

                      data1.Add("");
                      data1.Add("钢带厚度/钢丝直径\r\n（依据GB/T 12706.1假定外径对应选取）");
                      data1.Add("mm");
                      data1.Add("（投标人提供）");
                      data1.Add("0.2~0.5");  // 待输入
                      data1.Add("");
                      styleList.Add(0);
                      styleList.Add(0);
                      styleList.Add(0);
                      styleList.Add(0);
                      styleList.Add(1);
                      styleList.Add(0);
                      createRowsTreatDataStyleConstructionPara(data1, styleList, 17);
                      data1.Clear();            styleList.Clear();


                      data1.Add("");
                      data1.Add("钢带层数");
                      data1.Add("层");
                      data1.Add(2.0); // 待输入
                      data1.Add("2");
                      data1.Add("");
                      styleList.Add(0);
                      styleList.Add(0);
                      styleList.Add(0);
                      styleList.Add(1);
                      styleList.Add(1);
                      styleList.Add(0);
                      createRowsTreatDataStyleConstructionPara(data1, styleList, 18);
                      data1.Clear();            styleList.Clear();

                      data1.Add("");
                      data1.Add("钢带宽度");
                      data1.Add("mm");
                      data1.Add("（投标人提供）");
                      data1.Add(25); // 待输入
                      data1.Add("");
                      styleList.Add(0);
                      styleList.Add(0);
                      styleList.Add(0);
                      styleList.Add(0);
                      styleList.Add(1);
                      styleList.Add(0);
                      createRowsTreatDataStyleConstructionPara(data1, styleList, 19);
                      data1.Clear();            styleList.Clear();

                      data1.Add("外护套");
                      data1.Add("材料");
                      data1.Add("");
                      data1.Add("（投标人提供）");
                      data1.Add("ZH-90 PVC护套料"); // 待输入
                      data1.Add("");
                      styleList.Add(0);
                      styleList.Add(0);
                      styleList.Add(0);
                      styleList.Add(0);
                      styleList.Add(1);
                      styleList.Add(0);
                      createRowsTreatDataStyleConstructionPara(data1, styleList, 20);
                      data1.Clear();            styleList.Clear();

                      data1.Add("");
                      data1.Add("颜色");
                      data1.Add("（投标人提供）");
                      data1.Add("黑色");  //高亮颜色
                      data1.Add("黑色"); //高亮颜色
                      data1.Add("");
                      styleList.Add(0);
                      styleList.Add(0);
                      styleList.Add(0);
                      styleList.Add(1);
                      styleList.Add(1);
                      styleList.Add(0);
                      createRowsTreatDataStyleConstructionPara(data1, styleList, 21);
                      data1.Clear();            styleList.Clear();


                      data1.Add("");
                      data1.Add("标称厚度t（无铠装）");
                      data1.Add("mm");
                      data1.Add("（投标人提供）");
                      data1.Add(0.8); // 待输入
                      data1.Add("Z - YJV");// 待输入
                      styleList.Add(0);
                      styleList.Add(0);
                      styleList.Add(0);
                      styleList.Add(1);
                      styleList.Add(1);
                      styleList.Add(1);
                      createRowsTreatDataStyleConstructionPara(data1, styleList, 22);
                      data1.Clear();            styleList.Clear();

                      data1.Add("最薄点厚度不小于");
                      data1.Add("");
                      data1.Add("mm");
                      data1.Add("无铠85%，铠装80%");// 待输入
                      data1.Add("无铠85%，铠装80%");// 待输入
                      data1.Add("");
                      styleList.Add(0);
                      styleList.Add(0);
                      styleList.Add(0);
                      styleList.Add(0);
                      styleList.Add(0);
                      styleList.Add(0);
                      createRowsTreatDataStyleConstructionPara(data1, styleList, 23);
                      data1.Clear();            styleList.Clear();

                      data1.Add(" 电缆外径D：");
                      data1.Add("");
                      data1.Add("mm");
                      data1.Add("（投标人提供）");
                      data1.Add(19.74);// 待输入
                      data1.Add("Z - YJV22 0.6 / 1 3×10＋1×6");// 待输入
                      styleList.Add(0);
                      styleList.Add(0);
                      styleList.Add(0);
                      styleList.Add(0);
                      styleList.Add(1);
                      styleList.Add(1);
                      createRowsTreatDataStyleConstructionPara(data1, styleList, 24);
                      data1.Clear();            styleList.Clear();

                      data1.Add("20℃时铜导体最大直流电阻?");
                      data1.Add("");
                      data1.Add("欧姆/km");
                      data1.Add(3.08);// 待输入
                      data1.Add(3.08);// 待输入
                      data1.Add("对应6mm²截面");// 待输入
                      styleList.Add(0);
                      styleList.Add(0);
                      styleList.Add(0);
                      styleList.Add(1);
                      styleList.Add(1);
                      styleList.Add(1);
                      createRowsTreatDataStyleConstructionPara(data1, styleList, 25);
                      data1.Clear();            styleList.Clear();

                      data1.Add("");
                      data1.Add("");
                      data1.Add("");
                      data1.Add(1.83);// 待输入
                      data1.Add(1.83);// 待输入
                      data1.Add("错误对应10mm²截面");// 待输入
                      styleList.Add(0);
                      styleList.Add(0);
                      styleList.Add(0);
                      styleList.Add(1);
                      styleList.Add(1);
                      styleList.Add(1);
                      createRowsTreatDataStyleConstructionPara(data1, styleList, 26);
                      data1.Clear();            styleList.Clear();

                      data1.Add("90℃时铜导体最大交流电阻");
                      data1.Add("");
                      data1.Add("欧姆/千？");// 高亮
                      data1.Add("（投标人提供）");
                      data1.Add(3.9273);// 待输入
                      data1.Add("对应6mm²截面");// 待输入
                      styleList.Add(0);
                      styleList.Add(0);
                      styleList.Add(1);
                      styleList.Add(0);
                      styleList.Add(1);
                      styleList.Add(1);
                      createRowsTreatDataStyleConstructionPara(data1, styleList, 27);
                      data1.Clear();            styleList.Clear();

                      data1.Add("");
                      data1.Add("");
                      data1.Add("");
                      data1.Add("");
                      data1.Add(2.3334);// 待输入
                      data1.Add("对应6mm²截面");// 待输入
                      styleList.Add(0);
                      styleList.Add(0);
                      styleList.Add(0);
                      styleList.Add(0);
                      styleList.Add(1);
                      styleList.Add(1);
                      createRowsTreatDataStyleConstructionPara(data1, styleList, 28);
                      data1.Clear();            styleList.Clear();

                      data1.Add("电缆长期允许载流量\r\n（计算值，空气中40℃敷设）");
                      data1.Add("");
                      data1.Add("A");
                      data1.Add("（投标人提供）");
                      data1.Add(272);// 待输入
                      data1.Add("ZC-YJLV 0.6/1 4×150");// 待输入
                      styleList.Add(0);
                      styleList.Add(0);
                      styleList.Add(0);
                      styleList.Add(0);
                      styleList.Add(1);
                      styleList.Add(1);
                      createRowsTreatDataStyleConstructionPara(data1, styleList, 29);
                      data1.Clear();            styleList.Clear();

                      data1.Add("");
                      data1.Add("");
                      data1.Add("");
                      data1.Add("");
                      data1.Add("65");// 待输入
                      data1.Add("ZC-YJLV 0.6/1 4×16");// 待输入
                      styleList.Add(0);
                      styleList.Add(0);
                      styleList.Add(0);
                      styleList.Add(0);
                      styleList.Add(1);
                      styleList.Add(1);
                      createRowsTreatDataStyleConstructionPara(data1, styleList, 30); //需要2行。数据不全
                      data1.Clear();            styleList.Clear();

                      data1.Add("出厂工频电压试验");
                      data1.Add("");
                      data1.Add("kV/min");
                      data1.Add("3.5 U0/5");
                      data1.Add("3.5/5");
                      data1.Add("");
                      styleList.Add(0);
                      styleList.Add(0);
                      styleList.Add(0);
                      styleList.Add(0);
                      styleList.Add(0);
                      styleList.Add(0);
                      createRowsTreatDataStyleConstructionPara(data1, styleList, 31);
                      data1.Clear();            styleList.Clear();

                      data1.Add("电缆盘尺寸");
                      data1.Add("");
                      data1.Add("mm");
                      data1.Add("（投标人提供）");
                      data1.Add("根据订单长度选择");
                      data1.Add("");
                      data1.Add("");
                      styleList.Add(0);
                      styleList.Add(0);
                      styleList.Add(0);
                      styleList.Add(0);
                      styleList.Add(0);
                      styleList.Add(0);
                      createRowsTreatDataStyleConstructionPara(data1, styleList, 32);
                      data1.Clear();            styleList.Clear();

                      data1.Add("电缆敷设时的最大牵引力");
                      data1.Add("");
                      data1.Add("N/mm²");
                      data1.Add("70");// 待输入
                      data1.Add("70");// 待输入
                      data1.Add("铜芯，牵引头");// 待输入
                      styleList.Add(0);
                      styleList.Add(0);
                      styleList.Add(0);
                      styleList.Add(1);
                      styleList.Add(1);
                      styleList.Add(1);
                      createRowsTreatDataStyleConstructionPara(data1, styleList, 33);
                      data1.Clear();            styleList.Clear();

                      data1.Add("电缆敷设时的最大侧压力");
                      data1.Add("");
                      data1.Add("N/m");
                      data1.Add("5000");
                      data1.Add("5000");
                      data1.Add("");
                      styleList.Add(0);
                      styleList.Add(0);
                      styleList.Add(0);
                      styleList.Add(0);
                      styleList.Add(0);
                      styleList.Add(0);
                      createRowsTreatDataStyleConstructionPara(data1, styleList, 34);
                      data1.Clear();            styleList.Clear();

                      data1.Add("电缆质量（近似值）");
                      data1.Add("");
                      data1.Add("kg/m");
                      data1.Add("（投标人提供）");
                      data1.Add("2.7");// 待输入
                      data1.Add("ZC-YJLV 0.6/1 4×150");// 待输入
                      styleList.Add(0);
                      styleList.Add(0);
                      styleList.Add(0);
                      styleList.Add(0);
                      styleList.Add(1);
                      styleList.Add(1);
                      createRowsTreatDataStyleConstructionPara(data1, styleList, 35);
                      data1.Clear();            styleList.Clear();

                      data1.Add("");
                      data1.Add("");
                      data1.Add("");
                      data1.Add("");
                      data1.Add("14.2");// 待输入
                      data1.Add("ZC-YJV22 0.6/1 4×300");// 待输入
                      styleList.Add(0);
                      styleList.Add(0);
                      styleList.Add(0);
                      styleList.Add(0);
                      styleList.Add(0);
                      styleList.Add(0);
                      createRowsTreatDataStyleConstructionPara(data1, styleList, 36);
                      data1.Clear();            styleList.Clear();

                      data1.Add("电缆敷设时允许环境温度");
                      data1.Add("");
                      data1.Add("℃");
                      data1.Add("-5～＋40");
                      data1.Add("-5～＋40");
                      data1.Add("");
                      styleList.Add(0);
                      styleList.Add(0);
                      styleList.Add(0);
                      styleList.Add(0);
                      styleList.Add(0);
                      styleList.Add(0);
                      createRowsTreatDataStyleConstructionPara(data1, styleList, 37);
                      data1.Clear();            styleList.Clear();

                      data1.Add("电缆在正常使用条件下的寿命");
                      data1.Add("");
                      data1.Add("年");
                      data1.Add("≥30");
                      data1.Add("≥30");
                      data1.Add("");
                      styleList.Add(0);
                      styleList.Add(0);
                      styleList.Add(0);
                      styleList.Add(0);
                      styleList.Add(0);
                      styleList.Add(0);
                      createRowsTreatDataStyleConstructionPara(data1, styleList, 38);
                      data1.Clear();            styleList.Clear();


                      data1.Add("电缆阻燃级别");
                      data1.Add("");
                      data1.Add("");
                      data1.Add("按供货需求表");
                      data1.Add("ZC");
                      data1.Add("");
                      styleList.Add(0);
                      styleList.Add(0);
                      styleList.Add(0);
                      styleList.Add(0);
                      styleList.Add(0);
                      styleList.Add(0);
                      createRowsTreatDataStyleConstructionPara(data1, styleList, 39);
                      data1.Clear();            styleList.Clear();

                      data1.Add("电缆的无卤性能");
                      data1.Add("");
                      data1.Add("");
                      data1.Add("按供货需求表");
                      data1.Add("不适用");// 待输入
                      data1.Add("对应无卤、C级阻燃电缆");// 待输入
                      styleList.Add(0);
                      styleList.Add(0);
                      styleList.Add(0);
                      styleList.Add(0);
                      styleList.Add(1);
                      styleList.Add(1);
                      createRowsTreatDataStyleConstructionPara(data1, styleList, 40);
                      data1.Clear();            styleList.Clear();

                      data1.Add("电缆的低烟性能");
                      data1.Add("");
                      data1.Add("");
                      data1.Add("按供货需求表");
                      data1.Add("不适用");// 待输入
                      data1.Add("对应低烟、C级阻燃电缆");// 待输入
                      styleList.Add(0);
                      styleList.Add(0);
                      styleList.Add(0);
                      styleList.Add(0);
                      styleList.Add(1);
                      styleList.Add(1);
                      createRowsTreatDataStyleConstructionPara(data1, styleList, 41);
                      data1.Clear(); styleList.Clear();
                      */
                }
                void feedDataByRows3()// 未使用
                {
                    /*
                    // ICellStyle stringStyle = CreateStringStyle();
                    List<IRow> rows = [];
                    List<List<ICell>> sheetCells = [];

                    for (int r = 0; r <= 41; r++)  // 假设创建42行
                    {
                        // 创建新行并添加到rows列表
                        IRow iRow = sheetConstructionPara.CreateRow(r);
                        iRow.HeightInPoints = 15;
                        rows.Add(iRow);

                        // 为当前行创建新的单元格列表
                        List<ICell> rowCells = [];

                        for (int c = 0; c <= 5; c++)  // 假设每行6个单元格
                        {
                            // 创建单元格并添加到当前行的单元格列表
                            ICell cell = iRow.CreateCell(c);
                            cell.CellStyle = stringCenterStyle;
                            rowCells.Add(cell);
                        }
                        sheetCells.Add(rowCells);

                    }


                    //rows[1].GetCell(3).SetCellValue("电缆结构技术参数1");
                    //rows[1].CreateCell(0).SetCellValue("电缆型号");          
                    sheetCells[0][0].SetCellValue("电缆结构技术参数");

                    sheetCells[1][0].SetCellValue("电缆型号");
                    sheetCells[1][2].SetCellValue("Z-YJV22 0.6/1 3×10＋1×6");//待输入
                    sheetCells[1][2].CellStyle = warnStyle;
                    //
                    sheetCells[2][0].SetCellValue("项　　目");
                    sheetCells[2][2].SetCellValue("单位");
                    sheetCells[2][3].SetCellValue("标准参数值");
                    sheetCells[2][4].SetCellValue("投标人响应值");
                    sheetCells[2][5].SetCellValue("备注");

                    sheetCells[3][0].SetCellValue("铜导体");
                    sheetCells[3][2].SetCellValue("材料");
                    sheetCells[3][3].SetCellValue("铜");
                    sheetCells[3][4].SetCellValue("铜");

                    sheetCells[4][1].SetCellValue("芯数×标称截面");
                    sheetCells[4][2].SetCellValue("芯×mm²");
                    sheetCells[4][3].SetCellValue("3×10＋1×6"); // 待输入
                    sheetCells[4][4].SetCellValue("3×10＋1×6"); // 待输入

                    sheetCells[5][1].SetCellValue("结构形式");
                    sheetCells[5][2].SetCellValue("芯×mm²");
                    sheetCells[5][3].SetCellValue("紧压圆形 / 实心导体");
                    sheetCells[5][4].SetCellValue("紧压圆形 / 实心导体");

                    sheetCells[6][1].SetCellValue("最少单线根数");
                    sheetCells[6][2].SetCellValue("根");
                    sheetCells[6][3].SetCellValue(1); // 待输入
                    sheetCells[6][4].SetCellValue(1); // 待输入
                    sheetCells[6][5].SetCellValue("对应6mm²截面");  // 待输入

                    sheetCells[7][1].SetCellValue("导体外径（近似值）");
                    sheetCells[7][2].SetCellValue("mm");
                    sheetCells[7][3].SetCellValue("（投标人提供）");
                    sheetCells[7][4].SetCellValue(2.2);  // 待输入
                    sheetCells[7][5].SetCellValue("对应4mm²截面");  // 待输入


                    sheetCells[8][1].SetCellValue("紧压系数");
                    sheetCells[8][3].SetCellValue("≥0.9");
                    sheetCells[8][4].SetCellValue("≥0.9（对应紧压圆形导体结构）");

                    sheetCells[9][0].SetCellValue("绝缘");
                    sheetCells[9][1].SetCellValue("材料");
                    sheetCells[9][3].SetCellValue("XLPE"); // 待输入
                    sheetCells[9][4].SetCellValue("XLPE"); // 待输入

                    sheetCells[10][1].SetCellValue("平均厚度不小于标称厚度 t");
                    sheetCells[10][2].SetCellValue("mm");
                    sheetCells[10][3].SetCellValue("（投标人提供）");
                    sheetCells[10][4].SetCellValue(0.7); // 待输入
                    sheetCells[10][5].SetCellValue("对应6、10截面");  // 待输入

                    sheetCells[11][1].SetCellValue("最薄点厚度不小于标称值");
                    sheetCells[11][2].SetCellValue("mm");
                    sheetCells[11][3].SetCellValue("90 % t");
                    sheetCells[11][4].SetCellValue("90 % t");

                    sheetCells[12][1].SetCellValue("偏心度");
                    sheetCells[12][2].SetCellValue("%");
                    sheetCells[12][3].SetCellValue("10");
                    sheetCells[12][4].SetCellValue("≤10");

                    sheetCells[13][0].SetCellValue("填充层");
                    sheetCells[13][1].SetCellValue("填充材料");
                    sheetCells[13][3].SetCellValue("（投标人提供）");
                    sheetCells[13][4].SetCellValue("PP填充绳");

                    sheetCells[14][0].SetCellValue("内衬层");
                    sheetCells[14][1].SetCellValue("材料");
                    sheetCells[14][3].SetCellValue("（投标人提供）");
                    sheetCells[14][4].SetCellValue("H - 90 PVC护套料");  // 待输入

                    sheetCells[15][1].SetCellValue("厚度\r\n（依据GB/T 12706.1假定外径对应选取）");
                    sheetCells[15][2].SetCellValue("mm");
                    sheetCells[15][3].SetCellValue("（投标人提供）");
                    sheetCells[15][4].SetCellValue("1.0 - 2.0");  // 待输入


                    sheetCells[16][0].SetCellValue("铠装层");
                    sheetCells[16][1].SetCellValue("材料");
                    sheetCells[16][3].SetCellValue("（投标人提供）");
                    sheetCells[16][4].SetCellValue("单芯采用不锈钢带，\r\n多芯采用镀锌钢带");  // 待输入
                    sheetCells[16][5].SetCellValue("与供货需求表一致");


                    sheetCells[17][1].SetCellValue("钢带厚度/钢丝直径\r\n（依据GB/T 12706.1假定外径对应选取）");
                    sheetCells[17][2].SetCellValue("mm");
                    sheetCells[17][3].SetCellValue("（投标人提供）");
                    sheetCells[17][4].SetCellValue("0.2~0.5");  // 待输入


                    sheetCells[18][1].SetCellValue("钢带层数");
                    sheetCells[18][2].SetCellValue("层");
                    sheetCells[18][3].SetCellValue(2); // 待输入
                    sheetCells[18][4].SetCellValue(2); // 待输入

                    sheetCells[19][1].SetCellValue("钢带宽度");
                    sheetCells[19][2].SetCellValue("mm");
                    sheetCells[19][3].SetCellValue("（投标人提供）");
                    sheetCells[19][4].SetCellValue(25); // 待输入

                    sheetCells[20][0].SetCellValue("外护套");
                    sheetCells[20][1].SetCellValue("材料");
                    sheetCells[20][3].SetCellValue("（投标人提供）");
                    sheetCells[20][4].SetCellValue("PVC"); // 待输入
                    sheetCells[20][5].SetCellValue ("ZH-90 PVC护套料（氧指数大于等于30 %）"); // 待输入

                    sheetCells[21][1].SetCellValue("颜色");
                    sheetCells[21][3].SetCellValue("（投标人提供）");
                    sheetCells[21][4].SetCellValue("黑色");  //高亮颜色
                    sheetCells[21][5].SetCellValue("黑色"); //高亮颜色

                    sheetCells[22][1].SetCellValue("标称厚度t（无铠装）");
                    sheetCells[22][2].SetCellValue("mm");
                    sheetCells[22][3].SetCellValue("（投标人提供）");
                    sheetCells[22][4].SetCellValue(0.8); // 待输入
                    sheetCells[22][5].SetCellValue("Z - YJV"); //高亮颜色

                    sheetCells[23][0].SetCellValue("最薄点厚度不小于");
                    sheetCells[23][2].SetCellValue("mm");
                    sheetCells[23][3].SetCellValue("无铠85%，铠装80%");// 待输入
                    sheetCells[23][4].SetCellValue("无铠85%，铠装80%");// 待输入


                    sheetCells[24][0].SetCellValue(" 电缆外径D：");
                    sheetCells[24][2].SetCellValue("mm");
                    sheetCells[24][3].SetCellValue("（投标人提供）");
                    sheetCells[24][4].SetCellValue(19.74);// 待输入
                    sheetCells[24][5].SetCellValue(" Z - YJV22 0.6 / 1 3×10＋1×6");

                    sheetCells[25][0].SetCellValue("20℃时铜导体最大直流电阻");
                    sheetCells[25][2].SetCellValue("欧姆/km");
                    sheetCells[25][3].SetCellValue(3.08);// 待输入
                    sheetCells[25][4].SetCellValue(3.08);// 待输入
                    sheetCells[25][5].SetCellValue("对应6mm²截面");// 待输入

                    sheetCells[26][3].SetCellValue(1.83);// 待输入
                    sheetCells[26][4].SetCellValue(1.83);// 待输入
                    sheetCells[26][5].SetCellValue("对应10mm²截面");// 待输入

                    sheetCells[27][0].SetCellValue("90℃时铜导体最大交流电阻");
                    sheetCells[27][2].SetCellValue("欧姆/千？");// 高亮
                    sheetCells[27][3].SetCellValue("（投标人提供）");
                    sheetCells[27][4].SetCellValue(3.9273);// 待输入
                    sheetCells[27][5].SetCellValue("对应6mm²截面");// 待输入

                    sheetCells[28][4].SetCellValue(2.3334);// 待输入
                    sheetCells[28][5].SetCellValue("对应6mm²截面");// 待输入

                    sheetCells[29][0].SetCellValue("电缆长期允许载流量\r\n（计算值，空气中40℃敷设）");
                    sheetCells[29][1].SetCellValue("");
                    sheetCells[29][2].SetCellValue("A");
                    sheetCells[29][3].SetCellValue("（投标人提供）");
                    sheetCells[29][4].SetCellValue(272);// 待输入
                    sheetCells[29][5].SetCellValue("ZC-YJLV 0.6/1 4×150");// 待输入

                    sheetCells[30][4].SetCellValue("65");// 待输入
                    sheetCells[30][5].SetCellValue("ZC-YJLV 0.6/1 4×16");// 待输入


                    sheetCells[31][0].SetCellValue("出厂工频电压试验");
                    sheetCells[31][2].SetCellValue("kV/min");
                    sheetCells[31][3].SetCellValue("3.5 U0/5");
                    sheetCells[31][4].SetCellValue("3.5/5");

                    sheetCells[32][0].SetCellValue("电缆盘尺寸");
                    sheetCells[32][2].SetCellValue("mm");
                    sheetCells[32][3].SetCellValue("（投标人提供）");
                    sheetCells[32][4].SetCellValue("根据订单长度选择");

                    sheetCells[33][0].SetCellValue("电缆敷设时的最大牵引力");
                    sheetCells[33][2].SetCellValue("N/mm²");
                    sheetCells[33][3].SetCellValue("70");// 待输入
                    sheetCells[33][4].SetCellValue("70");// 待输入
                    sheetCells[33][5].SetCellValue("铜芯，牵引头");// 待输入

                    sheetCells[34][0].SetCellValue("电缆敷设时的最大侧压力");
                    sheetCells[34][2].SetCellValue("N/m");
                    sheetCells[34][3].SetCellValue("5000");
                    sheetCells[34][4].SetCellValue("5000");

                    sheetCells[35][0].SetCellValue("电缆质量（近似值）");
                    sheetCells[35][2].SetCellValue("kg/m");
                    sheetCells[35][3].SetCellValue("（投标人提供）");
                    sheetCells[35][4].SetCellValue("2.7");// 待输入
                    sheetCells[35][5].SetCellValue("ZC-YJLV 0.6/1 4×150");// 待输入

                    sheetCells[36][4].SetCellValue("14.2");// 待输入
                    sheetCells[36][5].SetCellValue("ZC-YJV22 0.6/1 4×300");// 待输入

                    sheetCells[37][0].SetCellValue("电缆敷设时允许环境温度");
                    sheetCells[37][2].SetCellValue("℃");
                    sheetCells[37][3].SetCellValue("-5～＋40");
                    sheetCells[37][4].SetCellValue("-5～＋40");

                    sheetCells[38][0].SetCellValue("电缆在正常使用条件下的寿命");
                    sheetCells[38][2].SetCellValue("年");
                    sheetCells[38][3].SetCellValue("≥30");
                    sheetCells[38][4].SetCellValue("≥30");

                    sheetCells[39][0].SetCellValue("电缆阻燃级别");
                    sheetCells[39][3].SetCellValue("按供货需求表");
                    sheetCells[39][4].SetCellValue("ZC");

                    sheetCells[40][0].SetCellValue("电缆的无卤性能");
                    sheetCells[40][3].SetCellValue("按供货需求表");
                    sheetCells[40][4].SetCellValue("不适用");// 待输入
                    sheetCells[40][5].SetCellValue("对应无卤、C级阻燃电缆");// 待输入

                    sheetCells[41][0].SetCellValue("电缆的低烟性能");
                    sheetCells[41][3].SetCellValue("按供货需求表");
                    sheetCells[41][4].SetCellValue("不适用");// 待输入
                    sheetCells[41][5].SetCellValue("对应低烟、C级阻燃电缆");// 待输入
                    //*/
                }

                void createRowsTreatDataStyleConstructionPara(List<KeyValuePair<int, object>> inputDataList, int rowNum) //方法1        
                {

                    // 3. 添加标题行
                    IRow iRow = sheetConstructionPara.CreateRow(rowNum);
                    sheetConstructionPara.AutoSizeRow(rowNum);
                    int colIdx = 0;
                    //for (int colIdx = 0; colIdx < 6; colIdx++)//方法2,3
                    foreach (var keyValuePair in inputDataList)//方法1
                    {
                        try
                        {
                            //ICell cell = iRow.GetCell(colIdx);//如果获得已有单元格，则这样写
                            ICell cell = iRow.CreateCell(colIdx);//IRow是地址引用，像指针，反过来赋值
                                                                 // cell.CellStyle = stringCenterStyle;
                                                                 //int styleInt = styleList[colIdx];//方法2
                            int styleInt = keyValuePair.Key;//方法1
                                                            // object value = inputDataList[colIdx];
                            object value = keyValuePair.Value;

                            // 根据数据类型应用样式
                            if (value == DBNull.Value || value == "")
                            {
                                // cell.SetCellValue("ut");// string.Empty;
                                // cell.CellStyle = nullStyle;// CreateNullStyle();
                                cell.SetCellValue("数据未提供");
                                cell.CellStyle = warnStyle;
                            }
                            else
                            {
                                switch (value)
                                {
                                    case string string1:
                                        cell.SetCellValue(Convert.ToString(value));
                                        // cell.CellStyle = stringStyle;
                                        break;
                                    case int int1:
                                        cell.SetCellValue(Convert.ToString(value));
                                        // cell.CellStyle = itemStyle;
                                        break;
                                    case DateTime dateTime1:
                                        cell.SetCellValue(Convert.ToDateTime(value));
                                        // cell.CellStyle = dateTimeStyle;
                                        break;
                                    case double double1:
                                        cell.SetCellValue(Convert.ToDouble(value));
                                        if (Convert.ToDouble(value) < 0)
                                        {
                                            // cell.CellStyle = warnNumStyle;//numberStyle; //
                                            break;
                                        }
                                        // cell.CellStyle = numberStyle;
                                        break;
                                    default:
                                        cell.SetCellValue(value.ToString());
                                        // cell.CellStyle = stringStyle; //文本自动换行
                                        break;
                                }
                                if (rowNum == 0) cell.CellStyle = titleStyle;
                                else if (faultMarkArray.Any(term => value.ToString().Contains(term, StringComparison.OrdinalIgnoreCase)))
                                    cell.CellStyle = warnStyle;
                                else if (styleInt == 1)//方法2
                                    cell.CellStyle = stringBlueStyle;
                                else if (value.ToString().Contains("电缆长期允许载流量", StringComparison.OrdinalIgnoreCase))
                                {
                                    //合并单元格，文字超出单元格范围，行高不会自动变化
                                    iRow.Height = 576;// 14.4*2 *20;1/20个点为最小单位
                                    cell.CellStyle = stringCenterStyle;
                                }
                                else
                                    cell.CellStyle = stringCenterStyle;

                            }
                            Trace.Write($"{value.ToString()}  ");//20250519 打印单元格数据                
                        }                             //                   }
                        catch (Exception excep1)
                        {
                            Trace.WriteLine(excep1.Message);
                            Trace.WriteLine($"问题在第{rowNum.ToString()}行  ");
                        }
                        ++colIdx;
                    }
                    Trace.WriteLine(""); ;
                }
                //void createRowsTreatDataStyleConstructionPara(List<Object> inputDataList, int rowIndex)//方法3
                //void createRowsTreatDataStyleConstructionPara(List<Object> inputDataList, List<int> styleList, int rowIndex) //方法2

                void setColumnWidthConstructionPara()
                {
                    sheetConstructionPara.SetColumnWidth(0, (7.11 + 0.78) * 256);
                    sheetConstructionPara.SetColumnWidth(1, (15.89 + 0.78) * 256);
                    sheetConstructionPara.SetColumnWidth(2, (7.11 + 0.78) * 256);
                    sheetConstructionPara.SetColumnWidth(3, (17.22 + 0.78) * 256);
                    sheetConstructionPara.SetColumnWidth(4, (19.22 + 0.78) * 256);
                    sheetConstructionPara.SetColumnWidth(5, (15.22 + 0.78) * 256);
                }

                //void mergeCellsConstructionPara(bool isDoubleCable) //合并单元格
                void mergeCellsConstructionPara() //合并单元格
                {

                    if (isDoubleCable)
                    {
                        if (isArmoured)
                        {
                            CellRangeAddress cellMerge = new CellRangeAddress(0, 0, 0, 5);//（起始行，结束行，起始列，结束列）
                            sheetConstructionPara.AddMergedRegion(cellMerge);
                            CellRangeAddress cellMerge1 = new CellRangeAddress(1, 1, 0, 1);//（起始行，结束行，起始列，结束列）         
                            sheetConstructionPara.AddMergedRegion(cellMerge1);
                            CellRangeAddress cellMerge2 = new CellRangeAddress(1, 1, 2, 4);//（起始行，结束行，起始列，结束列）
                            sheetConstructionPara.AddMergedRegion(cellMerge2);
                            /**/
                            CellRangeAddress cellMerge3 = new CellRangeAddress(2, 2, 0, 1);//（起始行，结束行，起始列，结束列）
                            sheetConstructionPara.AddMergedRegion(cellMerge3);
                            CellRangeAddress cellMerge4 = new CellRangeAddress(3, 10, 0, 0);//（起始行，结束行，起始列，结束列）
                            sheetConstructionPara.AddMergedRegion(cellMerge4);
                            CellRangeAddress cellMerge81 = new CellRangeAddress(6, 7, 1, 1);
                            sheetConstructionPara.AddMergedRegion(cellMerge81);
                            CellRangeAddress cellMerge82 = new CellRangeAddress(6, 7, 2, 2);
                            sheetConstructionPara.AddMergedRegion(cellMerge82);
                            CellRangeAddress cellMerge83 = new CellRangeAddress(8, 9, 1, 1);
                            sheetConstructionPara.AddMergedRegion(cellMerge83);
                            CellRangeAddress cellMerge84 = new CellRangeAddress(8, 9, 2, 2);
                            sheetConstructionPara.AddMergedRegion(cellMerge84);
                            CellRangeAddress cellMerge85 = new CellRangeAddress(8, 9, 3, 3);
                            sheetConstructionPara.AddMergedRegion(cellMerge85);
                            CellRangeAddress cellMerge86 = new CellRangeAddress(11, 15, 0, 0);
                            sheetConstructionPara.AddMergedRegion(cellMerge86);
                            CellRangeAddress cellMerge87 = new CellRangeAddress(12, 13, 1, 1);
                            sheetConstructionPara.AddMergedRegion(cellMerge87);
                            CellRangeAddress cellMerge88 = new CellRangeAddress(12, 13, 2, 2);
                            sheetConstructionPara.AddMergedRegion(cellMerge88);
                            CellRangeAddress cellMerge89 = new CellRangeAddress(12, 13, 3, 3);
                            sheetConstructionPara.AddMergedRegion(cellMerge89);
                            CellRangeAddress cellMerge6 = new CellRangeAddress(17, 18, 0, 0);
                            sheetConstructionPara.AddMergedRegion(cellMerge6);
                            CellRangeAddress cellMerge7 = new CellRangeAddress(19, 22, 0, 0);
                            sheetConstructionPara.AddMergedRegion(cellMerge7);
                            CellRangeAddress cellMerge8 = new CellRangeAddress(23, 26, 0, 0);
                            sheetConstructionPara.AddMergedRegion(cellMerge8);
                            // CellRangeAddress cellMerge9 = new CellRangeAddress(26, 26, 0, 1);
                            // sheetConstructionPara.AddMergedRegion(cellMerge9);
                            CellRangeAddress cellMerge10 = new CellRangeAddress(27, 27, 0, 1);
                            sheetConstructionPara.AddMergedRegion(cellMerge10);
                            CellRangeAddress cellMerge11 = new CellRangeAddress(28, 29, 0, 1);
                            sheetConstructionPara.AddMergedRegion(cellMerge11);
                            CellRangeAddress cellMerge12 = new CellRangeAddress(28, 29, 2, 2);
                            sheetConstructionPara.AddMergedRegion(cellMerge12);
                            CellRangeAddress cellMerge13 = new CellRangeAddress(30, 31, 0, 1);
                            sheetConstructionPara.AddMergedRegion(cellMerge13);
                            CellRangeAddress cellMerge14 = new CellRangeAddress(30, 31, 2, 2);
                            sheetConstructionPara.AddMergedRegion(cellMerge14);
                            CellRangeAddress cellMerge15 = new CellRangeAddress(30, 31, 3, 3);
                            sheetConstructionPara.AddMergedRegion(cellMerge15);
                            CellRangeAddress cellMerge17 = new CellRangeAddress(32, 32, 0, 1);
                            sheetConstructionPara.AddMergedRegion(cellMerge17);
                            /*     CellRangeAddress cellMerge18 = new CellRangeAddress(32, 32, 2, 2);
                                 sheetConstructionPara.AddMergedRegion(cellMerge18);
                                 CellRangeAddress cellMerge16 = new CellRangeAddress(32, 33, 3, 3);
                                 sheetConstructionPara.AddMergedRegion(cellMerge16);
                            */
                            CellRangeAddress cellMerge20 = new CellRangeAddress(33, 33, 0, 1);
                            sheetConstructionPara.AddMergedRegion(cellMerge20);
                            CellRangeAddress cellMerge21 = new CellRangeAddress(34, 34, 0, 1);
                            sheetConstructionPara.AddMergedRegion(cellMerge21);
                            CellRangeAddress cellMerge23 = new CellRangeAddress(35, 35, 0, 1);
                            sheetConstructionPara.AddMergedRegion(cellMerge23);
                            //   CellRangeAddress cellMerge24 = new CellRangeAddress(35, 36, 2, 2);
                            //    sheetConstructionPara.AddMergedRegion(cellMerge24);
                            //   CellRangeAddress cellMerge25 = new CellRangeAddress(35, 36, 3, 3);
                            //   sheetConstructionPara.AddMergedRegion(cellMerge25);
                            CellRangeAddress cellMerge26 = new CellRangeAddress(36, 36, 0, 1);
                            sheetConstructionPara.AddMergedRegion(cellMerge26);
                            CellRangeAddress cellMerge27 = new CellRangeAddress(37, 37, 0, 1);
                            sheetConstructionPara.AddMergedRegion(cellMerge27);
                            CellRangeAddress cellMerge28 = new CellRangeAddress(38, 38, 0, 1);
                            sheetConstructionPara.AddMergedRegion(cellMerge28);
                            CellRangeAddress cellMerge29 = new CellRangeAddress(39, 39, 0, 1);
                            sheetConstructionPara.AddMergedRegion(cellMerge29);
                            CellRangeAddress cellMerge30 = new CellRangeAddress(40, 40, 0, 1);
                            sheetConstructionPara.AddMergedRegion(cellMerge30);
                            CellRangeAddress cellMerge22 = new CellRangeAddress(41, 41, 0, 1);
                            sheetConstructionPara.AddMergedRegion(cellMerge22);
                            CellRangeAddress cellMerge19 = new CellRangeAddress(42, 42, 0, 1);
                            sheetConstructionPara.AddMergedRegion(cellMerge19);
                        }
                        else
                        {
                            CellRangeAddress cellMerge = new CellRangeAddress(0, 0, 0, 5);//（起始行，结束行，起始列，结束列）
                            sheetConstructionPara.AddMergedRegion(cellMerge);
                            CellRangeAddress cellMerge1 = new CellRangeAddress(1, 1, 0, 1);//（起始行，结束行，起始列，结束列）         
                            sheetConstructionPara.AddMergedRegion(cellMerge1);
                            CellRangeAddress cellMerge2 = new CellRangeAddress(1, 1, 2, 4);//（起始行，结束行，起始列，结束列）
                            sheetConstructionPara.AddMergedRegion(cellMerge2);
                            CellRangeAddress cellMerge3 = new CellRangeAddress(2, 2, 0, 1);//（起始行，结束行，起始列，结束列）
                            sheetConstructionPara.AddMergedRegion(cellMerge3);
                            CellRangeAddress cellMerge4 = new CellRangeAddress(3, 10, 0, 0);//（起始行，结束行，起始列，结束列）
                            sheetConstructionPara.AddMergedRegion(cellMerge4);
                            CellRangeAddress cellMerge81 = new CellRangeAddress(6, 7, 1, 1);
                            sheetConstructionPara.AddMergedRegion(cellMerge81);
                            CellRangeAddress cellMerge82 = new CellRangeAddress(6, 7, 2, 2);
                            sheetConstructionPara.AddMergedRegion(cellMerge82);
                            CellRangeAddress cellMerge83 = new CellRangeAddress(8, 9, 1, 1);
                            sheetConstructionPara.AddMergedRegion(cellMerge83);
                            CellRangeAddress cellMerge84 = new CellRangeAddress(8, 9, 2, 2);
                            sheetConstructionPara.AddMergedRegion(cellMerge84);
                            CellRangeAddress cellMerge85 = new CellRangeAddress(8, 9, 3, 3);
                            sheetConstructionPara.AddMergedRegion(cellMerge85);
                            CellRangeAddress cellMerge86 = new CellRangeAddress(11, 15, 0, 0);
                            sheetConstructionPara.AddMergedRegion(cellMerge86);
                            CellRangeAddress cellMerge87 = new CellRangeAddress(12, 13, 1, 1);
                            sheetConstructionPara.AddMergedRegion(cellMerge87);
                            CellRangeAddress cellMerge88 = new CellRangeAddress(12, 13, 2, 2);
                            sheetConstructionPara.AddMergedRegion(cellMerge88);
                            CellRangeAddress cellMerge89 = new CellRangeAddress(12, 13, 3, 3);
                            sheetConstructionPara.AddMergedRegion(cellMerge89);
                            CellRangeAddress cellMerge6 = new CellRangeAddress(17, 20, 0, 0);
                            sheetConstructionPara.AddMergedRegion(cellMerge6);
                            CellRangeAddress cellMerge27 = new CellRangeAddress(21, 21, 0, 1);
                            sheetConstructionPara.AddMergedRegion(cellMerge27);
                            CellRangeAddress cellMerge28 = new CellRangeAddress(22, 23, 0, 1);
                            sheetConstructionPara.AddMergedRegion(cellMerge28);
                            CellRangeAddress cellMerge29 = new CellRangeAddress(24, 25, 0, 1);
                            sheetConstructionPara.AddMergedRegion(cellMerge29);
                            CellRangeAddress cellMerge30 = new CellRangeAddress(26, 26, 0, 1);
                            sheetConstructionPara.AddMergedRegion(cellMerge30);
                            CellRangeAddress cellMerge22 = new CellRangeAddress(27, 27, 0, 1);
                            sheetConstructionPara.AddMergedRegion(cellMerge22);
                            CellRangeAddress cellMerge19 = new CellRangeAddress(28, 28, 0, 1);
                            sheetConstructionPara.AddMergedRegion(cellMerge19);
                            CellRangeAddress cellMerge7 = new CellRangeAddress(29, 29, 0, 1);
                            sheetConstructionPara.AddMergedRegion(cellMerge7);
                            CellRangeAddress cellMerge8 = new CellRangeAddress(30, 30, 0, 1);
                            sheetConstructionPara.AddMergedRegion(cellMerge8);
                            CellRangeAddress cellMerge10 = new CellRangeAddress(31, 31, 0, 1);
                            sheetConstructionPara.AddMergedRegion(cellMerge10);
                            CellRangeAddress cellMerge17 = new CellRangeAddress(32, 32, 0, 1);
                            sheetConstructionPara.AddMergedRegion(cellMerge17);
                            CellRangeAddress cellMerge20 = new CellRangeAddress(33, 33, 0, 1);
                            sheetConstructionPara.AddMergedRegion(cellMerge20);
                            CellRangeAddress cellMerge21 = new CellRangeAddress(34, 34, 0, 1);
                            sheetConstructionPara.AddMergedRegion(cellMerge21);
                            CellRangeAddress cellMerge23 = new CellRangeAddress(35, 35, 0, 1);
                            sheetConstructionPara.AddMergedRegion(cellMerge23);
                            CellRangeAddress cellMerge26 = new CellRangeAddress(36, 36, 0, 1);
                            sheetConstructionPara.AddMergedRegion(cellMerge26);
                            /*
                             *                         //   CellRangeAddress cellMerge24 = new CellRangeAddress(35, 36, 2, 2);
                            //    sheetConstructionPara.AddMergedRegion(cellMerge24);
                            //   CellRangeAddress cellMerge25 = new CellRangeAddress(35, 36, 3, 3);
                            //   sheetConstructionPara.AddMergedRegion(cellMerge25);
                            CellRangeAddress cellMerge11 = new CellRangeAddress(32, 32, 0, 1);
                            sheetConstructionPara.AddMergedRegion(cellMerge11);
                            CellRangeAddress cellMerge12 = new CellRangeAddress(28, 29, 2, 2);
                            sheetConstructionPara.AddMergedRegion(cellMerge12);
                            CellRangeAddress cellMerge13 = new CellRangeAddress(30, 31, 0, 1);
                            sheetConstructionPara.AddMergedRegion(cellMerge13);
                            CellRangeAddress cellMerge14 = new CellRangeAddress(30, 31, 2, 2);
                            sheetConstructionPara.AddMergedRegion(cellMerge14);
                            CellRangeAddress cellMerge15 = new CellRangeAddress(30, 31, 3, 3);
                            sheetConstructionPara.AddMergedRegion(cellMerge15);
                                                    /*     CellRangeAddress cellMerge18 = new CellRangeAddress(32, 32, 2, 2);
                                 sheetConstructionPara.AddMergedRegion(cellMerge18);
                                 CellRangeAddress cellMerge16 = new CellRangeAddress(32, 33, 3, 3);
                                 sheetConstructionPara.AddMergedRegion(cellMerge16);
                                                    // CellRangeAddress cellMerge9 = new CellRangeAddress(26, 26, 0, 1);
                            // sheetConstructionPara.AddMergedRegion(cellMerge9);
                            */

                        }
                    }
                    else if (isMultiCore)
                    {
                        if (isArmoured)
                        {
                            CellRangeAddress cellMerge = new CellRangeAddress(0, 0, 0, 5);//（起始行，结束行，起始列，结束列）
                            sheetConstructionPara.AddMergedRegion(cellMerge);
                            CellRangeAddress cellMerge1 = new CellRangeAddress(1, 1, 0, 1);//（起始行，结束行，起始列，结束列）         
                            sheetConstructionPara.AddMergedRegion(cellMerge1);
                            CellRangeAddress cellMerge2 = new CellRangeAddress(1, 1, 2, 4);//（起始行，结束行，起始列，结束列）
                            sheetConstructionPara.AddMergedRegion(cellMerge2);
                            CellRangeAddress cellMerge3 = new CellRangeAddress(2, 2, 0, 1);//（起始行，结束行，起始列，结束列）
                            sheetConstructionPara.AddMergedRegion(cellMerge3);
                            CellRangeAddress cellMerge4 = new CellRangeAddress(3, 8, 0, 0);//（起始行，结束行，起始列，结束列）
                            sheetConstructionPara.AddMergedRegion(cellMerge4);
                            CellRangeAddress cellMerge81 = new CellRangeAddress(9, 12, 0, 0);
                            sheetConstructionPara.AddMergedRegion(cellMerge81);
                            CellRangeAddress cellMerge82 = new CellRangeAddress(14, 15, 0, 0);
                            sheetConstructionPara.AddMergedRegion(cellMerge82);
                            CellRangeAddress cellMerge83 = new CellRangeAddress(16, 19, 0, 0);
                            sheetConstructionPara.AddMergedRegion(cellMerge83);
                            CellRangeAddress cellMerge84 = new CellRangeAddress(20, 23, 0, 0);
                            sheetConstructionPara.AddMergedRegion(cellMerge84);
                            CellRangeAddress cellMerge28 = new CellRangeAddress(24, 24, 0, 1);
                            sheetConstructionPara.AddMergedRegion(cellMerge28);
                            CellRangeAddress cellMerge29 = new CellRangeAddress(25, 25, 0, 1);
                            sheetConstructionPara.AddMergedRegion(cellMerge29);
                            CellRangeAddress cellMerge22 = new CellRangeAddress(26, 26, 0, 1);
                            sheetConstructionPara.AddMergedRegion(cellMerge22);
                            CellRangeAddress cellMerge19 = new CellRangeAddress(27, 27, 0, 1);
                            sheetConstructionPara.AddMergedRegion(cellMerge19);
                            CellRangeAddress cellMerge20 = new CellRangeAddress(28, 28, 0, 1);
                            sheetConstructionPara.AddMergedRegion(cellMerge20);
                            CellRangeAddress cellMerge10 = new CellRangeAddress(29, 29, 0, 1);
                            sheetConstructionPara.AddMergedRegion(cellMerge10);
                            CellRangeAddress cellMerge11 = new CellRangeAddress(30, 30, 0, 1);
                            sheetConstructionPara.AddMergedRegion(cellMerge11);
                            CellRangeAddress cellMerge9 = new CellRangeAddress(31, 31, 0, 1);
                            sheetConstructionPara.AddMergedRegion(cellMerge9);
                            CellRangeAddress cellMerge17 = new CellRangeAddress(32, 32, 0, 1);
                            sheetConstructionPara.AddMergedRegion(cellMerge17);
                            CellRangeAddress cellMerge13 = new CellRangeAddress(33, 33, 0, 1);
                            sheetConstructionPara.AddMergedRegion(cellMerge13);
                            CellRangeAddress cellMerge21 = new CellRangeAddress(34, 34, 0, 1);
                            sheetConstructionPara.AddMergedRegion(cellMerge21);
                            CellRangeAddress cellMerge23 = new CellRangeAddress(35, 35, 0, 1);
                            sheetConstructionPara.AddMergedRegion(cellMerge23);
                            CellRangeAddress cellMerge26 = new CellRangeAddress(36, 36, 0, 1);
                            sheetConstructionPara.AddMergedRegion(cellMerge26);
                            CellRangeAddress cellMerge30 = new CellRangeAddress(37, 37, 0, 1);
                            sheetConstructionPara.AddMergedRegion(cellMerge30);
                        }
                        else
                        {
                            CellRangeAddress cellMerge = new CellRangeAddress(0, 0, 0, 5);//（起始行，结束行，起始列，结束列）
                            sheetConstructionPara.AddMergedRegion(cellMerge);
                            CellRangeAddress cellMerge1 = new CellRangeAddress(1, 1, 0, 1);//（起始行，结束行，起始列，结束列）         
                            sheetConstructionPara.AddMergedRegion(cellMerge1);
                            CellRangeAddress cellMerge2 = new CellRangeAddress(1, 1, 2, 4);//（起始行，结束行，起始列，结束列）
                            sheetConstructionPara.AddMergedRegion(cellMerge2);
                            CellRangeAddress cellMerge3 = new CellRangeAddress(2, 2, 0, 1);//（起始行，结束行，起始列，结束列）
                            sheetConstructionPara.AddMergedRegion(cellMerge3);
                            CellRangeAddress cellMerge4 = new CellRangeAddress(3, 8, 0, 0);//（起始行，结束行，起始列，结束列）
                            sheetConstructionPara.AddMergedRegion(cellMerge4);
                            CellRangeAddress cellMerge81 = new CellRangeAddress(9, 12, 0, 0);
                            sheetConstructionPara.AddMergedRegion(cellMerge81);
                            CellRangeAddress cellMerge82 = new CellRangeAddress(14, 17, 0, 0);
                            sheetConstructionPara.AddMergedRegion(cellMerge82);

                            CellRangeAddress cellMerge9 = new CellRangeAddress(18, 18, 0, 1);
                            sheetConstructionPara.AddMergedRegion(cellMerge9);

                            CellRangeAddress cellMerge13 = new CellRangeAddress(19, 19, 0, 1);
                            sheetConstructionPara.AddMergedRegion(cellMerge13);
                            CellRangeAddress cellMerge21 = new CellRangeAddress(20, 20, 0, 1);
                            sheetConstructionPara.AddMergedRegion(cellMerge21);
                            CellRangeAddress cellMerge23 = new CellRangeAddress(21, 21, 0, 1);
                            sheetConstructionPara.AddMergedRegion(cellMerge23);
                            CellRangeAddress cellMerge26 = new CellRangeAddress(22, 22, 0, 1);
                            sheetConstructionPara.AddMergedRegion(cellMerge26);
                            CellRangeAddress cellMerge27 = new CellRangeAddress(23, 23, 0, 1);
                            sheetConstructionPara.AddMergedRegion(cellMerge27);

                            CellRangeAddress cellMerge29 = new CellRangeAddress(24, 24, 0, 1);
                            sheetConstructionPara.AddMergedRegion(cellMerge29);
                            CellRangeAddress cellMerge30 = new CellRangeAddress(25, 25, 0, 1);
                            sheetConstructionPara.AddMergedRegion(cellMerge30);
                            CellRangeAddress cellMerge22 = new CellRangeAddress(26, 26, 0, 1);
                            sheetConstructionPara.AddMergedRegion(cellMerge22);
                            CellRangeAddress cellMerge19 = new CellRangeAddress(27, 27, 0, 1);
                            sheetConstructionPara.AddMergedRegion(cellMerge19);
                            CellRangeAddress cellMerge20 = new CellRangeAddress(28, 28, 0, 1);
                            sheetConstructionPara.AddMergedRegion(cellMerge20);
                            CellRangeAddress cellMerge10 = new CellRangeAddress(29, 29, 0, 1);
                            sheetConstructionPara.AddMergedRegion(cellMerge10);
                            CellRangeAddress cellMerge11 = new CellRangeAddress(30, 30, 0, 1);
                            sheetConstructionPara.AddMergedRegion(cellMerge11);
                            CellRangeAddress cellMerge17 = new CellRangeAddress(31, 31, 0, 1);
                            sheetConstructionPara.AddMergedRegion(cellMerge17);

                        }
                    }
                    else
                    {
                        if (isArmoured)
                        {
                            CellRangeAddress cellMerge = new CellRangeAddress(0, 0, 0, 5);//（起始行，结束行，起始列，结束列）
                            sheetConstructionPara.AddMergedRegion(cellMerge);
                            CellRangeAddress cellMerge1 = new CellRangeAddress(1, 1, 0, 1);//（起始行，结束行，起始列，结束列）         
                            sheetConstructionPara.AddMergedRegion(cellMerge1);
                            CellRangeAddress cellMerge2 = new CellRangeAddress(1, 1, 2, 4);//（起始行，结束行，起始列，结束列）
                            sheetConstructionPara.AddMergedRegion(cellMerge2);
                            CellRangeAddress cellMerge3 = new CellRangeAddress(2, 2, 0, 1);//（起始行，结束行，起始列，结束列）
                            sheetConstructionPara.AddMergedRegion(cellMerge3);
                            CellRangeAddress cellMerge4 = new CellRangeAddress(3, 8, 0, 0);//（起始行，结束行，起始列，结束列）
                            sheetConstructionPara.AddMergedRegion(cellMerge4);
                            CellRangeAddress cellMerge81 = new CellRangeAddress(9, 12, 0, 0);
                            sheetConstructionPara.AddMergedRegion(cellMerge81);
                            CellRangeAddress cellMerge82 = new CellRangeAddress(13, 14, 0, 0);
                            sheetConstructionPara.AddMergedRegion(cellMerge82);
                            CellRangeAddress cellMerge83 = new CellRangeAddress(15, 18, 0, 0);
                            sheetConstructionPara.AddMergedRegion(cellMerge83);
                            CellRangeAddress cellMerge84 = new CellRangeAddress(19, 22, 0, 0);
                            sheetConstructionPara.AddMergedRegion(cellMerge84);
                            CellRangeAddress cellMerge28 = new CellRangeAddress(23, 23, 0, 1);
                            sheetConstructionPara.AddMergedRegion(cellMerge28);
                            CellRangeAddress cellMerge29 = new CellRangeAddress(24, 24, 0, 1);
                            sheetConstructionPara.AddMergedRegion(cellMerge29);
                            CellRangeAddress cellMerge30 = new CellRangeAddress(25, 25, 0, 1);
                            sheetConstructionPara.AddMergedRegion(cellMerge30);
                            CellRangeAddress cellMerge22 = new CellRangeAddress(26, 26, 0, 1);
                            sheetConstructionPara.AddMergedRegion(cellMerge22);
                            CellRangeAddress cellMerge19 = new CellRangeAddress(27, 27, 0, 1);
                            sheetConstructionPara.AddMergedRegion(cellMerge19);
                            CellRangeAddress cellMerge20 = new CellRangeAddress(28, 28, 0, 1);
                            sheetConstructionPara.AddMergedRegion(cellMerge20);
                            CellRangeAddress cellMerge10 = new CellRangeAddress(29, 29, 0, 1);
                            sheetConstructionPara.AddMergedRegion(cellMerge10);
                            CellRangeAddress cellMerge11 = new CellRangeAddress(30, 30, 0, 1);
                            sheetConstructionPara.AddMergedRegion(cellMerge11);
                            CellRangeAddress cellMerge9 = new CellRangeAddress(31, 31, 0, 1);
                            sheetConstructionPara.AddMergedRegion(cellMerge9);
                            CellRangeAddress cellMerge17 = new CellRangeAddress(32, 32, 0, 1);
                            sheetConstructionPara.AddMergedRegion(cellMerge17);
                            CellRangeAddress cellMerge13 = new CellRangeAddress(33, 33, 0, 1);
                            sheetConstructionPara.AddMergedRegion(cellMerge13);
                            CellRangeAddress cellMerge21 = new CellRangeAddress(34, 34, 0, 1);
                            sheetConstructionPara.AddMergedRegion(cellMerge21);
                            CellRangeAddress cellMerge23 = new CellRangeAddress(35, 35, 0, 1);
                            sheetConstructionPara.AddMergedRegion(cellMerge23);
                            CellRangeAddress cellMerge26 = new CellRangeAddress(36, 36, 0, 1);
                            sheetConstructionPara.AddMergedRegion(cellMerge26);
                            //CellRangeAddress cellMerge27 = new CellRangeAddress(37, 37, 0, 1);
                            //sheetConstructionPara.AddMergedRegion(cellMerge27);
                        }
                        else
                        {
                            CellRangeAddress cellMerge = new CellRangeAddress(0, 0, 0, 5);//（起始行，结束行，起始列，结束列）
                            sheetConstructionPara.AddMergedRegion(cellMerge);
                            CellRangeAddress cellMerge1 = new CellRangeAddress(1, 1, 0, 1);//（起始行，结束行，起始列，结束列）         
                            sheetConstructionPara.AddMergedRegion(cellMerge1);
                            CellRangeAddress cellMerge2 = new CellRangeAddress(1, 1, 2, 4);//（起始行，结束行，起始列，结束列）
                            sheetConstructionPara.AddMergedRegion(cellMerge2);
                            CellRangeAddress cellMerge3 = new CellRangeAddress(2, 2, 0, 1);//（起始行，结束行，起始列，结束列）
                            sheetConstructionPara.AddMergedRegion(cellMerge3);
                            CellRangeAddress cellMerge4 = new CellRangeAddress(3, 8, 0, 0);//（起始行，结束行，起始列，结束列）
                            sheetConstructionPara.AddMergedRegion(cellMerge4);
                            CellRangeAddress cellMerge81 = new CellRangeAddress(9, 12, 0, 0);
                            sheetConstructionPara.AddMergedRegion(cellMerge81);
                            CellRangeAddress cellMerge82 = new CellRangeAddress(13, 16, 0, 0);
                            sheetConstructionPara.AddMergedRegion(cellMerge82);

                            CellRangeAddress cellMerge9 = new CellRangeAddress(17, 17, 0, 1);
                            sheetConstructionPara.AddMergedRegion(cellMerge9);
                            CellRangeAddress cellMerge17 = new CellRangeAddress(18, 18, 0, 1);
                            sheetConstructionPara.AddMergedRegion(cellMerge17);
                            CellRangeAddress cellMerge13 = new CellRangeAddress(19, 19, 0, 1);
                            sheetConstructionPara.AddMergedRegion(cellMerge13);
                            CellRangeAddress cellMerge21 = new CellRangeAddress(20, 20, 0, 1);
                            sheetConstructionPara.AddMergedRegion(cellMerge21);
                            CellRangeAddress cellMerge23 = new CellRangeAddress(21, 21, 0, 1);
                            sheetConstructionPara.AddMergedRegion(cellMerge23);
                            CellRangeAddress cellMerge26 = new CellRangeAddress(22, 22, 0, 1);
                            sheetConstructionPara.AddMergedRegion(cellMerge26);
                            CellRangeAddress cellMerge27 = new CellRangeAddress(23, 23, 0, 1);
                            sheetConstructionPara.AddMergedRegion(cellMerge27);

                            CellRangeAddress cellMerge29 = new CellRangeAddress(24, 24, 0, 1);
                            sheetConstructionPara.AddMergedRegion(cellMerge29);
                            CellRangeAddress cellMerge30 = new CellRangeAddress(25, 25, 0, 1);
                            sheetConstructionPara.AddMergedRegion(cellMerge30);
                            CellRangeAddress cellMerge22 = new CellRangeAddress(26, 26, 0, 1);
                            sheetConstructionPara.AddMergedRegion(cellMerge22);
                            CellRangeAddress cellMerge19 = new CellRangeAddress(27, 27, 0, 1);
                            sheetConstructionPara.AddMergedRegion(cellMerge19);
                            CellRangeAddress cellMerge20 = new CellRangeAddress(28, 28, 0, 1);
                            sheetConstructionPara.AddMergedRegion(cellMerge20);
                            CellRangeAddress cellMerge10 = new CellRangeAddress(29, 29, 0, 1);
                            sheetConstructionPara.AddMergedRegion(cellMerge10);
                            CellRangeAddress cellMerge11 = new CellRangeAddress(30, 30, 0, 1);
                            sheetConstructionPara.AddMergedRegion(cellMerge11);


                        }
                        /*
                         * 

                            CellRangeAddress cellMerge83 = new CellRangeAddress(16, 19, 0, 0);
                            sheetConstructionPara.AddMergedRegion(cellMerge83);
                            CellRangeAddress cellMerge84 = new CellRangeAddress(20, 23, 0, 0);
                            sheetConstructionPara.AddMergedRegion(cellMerge84);
                            // CellRangeAddress cellMerge28 = new CellRangeAddress(23, 23, 0, 1);
                            // sheetConstructionPara.AddMergedRegion(cellMerge28);
                        CellRangeAddress cellMerge85 = new CellRangeAddress(8, 9, 3, 3);
                        sheetConstructionPara.AddMergedRegion(cellMerge85);
                        CellRangeAddress cellMerge86 = new CellRangeAddress(11, 15, 0, 0);
                        sheetConstructionPara.AddMergedRegion(cellMerge86);
                        CellRangeAddress cellMerge87 = new CellRangeAddress(12, 13, 1, 1);
                        sheetConstructionPara.AddMergedRegion(cellMerge87);
                        CellRangeAddress cellMerge88 = new CellRangeAddress(12, 13, 2, 2);
                        sheetConstructionPara.AddMergedRegion(cellMerge88);
                        CellRangeAddress cellMerge89 = new CellRangeAddress(12, 13, 3, 3);
                        sheetConstructionPara.AddMergedRegion(cellMerge89);
                        CellRangeAddress cellMerge6 = new CellRangeAddress(17, 18, 0, 0);
                        sheetConstructionPara.AddMergedRegion(cellMerge6);
                        CellRangeAddress cellMerge7 = new CellRangeAddress(19, 22, 0, 0);
                        sheetConstructionPara.AddMergedRegion(cellMerge7);
                        CellRangeAddress cellMerge8 = new CellRangeAddress(23, 25, 0, 0);
                        sheetConstructionPara.AddMergedRegion(cellMerge8);


                        CellRangeAddress cellMerge12 = new CellRangeAddress(28, 29, 2, 2);
                        sheetConstructionPara.AddMergedRegion(cellMerge12);

                        CellRangeAddress cellMerge14 = new CellRangeAddress(30, 31, 2, 2);
                        sheetConstructionPara.AddMergedRegion(cellMerge14);
                        CellRangeAddress cellMerge15 = new CellRangeAddress(30, 31, 3, 3);
                        sheetConstructionPara.AddMergedRegion(cellMerge15);

                        CellRangeAddress cellMerge18 = new CellRangeAddress(32, 33, 2, 2);
                        sheetConstructionPara.AddMergedRegion(cellMerge18);
                        CellRangeAddress cellMerge16 = new CellRangeAddress(32, 33, 3, 3);
                        sheetConstructionPara.AddMergedRegion(cellMerge16);

                        //   CellRangeAddress cellMerge24 = new CellRangeAddress(35, 36, 2, 2);
                        //    sheetConstructionPara.AddMergedRegion(cellMerge24);
                        //   CellRangeAddress cellMerge25 = new CellRangeAddress(35, 36, 3, 3);
                        //   sheetConstructionPara.AddMergedRegion(cellMerge25);

                        */
                    }
                }

            }



            void createSheetNonElectricPara()
            {

                feedDataByRowsForNonElectricPara();
                setColumnWidthNonElectricPara();
                mergeCellsNonElectricPara();

                void feedDataByRowsForNonElectricPara()
                {
                    // 待删除
                    #region
                    /*
                    string forhead = Regex.Match(theType, @"(\w+)\-").Groups[1].Value;
                    string flameRedartant = Regex.Replace(forhead, @"w|d|W|D", "");
                    if (flameRedartant == "") flameRedartant = "不适用";
                    Trace.WriteLine($"flameRedartant: {flameRedartant}   ");

                    string halogenFree = (Regex.Match(forhead, @"(w|W)").Value);
                    if (halogenFree == "") halogenFree = "不适用";
                    Trace.WriteLine($"halogenFree: {halogenFree}");
                    string smokeFree = Regex.Match(forhead, @"(d|D)").Value;
                    if (smokeFree == "") smokeFree = "不适用";
                    Trace.WriteLine($"smokingFree: {smokeFree}");
                    */
                    /*
                                    MatchCollection matchCollection = Regex.Matches(spec, @"(\d×)|(\dX)|(\dx)");
                                    foreach (MatchRegex match1 in matchCollection)
                                    {
                                        MatchRegex theMatch = Regex.Match(match1.Value, @"\d");
                                        //   Trace.WriteLine($"数字：{theMatch.Value}");
                                        isMultiCore = (Convert.ToInt16(theMatch.Value) > 1);
                                        Trace.WriteLine($"theMatch.Value: {Convert.ToInt16(theMatch.Value)}");
                                        // specMini.Add(match1.Value);
                                    }*/
                    /*   string outer_sheathMaterialFront = Regex.Match(outer_sheathMaterialSelected, @"(.+)（").Groups[1].Value;
                       Trace.WriteLine($"outer_sheathMaterialFront: {outer_sheathMaterialFront}");
                    */
                    /*     MatchCollection matchCollectionRear = Regex.Matches(spec, @"(×\d+)|(X\d+)|(x\d+)");
                         //int k = 0;
                         List<string> areaConductor = [];
                         foreach (MatchRegex match1 in matchCollectionRear)
                         {
                             MatchRegex theMatch = Regex.Match(match1.Value, @"\d+");
                             areaConductor.Add($"对应{theMatch.Value} mm²截面");
                             //Trace.WriteLine($"{areaConductor[k++]}");
                         }

                         */
                    #endregion


                    int rowOutNum = 0;
                    var pairList = new List<KeyValuePair<int, object>>();
                    // 添加元素					
                    pairList.Add(new KeyValuePair<int, object>(0, "电缆非电气技术参数"));
                    pairList.Add(new KeyValuePair<int, object>(0, "——"));
                    pairList.Add(new KeyValuePair<int, object>(0, "——"));
                    pairList.Add(new KeyValuePair<int, object>(0, "——"));
                    pairList.Add(new KeyValuePair<int, object>(0, "——"));
                    pairList.Add(new KeyValuePair<int, object>(0, "——"));
                    pairList.Add(new KeyValuePair<int, object>(0, "——"));
                    pairList.Add(new KeyValuePair<int, object>(0, "——"));
                    pairList.Add(new KeyValuePair<int, object>(0, "——"));
                    createRowsTreatDataStyleNonElectricPara(pairList, rowOutNum++);
                    pairList.Clear();


                    pairList.Add(new KeyValuePair<int, object>(0, "项　　目")); //0
                    pairList.Add(new KeyValuePair<int, object>(0, "——")); //1
                    pairList.Add(new KeyValuePair<int, object>(0, "——")); //2
                    pairList.Add(new KeyValuePair<int, object>(0, "单位")); //3
                    pairList.Add(new KeyValuePair<int, object>(0, "标准参数值")); //4
                    pairList.Add(new KeyValuePair<int, object>(0, "——")); //5
                    pairList.Add(new KeyValuePair<int, object>(0, "投标人响应值")); //6
                    pairList.Add(new KeyValuePair<int, object>(0, "——")); //7
                    pairList.Add(new KeyValuePair<int, object>(0, "备注")); //8
                    createRowsTreatDataStyleNonElectricPara(pairList, rowOutNum++);
                    pairList.Clear();


                    pairList.Add(new KeyValuePair<int, object>(0, "绝缘\r\nXLPE")); //0
                    pairList.Add(new KeyValuePair<int, object>(0, "老化前断裂伸长率不小于")); //1
                    pairList.Add(new KeyValuePair<int, object>(0, "——")); //2
                    pairList.Add(new KeyValuePair<int, object>(0, "%")); //3
                    pairList.Add(new KeyValuePair<int, object>(0, "200")); //4
                    pairList.Add(new KeyValuePair<int, object>(0, "——")); //5
                    pairList.Add(new KeyValuePair<int, object>(0, "200")); //6
                    pairList.Add(new KeyValuePair<int, object>(0, "——")); //7
                    pairList.Add(new KeyValuePair<int, object>(0, "——")); //8
                    createRowsTreatDataStyleNonElectricPara(pairList, rowOutNum++);
                    pairList.Clear();

                    pairList.Add(new KeyValuePair<int, object>(0, "——")); //0
                    pairList.Add(new KeyValuePair<int, object>(0, "老化后抗张强度变化率不超过")); //1
                    pairList.Add(new KeyValuePair<int, object>(0, "——")); //2
                    pairList.Add(new KeyValuePair<int, object>(0, "%")); //3
                    pairList.Add(new KeyValuePair<int, object>(0, "±25")); //4
                    pairList.Add(new KeyValuePair<int, object>(0, "——")); //5
                    pairList.Add(new KeyValuePair<int, object>(0, "±25")); //6
                    pairList.Add(new KeyValuePair<int, object>(0, "——")); //7
                    pairList.Add(new KeyValuePair<int, object>(0, "——")); //8
                    createRowsTreatDataStyleNonElectricPara(pairList, rowOutNum++);
                    pairList.Clear();

                    pairList.Add(new KeyValuePair<int, object>(0, "——")); //0
                    pairList.Add(new KeyValuePair<int, object>(0, "老化后断裂伸长率变化率不超过")); //1
                    pairList.Add(new KeyValuePair<int, object>(0, "——")); //2
                    pairList.Add(new KeyValuePair<int, object>(0, "%")); //3
                    pairList.Add(new KeyValuePair<int, object>(0, "±25")); //4
                    pairList.Add(new KeyValuePair<int, object>(0, "——")); //5
                    pairList.Add(new KeyValuePair<int, object>(0, "±25")); //6
                    pairList.Add(new KeyValuePair<int, object>(0, "——")); //7
                    pairList.Add(new KeyValuePair<int, object>(0, "——")); //8
                    createRowsTreatDataStyleNonElectricPara(pairList, rowOutNum++);
                    pairList.Clear();

                    pairList.Add(new KeyValuePair<int, object>(0, "——")); //0
                    pairList.Add(new KeyValuePair<int, object>(0, "电缆段老化后抗张强度变化率不超过")); //1
                    pairList.Add(new KeyValuePair<int, object>(0, "——")); //2
                    pairList.Add(new KeyValuePair<int, object>(0, "%")); //3
                    pairList.Add(new KeyValuePair<int, object>(0, "±25")); //4
                    pairList.Add(new KeyValuePair<int, object>(0, "——")); //5
                    pairList.Add(new KeyValuePair<int, object>(0, "±25")); //6
                    pairList.Add(new KeyValuePair<int, object>(0, "——")); //7
                    pairList.Add(new KeyValuePair<int, object>(0, "——")); //8
                    createRowsTreatDataStyleNonElectricPara(pairList, rowOutNum++);
                    pairList.Clear();

                    pairList.Add(new KeyValuePair<int, object>(0, "——")); //0
                    pairList.Add(new KeyValuePair<int, object>(0, "电缆段老化后\r\n断裂伸长率变化率不超过")); //1
                    pairList.Add(new KeyValuePair<int, object>(0, "——")); //2
                    pairList.Add(new KeyValuePair<int, object>(0, "%")); //3
                    pairList.Add(new KeyValuePair<int, object>(0, "±25")); //4
                    pairList.Add(new KeyValuePair<int, object>(0, "——")); //5
                    pairList.Add(new KeyValuePair<int, object>(0, "±25")); //6
                    pairList.Add(new KeyValuePair<int, object>(0, "——")); //7
                    pairList.Add(new KeyValuePair<int, object>(0, "——")); //8
                    createRowsTreatDataStyleNonElectricPara(pairList, rowOutNum++);
                    pairList.Clear();


                    pairList.Add(new KeyValuePair<int, object>(0, "——")); //0
                    pairList.Add(new KeyValuePair<int, object>(0, "绝缘收缩试验不大于")); //1
                    pairList.Add(new KeyValuePair<int, object>(0, "——")); //2
                    pairList.Add(new KeyValuePair<int, object>(0, "%")); //3
                    pairList.Add(new KeyValuePair<int, object>(0, "4")); //4
                    pairList.Add(new KeyValuePair<int, object>(0, "——")); //5
                    pairList.Add(new KeyValuePair<int, object>(0, "4")); //6
                    pairList.Add(new KeyValuePair<int, object>(0, "——")); //7
                    pairList.Add(new KeyValuePair<int, object>(0, "——")); //8
                    createRowsTreatDataStyleNonElectricPara(pairList, rowOutNum++);
                    pairList.Clear();

                    pairList.Add(new KeyValuePair<int, object>(0, "绝缘")); //0
                    pairList.Add(new KeyValuePair<int, object>(0, "热延伸")); //1
                    pairList.Add(new KeyValuePair<int, object>(0, "负荷下伸长率不大于")); //2
                    pairList.Add(new KeyValuePair<int, object>(0, "%")); //3
                    pairList.Add(new KeyValuePair<int, object>(0, "125")); //4
                    pairList.Add(new KeyValuePair<int, object>(0, "——")); //5
                    pairList.Add(new KeyValuePair<int, object>(0, "125")); //6
                    pairList.Add(new KeyValuePair<int, object>(0, "——")); //7
                    pairList.Add(new KeyValuePair<int, object>(0, "——")); //8
                    createRowsTreatDataStyleNonElectricPara(pairList, rowOutNum++);
                    pairList.Clear();

                    pairList.Add(new KeyValuePair<int, object>(0, "——")); //0
                    pairList.Add(new KeyValuePair<int, object>(0, "——")); //1
                    pairList.Add(new KeyValuePair<int, object>(0, "冷却后永久伸长率不大于")); //2
                    pairList.Add(new KeyValuePair<int, object>(0, "%")); //3
                    pairList.Add(new KeyValuePair<int, object>(0, "10")); //4
                    pairList.Add(new KeyValuePair<int, object>(0, "——")); //5
                    pairList.Add(new KeyValuePair<int, object>(0, "10")); //6
                    pairList.Add(new KeyValuePair<int, object>(0, "——")); //7
                    pairList.Add(new KeyValuePair<int, object>(0, "——")); //8
                    createRowsTreatDataStyleNonElectricPara(pairList, rowOutNum++);
                    pairList.Clear();

                    pairList.Add(new KeyValuePair<int, object>(0, "外护套")); //0
                    pairList.Add(new KeyValuePair<int, object>(0, "——")); //1
                    pairList.Add(new KeyValuePair<int, object>(0, "——")); //2
                    pairList.Add(new KeyValuePair<int, object>(0, "——")); //3
                    pairList.Add(new KeyValuePair<int, object>(0, "PE")); //4
                    pairList.Add(new KeyValuePair<int, object>(0, "PVC")); //5
                    pairList.Add(new KeyValuePair<int, object>(0, "无卤低烟\r\n阻燃护套")); //6
                    pairList.Add(new KeyValuePair<int, object>(0, "PVC")); //7
                    pairList.Add(new KeyValuePair<int, object>(0, "——")); //8
                    createRowsTreatDataStyleNonElectricPara(pairList, rowOutNum++);
                    pairList.Clear();

                    pairList.Add(new KeyValuePair<int, object>(0, "——")); //0
                    pairList.Add(new KeyValuePair<int, object>(0, "老化前抗张强度不小于")); //1
                    pairList.Add(new KeyValuePair<int, object>(0, "——")); //2
                    pairList.Add(new KeyValuePair<int, object>(0, "MPa")); //3
                    pairList.Add(new KeyValuePair<int, object>(0, "10")); //4
                    pairList.Add(new KeyValuePair<int, object>(0, "12.5")); //5
                    pairList.Add(new KeyValuePair<int, object>(0, "9")); //6
                    pairList.Add(new KeyValuePair<int, object>(0, "12.5")); //7
                    pairList.Add(new KeyValuePair<int, object>(0, "——")); //8
                    createRowsTreatDataStyleNonElectricPara(pairList, rowOutNum++);
                    pairList.Clear();

                    pairList.Add(new KeyValuePair<int, object>(0, "——")); //0
                    pairList.Add(new KeyValuePair<int, object>(0, "老化前断裂伸长率不小于")); //1
                    pairList.Add(new KeyValuePair<int, object>(0, "——")); //2
                    pairList.Add(new KeyValuePair<int, object>(0, "%")); //3
                    pairList.Add(new KeyValuePair<int, object>(0, "300")); //4
                    pairList.Add(new KeyValuePair<int, object>(0, "150")); //5
                    pairList.Add(new KeyValuePair<int, object>(0, "125")); //6
                    pairList.Add(new KeyValuePair<int, object>(0, "150")); //7
                    pairList.Add(new KeyValuePair<int, object>(0, "——")); //8
                    createRowsTreatDataStyleNonElectricPara(pairList, rowOutNum++);
                    pairList.Clear();

                    pairList.Add(new KeyValuePair<int, object>(0, "——")); //0
                    pairList.Add(new KeyValuePair<int, object>(0, "老化后抗张强度不小于")); //1
                    pairList.Add(new KeyValuePair<int, object>(0, "——")); //2
                    pairList.Add(new KeyValuePair<int, object>(0, "MPa")); //3
                    pairList.Add(new KeyValuePair<int, object>(0, "——")); //4
                    pairList.Add(new KeyValuePair<int, object>(0, "12.5")); //5
                    pairList.Add(new KeyValuePair<int, object>(0, "9")); //6
                    pairList.Add(new KeyValuePair<int, object>(0, "12.5")); //7
                    pairList.Add(new KeyValuePair<int, object>(0, "——")); //8
                    createRowsTreatDataStyleNonElectricPara(pairList, rowOutNum++);
                    pairList.Clear();

                    pairList.Add(new KeyValuePair<int, object>(0, "——")); //0
                    pairList.Add(new KeyValuePair<int, object>(0, "老化后断裂伸长率不小于")); //1
                    pairList.Add(new KeyValuePair<int, object>(0, "——")); //2
                    pairList.Add(new KeyValuePair<int, object>(0, "%")); //3
                    pairList.Add(new KeyValuePair<int, object>(0, "300")); //4
                    pairList.Add(new KeyValuePair<int, object>(0, "150")); //5
                    pairList.Add(new KeyValuePair<int, object>(0, "100")); //6
                    pairList.Add(new KeyValuePair<int, object>(0, "150")); //7
                    pairList.Add(new KeyValuePair<int, object>(0, "——")); //8
                    createRowsTreatDataStyleNonElectricPara(pairList, rowOutNum++);
                    pairList.Clear();

                    pairList.Add(new KeyValuePair<int, object>(0, "——")); //0
                    pairList.Add(new KeyValuePair<int, object>(0, "老化后抗张强度变化率不超过")); //1
                    pairList.Add(new KeyValuePair<int, object>(0, "——")); //2
                    pairList.Add(new KeyValuePair<int, object>(0, "%")); //3
                    pairList.Add(new KeyValuePair<int, object>(0, "——")); //4
                    pairList.Add(new KeyValuePair<int, object>(0, "±25")); //5
                    pairList.Add(new KeyValuePair<int, object>(0, "±40")); //6
                    pairList.Add(new KeyValuePair<int, object>(0, "±25")); //7
                    pairList.Add(new KeyValuePair<int, object>(0, "——")); //8
                    createRowsTreatDataStyleNonElectricPara(pairList, rowOutNum++);
                    pairList.Clear();

                    pairList.Add(new KeyValuePair<int, object>(0, "——")); //0
                    pairList.Add(new KeyValuePair<int, object>(0, "老化后断裂伸长率变化率不超过")); //1
                    pairList.Add(new KeyValuePair<int, object>(0, "——")); //2
                    pairList.Add(new KeyValuePair<int, object>(0, "%")); //3
                    pairList.Add(new KeyValuePair<int, object>(0, "——")); //4
                    pairList.Add(new KeyValuePair<int, object>(0, "±25")); //5
                    pairList.Add(new KeyValuePair<int, object>(0, "±40")); //6
                    pairList.Add(new KeyValuePair<int, object>(0, "±25")); //7
                    pairList.Add(new KeyValuePair<int, object>(0, "——")); //8
                    createRowsTreatDataStyleNonElectricPara(pairList, rowOutNum++);
                    pairList.Clear();

                    pairList.Add(new KeyValuePair<int, object>(0, "——")); //0
                    pairList.Add(new KeyValuePair<int, object>(0, "电缆段老化后抗张强度变化率不超过")); //1
                    pairList.Add(new KeyValuePair<int, object>(0, "——")); //2
                    pairList.Add(new KeyValuePair<int, object>(0, "%")); //3
                    pairList.Add(new KeyValuePair<int, object>(0, "——")); //4
                    pairList.Add(new KeyValuePair<int, object>(0, "±25")); //5
                    pairList.Add(new KeyValuePair<int, object>(0, "——")); //6
                    pairList.Add(new KeyValuePair<int, object>(0, "±25")); //7
                    pairList.Add(new KeyValuePair<int, object>(0, "——")); //8
                    createRowsTreatDataStyleNonElectricPara(pairList, rowOutNum++);
                    pairList.Clear();

                    pairList.Add(new KeyValuePair<int, object>(0, "——")); //0
                    pairList.Add(new KeyValuePair<int, object>(0, "电缆段老化后\r\n断裂伸长率变化率不超过")); //1
                    pairList.Add(new KeyValuePair<int, object>(0, "——")); //2
                    pairList.Add(new KeyValuePair<int, object>(0, "%")); //3
                    pairList.Add(new KeyValuePair<int, object>(0, "——")); //4
                    pairList.Add(new KeyValuePair<int, object>(0, "±25")); //5
                    pairList.Add(new KeyValuePair<int, object>(0, "——")); //6
                    pairList.Add(new KeyValuePair<int, object>(0, "±25")); //7
                    pairList.Add(new KeyValuePair<int, object>(0, "——")); //8
                    createRowsTreatDataStyleNonElectricPara(pairList, rowOutNum++);
                    pairList.Clear();

                    pairList.Add(new KeyValuePair<int, object>(0, "——")); //0
                    pairList.Add(new KeyValuePair<int, object>(0, "高温压力试验，压痕深度不大于")); //1
                    pairList.Add(new KeyValuePair<int, object>(0, "——")); //2
                    pairList.Add(new KeyValuePair<int, object>(0, "%")); //3
                    pairList.Add(new KeyValuePair<int, object>(0, "50")); //4
                    pairList.Add(new KeyValuePair<int, object>(0, "50")); //5
                    pairList.Add(new KeyValuePair<int, object>(0, "50")); //6
                    pairList.Add(new KeyValuePair<int, object>(0, "50")); //7
                    pairList.Add(new KeyValuePair<int, object>(0, "——")); //8
                    createRowsTreatDataStyleNonElectricPara(pairList, rowOutNum++);
                    pairList.Clear();

                    pairList.Add(new KeyValuePair<int, object>(0, "——")); //0
                    pairList.Add(new KeyValuePair<int, object>(0, "热冲击试验")); //1
                    pairList.Add(new KeyValuePair<int, object>(0, "——")); //2
                    pairList.Add(new KeyValuePair<int, object>(0, "——")); //3
                    pairList.Add(new KeyValuePair<int, object>(0, "——")); //4
                    pairList.Add(new KeyValuePair<int, object>(0, "不开裂")); //5
                    pairList.Add(new KeyValuePair<int, object>(0, "——")); //6
                    pairList.Add(new KeyValuePair<int, object>(0, "不开裂")); //7
                    pairList.Add(new KeyValuePair<int, object>(0, "——")); //8
                    createRowsTreatDataStyleNonElectricPara(pairList, rowOutNum++);
                    pairList.Clear();

                    pairList.Add(new KeyValuePair<int, object>(0, "——")); //0
                    pairList.Add(new KeyValuePair<int, object>(0, "低温冲击试验")); //1
                    pairList.Add(new KeyValuePair<int, object>(0, "——")); //2
                    pairList.Add(new KeyValuePair<int, object>(0, "——")); //3
                    pairList.Add(new KeyValuePair<int, object>(0, "——")); //4
                    pairList.Add(new KeyValuePair<int, object>(0, "不开裂")); //5
                    pairList.Add(new KeyValuePair<int, object>(0, "不开裂")); //6
                    pairList.Add(new KeyValuePair<int, object>(0, "不开裂")); //7
                    pairList.Add(new KeyValuePair<int, object>(0, "——")); //8
                    createRowsTreatDataStyleNonElectricPara(pairList, rowOutNum++);
                    pairList.Clear();

                    pairList.Add(new KeyValuePair<int, object>(0, "——")); //0
                    pairList.Add(new KeyValuePair<int, object>(0, "低温拉伸，断裂伸长率不小于")); //1
                    pairList.Add(new KeyValuePair<int, object>(0, "——")); //2
                    pairList.Add(new KeyValuePair<int, object>(0, "%")); //3
                    pairList.Add(new KeyValuePair<int, object>(0, "——")); //4
                    pairList.Add(new KeyValuePair<int, object>(0, "20")); //5
                    pairList.Add(new KeyValuePair<int, object>(0, "20")); //6
                    pairList.Add(new KeyValuePair<int, object>(0, "20")); //7
                    pairList.Add(new KeyValuePair<int, object>(0, "——")); //8
                    createRowsTreatDataStyleNonElectricPara(pairList, rowOutNum++);
                    pairList.Clear();

                    pairList.Add(new KeyValuePair<int, object>(0, "——")); //0
                    pairList.Add(new KeyValuePair<int, object>(0, "热失重，最大允许失重　")); //1
                    pairList.Add(new KeyValuePair<int, object>(0, "——")); //2
                    pairList.Add(new KeyValuePair<int, object>(0, "mg/cm2")); //3
                    pairList.Add(new KeyValuePair<int, object>(0, "——")); //4
                    pairList.Add(new KeyValuePair<int, object>(0, "1.5")); //5
                    pairList.Add(new KeyValuePair<int, object>(0, "——")); //6
                    pairList.Add(new KeyValuePair<int, object>(0, "1.5")); //7
                    pairList.Add(new KeyValuePair<int, object>(0, "——")); //8
                    createRowsTreatDataStyleNonElectricPara(pairList, rowOutNum++);
                    pairList.Clear();

                    pairList.Add(new KeyValuePair<int, object>(0, "——")); //0
                    pairList.Add(new KeyValuePair<int, object>(0, "酸气含量试验（GB/T17650）")); //1
                    pairList.Add(new KeyValuePair<int, object>(0, "——")); //2
                    pairList.Add(new KeyValuePair<int, object>(0, "%")); //3
                    pairList.Add(new KeyValuePair<int, object>(0, "——")); //4
                    pairList.Add(new KeyValuePair<int, object>(0, "——")); //5
                    pairList.Add(new KeyValuePair<int, object>(0, "0.5")); //6
                    pairList.Add(new KeyValuePair<int, object>(0, "不适用")); //7
                    pairList.Add(new KeyValuePair<int, object>(0, "——")); //8
                    createRowsTreatDataStyleNonElectricPara(pairList, rowOutNum++);
                    pairList.Clear();

                    pairList.Add(new KeyValuePair<int, object>(0, "——")); //0
                    pairList.Add(new KeyValuePair<int, object>(0, "最大值")); //1
                    pairList.Add(new KeyValuePair<int, object>(0, "——")); //2
                    pairList.Add(new KeyValuePair<int, object>(0, "——")); //3
                    pairList.Add(new KeyValuePair<int, object>(0, "——")); //4
                    pairList.Add(new KeyValuePair<int, object>(0, "——")); //5
                    pairList.Add(new KeyValuePair<int, object>(0, "——")); //6
                    pairList.Add(new KeyValuePair<int, object>(0, "——")); //7
                    pairList.Add(new KeyValuePair<int, object>(0, "——")); //8
                    createRowsTreatDataStyleNonElectricPara(pairList, rowOutNum++);
                    pairList.Clear();

                    pairList.Add(new KeyValuePair<int, object>(0, "——")); //0
                    pairList.Add(new KeyValuePair<int, object>(0, "氟含量试验（IEC60684）")); //1
                    pairList.Add(new KeyValuePair<int, object>(0, "——")); //2
                    pairList.Add(new KeyValuePair<int, object>(0, "%")); //3
                    pairList.Add(new KeyValuePair<int, object>(0, "——")); //4
                    pairList.Add(new KeyValuePair<int, object>(0, "——")); //5
                    pairList.Add(new KeyValuePair<int, object>(0, "0.1")); //6
                    pairList.Add(new KeyValuePair<int, object>(0, "不适用")); //7
                    pairList.Add(new KeyValuePair<int, object>(0, "——")); //8
                    createRowsTreatDataStyleNonElectricPara(pairList, rowOutNum++);
                    pairList.Clear();

                    pairList.Add(new KeyValuePair<int, object>(0, "——")); //0
                    pairList.Add(new KeyValuePair<int, object>(0, "最大值")); //1
                    pairList.Add(new KeyValuePair<int, object>(0, "——")); //2
                    pairList.Add(new KeyValuePair<int, object>(0, "——")); //3
                    pairList.Add(new KeyValuePair<int, object>(0, "——")); //4
                    pairList.Add(new KeyValuePair<int, object>(0, "——")); //5
                    pairList.Add(new KeyValuePair<int, object>(0, "——")); //6
                    pairList.Add(new KeyValuePair<int, object>(0, "——")); //7
                    pairList.Add(new KeyValuePair<int, object>(0, "——")); //8
                    createRowsTreatDataStyleNonElectricPara(pairList, rowOutNum++);
                    pairList.Clear();

                    pairList.Add(new KeyValuePair<int, object>(0, "——")); //0
                    pairList.Add(new KeyValuePair<int, object>(0, "PH值 最小值")); //1
                    pairList.Add(new KeyValuePair<int, object>(0, "——")); //2
                    pairList.Add(new KeyValuePair<int, object>(0, "——")); //3
                    pairList.Add(new KeyValuePair<int, object>(0, "——")); //4
                    pairList.Add(new KeyValuePair<int, object>(0, "——")); //5
                    pairList.Add(new KeyValuePair<int, object>(0, "4.3")); //6
                    pairList.Add(new KeyValuePair<int, object>(0, "不适用")); //7
                    pairList.Add(new KeyValuePair<int, object>(0, "——")); //8
                    createRowsTreatDataStyleNonElectricPara(pairList, rowOutNum++);
                    pairList.Clear();

                    pairList.Add(new KeyValuePair<int, object>(0, "——")); //0
                    pairList.Add(new KeyValuePair<int, object>(0, "电导率 最大值")); //1
                    pairList.Add(new KeyValuePair<int, object>(0, "——")); //2
                    pairList.Add(new KeyValuePair<int, object>(0, "µS/mm")); //3
                    pairList.Add(new KeyValuePair<int, object>(0, "——")); //4
                    pairList.Add(new KeyValuePair<int, object>(0, "——")); //5
                    pairList.Add(new KeyValuePair<int, object>(0, "10")); //6
                    pairList.Add(new KeyValuePair<int, object>(0, "不适用")); //7
                    pairList.Add(new KeyValuePair<int, object>(0, "——")); //8
                    createRowsTreatDataStyleNonElectricPara(pairList, rowOutNum++);
                    pairList.Clear();

                    pairList.Add(new KeyValuePair<int, object>(0, "——")); //0
                    pairList.Add(new KeyValuePair<int, object>(0, "吸水试验 最大增重")); //1
                    pairList.Add(new KeyValuePair<int, object>(0, "——")); //2
                    pairList.Add(new KeyValuePair<int, object>(0, "mg/cm2")); //3
                    pairList.Add(new KeyValuePair<int, object>(0, "——")); //4
                    pairList.Add(new KeyValuePair<int, object>(0, "——")); //5
                    pairList.Add(new KeyValuePair<int, object>(0, "10")); //6
                    pairList.Add(new KeyValuePair<int, object>(0, "不适用")); //7
                    pairList.Add(new KeyValuePair<int, object>(0, "——")); //8
                    createRowsTreatDataStyleNonElectricPara(pairList, rowOutNum++);
                    pairList.Clear();

                    pairList.Add(new KeyValuePair<int, object>(0, "——")); //0
                    pairList.Add(new KeyValuePair<int, object>(0, "炭黑含量")); //1
                    pairList.Add(new KeyValuePair<int, object>(0, "——")); //2
                    pairList.Add(new KeyValuePair<int, object>(0, "%")); //3
                    pairList.Add(new KeyValuePair<int, object>(0, "2.0～3.0")); //4
                    pairList.Add(new KeyValuePair<int, object>(0, "——")); //5
                    pairList.Add(new KeyValuePair<int, object>(0, "——")); //6
                    pairList.Add(new KeyValuePair<int, object>(0, "——")); //7
                    pairList.Add(new KeyValuePair<int, object>(0, "——")); //8
                    createRowsTreatDataStyleNonElectricPara(pairList, rowOutNum++);
                    pairList.Clear();

                }

                void createRowsTreatDataStyleNonElectricPara(List<KeyValuePair<int, object>> inputDataList, int rowNum)
                {

                    // 3. 添加标题行
                    IRow iRow = sheetNonElectricPara.CreateRow(rowNum);
                    sheetNonElectricPara.AutoSizeRow(rowNum);
                    int colIdx = 0;
                    //for (int colIdx = 0; colIdx < 6; colIdx++)//方法2,3
                    foreach (var keyValuePair in inputDataList)//方法1
                    {
                        try
                        {
                            //ICell cell = iRow.GetCell(colIdx);//如果获得已有单元格，则这样写
                            ICell cell = iRow.CreateCell(colIdx);//IRow是地址引用，像指针，反过来赋值
                                                                 // cell.CellStyle = stringCenterStyle;
                                                                 //int styleInt = styleList[colIdx];//方法2
                            int styleInt = keyValuePair.Key;//方法1
                                                            // object value = inputDataList[colIdx];
                            object value = keyValuePair.Value;

                            // 根据数据类型应用样式
                            if (value == DBNull.Value || value == "")
                            {
                                // cell.SetCellValue("ut");// string.Empty;
                                // cell.CellStyle = nullStyle;// CreateNullStyle();
                                cell.SetCellValue("数据未提供");
                                cell.CellStyle = warnStyle;
                            }
                            else
                            {
                                switch (value)
                                {
                                    case string string1:
                                        cell.SetCellValue(Convert.ToString(value));
                                        // cell.CellStyle = stringCenterStyle;
                                        break;
                                    case int int1:
                                        cell.SetCellValue(Convert.ToString(value));
                                        //  cell.CellStyle = cell.CellStyle = stringCenterStyle;                                        
                                        break;
                                    case DateTime dateTime1:
                                        cell.SetCellValue(Convert.ToDateTime(value));
                                        //  cell.CellStyle = dateTimeStyle;
                                        break;
                                    case double double1:
                                        cell.SetCellValue(Convert.ToDouble(value));
                                        if (Convert.ToDouble(value) < 0)
                                        {
                                            //  cell.CellStyle = warnNumStyle;//numberStyle; //
                                            break;
                                        }
                                        // cell.CellStyle = numberStyle;
                                        break;
                                    default:
                                        cell.SetCellValue(value.ToString());
                                        //  cell.CellStyle = stringCenterStyle; //文本自动换行
                                        break;
                                }
                                if (rowNum == 0) cell.CellStyle = titleStyle;
                                else if (faultMarkArray.Any(term => value.ToString().Contains(term, StringComparison.OrdinalIgnoreCase)))
                                    cell.CellStyle = warnStyle;
                                else if (styleInt == 1)//方法2
                                    cell.CellStyle = stringBlueStyle;
                                else
                                    cell.CellStyle = stringCenterStyle;
                                if (value.ToString().Contains("电缆段老化后\r\n", StringComparison.OrdinalIgnoreCase))
                                {
                                    //合并单元格，文字超出单元格范围，行高不会自动变化
                                    iRow.Height = 576;// 14.4*2 *20;1/20个点为最小单位
                                }

                            }
                            Trace.Write($"{value.ToString()}  ");//20250519 打印单元格数据                
                        }                             //                   }
                        catch (Exception excep1)
                        {
                            Trace.WriteLine(excep1.Message);
                            Trace.WriteLine($"问题在第{rowNum.ToString()}行  ");
                        }
                        ++colIdx;
                    }
                    Trace.WriteLine(""); ;
                }

                void setColumnWidthNonElectricPara()
                {
                    sheetNonElectricPara.SetColumnWidth(0, (5.33 + 0.78) * 256);
                    sheetNonElectricPara.SetColumnWidth(1, (6.68 + 0.78) * 256);
                    sheetNonElectricPara.SetColumnWidth(2, (21.68 + 0.78) * 256);
                    sheetNonElectricPara.SetColumnWidth(3, (5.33 + 0.78) * 256);
                    sheetNonElectricPara.SetColumnWidth(4, (6.56 + 0.78) * 256);
                    sheetNonElectricPara.SetColumnWidth(5, (7.22 + 0.78) * 256);
                    sheetNonElectricPara.SetColumnWidth(6, (7.22 + 0.78) * 256);
                    sheetNonElectricPara.SetColumnWidth(7, (7.22 + 0.78) * 256);
                    sheetNonElectricPara.SetColumnWidth(8, (12.44 + 0.78) * 256);
                }

                void mergeCellsNonElectricPara() //合并单元格
                {
                    CellRangeAddress cellMerge = new CellRangeAddress(0, 0, 0, 8);//（起始行，结束行，起始列，结束列）
                    sheetNonElectricPara.AddMergedRegion(cellMerge);
                    CellRangeAddress cellMerge1 = new CellRangeAddress(1, 1, 0, 2);//（起始行，结束行，起始列，结束列）         
                    sheetNonElectricPara.AddMergedRegion(cellMerge1);
                    CellRangeAddress cellMerge2 = new CellRangeAddress(2, 7, 0, 0);//（起始行，结束行，起始列，结束列）
                    sheetNonElectricPara.AddMergedRegion(cellMerge2);
                    CellRangeAddress cellMerge3 = new CellRangeAddress(2, 2, 1, 2);//（起始行，结束行，起始列，结束列）
                    sheetNonElectricPara.AddMergedRegion(cellMerge3);
                    CellRangeAddress cellMerge4 = new CellRangeAddress(3, 3, 1, 2);//（起始行，结束行，起始列，结束列）
                    sheetNonElectricPara.AddMergedRegion(cellMerge4);
                    CellRangeAddress cellMerge81 = new CellRangeAddress(4, 4, 1, 2);
                    sheetNonElectricPara.AddMergedRegion(cellMerge81);
                    CellRangeAddress cellMerge82 = new CellRangeAddress(5, 5, 1, 2);
                    sheetNonElectricPara.AddMergedRegion(cellMerge82);
                    CellRangeAddress cellMerge83 = new CellRangeAddress(6, 6, 1, 2);
                    sheetNonElectricPara.AddMergedRegion(cellMerge83);
                    CellRangeAddress cellMerge84 = new CellRangeAddress(7, 7, 1, 2);
                    sheetNonElectricPara.AddMergedRegion(cellMerge84);
                    CellRangeAddress cellMerge85 = new CellRangeAddress(8, 9, 0, 0);
                    sheetNonElectricPara.AddMergedRegion(cellMerge85);
                    CellRangeAddress cellMerge86 = new CellRangeAddress(8, 9, 1, 1);
                    sheetNonElectricPara.AddMergedRegion(cellMerge86);
                    CellRangeAddress cellMerge87 = new CellRangeAddress(10, 31, 0, 0);
                    sheetNonElectricPara.AddMergedRegion(cellMerge87);
                    CellRangeAddress cellMerge88 = new CellRangeAddress(10, 10, 1, 3);
                    sheetNonElectricPara.AddMergedRegion(cellMerge88);
                    CellRangeAddress cellMerge89 = new CellRangeAddress(11, 11, 1, 2);
                    sheetNonElectricPara.AddMergedRegion(cellMerge89);
                    CellRangeAddress cellMerge6 = new CellRangeAddress(12, 12, 1, 2);
                    sheetNonElectricPara.AddMergedRegion(cellMerge6);
                    CellRangeAddress cellMerge7 = new CellRangeAddress(13, 13, 1, 2);
                    sheetNonElectricPara.AddMergedRegion(cellMerge7);
                    CellRangeAddress cellMerge8 = new CellRangeAddress(14, 14, 1, 2);
                    sheetNonElectricPara.AddMergedRegion(cellMerge8);
                    CellRangeAddress cellMerge9 = new CellRangeAddress(15, 15, 1, 2);
                    sheetNonElectricPara.AddMergedRegion(cellMerge9);
                    CellRangeAddress cellMerge10 = new CellRangeAddress(16, 16, 1, 2);
                    sheetNonElectricPara.AddMergedRegion(cellMerge10);
                    CellRangeAddress cellMerge11 = new CellRangeAddress(17, 17, 1, 2);
                    sheetNonElectricPara.AddMergedRegion(cellMerge11);
                    CellRangeAddress cellMerge12 = new CellRangeAddress(18, 18, 1, 2);
                    sheetNonElectricPara.AddMergedRegion(cellMerge12);
                    CellRangeAddress cellMerge13 = new CellRangeAddress(19, 19, 1, 2);
                    sheetNonElectricPara.AddMergedRegion(cellMerge13);
                    CellRangeAddress cellMerge14 = new CellRangeAddress(20, 20, 1, 2);
                    sheetNonElectricPara.AddMergedRegion(cellMerge14);
                    CellRangeAddress cellMerge15 = new CellRangeAddress(21, 21, 1, 2);
                    sheetNonElectricPara.AddMergedRegion(cellMerge15);
                    CellRangeAddress cellMerge17 = new CellRangeAddress(22, 22, 1, 2);
                    sheetNonElectricPara.AddMergedRegion(cellMerge17);
                    CellRangeAddress cellMerge18 = new CellRangeAddress(23, 23, 1, 2);
                    sheetNonElectricPara.AddMergedRegion(cellMerge18);
                    CellRangeAddress cellMerge16 = new CellRangeAddress(24, 24, 1, 2);
                    sheetNonElectricPara.AddMergedRegion(cellMerge16);





                    CellRangeAddress cellMerge25 = new CellRangeAddress(24, 25, 3, 3);
                    sheetNonElectricPara.AddMergedRegion(cellMerge25);
                    CellRangeAddress cellMerge24 = new CellRangeAddress(24, 25, 4, 4);
                    sheetNonElectricPara.AddMergedRegion(cellMerge24);
                    CellRangeAddress cellMerge30 = new CellRangeAddress(24, 25, 5, 5);
                    sheetNonElectricPara.AddMergedRegion(cellMerge30);
                    CellRangeAddress cellMerge34 = new CellRangeAddress(24, 25, 6, 6);
                    sheetNonElectricPara.AddMergedRegion(cellMerge34);
                    CellRangeAddress cellMerge31 = new CellRangeAddress(24, 25, 7, 7);
                    sheetNonElectricPara.AddMergedRegion(cellMerge31);
                    CellRangeAddress cellMerge32 = new CellRangeAddress(24, 25, 8, 8);
                    sheetNonElectricPara.AddMergedRegion(cellMerge32);


                    CellRangeAddress cellMerge33 = new CellRangeAddress(26, 27, 3, 3);
                    sheetNonElectricPara.AddMergedRegion(cellMerge33);
                    CellRangeAddress cellMerge5 = new CellRangeAddress(26, 27, 4, 4);
                    sheetNonElectricPara.AddMergedRegion(cellMerge5);
                    CellRangeAddress cellMerge35 = new CellRangeAddress(26, 27, 5, 5);
                    sheetNonElectricPara.AddMergedRegion(cellMerge35);
                    CellRangeAddress cellMerge36 = new CellRangeAddress(26, 27, 6, 6);
                    sheetNonElectricPara.AddMergedRegion(cellMerge36);
                    CellRangeAddress cellMerge37 = new CellRangeAddress(26, 27, 7, 7);
                    sheetNonElectricPara.AddMergedRegion(cellMerge37);
                    CellRangeAddress cellMerge38 = new CellRangeAddress(26, 27, 8, 8);
                    sheetNonElectricPara.AddMergedRegion(cellMerge38);






                    CellRangeAddress cellMerge20 = new CellRangeAddress(25, 25, 1, 2);
                    sheetNonElectricPara.AddMergedRegion(cellMerge20);
                    CellRangeAddress cellMerge21 = new CellRangeAddress(26, 26, 1, 2);
                    sheetNonElectricPara.AddMergedRegion(cellMerge21);
                    CellRangeAddress cellMerge23 = new CellRangeAddress(27, 27, 1, 2);
                    sheetNonElectricPara.AddMergedRegion(cellMerge23);
                    CellRangeAddress cellMerge26 = new CellRangeAddress(28, 28, 1, 2);
                    sheetNonElectricPara.AddMergedRegion(cellMerge26);
                    CellRangeAddress cellMerge27 = new CellRangeAddress(29, 29, 1, 2);
                    sheetNonElectricPara.AddMergedRegion(cellMerge27);
                    CellRangeAddress cellMerge28 = new CellRangeAddress(30, 30, 1, 2);
                    sheetNonElectricPara.AddMergedRegion(cellMerge28);
                    CellRangeAddress cellMerge29 = new CellRangeAddress(31, 31, 1, 2);
                    sheetNonElectricPara.AddMergedRegion(cellMerge29);


                }
            }


            void createSheetMaterialconfiguration()
            {

                feedDataByRowsForMaterialConfiguration();
                setColumnWidthMaterialConfiguration();
                mergeCellsMaterialConfiguration();

                void feedDataByRowsForMaterialConfiguration()
                {
                    int rowOutNum = 0;
                    var pairList = new List<KeyValuePair<int, object>>();
                    // 添加元素					



                    pairList.Add(new KeyValuePair<int, object>(0, "组件材料配置表")); //0
                    pairList.Add(new KeyValuePair<int, object>(0, " ")); //1
                    pairList.Add(new KeyValuePair<int, object>(0, " ")); //2
                    pairList.Add(new KeyValuePair<int, object>(0, " ")); //3
                    pairList.Add(new KeyValuePair<int, object>(0, " ")); //4
                    pairList.Add(new KeyValuePair<int, object>(0, " ")); //5
                    createRowsTreatDataStyleMaterialConfiguration(pairList, rowOutNum++);
                    pairList.Clear();


                    pairList.Add(new KeyValuePair<int, object>(0, "序号")); //0
                    pairList.Add(new KeyValuePair<int, object>(0, "名称")); //1
                    pairList.Add(new KeyValuePair<int, object>(0, "型式规格，参数")); //2
                    pairList.Add(new KeyValuePair<int, object>(0, "数量")); //3
                    pairList.Add(new KeyValuePair<int, object>(0, "制造商")); //4
                    pairList.Add(new KeyValuePair<int, object>(0, "原产地")); //5
                    createRowsTreatDataStyleMaterialConfiguration(pairList, rowOutNum++);
                    pairList.Clear();

                    pairList.Add(new KeyValuePair<int, object>(0, 1)); //0
                    pairList.Add(new KeyValuePair<int, object>(0, "电缆导体")); //1
                    pairList.Add(new KeyValuePair<int, object>(0, "")); //2
                    pairList.Add(new KeyValuePair<int, object>(0, "——")); //3
                    pairList.Add(new KeyValuePair<int, object>(0, "")); //4
                    pairList.Add(new KeyValuePair<int, object>(0, "")); //5
                    createRowsTreatDataStyleMaterialConfiguration(pairList, rowOutNum++);
                    pairList.Clear();

                    pairList.Add(new KeyValuePair<int, object>(0, "")); //0
                    pairList.Add(new KeyValuePair<int, object>(0, "")); //1
                    pairList.Add(new KeyValuePair<int, object>(0, "")); //2
                    pairList.Add(new KeyValuePair<int, object>(0, "——")); //3
                    pairList.Add(new KeyValuePair<int, object>(0, "")); //4
                    pairList.Add(new KeyValuePair<int, object>(0, "")); //5
                    createRowsTreatDataStyleMaterialConfiguration(pairList, rowOutNum++);
                    pairList.Clear();

                    pairList.Add(new KeyValuePair<int, object>(0, 2)); //0
                    pairList.Add(new KeyValuePair<int, object>(0, "交联聚\r\n乙烯绝缘")); //1
                    pairList.Add(new KeyValuePair<int, object>(0, "")); //2
                    pairList.Add(new KeyValuePair<int, object>(0, "——")); //3
                    pairList.Add(new KeyValuePair<int, object>(0, "")); //4
                    pairList.Add(new KeyValuePair<int, object>(0, "")); //5
                    createRowsTreatDataStyleMaterialConfiguration(pairList, rowOutNum++);
                    pairList.Clear();

                    pairList.Add(new KeyValuePair<int, object>(0, 3)); //0
                    pairList.Add(new KeyValuePair<int, object>(0, "填充层")); //1
                    pairList.Add(new KeyValuePair<int, object>(0, "")); //2
                    pairList.Add(new KeyValuePair<int, object>(0, "——")); //3
                    pairList.Add(new KeyValuePair<int, object>(0, "")); //4
                    pairList.Add(new KeyValuePair<int, object>(0, "")); //5
                    createRowsTreatDataStyleMaterialConfiguration(pairList, rowOutNum++);
                    pairList.Clear();

                    pairList.Add(new KeyValuePair<int, object>(0, 4)); //0
                    pairList.Add(new KeyValuePair<int, object>(0, "内衬层")); //1
                    pairList.Add(new KeyValuePair<int, object>(0, "")); //2
                    pairList.Add(new KeyValuePair<int, object>(0, "——")); //3
                    pairList.Add(new KeyValuePair<int, object>(0, "")); //4
                    pairList.Add(new KeyValuePair<int, object>(0, "")); //5
                    createRowsTreatDataStyleMaterialConfiguration(pairList, rowOutNum++);
                    pairList.Clear();

                    pairList.Add(new KeyValuePair<int, object>(0, 5)); //0
                    pairList.Add(new KeyValuePair<int, object>(0, "铠装层")); //1
                    pairList.Add(new KeyValuePair<int, object>(0, "")); //2
                    pairList.Add(new KeyValuePair<int, object>(0, "——")); //3
                    pairList.Add(new KeyValuePair<int, object>(0, "")); //4
                    pairList.Add(new KeyValuePair<int, object>(0, "")); //5
                    createRowsTreatDataStyleMaterialConfiguration(pairList, rowOutNum++);
                    pairList.Clear();

                    pairList.Add(new KeyValuePair<int, object>(0, 6)); //0
                    pairList.Add(new KeyValuePair<int, object>(0, "外护套")); //1
                    pairList.Add(new KeyValuePair<int, object>(0, "")); //2
                    pairList.Add(new KeyValuePair<int, object>(0, "——")); //3
                    pairList.Add(new KeyValuePair<int, object>(0, "")); //4
                    pairList.Add(new KeyValuePair<int, object>(0, "")); //5
                    createRowsTreatDataStyleMaterialConfiguration(pairList, rowOutNum++);
                    pairList.Clear();

                    pairList.Add(new KeyValuePair<int, object>(0, "说明")); //0
                    pairList.Add(new KeyValuePair<int, object>(0, "1、主要货物的名称、规格型式、参数、单位、数量要求请保持与商务部分货物清单一致。")); //1
                    pairList.Add(new KeyValuePair<int, object>(0, "——")); //2
                    pairList.Add(new KeyValuePair<int, object>(0, "——")); //3
                    pairList.Add(new KeyValuePair<int, object>(0, "——")); //4
                    pairList.Add(new KeyValuePair<int, object>(0, "——")); //5
                    createRowsTreatDataStyleMaterialConfiguration(pairList, rowOutNum++);
                    pairList.Clear();

                    pairList.Add(new KeyValuePair<int, object>(0, "  ")); //0
                    pairList.Add(new KeyValuePair<int, object>(0, "2、此表中的主要货物的招标要求请与技术专用其他部分保持一致，不得前后矛盾。")); //1
                    pairList.Add(new KeyValuePair<int, object>(0, "——")); //2
                    pairList.Add(new KeyValuePair<int, object>(0, "——")); //3
                    pairList.Add(new KeyValuePair<int, object>(0, "——")); //4
                    pairList.Add(new KeyValuePair<int, object>(0, "——")); //5
                    createRowsTreatDataStyleMaterialConfiguration(pairList, rowOutNum++);
                    pairList.Clear();


                    pairList.Add(new KeyValuePair<int, object>(0, "  ")); //0
                    pairList.Add(new KeyValuePair<int, object>(0, "3、请勿修改此表格式，画“—”处不需要填写。")); //1
                    pairList.Add(new KeyValuePair<int, object>(0, "——")); //2
                    pairList.Add(new KeyValuePair<int, object>(0, "——")); //3
                    pairList.Add(new KeyValuePair<int, object>(0, "——")); //4
                    pairList.Add(new KeyValuePair<int, object>(0, "——")); //5
                    createRowsTreatDataStyleMaterialConfiguration(pairList, rowOutNum++);
                    pairList.Clear();

                }

                void createRowsTreatDataStyleMaterialConfiguration(List<KeyValuePair<int, object>> inputDataList, int rowNum)
                {

                    // 3. 添加标题行
                    IRow iRow = sheetMaterialConfiguration.CreateRow(rowNum);
                    sheetMaterialConfiguration.AutoSizeRow(rowNum);
                    int colIdx = 0;
                    //for (int colIdx = 0; colIdx < 6; colIdx++)//方法2,3
                    foreach (var keyValuePair in inputDataList)//方法1
                    {
                        try
                        {
                            //ICell cell = iRow.GetCell(colIdx);//如果获得已有单元格，则这样写
                            ICell cell = iRow.CreateCell(colIdx);//IRow是地址引用，像指针，反过来赋值
                                                                 // cell.CellStyle = stringCenterStyle;
                                                                 //int styleInt = styleList[colIdx];//方法2
                            int styleInt = keyValuePair.Key;//方法1
                                                            // object value = inputDataList[colIdx];
                            object value = keyValuePair.Value;

                            // 根据数据类型应用样式
                            if (value == DBNull.Value)
                            {
                                // cell.SetCellValue("ut");// string.Empty;
                                // cell.CellStyle = nullStyle;// CreateNullStyle();
                                cell.SetCellValue("?数据未提供");
                                cell.CellStyle = warnStyle;
                            }
                            else if (value == "") cell.CellStyle = stringCenterStyle;
                            else
                            {
                                switch (value)
                                {
                                    case string string1:
                                        cell.SetCellValue(Convert.ToString(value));
                                        // cell.CellStyle = stringStyle;
                                        break;
                                    case int int1:
                                        cell.SetCellValue(Convert.ToString(value));
                                        // cell.CellStyle = itemStyle;
                                        break;
                                    case DateTime dateTime1:
                                        cell.SetCellValue(Convert.ToDateTime(value));
                                        // cell.CellStyle = dateTimeStyle;
                                        break;
                                    case double double1:
                                        cell.SetCellValue(Convert.ToDouble(value));
                                        if (Convert.ToDouble(value) < 0)
                                        {
                                            // cell.CellStyle = warnNumStyle;//numberStyle; //
                                            break;
                                        }
                                        // cell.CellStyle = numberStyle;
                                        break;
                                    default:
                                        cell.SetCellValue(value.ToString());
                                        // cell.CellStyle = stringStyle; //文本自动换行
                                        break;
                                }

                                if (rowNum == 0) cell.CellStyle = titleBlankBorderStyle;
                                else if (rowNum == 9 || rowNum == 10 || rowNum == 11) cell.CellStyle = stringLeftStyle;
                                else if (faultMarkArray.Any(term => value.ToString().Contains(term, StringComparison.OrdinalIgnoreCase)))
                                    cell.CellStyle = warnStyle;
                                else if (styleInt == 1)//方法2
                                    cell.CellStyle = stringBlueStyle;
                                else
                                    cell.CellStyle = stringCenterStyle;
                                if (value.ToString().Contains("说明", StringComparison.OrdinalIgnoreCase))
                                {
                                    //合并单元格，文字超出单元格范围，行高不会自动变化
                                    // iRow.Height = 576;// 14.4*2 *20;1/20个点为最小单位
                                    cell.CellStyle = stringCenterStyle;
                                }

                            }
                            Trace.Write($"{value.ToString()}  ");//20250519 打印单元格数据                
                        }                             //                   }
                        catch (Exception excep1)
                        {
                            Trace.WriteLine(excep1.Message);
                            Trace.WriteLine($"问题在第{rowNum.ToString()}行  ");
                        }
                        ++colIdx;
                    }
                    Trace.WriteLine(""); ;
                }

                void setColumnWidthMaterialConfiguration()
                {
                    sheetMaterialConfiguration.SetColumnWidth(0, (8.11 + 0.78) * 256);
                    sheetMaterialConfiguration.SetColumnWidth(1, (8.11 + 0.78) * 256);
                    sheetMaterialConfiguration.SetColumnWidth(2, (16.22 + 0.78) * 256);
                    sheetMaterialConfiguration.SetColumnWidth(3, (8.11 + 0.78) * 256);
                    sheetMaterialConfiguration.SetColumnWidth(4, (32.78 + 0.78) * 256);
                    sheetMaterialConfiguration.SetColumnWidth(5, (8.11 + 0.78) * 256);

                }

                void mergeCellsMaterialConfiguration() //合并单元格
                {
                    CellRangeAddress cellMerge = new CellRangeAddress(0, 0, 0, 5);//（起始行，结束行，起始列，结束列）
                    sheetMaterialConfiguration.AddMergedRegion(cellMerge);
                    CellRangeAddress cellMerge1 = new CellRangeAddress(2, 3, 0, 0);//（起始行，结束行，起始列，结束列）         
                    sheetMaterialConfiguration.AddMergedRegion(cellMerge1);
                    CellRangeAddress cellMerge2 = new CellRangeAddress(2, 3, 1, 1);//（起始行，结束行，起始列，结束列）
                    sheetMaterialConfiguration.AddMergedRegion(cellMerge2);
                    CellRangeAddress cellMerge3 = new CellRangeAddress(9, 11, 0, 0);//（起始行，结束行，起始列，结束列）
                    sheetMaterialConfiguration.AddMergedRegion(cellMerge3);
                    CellRangeAddress cellMerge4 = new CellRangeAddress(9, 9, 1, 5);//（起始行，结束行，起始列，结束列）
                    sheetMaterialConfiguration.AddMergedRegion(cellMerge4);
                    CellRangeAddress cellMerge81 = new CellRangeAddress(10, 10, 1, 5);
                    sheetMaterialConfiguration.AddMergedRegion(cellMerge81);
                    CellRangeAddress cellMerge82 = new CellRangeAddress(11, 11, 1, 5);
                    sheetMaterialConfiguration.AddMergedRegion(cellMerge82);

                }

            }



            void createSheetEnvironment()
            {

                feedDataByRowsForEnvironment();
                setColumnWidthEnvironment();
                mergeCellsEnvironment();

                void feedDataByRowsForEnvironment()
                {
                    int rowOutNum = 0;
                    var pairList = new List<KeyValuePair<int, object>>();
                    // 添加元素					



                    pairList.Add(new KeyValuePair<int, object>(0, "使用环境条件表")); //0
                    pairList.Add(new KeyValuePair<int, object>(0, "")); //1
                    pairList.Add(new KeyValuePair<int, object>(0, "")); //2
                    pairList.Add(new KeyValuePair<int, object>(0, "")); //3
                    createRowsTreatDataStyleEnvironment(pairList, rowOutNum++);
                    pairList.Clear();


                    pairList.Add(new KeyValuePair<int, object>(0, "名称")); //0
                    pairList.Add(new KeyValuePair<int, object>(0, "")); //1
                    pairList.Add(new KeyValuePair<int, object>(0, "参数值")); //2
                    pairList.Add(new KeyValuePair<int, object>(0, "投标人响应值")); //3
                    createRowsTreatDataStyleEnvironment(pairList, rowOutNum++);
                    pairList.Clear();

                    pairList.Add(new KeyValuePair<int, object>(0, "海拔高度（m）")); //0
                    pairList.Add(new KeyValuePair<int, object>(0, "")); //1
                    pairList.Add(new KeyValuePair<int, object>(0, "≤4000")); //2
                    pairList.Add(new KeyValuePair<int, object>(0, "≤4000")); //3
                    createRowsTreatDataStyleEnvironment(pairList, rowOutNum++);
                    pairList.Clear();

                    pairList.Add(new KeyValuePair<int, object>(0, "最高环境温度（℃）")); //0
                    pairList.Add(new KeyValuePair<int, object>(0, "")); //1
                    pairList.Add(new KeyValuePair<int, object>(0, "45")); //2
                    pairList.Add(new KeyValuePair<int, object>(0, "45")); //3
                    createRowsTreatDataStyleEnvironment(pairList, rowOutNum++);
                    pairList.Clear();

                    pairList.Add(new KeyValuePair<int, object>(0, "最低环境温度（℃）")); //0
                    pairList.Add(new KeyValuePair<int, object>(0, "")); //1
                    pairList.Add(new KeyValuePair<int, object>(0, "-40")); //2
                    pairList.Add(new KeyValuePair<int, object>(0, "-40")); //3
                    createRowsTreatDataStyleEnvironment(pairList, rowOutNum++);
                    pairList.Clear();

                    pairList.Add(new KeyValuePair<int, object>(0, "土壤最高环境温度（℃）")); //0
                    pairList.Add(new KeyValuePair<int, object>(0, "")); //1
                    pairList.Add(new KeyValuePair<int, object>(0, "35")); //2
                    pairList.Add(new KeyValuePair<int, object>(0, "35")); //3
                    createRowsTreatDataStyleEnvironment(pairList, rowOutNum++);
                    pairList.Clear();

                    pairList.Add(new KeyValuePair<int, object>(0, "土壤最低环境温度（℃）")); //0
                    pairList.Add(new KeyValuePair<int, object>(0, "")); //1
                    pairList.Add(new KeyValuePair<int, object>(0, "-20")); //2
                    pairList.Add(new KeyValuePair<int, object>(0, "-20")); //3
                    createRowsTreatDataStyleEnvironment(pairList, rowOutNum++);
                    pairList.Clear();

                    pairList.Add(new KeyValuePair<int, object>(0, "日照强度（W/cm2）")); //0
                    pairList.Add(new KeyValuePair<int, object>(0, "")); //1
                    pairList.Add(new KeyValuePair<int, object>(0, "0.1")); //2
                    pairList.Add(new KeyValuePair<int, object>(0, "0.1")); //3
                    createRowsTreatDataStyleEnvironment(pairList, rowOutNum++);
                    pairList.Clear();

                    pairList.Add(new KeyValuePair<int, object>(0, "湿度")); //0
                    pairList.Add(new KeyValuePair<int, object>(0, "日相对湿度平均值（％）")); //1
                    pairList.Add(new KeyValuePair<int, object>(0, "≤95")); //2
                    pairList.Add(new KeyValuePair<int, object>(0, "≤95")); //3
                    createRowsTreatDataStyleEnvironment(pairList, rowOutNum++);
                    pairList.Clear();

                    pairList.Add(new KeyValuePair<int, object>(0, "")); //0
                    pairList.Add(new KeyValuePair<int, object>(0, "月相对湿度平均值（％）")); //1
                    pairList.Add(new KeyValuePair<int, object>(0, "≤90")); //2
                    pairList.Add(new KeyValuePair<int, object>(0, "≤90")); //3
                    createRowsTreatDataStyleEnvironment(pairList, rowOutNum++);
                    pairList.Clear();

                    pairList.Add(new KeyValuePair<int, object>(0, "最大风速（户外）（m/s）/Pa")); //0
                    pairList.Add(new KeyValuePair<int, object>(0, "")); //1
                    pairList.Add(new KeyValuePair<int, object>(0, "35/700")); //2
                    pairList.Add(new KeyValuePair<int, object>(0, "35/700")); //3
                    createRowsTreatDataStyleEnvironment(pairList, rowOutNum++);
                    pairList.Clear();


                    pairList.Add(new KeyValuePair<int, object>(0, "电缆敷设方式\r\n(多种方式并存时，选择载流量最小的一种方式)")); //0
                    pairList.Add(new KeyValuePair<int, object>(0, "")); //1
                    pairList.Add(new KeyValuePair<int, object>(0, "直埋、排管、电缆沟、\r\n隧道、空气")); //2
                    pairList.Add(new KeyValuePair<int, object>(0, "直埋、排管、电缆沟、\r\n隧道、空气")); //3
                    createRowsTreatDataStyleEnvironment(pairList, rowOutNum++);
                    pairList.Clear();


                }

                void createRowsTreatDataStyleEnvironment(List<KeyValuePair<int, object>> inputDataList, int rowNum)
                {

                    // 3. 添加标题行
                    IRow iRow = sheetEnvironment.CreateRow(rowNum);
                    sheetEnvironment.AutoSizeRow(rowNum);
                    int colIdx = 0;
                    //for (int colIdx = 0; colIdx < 6; colIdx++)//方法2,3
                    foreach (var keyValuePair in inputDataList)//方法1
                    {
                        try
                        {
                            //ICell cell = iRow.GetCell(colIdx);//如果获得已有单元格，则这样写
                            ICell cell = iRow.CreateCell(colIdx);//IRow是地址引用，像指针，反过来赋值
                                                                 // cell.CellStyle = stringCenterStyle;
                                                                 //int styleInt = styleList[colIdx];//方法2
                            int styleInt = keyValuePair.Key;//方法1
                                                            // object value = inputDataList[colIdx];
                            object value = keyValuePair.Value;

                            // 根据数据类型应用样式
                            if (value == DBNull.Value)
                            {
                                // cell.SetCellValue("ut");// string.Empty;
                                // cell.CellStyle = nullStyle;// CreateNullStyle();
                                cell.SetCellValue("?数据未提供");
                                cell.CellStyle = warnStyle;
                            }
                            else if (value == "") cell.CellStyle = stringCenterStyle;
                            else
                            {
                                switch (value)
                                {
                                    case string string1:
                                        cell.SetCellValue(Convert.ToString(value));
                                        // cell.CellStyle = stringStyle;
                                        break;
                                    case int int1:
                                        cell.SetCellValue(Convert.ToString(value));
                                        // cell.CellStyle = itemStyle;
                                        break;
                                    case DateTime dateTime1:
                                        cell.SetCellValue(Convert.ToDateTime(value));
                                        // cell.CellStyle = dateTimeStyle;
                                        break;
                                    case double double1:
                                        cell.SetCellValue(Convert.ToDouble(value));
                                        if (Convert.ToDouble(value) < 0)
                                        {
                                            // cell.CellStyle = warnNumStyle;//numberStyle; //
                                            break;
                                        }
                                        // cell.CellStyle = numberStyle;
                                        break;
                                    default:
                                        cell.SetCellValue(value.ToString());
                                        // cell.CellStyle = stringStyle; //文本自动换行
                                        break;
                                }

                                if (rowNum == 0) cell.CellStyle = titleBlankBorderStyle;
                                else if (rowNum > 1 && colIdx < 2) cell.CellStyle = stringLeftStyle;
                                //else if (rowIndex == 11) cell.CellStyle = stringLeftStyle;
                                else
                                    cell.CellStyle = stringCenterStyle;

                            }
                            Trace.Write($"{value.ToString()}  ");//20250519 打印单元格数据                
                        }                             //                   }
                        catch (Exception excep1)
                        {
                            Trace.WriteLine(excep1.Message);
                            Trace.WriteLine($"问题在第{rowNum.ToString()}行  ");
                        }
                        ++colIdx;
                    }
                    Trace.WriteLine(""); ;
                }

                void setColumnWidthEnvironment()
                {
                    sheetEnvironment.SetColumnWidth(0, (6 + 0.78) * 256);
                    sheetEnvironment.SetColumnWidth(1, (30 + 0.78) * 256);
                    sheetEnvironment.SetColumnWidth(2, (23.5 + 0.78) * 256);
                    sheetEnvironment.SetColumnWidth(3, (23.5 + 0.78) * 256);

                }

                void mergeCellsEnvironment() //合并单元格
                {
                    CellRangeAddress cellMerge = new CellRangeAddress(0, 0, 0, 3);//（起始行，结束行，起始列，结束列）
                    sheetEnvironment.AddMergedRegion(cellMerge);
                    CellRangeAddress cellMerge1 = new CellRangeAddress(1, 1, 0, 1);//（起始行，结束行，起始列，结束列）         
                    sheetEnvironment.AddMergedRegion(cellMerge1);
                    CellRangeAddress cellMerge2 = new CellRangeAddress(2, 2, 0, 1);//（起始行，结束行，起始列，结束列）
                    sheetEnvironment.AddMergedRegion(cellMerge2);
                    CellRangeAddress cellMerge3 = new CellRangeAddress(3, 3, 0, 1);//（起始行，结束行，起始列，结束列）
                    sheetEnvironment.AddMergedRegion(cellMerge3);
                    CellRangeAddress cellMerge4 = new CellRangeAddress(4, 4, 0, 1);//（起始行，结束行，起始列，结束列）
                    sheetEnvironment.AddMergedRegion(cellMerge4);
                    CellRangeAddress cellMerge81 = new CellRangeAddress(5, 5, 0, 1);
                    sheetEnvironment.AddMergedRegion(cellMerge81);
                    CellRangeAddress cellMerge82 = new CellRangeAddress(6, 6, 0, 1);
                    sheetEnvironment.AddMergedRegion(cellMerge82);
                    CellRangeAddress cellMerge84 = new CellRangeAddress(7, 7, 0, 1);
                    sheetEnvironment.AddMergedRegion(cellMerge84);
                    CellRangeAddress cellMerge85 = new CellRangeAddress(8, 9, 0, 0);
                    sheetEnvironment.AddMergedRegion(cellMerge85);
                    CellRangeAddress cellMerge83 = new CellRangeAddress(10, 10, 0, 1);
                    sheetEnvironment.AddMergedRegion(cellMerge83);
                    CellRangeAddress cellMerge86 = new CellRangeAddress(11, 11, 0, 1);
                    sheetEnvironment.AddMergedRegion(cellMerge86);

                    /*
                    CellRangeAddress cellMerge87 = new CellRangeAddress(10, 31, 0, 0);
                    sheetEnvironment.AddMergedRegion(cellMerge87);
                    CellRangeAddress cellMerge88 = new CellRangeAddress(10, 10, 1, 3);
                    sheetEnvironment.AddMergedRegion(cellMerge88);
                    CellRangeAddress cellMerge89 = new CellRangeAddress(11, 11, 1, 2);
                    sheetEnvironment.AddMergedRegion(cellMerge89);
                    CellRangeAddress cellMerge6 = new CellRangeAddress(12, 12, 1, 2);
                    sheetEnvironment.AddMergedRegion(cellMerge6);
                    CellRangeAddress cellMerge7 = new CellRangeAddress(13, 13, 1, 2);
                    sheetEnvironment.AddMergedRegion(cellMerge7);
                    /*
                  CellRangeAddress cellMerge8 = new CellRangeAddress(14, 14, 1, 2);
                  sheetEnvironment.AddMergedRegion(cellMerge8);
                  CellRangeAddress cellMerge9 = new CellRangeAddress(15, 15, 1, 2);
                  sheetEnvironment.AddMergedRegion(cellMerge9);
                  CellRangeAddress cellMerge10 = new CellRangeAddress(16, 16, 1, 2);
                  sheetEnvironment.AddMergedRegion(cellMerge10);
                  CellRangeAddress cellMerge11 = new CellRangeAddress(17, 17, 1, 2);
                  sheetEnvironment.AddMergedRegion(cellMerge11);
                  CellRangeAddress cellMerge12 = new CellRangeAddress(18, 18, 1, 2);
                  sheetEnvironment.AddMergedRegion(cellMerge12);
                  CellRangeAddress cellMerge13 = new CellRangeAddress(19, 19, 1, 2);
                  sheetEnvironment.AddMergedRegion(cellMerge13);
                  CellRangeAddress cellMerge14 = new CellRangeAddress(20, 20, 1, 2);
                  sheetEnvironment.AddMergedRegion(cellMerge14);
                  CellRangeAddress cellMerge15 = new CellRangeAddress(21, 21, 1, 2);
                  sheetEnvironment.AddMergedRegion(cellMerge15);
                  CellRangeAddress cellMerge17 = new CellRangeAddress(22, 22, 1, 2);
                  sheetEnvironment.AddMergedRegion(cellMerge17);
                  CellRangeAddress cellMerge18 = new CellRangeAddress(23, 23, 1, 2);
                  sheetEnvironment.AddMergedRegion(cellMerge18);
                  CellRangeAddress cellMerge16 = new CellRangeAddress(24, 24, 1, 2);
                  sheetEnvironment.AddMergedRegion(cellMerge16);





                  CellRangeAddress cellMerge25 = new CellRangeAddress(24, 25, 3, 3);
                  sheetEnvironment.AddMergedRegion(cellMerge25);
                  CellRangeAddress cellMerge24 = new CellRangeAddress(24, 25, 4, 4);
                  sheetEnvironment.AddMergedRegion(cellMerge24);
                  CellRangeAddress cellMerge30 = new CellRangeAddress(24, 25, 5, 5);
                  sheetEnvironment.AddMergedRegion(cellMerge30);
                  CellRangeAddress cellMerge34 = new CellRangeAddress(24, 25, 6, 6);
                  sheetEnvironment.AddMergedRegion(cellMerge34);
                  CellRangeAddress cellMerge31 = new CellRangeAddress(24, 25, 7, 7);
                  sheetEnvironment.AddMergedRegion(cellMerge31);
                  CellRangeAddress cellMerge32 = new CellRangeAddress(24, 25, 8, 8);
                  sheetEnvironment.AddMergedRegion(cellMerge32);


                  CellRangeAddress cellMerge33 = new CellRangeAddress(26, 27, 3, 3);
                  sheetEnvironment.AddMergedRegion(cellMerge33);
                  CellRangeAddress cellMerge5 = new CellRangeAddress(26, 27, 4, 4);
                  sheetEnvironment.AddMergedRegion(cellMerge5);
                  CellRangeAddress cellMerge35 = new CellRangeAddress(26, 27, 5, 5);
                  sheetEnvironment.AddMergedRegion(cellMerge35);
                  CellRangeAddress cellMerge36 = new CellRangeAddress(26, 27, 6, 6);
                  sheetEnvironment.AddMergedRegion(cellMerge36);
                  CellRangeAddress cellMerge37 = new CellRangeAddress(26, 27, 7, 7);
                  sheetEnvironment.AddMergedRegion(cellMerge37);
                  CellRangeAddress cellMerge38 = new CellRangeAddress(26, 27, 8, 8);
                  sheetEnvironment.AddMergedRegion(cellMerge38);






                  CellRangeAddress cellMerge20 = new CellRangeAddress(25, 25, 1, 2);
                  sheetEnvironment.AddMergedRegion(cellMerge20);
                  CellRangeAddress cellMerge21 = new CellRangeAddress(26, 26, 1, 2);
                  sheetEnvironment.AddMergedRegion(cellMerge21);
                  CellRangeAddress cellMerge23 = new CellRangeAddress(27, 27, 1, 2);
                  sheetEnvironment.AddMergedRegion(cellMerge23);
                  CellRangeAddress cellMerge26 = new CellRangeAddress(28, 28, 1, 2);
                  sheetEnvironment.AddMergedRegion(cellMerge26);
                  CellRangeAddress cellMerge27 = new CellRangeAddress(29, 29, 1, 2);
                  sheetEnvironment.AddMergedRegion(cellMerge27);
                  CellRangeAddress cellMerge28 = new CellRangeAddress(30, 30, 1, 2);
                  sheetEnvironment.AddMergedRegion(cellMerge28);
                  CellRangeAddress cellMerge29 = new CellRangeAddress(31, 31, 1, 2);
                  sheetEnvironment.AddMergedRegion(cellMerge29);
                  */

                }

            }



            void saveExcel(XSSFWorkbook theWorkbook, string theFileName)
            {
                try
                {


                    using (SaveFileDialog saveFileDialog = new SaveFileDialog())//弹出对话框，可指定Excel 存储路径，文件名
                    {
                        saveFileDialog.Filter = "Excel文件|*.xlsx";
                        saveFileDialog.Title = "保存Excel文件";
                        saveFileDialog.FileName = theFileName + DateTime.Now.ToString("yyyy-MM-dd HH-mm-ss-fff") + ".xlsx";

                        // string filePath = "output.xlsx";//默认在...bin\里面
                        //string filePath = @"C:\Temp\MySqlOutput.xlsx";

                        //string tempPath = Path.GetTempPath();//C:\Users\Admin\AppData\Local\Temp
                        //string filePath = Path.Combine(tempPath, "MySQLOutput.xlsx");

                        if (saveFileDialog.ShowDialog() == DialogResult.OK)//选择保存路径
                        {
                            // string filePath = "output.xlsx";//默认在...bin\里面

                            //string filePath = @"C:\Temp\MySqlOutput.xlsx";

                            //string tempPath = Path.GetTempPath();//C:\Users\Admin\AppData\Local\Temp
                            //string filePath = Path.Combine(tempPath, "MySQLOutput.xlsx");

                            string filePath = saveFileDialog.FileName;

                            Directory.CreateDirectory(Path.GetDirectoryName(filePath)); //创建目录
                                                                                        // 4.1 （FileStream）
                            using (var fs = new FileStream(filePath, FileMode.Create))
                            // FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                            {
                                theWorkbook.Write(fs);
                            }
                        }
                    }
                    Trace.WriteLine("文件保存成功");
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"创建Excel文件失败: {ex.ToString()}");
                    Trace.WriteLine($"创建Excel文件失败: {ex.ToString()}");
                }
                finally
                {
                    // 5. 手动清理资源（根据NPOI版本可能需要）
                    if (workbook is IDisposable disposable)
                    {
                        disposable.Dispose();
                    }
                }
            }

        }




        public class CreateWord
        {

            //1 -----------------------------数据输入---------------------------------


            //2 --------------------------------创建word 含table -3种方式------------------------------------------------
            //2.1 方法1：用 List<List<KeyValuePair<int, object>>> 输入数据



            public void CreateWordTechKeyValuePair()
            {
                //1 创建word文档
                var aWord = new XWPFDocument();

                //2 添加内容


                //2.0 ----封面----
                //2.0.1 Word 布局
                theMargin(aWord, theTop: 1800); //页边距

                //2.0.2 添加标题, 文字行
                var titlePara = aWord.CreateParagraph();
                titlePara.Alignment = ParagraphAlignment.CENTER;
                rowDistance(titlePara, 50, (int)LineSpacingRule.EXACT);
                var titleRun = titlePara.CreateRun();
                titleRun.FontSize = 26;
                titleRun.SetText("国家电网公司集中规模招标采购");
                titleRun.IsBold = true;
                titleRun.FontFamily = "宋体";
                //titleRun.SetFontFamily("宋体", FontCharRange.CS); //这代码只改变汉字字体，英语自动

                //2.0.3 添加文字
                var XWPFParagraph1 = aWord.CreateParagraph();
                XWPFParagraph1.Alignment = ParagraphAlignment.CENTER;
                rowDistance(XWPFParagraph1, 50, (int)LineSpacingRule.EXACT);
                SetRunStyle(XWPFParagraph1, "国家电网公司总部", fontSize: 24);
                SetRunStyle(XWPFParagraph1, "配网普通固化ID编制", fontSize: 24);
                // SetRunStyle(XWPFParagraph1, $"{voltageLevel}电力电缆", fontSize: 24);
                SetRunStyle(XWPFParagraph1, $"{voltageLevel}", fontSize: 24, color: "Blue", breakChangeRow: false);
                SetRunStyle(XWPFParagraph1, "电力电缆", fontSize: 24);
                //SetRunStyle(XWPFParagraph1, "(A171-500135730-00001)", underline: UnderlinePatterns.Single, fontSize: 24, isBold: true);
                SetRunStyle(XWPFParagraph1, "(标书编号自行输入)", color: "Red", underline: UnderlinePatterns.Single, fontSize: 24, isBold: true);
                SetRunStyle(XWPFParagraph1, "投标文件", fontSize: 26, isBold: true);
                //2.0.4  插入分节符
                addBreak(XWPFParagraph1, isPageBreak: true);


                //if (type_spec.Contains("0.6/1")) voltageLevel = "低压"; else voltageLevel = "高压？中压？低压？";
                //if (isArmoured) armourString = "铠装"; else armourString = "无铠装";
                //if (isDoubleCable) coreString = "双缆"; else if (isMultiCore) coreString = "多芯"; else coreString = "单芯";

                //2.1 ----第一页----
                //2.1.1 页面布局
                ulong pageWidth = theMargin(aWord, theLeft: 567 * 2, theRight: 567 * 2); //页边距

                //2.1.2 插入文字
                var XWPFParagraph2 = aWord.CreateParagraph();
                // SetRunStyle(XWPFParagraph2, "1　标准技术参数", fontFamily: "宋体");
                SetRunStyle(XWPFParagraph2, "1　标准技术参数");
                // SetRunStyle(XWPFParagraph2, $"    技术参数特性表是国家电网公司对采购设备的基础技术参数要求，在招投标过程中，投标人应该依据招标文件，对技术参数特性表中标准参数值进行响应。{voltageLevel}{coreString}电力电缆技术参数特性见表1。");
                SetRunStyle(XWPFParagraph2, "   技术参数特性表是国家电网公司对采购设备的基础技术参数要求，在招投标过程中，投标人应该依据招标文件，对技术参数特性表中标准参数值进行响应。", breakChangeRow: false);
                SetRunStyle(XWPFParagraph2, $"{voltageLevel}{coreString}", color: "Blue", breakChangeRow: false);
                SetRunStyle(XWPFParagraph2, "电力电缆技术参数特性见表1。");
                SetRunStyle(XWPFParagraph2, "");
                //  SetRunStyle(XWPFParagraph2, "表1　技术参数特性表", fontFamily: "宋体", breakChangeRow: false);
                SetRunStyle(XWPFParagraph2, "   表1  技术参数特性表", breakChangeRow: false);
                thePageNumber(aWord, textForPage: 2);//页码
                                                     //addBreak(XWPFParagraph2);
                                                     //2.1.3 创建表格: 电缆结构技术参数
                XWPFTable tableConstructionPara = CreateTable(TableDataInputConstruction(), aWord);
                //2.1.3.1 设置表格宽度（1英寸=1440 twips, 1厘米≈567 twips, 1 磅 = 20 twips）
                // ColumnWidth(tableConstructionPara, 15 * 57, 32 * 57, 15 * 57, 34 * 57, 38 * 56, 34 * 57);
                ColumnWidthProportion(tableConstructionPara, pageWidth, 15, 32, 15, 34, 38, 34);
                //2.1.3.2 设置单元格对齐方式， 默认左上对齐
                // positioning(tableQuote);
                //2.1.3.3 合并单元格
                MergeCellConstruction();



                //2.2 ----第二页----
                var XWPFParagraph3 = aWord.CreateParagraph();
                //2.2.0 换页
                addBreak(XWPFParagraph3);

                //2.2.1 创建表格：电缆非电气技术参数
                XWPFTable tableNonElecticPara = CreateTable(TableDataInputNonElectric(), aWord);
                //2.2.1.2 设置表格宽度（1英寸=1440 twips, 1厘米≈567 twips, 1 磅 = 20 twips）
                // ColumnWidth(tableNonElecticPara, 11 * 57, 11 * 57, 44 * 57, 15 * 57, 14 * 56, 12 * 57, 15 * 57, 15 * 57, 32 * 57);             
                ColumnWidthProportion(tableNonElecticPara, pageWidth, 11, 11, 44, 15, 14, 12, 15, 15, 32);
                //2.2.1.3 设置单元格对齐方式， 默认左上对齐
                //positioning(tableNonElecticPara);
                //2.2.1.4 合并单元格
                MergeCellNonElectric();



                //2.3 ----第三页----
                var XWPFParagraph4 = aWord.CreateParagraph();
                //2.3.0 换页
                addBreak(XWPFParagraph4);
                //2.3.1 插入文字
                // var XWPFParagraph5 = aWord.CreateParagraph();
                SetRunStyle(XWPFParagraph4, "2　组件材料配置表");
                SetRunStyle(XWPFParagraph4, "");
                SetRunStyle(XWPFParagraph4, "       表2  组件材料配置表", breakChangeRow: false);

                //2.3.2 创建表格：材料配置
                XWPFTable tableMaterial = CreateTable(TableDataInputMaterial(), aWord, tableMark: "Material");
                //2.3.2.1 设置表格宽度（1英寸=1440 twips, 1厘米≈567 twips, 1 磅 = 20 twips）
                // ColumnWidth(tableMaterial, 17 * 57, 19 * 57, 33 * 57, 15 * 57, 64 * 57, 21 * 57);
                //ColumnWidthProportion(tableConstructionPara, pageWidth, 15 * 57, 32 * 57, 15 * 57, 34 * 57, 38 * 56, 34 * 57);
                ColumnWidthProportion(tableMaterial, pageWidth, 17, 19, 33, 15, 64, 21);
                //2.3.2.2 设置单元格对齐方式， 默认左上对齐
                //positioning(tableMaterial);
                //2.3.2.3 合并单元格
                MergeCellMaterial();


                //2.4 ----第四页----              
                var XWPFParagraph6 = aWord.CreateParagraph();
                //2.4.0 换页
                //addBreak(XWPFParagraph6);

                //2.4.1 插入文字
                SetRunStyle(XWPFParagraph6, ""); //空白白行
                SetRunStyle(XWPFParagraph6, "3　使用环境条件表");
                SetRunStyle(XWPFParagraph6, "");
                SetRunStyle(XWPFParagraph6, "    表3  使用环境条件表", breakChangeRow: false);

                //2.4.2 创建表格：环境条件
                XWPFTable tableEnvironment = CreateTable(TableDataInputEnvironment(), aWord, tableMark: "Environment");
                //2.4.2.1 设置表格宽度（1英寸=1440 twips, 1厘米≈567 twips, 1 磅 = 20 twips）
                // ColumnWidth(tableEnvironment, 17 * 57, 58 * 57, 47 * 57, 47 * 57);
                ColumnWidthProportion(tableEnvironment, pageWidth, 17, 58, 47, 47);
                //2.4.2.2 设置单元格对齐方式， 默认左上对齐
                //positioning(tableEnvironment);
                //2.4.2.3 合并单元格
                MergeCellEnvironment();



                //5  保存文档

                //saveWord(aWord, "Output工艺文件");               
                saveWord(aWord, $"工艺参数{type_specInFileName} ");


                //6   内部方法
                //6.1 输入数据至tableConstructionPara
                //6.1.1  至tableConstructionPara
                List<List<KeyValuePair<int, object>>> TableDataInputConstruction()
                {
                    var tableData = new List<List<KeyValuePair<int, object>>>();   //table数据

                    //string forhead = Regex.Match(theType, @"(\w+)\-").Groups[1].Value;
                    //string flameRedartant = Regex.Replace(forhead, @"w|d|W|D", "");
                    //if (flameRedartant == "") flameRedartant = "不适用";
                    //Trace.WriteLine($"flameRedartant: {flameRedartant}   ");

                    //string halogenFree = (Regex.Match(forhead, @"(w|W)").Value);
                    //if (halogenFree == "") halogenFree = "不适用";
                    //Trace.WriteLine($"halogenFree: {halogenFree}");
                    //string smokeFree = Regex.Match(forhead, @"(d|D)").Value;
                    //if (smokeFree == "") smokeFree = "不适用";
                    //Trace.WriteLine($"smokingFree: {smokeFree}");
                    ///*
                    //                MatchCollection matchCollection = Regex.Matches(spec, @"(\d×)|(\dX)|(\dx)");
                    //                foreach (MatchRegex match1 in matchCollection)
                    //                {
                    //                    MatchRegex theMatch = Regex.Match(match1.Value, @"\d");
                    //                    //   Trace.WriteLine($"数字：{theMatch.Value}");
                    //                    isMultiCore = (Convert.ToInt16(theMatch.Value) > 1);
                    //                    Trace.WriteLine($"theMatch.Value: {Convert.ToInt16(theMatch.Value)}");
                    //                    // specMini.Add(match1.Value);
                    //                }*/
                    //string outer_sheathMaterialFront = Regex.Match(outer_sheathMaterialSelected, @"(.+)（").Groups[1].Value;
                    //Trace.WriteLine($"outer_sheathMaterialFront: {outer_sheathMaterialFront}");




                    var rowList1 = new List<KeyValuePair<int, object>>();  //行数据
                    rowList1.Add(new KeyValuePair<int, object>(0, "电缆结构技术参数"));
                    rowList1.Add(new KeyValuePair<int, object>(0, "——"));
                    rowList1.Add(new KeyValuePair<int, object>(0, "——"));
                    rowList1.Add(new KeyValuePair<int, object>(0, "——"));
                    rowList1.Add(new KeyValuePair<int, object>(0, "——"));
                    rowList1.Add(new KeyValuePair<int, object>(0, "——"));
                    tableData.Add(rowList1);

                    var rowList2 = new List<KeyValuePair<int, object>>();  //行数据
                    rowList2.Add(new KeyValuePair<int, object>(0, "电缆型号"));
                    rowList2.Add(new KeyValuePair<int, object>(0, "——"));
                    rowList2.Add(new KeyValuePair<int, object>(1, type_spec));
                    rowList2.Add(new KeyValuePair<int, object>(0, "——"));
                    rowList2.Add(new KeyValuePair<int, object>(0, "——"));
                    rowList2.Add(new KeyValuePair<int, object>(0, "——"));
                    tableData.Add(rowList2);

                    var rowList3 = new List<KeyValuePair<int, object>>();  //行数据
                    rowList3.Add(new KeyValuePair<int, object>(0, "项　　目"));
                    rowList3.Add(new KeyValuePair<int, object>(0, "——"));
                    rowList3.Add(new KeyValuePair<int, object>(0, "单位"));
                    rowList3.Add(new KeyValuePair<int, object>(0, "标准参数值"));
                    rowList3.Add(new KeyValuePair<int, object>(0, "投标人响应值"));
                    rowList3.Add(new KeyValuePair<int, object>(0, "备注"));
                    tableData.Add(rowList3);

                    var rowList4 = new List<KeyValuePair<int, object>>();  //行数据
                    rowList4.Add(new KeyValuePair<int, object>(0, "铜导体"));
                    rowList4.Add(new KeyValuePair<int, object>(0, "材料"));
                    rowList4.Add(new KeyValuePair<int, object>(0, "——"));
                    rowList4.Add(new KeyValuePair<int, object>(1, conductorMaterial));  //type_spec.Contains('L', StringComparison.OrdinalIgnoreCase) ? "铝" : "铜")); // "铜"));//变色
                    rowList4.Add(new KeyValuePair<int, object>(1, conductorMaterial));  //type_spec.Contains('L', StringComparison.OrdinalIgnoreCase) ? "铝" : "铜")); // "铜"));//变色
                    rowList4.Add(new KeyValuePair<int, object>(0, "——"));
                    tableData.Add(rowList4);

                    var rowList5 = new List<KeyValuePair<int, object>>();  //行数据
                    rowList5.Add(new KeyValuePair<int, object>(0, "——"));
                    rowList5.Add(new KeyValuePair<int, object>(0, "芯数×标称截面"));
                    rowList5.Add(new KeyValuePair<int, object>(0, "芯×mm²"));
                    rowList5.Add(new KeyValuePair<int, object>(1, spec));// 待输入
                    rowList5.Add(new KeyValuePair<int, object>(1, spec));// 待输入
                    rowList5.Add(new KeyValuePair<int, object>(0, "——"));
                    tableData.Add(rowList5);

                    var rowList6 = new List<KeyValuePair<int, object>>();  //行数据
                    rowList6.Add(new KeyValuePair<int, object>(0, "——"));
                    rowList6.Add(new KeyValuePair<int, object>(0, "结构形式"));
                    rowList6.Add(new KeyValuePair<int, object>(0, "芯×mm²"));
                    rowList6.Add(new KeyValuePair<int, object>(0, "紧压圆形 / 实心导体"));
                    rowList6.Add(new KeyValuePair<int, object>(0, "紧压圆形 / 实心导体"));
                    rowList6.Add(new KeyValuePair<int, object>(0, "固定不变？"));
                    tableData.Add(rowList6);

                    var rowList7 = new List<KeyValuePair<int, object>>();  //行数据
                    rowList7.Add(new KeyValuePair<int, object>(0, "——"));
                    rowList7.Add(new KeyValuePair<int, object>(0, "最少单线根数"));
                    rowList7.Add(new KeyValuePair<int, object>(0, "根"));
                    rowList7.Add(new KeyValuePair<int, object>(1, pieces_1));// "?"));// 待输入  
                    rowList7.Add(new KeyValuePair<int, object>(1, pieces_1));// "?"));// 待输入 pieces_1)); //
                    rowList7.Add(new KeyValuePair<int, object>(1, areaConductor[0]));//"对应10mm²截面"));// 待输入
                    tableData.Add(rowList7);

                    if (isDoubleCable)
                    {
                        var rowList8 = new List<KeyValuePair<int, object>>();  //行数据
                        rowList8.Add(new KeyValuePair<int, object>(0, "——"));
                        rowList8.Add(new KeyValuePair<int, object>(0, "——"));
                        rowList8.Add(new KeyValuePair<int, object>(0, "——"));
                        rowList8.Add(new KeyValuePair<int, object>(1, pieces_2));//"3"));// 待输入
                        Trace.WriteLine($"pieces_2={pieces_2}");
                        rowList8.Add(new KeyValuePair<int, object>(1, pieces_2));//"2"));// 待输入
                        rowList8.Add(new KeyValuePair<int, object>(1, areaConductor[1])); //"对应6mm²截面"));// 待输入
                        tableData.Add(rowList8);
                    }

                    var rowList9 = new List<KeyValuePair<int, object>>();  //行数据
                    rowList9.Add(new KeyValuePair<int, object>(0, "——"));
                    rowList9.Add(new KeyValuePair<int, object>(0, "导体外径（近似值）"));
                    rowList9.Add(new KeyValuePair<int, object>(0, "mm"));
                    rowList9.Add(new KeyValuePair<int, object>(0, "（投标人提供）"));
                    rowList9.Add(new KeyValuePair<int, object>(1, conductDiameter_1));////4.1));  // 待输入
                    rowList9.Add(new KeyValuePair<int, object>(1, areaConductor[0])); //"对应10mm²截面")); // 待输入
                    tableData.Add(rowList9);

                    if (isDoubleCable)
                    {
                        var rowList10 = new List<KeyValuePair<int, object>>();  //行数据
                        rowList10.Add(new KeyValuePair<int, object>(0, "——"));
                        rowList10.Add(new KeyValuePair<int, object>(0, "——"));
                        rowList10.Add(new KeyValuePair<int, object>(0, "——"));
                        rowList10.Add(new KeyValuePair<int, object>(0, "——"));
                        rowList10.Add(new KeyValuePair<int, object>(1, conductDiameter_2));//2.2));  // 待输入
                        rowList10.Add(new KeyValuePair<int, object>(1, areaConductor[1])); //"对应4mm²截面")); // 待输入
                        tableData.Add(rowList10);
                    }

                    var rowList11 = new List<KeyValuePair<int, object>>();  //行数据
                    rowList11.Add(new KeyValuePair<int, object>(0, "——"));
                    rowList11.Add(new KeyValuePair<int, object>(0, "紧压系数"));
                    rowList11.Add(new KeyValuePair<int, object>(0, "——"));
                    rowList11.Add(new KeyValuePair<int, object>(0, "≥0.9"));
                    rowList11.Add(new KeyValuePair<int, object>(0, "≥0.9\r\n(对应紧压圆形导体结构)"));
                    rowList11.Add(new KeyValuePair<int, object>(0, "——"));
                    tableData.Add(rowList11);

                    Trace.WriteLine($"insulationMaterialSelected: {insulationMaterialSelected}");
                    var rowList12 = new List<KeyValuePair<int, object>>();  //行数据
                    rowList12.Add(new KeyValuePair<int, object>(0, "绝缘"));
                    rowList12.Add(new KeyValuePair<int, object>(0, "材料"));
                    rowList12.Add(new KeyValuePair<int, object>(0, "——"));
                    rowList12.Add(new KeyValuePair<int, object>(1, insulationMaterialSelected)); //insulation1Material));// "XLPE"));// 待输入
                    rowList12.Add(new KeyValuePair<int, object>(1, insulationMaterialSelected));//insulation1Material));// "XLPE"));// 待输入
                    rowList12.Add(new KeyValuePair<int, object>(0, "——"));
                    tableData.Add(rowList12);

                    var rowList13 = new List<KeyValuePair<int, object>>();  //行数据
                    rowList13.Add(new KeyValuePair<int, object>(0, "——"));
                    rowList13.Add(new KeyValuePair<int, object>(0, "平均厚度不小于标称厚度 t"));
                    rowList13.Add(new KeyValuePair<int, object>(0, "mm"));
                    rowList13.Add(new KeyValuePair<int, object>(0, "（投标人提供）"));
                    rowList13.Add(new KeyValuePair<int, object>(1, insulationThick_1));//"0.7"));// 待输入
                    rowList13.Add(new KeyValuePair<int, object>(1, areaConductor[0])); //"对应10截面"));// 待输入
                    tableData.Add(rowList13);

                    if (isDoubleCable)
                    {
                        var rowList14 = new List<KeyValuePair<int, object>>();  //行数据
                        rowList14.Add(new KeyValuePair<int, object>(0, "——"));
                        rowList14.Add(new KeyValuePair<int, object>(0, "——"));
                        rowList14.Add(new KeyValuePair<int, object>(0, "——"));
                        rowList14.Add(new KeyValuePair<int, object>(0, "——"));
                        rowList14.Add(new KeyValuePair<int, object>(1, insulationThick_2));//"0.7"));// 待输入
                        rowList14.Add(new KeyValuePair<int, object>(1, areaConductor[1])); //"对应6截面"));// 待输入
                        tableData.Add(rowList14);
                    }

                    var rowList15 = new List<KeyValuePair<int, object>>();  //行数据
                    rowList15.Add(new KeyValuePair<int, object>(0, "——"));
                    rowList15.Add(new KeyValuePair<int, object>(0, "最薄点厚度不小于标称值"));
                    rowList15.Add(new KeyValuePair<int, object>(0, "mm"));
                    rowList15.Add(new KeyValuePair<int, object>(0, "90 % t"));
                    rowList15.Add(new KeyValuePair<int, object>(0, "90 % t"));
                    rowList15.Add(new KeyValuePair<int, object>(0, "——"));
                    tableData.Add(rowList15);

                    var rowList16 = new List<KeyValuePair<int, object>>();  //行数据
                    rowList16.Add(new KeyValuePair<int, object>(0, "——"));
                    rowList16.Add(new KeyValuePair<int, object>(0, "偏心度"));
                    rowList16.Add(new KeyValuePair<int, object>(0, "%"));
                    rowList16.Add(new KeyValuePair<int, object>(0, "10"));
                    rowList16.Add(new KeyValuePair<int, object>(0, "≤10"));
                    rowList16.Add(new KeyValuePair<int, object>(0, "——"));
                    tableData.Add(rowList16);
                    if (isDoubleCable || isMultiCore)
                    {
                        var rowList17 = new List<KeyValuePair<int, object>>();  //行数据
                        rowList17.Add(new KeyValuePair<int, object>(0, "填充层"));
                        rowList17.Add(new KeyValuePair<int, object>(0, "填充材料"));
                        rowList17.Add(new KeyValuePair<int, object>(0, "——"));
                        rowList17.Add(new KeyValuePair<int, object>(0, "（投标人提供）"));
                        rowList17.Add(new KeyValuePair<int, object>(1, bufferMaterialSelected)); //buffer1Material));//buffer1Material ?? "无"));// 待输入
                        rowList17.Add(new KeyValuePair<int, object>(0, "——"));
                        tableData.Add(rowList17);
                    }
                    if (Convert.ToDouble(inner_sheathWeight) >= 1)
                    {
                        var rowList18 = new List<KeyValuePair<int, object>>();  //行数据
                        rowList18.Add(new KeyValuePair<int, object>(0, "内衬层"));
                        rowList18.Add(new KeyValuePair<int, object>(0, "材料"));
                        rowList18.Add(new KeyValuePair<int, object>(0, "——"));
                        rowList18.Add(new KeyValuePair<int, object>(0, "（投标人提供）"));
                        rowList18.Add(new KeyValuePair<int, object>(1, inner_sheathMaterial));//"H - 90 PVC护套料"));  // 待输入 
                        rowList18.Add(new KeyValuePair<int, object>(0, "——"));
                        tableData.Add(rowList18);

                        var rowList19 = new List<KeyValuePair<int, object>>();  //行数据
                        rowList19.Add(new KeyValuePair<int, object>(0, "——"));
                        rowList19.Add(new KeyValuePair<int, object>(0, "厚度\r\n（依据GB/T 12706.1假定外径对应选取）"));
                        rowList19.Add(new KeyValuePair<int, object>(0, "mm"));
                        rowList19.Add(new KeyValuePair<int, object>(0, "（投标人提供）"));
                        rowList19.Add(new KeyValuePair<int, object>(1, inner_thick)); //"1.0 - 2.0"));  // 待输入
                        rowList19.Add(new KeyValuePair<int, object>(0, "——"));
                        tableData.Add(rowList19);
                    }

                    if (isArmoured)
                    {
                        var rowList20 = new List<KeyValuePair<int, object>>();  //行数据
                        rowList20.Add(new KeyValuePair<int, object>(0, "铠装层"));
                        rowList20.Add(new KeyValuePair<int, object>(0, "材料"));
                        rowList20.Add(new KeyValuePair<int, object>(0, "——"));
                        rowList20.Add(new KeyValuePair<int, object>(0, "（投标人提供）"));
                        rowList20.Add(new KeyValuePair<int, object>(1, armourMaterialSelected)); //(isDoubleCable || isMultiCore) ? armour2Material : armour1Material)); //"多芯采用镀锌钢带" : "单芯采用不锈钢带"));  // 待输入
                        Trace.WriteLine($"isDoubleCable: {isDoubleCable}    isMultiCore:  {isMultiCore}");
                        rowList20.Add(new KeyValuePair<int, object>(0, "与供货需求表一致"));
                        tableData.Add(rowList20);

                        var rowList21 = new List<KeyValuePair<int, object>>();  //行数据
                        rowList21.Add(new KeyValuePair<int, object>(0, "——"));
                        rowList21.Add(new KeyValuePair<int, object>(0, "钢带厚度/钢丝直径\r\n（依据GB/T 12706.1假定外径对应选取）"));
                        rowList21.Add(new KeyValuePair<int, object>(0, "mm"));
                        rowList21.Add(new KeyValuePair<int, object>(0, "（投标人提供）"));
                        rowList21.Add(new KeyValuePair<int, object>(1, steel_thick)); //"0.2~0.5"));  // 待输入
                        rowList21.Add(new KeyValuePair<int, object>(0, "——"));
                        tableData.Add(rowList21);

                        var rowList22 = new List<KeyValuePair<int, object>>();  //行数据
                        rowList22.Add(new KeyValuePair<int, object>(0, "——"));
                        rowList22.Add(new KeyValuePair<int, object>(0, "钢带层数"));
                        rowList22.Add(new KeyValuePair<int, object>(0, "层"));
                        rowList22.Add(new KeyValuePair<int, object>(1, 2)); // 待输入
                        rowList22.Add(new KeyValuePair<int, object>(1, 2));  // 待输入
                        rowList22.Add(new KeyValuePair<int, object>(0, "固定不变？"));
                        tableData.Add(rowList22);

                        var rowList23 = new List<KeyValuePair<int, object>>();  //行数据
                        rowList23.Add(new KeyValuePair<int, object>(0, "——"));
                        rowList23.Add(new KeyValuePair<int, object>(0, "钢带宽度"));
                        rowList23.Add(new KeyValuePair<int, object>(0, "mm"));
                        rowList23.Add(new KeyValuePair<int, object>(0, "（投标人提供）"));
                        rowList23.Add(new KeyValuePair<int, object>(1, steel_width)); // 待输入
                        rowList23.Add(new KeyValuePair<int, object>(0, "——"));
                        tableData.Add(rowList23);

                    }

                    var rowList24 = new List<KeyValuePair<int, object>>();  //行数据
                    rowList24.Add(new KeyValuePair<int, object>(0, "外护套"));
                    rowList24.Add(new KeyValuePair<int, object>(0, "材料"));
                    rowList24.Add(new KeyValuePair<int, object>(0, "——"));
                    rowList24.Add(new KeyValuePair<int, object>(0, "（投标人提供）"));
                    rowList24.Add(new KeyValuePair<int, object>(1, outer_sheathMaterialFront));// outer_sheath1Material));//"ZH-90 PVC护套料")); // 待输入
                    rowList24.Add(new KeyValuePair<int, object>(0, "——"));
                    tableData.Add(rowList24);

                    var rowList25 = new List<KeyValuePair<int, object>>();  //行数据
                    rowList25.Add(new KeyValuePair<int, object>(0, "——"));
                    rowList25.Add(new KeyValuePair<int, object>(0, "颜色"));
                    rowList25.Add(new KeyValuePair<int, object>(0, "——"));
                    rowList25.Add(new KeyValuePair<int, object>(0, "（投标人提供）"));
                    rowList25.Add(new KeyValuePair<int, object>(1, "黑色"));  //高亮颜色
                    rowList25.Add(new KeyValuePair<int, object>(1, "黑色？留空？")); //高亮颜色
                    tableData.Add(rowList25);

                    var rowList26 = new List<KeyValuePair<int, object>>();  //行数据
                    rowList26.Add(new KeyValuePair<int, object>(0, "——"));
                    rowList26.Add(new KeyValuePair<int, object>(1, isArmoured ? "标称厚度t(有铠装)" : "标称厚度t(无铠装)"));//(0, "标称厚度t（无铠装）"));
                    rowList26.Add(new KeyValuePair<int, object>(0, "mm"));
                    rowList26.Add(new KeyValuePair<int, object>(0, "（投标人提供）"));
                    rowList26.Add(new KeyValuePair<int, object>(1, sheathThick)); //0.8)); // 待输入
                    rowList26.Add(new KeyValuePair<int, object>(0, "——")); //"Z - YJV"));// 待输入
                    tableData.Add(rowList26);

                    string armourString = isArmoured ? "铠装80%" : "无铠装85%";
                    var rowList27 = new List<KeyValuePair<int, object>>();  //行数据
                    rowList27.Add(new KeyValuePair<int, object>(0, "——"));
                    rowList27.Add(new KeyValuePair<int, object>(0, "最薄点厚度不小于"));
                    rowList27.Add(new KeyValuePair<int, object>(0, "mm"));
                    rowList27.Add(new KeyValuePair<int, object>(1, armourString));// 待输入armourWeight
                    rowList27.Add(new KeyValuePair<int, object>(1, armourString));// 待输入
                    rowList27.Add(new KeyValuePair<int, object>(0, "——"));
                    tableData.Add(rowList27);

                    var rowList28 = new List<KeyValuePair<int, object>>();  //行数据
                    rowList28.Add(new KeyValuePair<int, object>(0, " 电缆外径D："));
                    rowList28.Add(new KeyValuePair<int, object>(0, "——"));
                    rowList28.Add(new KeyValuePair<int, object>(0, "mm"));
                    rowList28.Add(new KeyValuePair<int, object>(0, "（投标人提供）"));
                    rowList28.Add(new KeyValuePair<int, object>(1, cableDiameter));// 19.74));// 待输入
                    rowList28.Add(new KeyValuePair<int, object>(0, "——")); //(1, type_spec)); // 待输入
                    tableData.Add(rowList28);

                    var rowList29 = new List<KeyValuePair<int, object>>();  //行数据
                    rowList29.Add(new KeyValuePair<int, object>(0, "20℃时铜导体最大直流电阻"));
                    rowList29.Add(new KeyValuePair<int, object>(0, "——"));
                    rowList29.Add(new KeyValuePair<int, object>(0, "Ω/km"));
                    rowList29.Add(new KeyValuePair<int, object>(1, resistant20_1));//1.83));// 待输入
                    rowList29.Add(new KeyValuePair<int, object>(1, resistant20_1));//1.83));// 待输入
                    rowList29.Add(new KeyValuePair<int, object>(1, areaConductor[0])); //"错误对应10mm²截面"));// 待输入
                    tableData.Add(rowList29);

                    if (isDoubleCable)
                    {
                        var rowList30 = new List<KeyValuePair<int, object>>();  //行数据
                        rowList30.Add(new KeyValuePair<int, object>(0, "——"));
                        rowList30.Add(new KeyValuePair<int, object>(0, "——"));
                        rowList30.Add(new KeyValuePair<int, object>(0, "——"));
                        rowList30.Add(new KeyValuePair<int, object>(1, resistant20_2));//3.08));// 待输入
                        rowList30.Add(new KeyValuePair<int, object>(1, resistant20_2));//3.08));// 待输入
                        rowList30.Add(new KeyValuePair<int, object>(1, areaConductor[1])); //"对应6mm²截面"));// 待输入
                        tableData.Add(rowList30);
                    }
                    var rowList31 = new List<KeyValuePair<int, object>>();  //行数据
                    rowList31.Add(new KeyValuePair<int, object>(0, "90℃时铜导体最大交流电阻"));
                    rowList31.Add(new KeyValuePair<int, object>(0, "——"));
                    rowList31.Add(new KeyValuePair<int, object>(0, "Ω/kμ"));
                    rowList31.Add(new KeyValuePair<int, object>(0, "（投标人提供）"));
                    rowList31.Add(new KeyValuePair<int, object>(1, resistant90_1));//2.3334));// 待输入
                    rowList31.Add(new KeyValuePair<int, object>(1, areaConductor[0])); //"对应10mm²截面"));// 待输入
                    tableData.Add(rowList31);


                    if (isDoubleCable)
                    {
                        var rowList32 = new List<KeyValuePair<int, object>>();  //行数据
                        rowList32.Add(new KeyValuePair<int, object>(0, "——"));
                        rowList32.Add(new KeyValuePair<int, object>(0, "——"));
                        rowList32.Add(new KeyValuePair<int, object>(0, "——"));
                        rowList32.Add(new KeyValuePair<int, object>(0, "——"));
                        rowList32.Add(new KeyValuePair<int, object>(1, resistant90_2));//3.9273));// 待输入
                        rowList32.Add(new KeyValuePair<int, object>(1, areaConductor[1])); //"对应6mm²截面"));// 待输入
                        tableData.Add(rowList32);
                    }

                    var rowList33 = new List<KeyValuePair<int, object>>();  //行数据
                    rowList33.Add(new KeyValuePair<int, object>(0, "电缆长期允许载流量\r\n（计算值，空气中40℃敷设）"));
                    rowList33.Add(new KeyValuePair<int, object>(0, "——"));
                    rowList33.Add(new KeyValuePair<int, object>(0, "A"));
                    rowList33.Add(new KeyValuePair<int, object>(0, "（投标人提供）"));
                    rowList33.Add(new KeyValuePair<int, object>(1, current40)); //272));// 待输入
                    rowList33.Add(new KeyValuePair<int, object>(1, areaConductor[0])); //type_spec)); // 待输入
                    tableData.Add(rowList33);

                    var rowList34 = new List<KeyValuePair<int, object>>();  //行数据
                    rowList34.Add(new KeyValuePair<int, object>(0, "出厂工频电压试验"));
                    rowList34.Add(new KeyValuePair<int, object>(0, "——"));
                    rowList34.Add(new KeyValuePair<int, object>(0, "kV/min"));
                    rowList34.Add(new KeyValuePair<int, object>(0, "3.5 U0/5"));
                    rowList34.Add(new KeyValuePair<int, object>(0, "3.5/5"));
                    rowList34.Add(new KeyValuePair<int, object>(0, "——"));
                    tableData.Add(rowList34);

                    var rowList35 = new List<KeyValuePair<int, object>>();  //行数据
                    rowList35.Add(new KeyValuePair<int, object>(0, "电缆盘尺寸"));
                    rowList35.Add(new KeyValuePair<int, object>(0, "——"));
                    rowList35.Add(new KeyValuePair<int, object>(0, "mm"));
                    rowList35.Add(new KeyValuePair<int, object>(0, "（投标人提供）"));
                    rowList35.Add(new KeyValuePair<int, object>(0, "根据订单长度选择"));
                    rowList35.Add(new KeyValuePair<int, object>(0, "——"));
                    tableData.Add(rowList35);

                    var rowList36 = new List<KeyValuePair<int, object>>();  //行数据
                    rowList36.Add(new KeyValuePair<int, object>(0, "电缆敷设时的最大牵引力"));
                    rowList36.Add(new KeyValuePair<int, object>(0, "——"));
                    rowList36.Add(new KeyValuePair<int, object>(0, "N/mm²"));
                    rowList36.Add(new KeyValuePair<int, object>(1, "70"));// 待输入
                    rowList36.Add(new KeyValuePair<int, object>(1, "70"));// 待输入
                    rowList36.Add(new KeyValuePair<int, object>(1, "铜芯，牵引头?"));// 待输入
                    tableData.Add(rowList36);

                    var rowList37 = new List<KeyValuePair<int, object>>();  //行数据
                    rowList37.Add(new KeyValuePair<int, object>(0, "电缆敷设时的最大侧压力"));
                    rowList37.Add(new KeyValuePair<int, object>(0, "——"));
                    rowList37.Add(new KeyValuePair<int, object>(0, "N/m"));
                    rowList37.Add(new KeyValuePair<int, object>(0, "5000"));
                    rowList37.Add(new KeyValuePair<int, object>(0, "5000"));
                    rowList37.Add(new KeyValuePair<int, object>(0, "——"));
                    tableData.Add(rowList37);

                    var rowList38 = new List<KeyValuePair<int, object>>();  //行数据
                    rowList38.Add(new KeyValuePair<int, object>(0, "电缆质量（近似值）"));
                    rowList38.Add(new KeyValuePair<int, object>(0, "——"));
                    rowList38.Add(new KeyValuePair<int, object>(0, "kg/m"));
                    rowList38.Add(new KeyValuePair<int, object>(0, "（投标人提供）"));
                    rowList38.Add(new KeyValuePair<int, object>(1, cableWeight));//(1, "2.7"));// 待输入  cableWeight
                    rowList38.Add(new KeyValuePair<int, object>(0, "——")); //type_spec)); // 待输入
                    tableData.Add(rowList38);

                    Trace.WriteLine($"cableWeight 等于： {cableWeight} ");

                    var rowList39 = new List<KeyValuePair<int, object>>();  //行数据
                    rowList39.Add(new KeyValuePair<int, object>(0, "电缆敷设时允许环境温度"));
                    rowList39.Add(new KeyValuePair<int, object>(0, "——"));
                    rowList39.Add(new KeyValuePair<int, object>(0, "℃"));
                    rowList39.Add(new KeyValuePair<int, object>(0, "-5～＋40"));
                    rowList39.Add(new KeyValuePair<int, object>(0, "-5～＋40"));
                    rowList39.Add(new KeyValuePair<int, object>(0, "——"));
                    tableData.Add(rowList39);

                    var rowList40 = new List<KeyValuePair<int, object>>();  //行数据
                    rowList40.Add(new KeyValuePair<int, object>(0, "电缆在正常使用条件下的寿命"));
                    rowList40.Add(new KeyValuePair<int, object>(0, "——"));
                    rowList40.Add(new KeyValuePair<int, object>(0, "年"));
                    rowList40.Add(new KeyValuePair<int, object>(0, "≥30"));
                    rowList40.Add(new KeyValuePair<int, object>(0, "≥30"));
                    rowList40.Add(new KeyValuePair<int, object>(0, "——"));
                    tableData.Add(rowList40);


                    var rowList41 = new List<KeyValuePair<int, object>>();  //行数据
                    rowList41.Add(new KeyValuePair<int, object>(0, "电缆阻燃级别"));
                    rowList41.Add(new KeyValuePair<int, object>(0, "——"));
                    rowList41.Add(new KeyValuePair<int, object>(0, "——"));
                    rowList41.Add(new KeyValuePair<int, object>(0, "按供货需求表"));
                    rowList41.Add(new KeyValuePair<int, object>(1, flameRedartant)); ////"ZC"));
                    rowList41.Add(new KeyValuePair<int, object>(0, "——"));
                    tableData.Add(rowList41);

                    var rowList42 = new List<KeyValuePair<int, object>>();  //行数据
                    rowList42.Add(new KeyValuePair<int, object>(0, "电缆的无卤性能"));
                    rowList42.Add(new KeyValuePair<int, object>(0, "——"));
                    rowList42.Add(new KeyValuePair<int, object>(0, "——"));
                    rowList42.Add(new KeyValuePair<int, object>(0, "按供货需求表"));
                    rowList42.Add(new KeyValuePair<int, object>(1, halogenFree));//"不适用"));// 待输入
                    rowList42.Add(new KeyValuePair<int, object>(0, "——"));// 待输入
                    tableData.Add(rowList42);

                    var rowList43 = new List<KeyValuePair<int, object>>();  //行数据
                    rowList43.Add(new KeyValuePair<int, object>(0, "电缆的低烟性能"));
                    rowList43.Add(new KeyValuePair<int, object>(0, "——"));
                    rowList43.Add(new KeyValuePair<int, object>(0, "——"));
                    rowList43.Add(new KeyValuePair<int, object>(0, "按供货需求表"));
                    rowList43.Add(new KeyValuePair<int, object>(1, smokeFree));// "不适用"));// 待输入
                    rowList43.Add(new KeyValuePair<int, object>(0, "——"));// 待输入
                    tableData.Add(rowList43);

                    return tableData;
                }

                //6.1.2  至tableNonElecticPara
                List<List<KeyValuePair<int, object>>> TableDataInputNonElectric()
                {
                    var tableData = new List<List<KeyValuePair<int, object>>>();   //table数据

                    var rowList1 = new List<KeyValuePair<int, object>>();  //行数据
                    rowList1.Add(new KeyValuePair<int, object>(0, "电缆非电气技术参数")); //0
                    rowList1.Add(new KeyValuePair<int, object>(0, "——")); //1
                    rowList1.Add(new KeyValuePair<int, object>(1, "——"));//2
                    rowList1.Add(new KeyValuePair<int, object>(0, "——"));//3
                    rowList1.Add(new KeyValuePair<int, object>(0, "——"));//4
                    rowList1.Add(new KeyValuePair<int, object>(0, "——"));//5
                    rowList1.Add(new KeyValuePair<int, object>(0, "——"));//6
                    rowList1.Add(new KeyValuePair<int, object>(0, "——"));//7
                    rowList1.Add(new KeyValuePair<int, object>(0, "——")); //8
                    tableData.Add(rowList1);

                    var rowList2 = new List<KeyValuePair<int, object>>();  //行数据
                    rowList2.Add(new KeyValuePair<int, object>(0, "项　　目")); //0
                    rowList2.Add(new KeyValuePair<int, object>(0, "——")); //1
                    rowList2.Add(new KeyValuePair<int, object>(1, "——"));//2
                    rowList2.Add(new KeyValuePair<int, object>(0, "单位"));//3
                    rowList2.Add(new KeyValuePair<int, object>(0, "标准参数值"));//4
                    rowList2.Add(new KeyValuePair<int, object>(0, "——"));//5
                    rowList2.Add(new KeyValuePair<int, object>(0, "投标人响应值"));//6
                    rowList2.Add(new KeyValuePair<int, object>(0, "——"));//7
                    rowList2.Add(new KeyValuePair<int, object>(0, "备注")); //8
                    tableData.Add(rowList2);

                    var rowList3 = new List<KeyValuePair<int, object>>();  //行数据
                    rowList3.Add(new KeyValuePair<int, object>(0, "绝缘\r\nXLPE")); //0
                    rowList3.Add(new KeyValuePair<int, object>(0, "老化前断裂伸长率不小于")); //1
                    rowList3.Add(new KeyValuePair<int, object>(0, "——")); //2
                    rowList3.Add(new KeyValuePair<int, object>(0, "%")); //3
                    rowList3.Add(new KeyValuePair<int, object>(0, "200")); //4
                    rowList3.Add(new KeyValuePair<int, object>(0, "——")); //5
                    rowList3.Add(new KeyValuePair<int, object>(0, "200")); //6
                    rowList3.Add(new KeyValuePair<int, object>(0, "——")); //7
                    rowList3.Add(new KeyValuePair<int, object>(0, "——")); //8
                    tableData.Add(rowList3);

                    var rowList4 = new List<KeyValuePair<int, object>>();  //行数据
                    rowList4.Add(new KeyValuePair<int, object>(0, "——")); //0
                    rowList4.Add(new KeyValuePair<int, object>(0, "老化后抗张强度变化率不超过")); //1
                    rowList4.Add(new KeyValuePair<int, object>(0, "——")); //2
                    rowList4.Add(new KeyValuePair<int, object>(0, "%")); //3
                    rowList4.Add(new KeyValuePair<int, object>(0, "±25")); //4
                    rowList4.Add(new KeyValuePair<int, object>(0, "——")); //5
                    rowList4.Add(new KeyValuePair<int, object>(0, "±25")); //6
                    rowList4.Add(new KeyValuePair<int, object>(0, "——")); //7
                    rowList4.Add(new KeyValuePair<int, object>(0, "——")); //8
                    tableData.Add(rowList4);

                    var rowList5 = new List<KeyValuePair<int, object>>();  //行数据
                    rowList5.Add(new KeyValuePair<int, object>(0, "——")); //0
                    rowList5.Add(new KeyValuePair<int, object>(0, "老化后断裂伸长率变化率不超过")); //1
                    rowList5.Add(new KeyValuePair<int, object>(0, "——")); //2
                    rowList5.Add(new KeyValuePair<int, object>(0, "%")); //3
                    rowList5.Add(new KeyValuePair<int, object>(0, "±25")); //4
                    rowList5.Add(new KeyValuePair<int, object>(0, "——")); //5
                    rowList5.Add(new KeyValuePair<int, object>(0, "±25")); //6
                    rowList5.Add(new KeyValuePair<int, object>(0, "——")); //7
                    rowList5.Add(new KeyValuePair<int, object>(0, "——")); //8
                    tableData.Add(rowList5);

                    var rowList6 = new List<KeyValuePair<int, object>>();  //行数据
                    rowList6.Add(new KeyValuePair<int, object>(0, "——")); //0
                    rowList6.Add(new KeyValuePair<int, object>(0, "电缆段老化后抗张强度变化率不超过")); //1
                    rowList6.Add(new KeyValuePair<int, object>(0, "——")); //2
                    rowList6.Add(new KeyValuePair<int, object>(0, "%")); //3
                    rowList6.Add(new KeyValuePair<int, object>(0, "±25")); //4
                    rowList6.Add(new KeyValuePair<int, object>(0, "——")); //5
                    rowList6.Add(new KeyValuePair<int, object>(0, "±25")); //6
                    rowList6.Add(new KeyValuePair<int, object>(0, "——")); //7
                    rowList6.Add(new KeyValuePair<int, object>(0, "——")); //8
                    tableData.Add(rowList6);

                    var rowList7 = new List<KeyValuePair<int, object>>();  //行数据
                    rowList7.Add(new KeyValuePair<int, object>(0, "——")); //0
                    rowList7.Add(new KeyValuePair<int, object>(0, "电缆段老化后断裂伸长率变化率不超过")); //1
                    rowList7.Add(new KeyValuePair<int, object>(0, "——")); //2
                    rowList7.Add(new KeyValuePair<int, object>(0, "%")); //3
                    rowList7.Add(new KeyValuePair<int, object>(0, "±25")); //4
                    rowList7.Add(new KeyValuePair<int, object>(0, "——")); //5
                    rowList7.Add(new KeyValuePair<int, object>(0, "±25")); //6
                    rowList7.Add(new KeyValuePair<int, object>(0, "——")); //7
                    rowList7.Add(new KeyValuePair<int, object>(0, "——")); //8
                    tableData.Add(rowList7);

                    var rowList8 = new List<KeyValuePair<int, object>>();  //行数据
                    rowList8.Add(new KeyValuePair<int, object>(0, "——")); //0
                    rowList8.Add(new KeyValuePair<int, object>(0, "绝缘收缩试验不大于")); //1
                    rowList8.Add(new KeyValuePair<int, object>(0, "——")); //2
                    rowList8.Add(new KeyValuePair<int, object>(0, "%")); //3
                    rowList8.Add(new KeyValuePair<int, object>(0, "4")); //4
                    rowList8.Add(new KeyValuePair<int, object>(0, "——")); //5
                    rowList8.Add(new KeyValuePair<int, object>(0, "4")); //6
                    rowList8.Add(new KeyValuePair<int, object>(0, "——")); //7
                    rowList8.Add(new KeyValuePair<int, object>(0, "——")); //8
                    tableData.Add(rowList8);

                    var rowList9 = new List<KeyValuePair<int, object>>();  //行数据
                    rowList9.Add(new KeyValuePair<int, object>(0, "绝缘")); //0
                    rowList9.Add(new KeyValuePair<int, object>(0, "热延伸")); //1
                    rowList9.Add(new KeyValuePair<int, object>(0, "负荷下伸长率不大于")); //2
                    rowList9.Add(new KeyValuePair<int, object>(0, "%")); //3
                    rowList9.Add(new KeyValuePair<int, object>(0, "125")); //4
                    rowList9.Add(new KeyValuePair<int, object>(0, "——")); //5
                    rowList9.Add(new KeyValuePair<int, object>(0, "125")); //6
                    rowList9.Add(new KeyValuePair<int, object>(0, "——")); //7
                    rowList9.Add(new KeyValuePair<int, object>(0, "——")); //8
                    tableData.Add(rowList9);

                    var rowList10 = new List<KeyValuePair<int, object>>();  //行数据
                    rowList10.Add(new KeyValuePair<int, object>(0, "——")); //0
                    rowList10.Add(new KeyValuePair<int, object>(0, "——")); //1
                    rowList10.Add(new KeyValuePair<int, object>(0, "冷却后永久伸长率不大于")); //2
                    rowList10.Add(new KeyValuePair<int, object>(0, "%")); //3
                    rowList10.Add(new KeyValuePair<int, object>(0, "10")); //4
                    rowList10.Add(new KeyValuePair<int, object>(0, "——")); //5
                    rowList10.Add(new KeyValuePair<int, object>(0, "10")); //6
                    rowList10.Add(new KeyValuePair<int, object>(0, "——")); //7
                    rowList10.Add(new KeyValuePair<int, object>(0, "——")); //8
                    tableData.Add(rowList10);

                    var rowList11 = new List<KeyValuePair<int, object>>();  //行数据
                    rowList11.Add(new KeyValuePair<int, object>(0, "外护套")); //0
                    rowList11.Add(new KeyValuePair<int, object>(0, "——")); //1
                    rowList11.Add(new KeyValuePair<int, object>(0, "——")); //2
                    rowList11.Add(new KeyValuePair<int, object>(0, "——")); //3
                    rowList11.Add(new KeyValuePair<int, object>(0, "PE")); //4
                    rowList11.Add(new KeyValuePair<int, object>(0, "PVC")); //5
                    rowList11.Add(new KeyValuePair<int, object>(0, "无卤低烟\r\n阻燃护套")); //6
                    rowList11.Add(new KeyValuePair<int, object>(0, "PVC")); //7
                    rowList11.Add(new KeyValuePair<int, object>(0, "——")); //8
                    tableData.Add(rowList11);

                    var rowList12 = new List<KeyValuePair<int, object>>();  //行数据
                    rowList12.Add(new KeyValuePair<int, object>(0, "——")); //0
                    rowList12.Add(new KeyValuePair<int, object>(0, "老化前抗张强度不小于")); //1
                    rowList12.Add(new KeyValuePair<int, object>(0, "——")); //2
                    rowList12.Add(new KeyValuePair<int, object>(0, "MPa")); //3
                    rowList12.Add(new KeyValuePair<int, object>(0, "10")); //4
                    rowList12.Add(new KeyValuePair<int, object>(0, "12.5")); //5
                    rowList12.Add(new KeyValuePair<int, object>(0, "9")); //6
                    rowList12.Add(new KeyValuePair<int, object>(0, "12.5")); //7
                    rowList12.Add(new KeyValuePair<int, object>(0, "——")); //8
                    tableData.Add(rowList12);

                    var rowList13 = new List<KeyValuePair<int, object>>();  //行数据
                    rowList13.Add(new KeyValuePair<int, object>(0, "——")); //0
                    rowList13.Add(new KeyValuePair<int, object>(0, "老化前断裂伸长率不小于")); //1
                    rowList13.Add(new KeyValuePair<int, object>(0, "——")); //2
                    rowList13.Add(new KeyValuePair<int, object>(0, "%")); //3
                    rowList13.Add(new KeyValuePair<int, object>(0, "300")); //4
                    rowList13.Add(new KeyValuePair<int, object>(0, "150")); //5
                    rowList13.Add(new KeyValuePair<int, object>(0, "125")); //6
                    rowList13.Add(new KeyValuePair<int, object>(0, "150")); //7
                    rowList13.Add(new KeyValuePair<int, object>(0, "——")); //8
                    tableData.Add(rowList13);

                    var rowList14 = new List<KeyValuePair<int, object>>();  //行数据
                    rowList14.Add(new KeyValuePair<int, object>(0, "——")); //0
                    rowList14.Add(new KeyValuePair<int, object>(0, "老化后抗张强度不小于")); //1
                    rowList14.Add(new KeyValuePair<int, object>(0, "——")); //2
                    rowList14.Add(new KeyValuePair<int, object>(0, "MPa")); //3
                    rowList14.Add(new KeyValuePair<int, object>(0, "——")); //4
                    rowList14.Add(new KeyValuePair<int, object>(0, "12.5")); //5
                    rowList14.Add(new KeyValuePair<int, object>(0, "9")); //6
                    rowList14.Add(new KeyValuePair<int, object>(0, "12.5")); //7
                    rowList14.Add(new KeyValuePair<int, object>(0, "——")); //8
                    tableData.Add(rowList14);

                    var rowList15 = new List<KeyValuePair<int, object>>();  //行数据
                    rowList15.Add(new KeyValuePair<int, object>(0, "——")); //0
                    rowList15.Add(new KeyValuePair<int, object>(0, "老化后断裂伸长率不小于")); //1
                    rowList15.Add(new KeyValuePair<int, object>(0, "——")); //2
                    rowList15.Add(new KeyValuePair<int, object>(0, "%")); //3
                    rowList15.Add(new KeyValuePair<int, object>(0, "300")); //4
                    rowList15.Add(new KeyValuePair<int, object>(0, "150")); //5
                    rowList15.Add(new KeyValuePair<int, object>(0, "100")); //6
                    rowList15.Add(new KeyValuePair<int, object>(0, "150")); //7
                    rowList15.Add(new KeyValuePair<int, object>(0, "——")); //8
                    tableData.Add(rowList15);

                    var rowList16 = new List<KeyValuePair<int, object>>();  //行数据
                    rowList16.Add(new KeyValuePair<int, object>(0, "——")); //0
                    rowList16.Add(new KeyValuePair<int, object>(0, "老化后抗张强度变化率不超过")); //1
                    rowList16.Add(new KeyValuePair<int, object>(0, "——")); //2
                    rowList16.Add(new KeyValuePair<int, object>(0, "%")); //3
                    rowList16.Add(new KeyValuePair<int, object>(0, "——")); //4
                    rowList16.Add(new KeyValuePair<int, object>(0, "±25")); //5
                    rowList16.Add(new KeyValuePair<int, object>(0, "±40")); //6
                    rowList16.Add(new KeyValuePair<int, object>(0, "±25")); //7
                    rowList16.Add(new KeyValuePair<int, object>(0, "——")); //8
                    tableData.Add(rowList16);

                    var rowList17 = new List<KeyValuePair<int, object>>();  //行数据
                    rowList17.Add(new KeyValuePair<int, object>(0, "——")); //0
                    rowList17.Add(new KeyValuePair<int, object>(0, "老化后断裂伸长率变化率不超过")); //1
                    rowList17.Add(new KeyValuePair<int, object>(0, "——")); //2
                    rowList17.Add(new KeyValuePair<int, object>(0, "%")); //3
                    rowList17.Add(new KeyValuePair<int, object>(0, "——")); //4
                    rowList17.Add(new KeyValuePair<int, object>(0, "±25")); //5
                    rowList17.Add(new KeyValuePair<int, object>(0, "±40")); //6
                    rowList17.Add(new KeyValuePair<int, object>(0, "±25")); //7
                    rowList17.Add(new KeyValuePair<int, object>(0, "——")); //8
                    tableData.Add(rowList17);

                    var rowList18 = new List<KeyValuePair<int, object>>();  //行数据
                    rowList18.Add(new KeyValuePair<int, object>(0, "——")); //0
                    rowList18.Add(new KeyValuePair<int, object>(0, "电缆段老化后抗张强度变化率不超过")); //1
                    rowList18.Add(new KeyValuePair<int, object>(0, "——")); //2
                    rowList18.Add(new KeyValuePair<int, object>(0, "%")); //3
                    rowList18.Add(new KeyValuePair<int, object>(0, "——")); //4
                    rowList18.Add(new KeyValuePair<int, object>(0, "±25")); //5
                    rowList18.Add(new KeyValuePair<int, object>(0, "——")); //6
                    rowList18.Add(new KeyValuePair<int, object>(0, "±25")); //7
                    rowList18.Add(new KeyValuePair<int, object>(0, "——")); //8
                    tableData.Add(rowList18);

                    var rowList19 = new List<KeyValuePair<int, object>>();  //行数据
                    rowList19.Add(new KeyValuePair<int, object>(0, "——")); //0
                    rowList19.Add(new KeyValuePair<int, object>(0, "电缆段老化后断裂伸长率变化率不超过")); //1
                    rowList19.Add(new KeyValuePair<int, object>(0, "——")); //2
                    rowList19.Add(new KeyValuePair<int, object>(0, "%")); //3
                    rowList19.Add(new KeyValuePair<int, object>(0, "——")); //4
                    rowList19.Add(new KeyValuePair<int, object>(0, "±25")); //5
                    rowList19.Add(new KeyValuePair<int, object>(0, "——")); //6
                    rowList19.Add(new KeyValuePair<int, object>(0, "±25")); //7
                    rowList19.Add(new KeyValuePair<int, object>(0, "——")); //8
                    tableData.Add(rowList19);

                    var rowList20 = new List<KeyValuePair<int, object>>();  //行数据
                    rowList20.Add(new KeyValuePair<int, object>(0, "——")); //0
                    rowList20.Add(new KeyValuePair<int, object>(0, "高温压力试验，压痕深度不大于")); //1
                    rowList20.Add(new KeyValuePair<int, object>(0, "——")); //2
                    rowList20.Add(new KeyValuePair<int, object>(0, "%")); //3
                    rowList20.Add(new KeyValuePair<int, object>(0, "50")); //4
                    rowList20.Add(new KeyValuePair<int, object>(0, "50")); //5
                    rowList20.Add(new KeyValuePair<int, object>(0, "50")); //6
                    rowList20.Add(new KeyValuePair<int, object>(0, "50")); //7
                    rowList20.Add(new KeyValuePair<int, object>(0, "——")); //8
                    tableData.Add(rowList20);

                    var rowList21 = new List<KeyValuePair<int, object>>();  //行数据
                    rowList21.Add(new KeyValuePair<int, object>(0, "——")); //0
                    rowList21.Add(new KeyValuePair<int, object>(0, "热冲击试验")); //1
                    rowList21.Add(new KeyValuePair<int, object>(0, "——")); //2
                    rowList21.Add(new KeyValuePair<int, object>(0, "——")); //3
                    rowList21.Add(new KeyValuePair<int, object>(0, "——")); //4
                    rowList21.Add(new KeyValuePair<int, object>(0, "不开裂")); //5
                    rowList21.Add(new KeyValuePair<int, object>(0, "——")); //6
                    rowList21.Add(new KeyValuePair<int, object>(0, "不开裂")); //7
                    rowList21.Add(new KeyValuePair<int, object>(0, "——")); //8
                    tableData.Add(rowList21);

                    var rowList22 = new List<KeyValuePair<int, object>>();  //行数据
                    rowList22.Add(new KeyValuePair<int, object>(0, "——")); //0
                    rowList22.Add(new KeyValuePair<int, object>(0, "低温冲击试验")); //1
                    rowList22.Add(new KeyValuePair<int, object>(0, "——")); //2
                    rowList22.Add(new KeyValuePair<int, object>(0, "——")); //3
                    rowList22.Add(new KeyValuePair<int, object>(0, "——")); //4
                    rowList22.Add(new KeyValuePair<int, object>(0, "不开裂")); //5
                    rowList22.Add(new KeyValuePair<int, object>(0, "不开裂")); //6
                    rowList22.Add(new KeyValuePair<int, object>(0, "不开裂")); //7
                    rowList22.Add(new KeyValuePair<int, object>(0, "——")); //8
                    tableData.Add(rowList22);

                    var rowList23 = new List<KeyValuePair<int, object>>();  //行数据
                    rowList23.Add(new KeyValuePair<int, object>(0, "——")); //0
                    rowList23.Add(new KeyValuePair<int, object>(0, "低温拉伸，断裂伸长率不小于")); //1
                    rowList23.Add(new KeyValuePair<int, object>(0, "——")); //2
                    rowList23.Add(new KeyValuePair<int, object>(0, "%")); //3
                    rowList23.Add(new KeyValuePair<int, object>(0, "——")); //4
                    rowList23.Add(new KeyValuePair<int, object>(0, "20")); //5
                    rowList23.Add(new KeyValuePair<int, object>(0, "20")); //6
                    rowList23.Add(new KeyValuePair<int, object>(0, "20")); //7
                    rowList23.Add(new KeyValuePair<int, object>(0, "——")); //8
                    tableData.Add(rowList23);

                    var rowList24 = new List<KeyValuePair<int, object>>();  //行数据
                    rowList24.Add(new KeyValuePair<int, object>(0, "——")); //0
                    rowList24.Add(new KeyValuePair<int, object>(0, "热失重，最大允许失重　")); //1
                    rowList24.Add(new KeyValuePair<int, object>(0, "——")); //2
                    rowList24.Add(new KeyValuePair<int, object>(0, "mg/cm2")); //3
                    rowList24.Add(new KeyValuePair<int, object>(0, "——")); //4
                    rowList24.Add(new KeyValuePair<int, object>(0, "1.5")); //5
                    rowList24.Add(new KeyValuePair<int, object>(0, "——")); //6
                    rowList24.Add(new KeyValuePair<int, object>(0, "1.5")); //7
                    rowList24.Add(new KeyValuePair<int, object>(0, "——")); //8
                    tableData.Add(rowList24);

                    var rowList25 = new List<KeyValuePair<int, object>>();  //行数据
                    rowList25.Add(new KeyValuePair<int, object>(0, "——")); //0
                    rowList25.Add(new KeyValuePair<int, object>(0, "酸气含量试验（GB/T17650）")); //1
                    rowList25.Add(new KeyValuePair<int, object>(0, "——")); //2
                    rowList25.Add(new KeyValuePair<int, object>(0, "%")); //3
                    rowList25.Add(new KeyValuePair<int, object>(0, "——")); //4
                    rowList25.Add(new KeyValuePair<int, object>(0, "——")); //5
                    rowList25.Add(new KeyValuePair<int, object>(0, "0.5")); //6
                    rowList25.Add(new KeyValuePair<int, object>(0, "不适用")); //7
                    rowList25.Add(new KeyValuePair<int, object>(0, "——")); //8
                    tableData.Add(rowList25);

                    var rowList26 = new List<KeyValuePair<int, object>>();  //行数据
                    rowList26.Add(new KeyValuePair<int, object>(0, "——")); //0
                    rowList26.Add(new KeyValuePair<int, object>(0, "最大值")); //1
                    rowList26.Add(new KeyValuePair<int, object>(0, "——")); //2
                    rowList26.Add(new KeyValuePair<int, object>(0, "——")); //3
                    rowList26.Add(new KeyValuePair<int, object>(0, "——")); //4
                    rowList26.Add(new KeyValuePair<int, object>(0, "——")); //5
                    rowList26.Add(new KeyValuePair<int, object>(0, "——")); //6
                    rowList26.Add(new KeyValuePair<int, object>(0, "——")); //7
                    rowList26.Add(new KeyValuePair<int, object>(0, "——")); //8
                    tableData.Add(rowList26);

                    var rowList27 = new List<KeyValuePair<int, object>>();  //行数据
                    rowList27.Add(new KeyValuePair<int, object>(0, "——")); //0
                    rowList27.Add(new KeyValuePair<int, object>(0, "氟含量试验（IEC60684）")); //1
                    rowList27.Add(new KeyValuePair<int, object>(0, "——")); //2
                    rowList27.Add(new KeyValuePair<int, object>(0, "%")); //3
                    rowList27.Add(new KeyValuePair<int, object>(0, "——")); //4
                    rowList27.Add(new KeyValuePair<int, object>(0, "——")); //5
                    rowList27.Add(new KeyValuePair<int, object>(0, "0.1")); //6
                    rowList27.Add(new KeyValuePair<int, object>(0, "不适用")); //7
                    rowList27.Add(new KeyValuePair<int, object>(0, "——")); //8
                    tableData.Add(rowList27);

                    var rowList28 = new List<KeyValuePair<int, object>>();  //行数据
                    rowList28.Add(new KeyValuePair<int, object>(0, "——")); //0
                    rowList28.Add(new KeyValuePair<int, object>(0, "最大值")); //1
                    rowList28.Add(new KeyValuePair<int, object>(0, "——")); //2
                    rowList28.Add(new KeyValuePair<int, object>(0, "——")); //3
                    rowList28.Add(new KeyValuePair<int, object>(0, "——")); //4
                    rowList28.Add(new KeyValuePair<int, object>(0, "——")); //5
                    rowList28.Add(new KeyValuePair<int, object>(0, "——")); //6
                    rowList28.Add(new KeyValuePair<int, object>(0, "——")); //7
                    rowList28.Add(new KeyValuePair<int, object>(0, "——")); //8
                    tableData.Add(rowList28);

                    var rowList29 = new List<KeyValuePair<int, object>>();  //行数据
                    rowList29.Add(new KeyValuePair<int, object>(0, "——")); //0
                    rowList29.Add(new KeyValuePair<int, object>(0, "PH值    最小值")); //1
                    rowList29.Add(new KeyValuePair<int, object>(0, "——")); //2
                    rowList29.Add(new KeyValuePair<int, object>(0, "——")); //3
                    rowList29.Add(new KeyValuePair<int, object>(0, "——")); //4
                    rowList29.Add(new KeyValuePair<int, object>(0, "——")); //5
                    rowList29.Add(new KeyValuePair<int, object>(0, "4.3")); //6
                    rowList29.Add(new KeyValuePair<int, object>(0, "不适用")); //7
                    rowList29.Add(new KeyValuePair<int, object>(0, "——")); //8
                    tableData.Add(rowList29);

                    var rowList30 = new List<KeyValuePair<int, object>>();  //行数据
                    rowList30.Add(new KeyValuePair<int, object>(0, "——")); //0
                    rowList30.Add(new KeyValuePair<int, object>(0, "电导率   最大值")); //1
                    rowList30.Add(new KeyValuePair<int, object>(0, "——")); //2
                    rowList30.Add(new KeyValuePair<int, object>(0, "µS/mm")); //3
                    rowList30.Add(new KeyValuePair<int, object>(0, "——")); //4
                    rowList30.Add(new KeyValuePair<int, object>(0, "——")); //5
                    rowList30.Add(new KeyValuePair<int, object>(0, "10")); //6
                    rowList30.Add(new KeyValuePair<int, object>(0, "不适用")); //7
                    rowList30.Add(new KeyValuePair<int, object>(0, "——")); //8
                    tableData.Add(rowList30);

                    var rowList31 = new List<KeyValuePair<int, object>>();  //行数据
                    rowList31.Add(new KeyValuePair<int, object>(0, "——")); //0
                    rowList31.Add(new KeyValuePair<int, object>(0, "吸水试验  最大增重")); //1
                    rowList31.Add(new KeyValuePair<int, object>(0, "——")); //2
                    rowList31.Add(new KeyValuePair<int, object>(0, "mg/cm2")); //3
                    rowList31.Add(new KeyValuePair<int, object>(0, "——")); //4
                    rowList31.Add(new KeyValuePair<int, object>(0, "——")); //5
                    rowList31.Add(new KeyValuePair<int, object>(0, "10")); //6
                    rowList31.Add(new KeyValuePair<int, object>(0, "不适用")); //7
                    rowList31.Add(new KeyValuePair<int, object>(0, "——")); //8
                    tableData.Add(rowList31);

                    var rowList32 = new List<KeyValuePair<int, object>>();  //行数据
                    rowList32.Add(new KeyValuePair<int, object>(0, "——")); //0
                    rowList32.Add(new KeyValuePair<int, object>(0, "炭黑含量")); //1
                    rowList32.Add(new KeyValuePair<int, object>(0, "——")); //2
                    rowList32.Add(new KeyValuePair<int, object>(0, "%")); //3
                    rowList32.Add(new KeyValuePair<int, object>(0, "2.0～3.0")); //4
                    rowList32.Add(new KeyValuePair<int, object>(0, "——")); //5
                    rowList32.Add(new KeyValuePair<int, object>(0, "——")); //6
                    rowList32.Add(new KeyValuePair<int, object>(0, "——")); //7
                    rowList32.Add(new KeyValuePair<int, object>(0, "——")); //8
                    tableData.Add(rowList32);

                    return tableData;
                }

                //6.1.3  至tableMaterial
                List<List<KeyValuePair<int, object>>> TableDataInputMaterial()
                {
                    var tableData = new List<List<KeyValuePair<int, object>>>();   //table数据

                    var rowList1 = new List<KeyValuePair<int, object>>();  //行数据
                    rowList1.Add(new KeyValuePair<int, object>(0, "序号")); //0
                    rowList1.Add(new KeyValuePair<int, object>(0, "名称")); //1
                    rowList1.Add(new KeyValuePair<int, object>(0, "型式规格，参数"));//2
                    rowList1.Add(new KeyValuePair<int, object>(0, "数量"));//3
                    rowList1.Add(new KeyValuePair<int, object>(0, "制造商"));//4
                    rowList1.Add(new KeyValuePair<int, object>(0, "原产地"));//5
                    tableData.Add(rowList1);

                    var rowList2 = new List<KeyValuePair<int, object>>();  //行数据
                    rowList2.Add(new KeyValuePair<int, object>(0, "1")); //0
                    rowList2.Add(new KeyValuePair<int, object>(0, "电缆导体")); //1
                    rowList2.Add(new KeyValuePair<int, object>(0, " "));//2
                    rowList2.Add(new KeyValuePair<int, object>(0, "——"));//3
                    rowList2.Add(new KeyValuePair<int, object>(0, " "));//4
                    rowList2.Add(new KeyValuePair<int, object>(0, " "));//5
                    tableData.Add(rowList2);

                    var rowList3 = new List<KeyValuePair<int, object>>();  //行数据
                    rowList3.Add(new KeyValuePair<int, object>(0, " ")); //0
                    rowList3.Add(new KeyValuePair<int, object>(0, " ")); //1
                    rowList3.Add(new KeyValuePair<int, object>(0, " "));//2
                    rowList3.Add(new KeyValuePair<int, object>(0, "——"));//3
                    rowList3.Add(new KeyValuePair<int, object>(0, " "));//4
                    rowList3.Add(new KeyValuePair<int, object>(0, " "));//5
                    tableData.Add(rowList3);

                    var rowList4 = new List<KeyValuePair<int, object>>();  //行数据
                    rowList4.Add(new KeyValuePair<int, object>(0, "2")); //0
                    rowList4.Add(new KeyValuePair<int, object>(0, "交联聚\r\n乙烯绝缘")); //1
                    rowList4.Add(new KeyValuePair<int, object>(0, " "));//2
                    rowList4.Add(new KeyValuePair<int, object>(0, "——"));//3
                    rowList4.Add(new KeyValuePair<int, object>(0, " "));//4
                    rowList4.Add(new KeyValuePair<int, object>(0, " "));//5
                    tableData.Add(rowList4);


                    var rowList5 = new List<KeyValuePair<int, object>>();  //行数据
                    rowList5.Add(new KeyValuePair<int, object>(0, "3")); //0
                    rowList5.Add(new KeyValuePair<int, object>(0, "填充层")); //1
                    rowList5.Add(new KeyValuePair<int, object>(0, " "));//2
                    rowList5.Add(new KeyValuePair<int, object>(0, "——"));//3
                    rowList5.Add(new KeyValuePair<int, object>(0, " "));//4
                    rowList5.Add(new KeyValuePair<int, object>(0, " "));//5
                    tableData.Add(rowList5);

                    var rowList6 = new List<KeyValuePair<int, object>>();  //行数据
                    rowList6.Add(new KeyValuePair<int, object>(0, "4")); //0
                    rowList6.Add(new KeyValuePair<int, object>(0, "内衬层")); //1
                    rowList6.Add(new KeyValuePair<int, object>(0, " "));//2
                    rowList6.Add(new KeyValuePair<int, object>(0, "——"));//3
                    rowList6.Add(new KeyValuePair<int, object>(0, " "));//4
                    rowList6.Add(new KeyValuePair<int, object>(0, " "));//5
                    tableData.Add(rowList6);

                    var rowList7 = new List<KeyValuePair<int, object>>();  //行数据
                    rowList7.Add(new KeyValuePair<int, object>(0, "6")); //0
                    rowList7.Add(new KeyValuePair<int, object>(0, "铠装层")); //1
                    rowList7.Add(new KeyValuePair<int, object>(0, " "));//2
                    rowList7.Add(new KeyValuePair<int, object>(0, "——"));//3
                    rowList7.Add(new KeyValuePair<int, object>(0, " "));//4
                    rowList7.Add(new KeyValuePair<int, object>(0, " "));//5
                    tableData.Add(rowList7);

                    var rowList8 = new List<KeyValuePair<int, object>>();  //行数据
                    rowList8.Add(new KeyValuePair<int, object>(0, "6")); //0
                    rowList8.Add(new KeyValuePair<int, object>(0, "外护套")); //1
                    rowList8.Add(new KeyValuePair<int, object>(0, " "));//2
                    rowList8.Add(new KeyValuePair<int, object>(0, "——"));//3
                    rowList8.Add(new KeyValuePair<int, object>(0, " "));//4
                    rowList8.Add(new KeyValuePair<int, object>(0, " "));//5
                    tableData.Add(rowList8);

                    var rowList9 = new List<KeyValuePair<int, object>>();  //行数据
                    rowList9.Add(new KeyValuePair<int, object>(0, "说明")); //0
                    rowList9.Add(new KeyValuePair<int, object>(0, "  1、主要货物的名称、规格型式、参数、单位、数量要求请保持与商务部分货物清单一致。")); //1
                    rowList9.Add(new KeyValuePair<int, object>(1, " "));//2
                    rowList9.Add(new KeyValuePair<int, object>(0, " "));//3
                    rowList9.Add(new KeyValuePair<int, object>(0, " "));//4
                    rowList9.Add(new KeyValuePair<int, object>(0, " "));//5
                    tableData.Add(rowList9);

                    var rowList10 = new List<KeyValuePair<int, object>>();  //行数据
                    rowList10.Add(new KeyValuePair<int, object>(0, " ")); //0
                    rowList10.Add(new KeyValuePair<int, object>(0, "  2、此表中的主要货物的招标要求请与技术专用其他部分保持一致，不得前后矛盾。")); //1
                    rowList10.Add(new KeyValuePair<int, object>(0, ""));//2
                    rowList10.Add(new KeyValuePair<int, object>(0, ""));//3
                    rowList10.Add(new KeyValuePair<int, object>(0, ""));//4
                    rowList10.Add(new KeyValuePair<int, object>(0, ""));//5
                    tableData.Add(rowList10);

                    var rowList11 = new List<KeyValuePair<int, object>>();  //行数据
                    rowList11.Add(new KeyValuePair<int, object>(0, " ")); //0
                    rowList11.Add(new KeyValuePair<int, object>(0, "  3、请勿修改此表格式，画“—”处不需要填写。")); //1
                    rowList11.Add(new KeyValuePair<int, object>(0, ""));//2
                    rowList11.Add(new KeyValuePair<int, object>(0, ""));//3
                    rowList11.Add(new KeyValuePair<int, object>(0, ""));//4
                    rowList11.Add(new KeyValuePair<int, object>(0, ""));//5
                    tableData.Add(rowList11);



                    return tableData;
                }

                //6.1.3  至tableMEnvironment
                List<List<KeyValuePair<int, object>>> TableDataInputEnvironment()
                {
                    var tableData = new List<List<KeyValuePair<int, object>>>();   //table数据

                    var rowList1 = new List<KeyValuePair<int, object>>();  //行数据
                    rowList1.Add(new KeyValuePair<int, object>(0, "名称")); //0
                    rowList1.Add(new KeyValuePair<int, object>(0, "")); //1
                    rowList1.Add(new KeyValuePair<int, object>(0, "参数值"));//2
                    rowList1.Add(new KeyValuePair<int, object>(0, "投标人响应值"));//3
                    tableData.Add(rowList1);

                    var rowList2 = new List<KeyValuePair<int, object>>();  //行数据
                    rowList2.Add(new KeyValuePair<int, object>(0, "  海拔高度（m）")); //0
                    rowList2.Add(new KeyValuePair<int, object>(0, "")); //1
                    rowList2.Add(new KeyValuePair<int, object>(0, "≤4000"));//2
                    rowList2.Add(new KeyValuePair<int, object>(0, "≤4000"));//3
                    tableData.Add(rowList2);


                    var rowList3 = new List<KeyValuePair<int, object>>();  //行数据
                    rowList3.Add(new KeyValuePair<int, object>(0, "  最高环境温度（℃）")); //0
                    rowList3.Add(new KeyValuePair<int, object>(0, "")); //1
                    rowList3.Add(new KeyValuePair<int, object>(0, "45"));//2
                    rowList3.Add(new KeyValuePair<int, object>(0, "45"));//3
                    tableData.Add(rowList3);


                    var rowList4 = new List<KeyValuePair<int, object>>();  //行数据
                    rowList4.Add(new KeyValuePair<int, object>(0, "  最低环境温度（℃）")); //0
                    rowList4.Add(new KeyValuePair<int, object>(0, "")); //1
                    rowList4.Add(new KeyValuePair<int, object>(0, "-40"));//2
                    rowList4.Add(new KeyValuePair<int, object>(0, "-40"));//3
                    tableData.Add(rowList4);

                    var rowList5 = new List<KeyValuePair<int, object>>();  //行数据
                    rowList5.Add(new KeyValuePair<int, object>(0, "  土壤最高环境温度（℃）")); //0
                    rowList5.Add(new KeyValuePair<int, object>(0, "")); //1
                    rowList5.Add(new KeyValuePair<int, object>(0, "35"));//2
                    rowList5.Add(new KeyValuePair<int, object>(0, "35"));//3
                    tableData.Add(rowList5);

                    var rowList6 = new List<KeyValuePair<int, object>>();  //行数据
                    rowList6.Add(new KeyValuePair<int, object>(0, "  土壤最低环境温度（℃）")); //0
                    rowList6.Add(new KeyValuePair<int, object>(0, "")); //1
                    rowList6.Add(new KeyValuePair<int, object>(0, "-20"));//2
                    rowList6.Add(new KeyValuePair<int, object>(0, "-20"));//3
                    tableData.Add(rowList6);

                    var rowList7 = new List<KeyValuePair<int, object>>();  //行数据
                    rowList7.Add(new KeyValuePair<int, object>(0, "  日照强度（W / cm2）")); //0
                    rowList7.Add(new KeyValuePair<int, object>(0, "")); //1
                    rowList7.Add(new KeyValuePair<int, object>(0, "0.1"));//2
                    rowList7.Add(new KeyValuePair<int, object>(0, "0.1"));//3
                    tableData.Add(rowList7);

                    var rowList8 = new List<KeyValuePair<int, object>>();  //行数据
                    rowList8.Add(new KeyValuePair<int, object>(0, "  湿度")); //0
                    rowList8.Add(new KeyValuePair<int, object>(0, "日相对湿度平均值（％）")); //1
                    rowList8.Add(new KeyValuePair<int, object>(0, "≤95"));//2
                    rowList8.Add(new KeyValuePair<int, object>(0, "≤95"));//3
                    tableData.Add(rowList8);

                    var rowList9 = new List<KeyValuePair<int, object>>();  //行数据
                    rowList9.Add(new KeyValuePair<int, object>(0, " ")); //0
                    rowList9.Add(new KeyValuePair<int, object>(0, "月相对湿度平均值（％）")); //1
                    rowList9.Add(new KeyValuePair<int, object>(0, "≤90"));//2
                    rowList9.Add(new KeyValuePair<int, object>(0, "≤90"));//3
                    tableData.Add(rowList9);

                    var rowList10 = new List<KeyValuePair<int, object>>();  //行数据
                    rowList10.Add(new KeyValuePair<int, object>(0, "  最大风速（户外）（m / s）/ Pa")); //0
                    rowList10.Add(new KeyValuePair<int, object>(0, "")); //1
                    rowList10.Add(new KeyValuePair<int, object>(0, "35 / 700"));//2
                    rowList10.Add(new KeyValuePair<int, object>(0, "35 / 700"));//3
                    tableData.Add(rowList10);

                    var rowList11 = new List<KeyValuePair<int, object>>();  //行数据
                    rowList11.Add(new KeyValuePair<int, object>(0, "  电缆敷设方式\r\n  (多种方式并存时，选择载流量最小的一种方式)")); //0
                    rowList11.Add(new KeyValuePair<int, object>(0, "")); //1
                    rowList11.Add(new KeyValuePair<int, object>(0, "埋、排管、电缆沟、\r\n隧道、空气"));//2
                    rowList11.Add(new KeyValuePair<int, object>(0, "埋、排管、电缆沟、\r\n隧道、空气"));//3
                    tableData.Add(rowList11);

                    return tableData;
                }

                //6.2 创建table for word
                XWPFTable CreateTable(List<List<KeyValuePair<int, object>>> inputDataList, XWPFDocument theWord, string tableMark = "Normal")
                {
                    // 创建表格
                    var theTable = theWord.CreateTable(inputDataList.Count, inputDataList[0].Count); //（行数，列数） 行数+1用于表头
                                                                                                     //  Trace.WriteLine("读取Word 结构参数 表格");
                    try
                    {
                        // 添加数据行
                        for (int rowIndex = 0; rowIndex < inputDataList.Count; rowIndex++)
                        {
                            //var row = theTable.GetRow(rowIndex + 1);
                            var row = theTable.GetRow(rowIndex);
                            //row.Height = 800;//高度.默认为自动高度
                            int colIndex = 0;
                            Trace.WriteLine("行循环");
                            foreach (var keyValuePair in inputDataList[rowIndex])
                            {
                                Trace.WriteLine("列循环");
                                int styleInt = keyValuePair.Key;//方法1
                                                                // object value = inputDataList[colIdx];
                                object value = keyValuePair.Value;

                                //  row.GetCell(colIndex).SetText(value);
                                var cell = row.GetCell(colIndex);
                                cell.SetVerticalAlignment(XWPFTableCell.XWPFVertAlign.CENTER); // //单元格内文字，垂直位置，默认靠上
                                cell.RemoveParagraph(0);//去掉段落,否则单元格前面会多一个空行
                                                        // var cellPara = cell.AddParagraph();
                                                        //cellPara.Alignment = ParagraphAlignment.CENTER;//水平靠中
                                                        //cellPara.SpacingBefore = 0; // 设置段落间距：段前为0
                                                        // cellPara.SpacingAfter = 0; // 设置段落间距：段后为0
                                                        // cellPara.SpacingBetween=0; // 设置行距单倍
                                                        // cellPara.SetSpacingBetween (0); // 设置行距1倍
                                                        // var run = cellPara.CreateRun();                                

                                if (Convert.ToString(value).Contains("\r\n"))
                                {
                                    string valueString = Convert.ToString(value);
                                    string value1 = valueString.Split("\r\n")[0];
                                    string value2 = valueString.Split("\r\n")[1];

                                    dataInput(value1);
                                    dataInput(value2);

                                }
                                else dataInput(value);

                                void dataInput(object value)
                                {
                                    var aParagraph = cell.AddParagraph();
                                    //Trace.WriteLine("run this ");
                                    switch (tableMark)
                                    {
                                        case "Normal":
                                            aParagraph.Alignment = ParagraphAlignment.CENTER; //单元格内文字，水平位置，默认靠左
                                            break;
                                        case "Material":
                                            if (rowIndex < 8 || colIndex == 0) aParagraph.Alignment = ParagraphAlignment.CENTER;//单元格内文字，水平位置，默认靠左
                                                                                                                                // else aParagraph.Alignment = ParagraphAlignment.LEFT;
                                            break;
                                        case "Environment":
                                            if (rowIndex == 0 || colIndex > 0) aParagraph.Alignment = ParagraphAlignment.CENTER;//单元格内文字，水平位置，默认靠左
                                                                                                                                // else aParagraph.Alignment = ParagraphAlignment.LEFT;
                                            break;
                                    }
                                    var run = aParagraph.CreateRun();
                                    run.FontSize = 9;
                                    run.FontFamily = "宋体";
                                    //run.SetFontFamily("宋体", FontCharRange.CS); //这代码只改变汉字字体，英语自动

                                    // 根据数据类型应用样式
                                    if (value == DBNull.Value || value == "")
                                    {
                                        //cell.SetText("数据未提供");
                                        run.SetText("数据未提供");//要设置文字颜色，文字输入也须用run设置。
                                        run.SetColor(ColorConverterWord.ToHexColor("Red")); //文字颜色
                                                                                            // cell.SetColor(ColorConverterWord.ToHexColor("Red"));//填充色
                                        Trace.WriteLine("数据未提供");

                                    }
                                    else
                                    {
                                        //cell.SetText(Convert.ToString(value));//这个只能设置文字，无法设置文字颜色      
                                        run.SetText(Convert.ToString(value));//用Run可以改变文字颜色，字体，大小 
                                        Trace.WriteLine($"Value: {Convert.ToString(value)}");
                                        if (faultMarkArray.Any(term => value.ToString().Contains(term, StringComparison.OrdinalIgnoreCase)))
                                        {
                                            run.SetColor(ColorConverterWord.ToHexColor("Red")); //文字颜色
                                                                                                //cell.SetColor(ColorConverterWord.ToHexColor("Yellow")); //填充色
                                        }
                                        else if (styleInt == 1)
                                        {
                                            run.SetColor(ColorConverterWord.ToHexColor("Blue")); //文字颜色
                                                                                                 //cell.SetColor(ColorConverterWord.ToHexColor("Yellow")); //填充色
                                        }

                                        #region 类型判断
                                        //if (rowIndex == 0) cell.CellStyle = titleStyle;
                                        //else if (faultMarkArray.Any(term => value.ToString().Contains(term, StringComparison.OrdinalIgnoreCase)))
                                        //    cell.CellStyle = warnStyle;
                                        //else if (styleInt == 1)//方法2
                                        //    cell.CellStyle = stringBlueStyle;
                                        //else if (value.ToString().Contains("电缆长期允许载流量", StringComparison.OrdinalIgnoreCase))
                                        //{
                                        //    //合并单元格，文字超出单元格范围，行高不会自动变化
                                        //    iRow.Height = 576;// 14.4*2 *20;1/20个点为最小单位
                                        //    cell.CellStyle = stringCenterStyle;
                                        //}
                                        //else
                                        //    cell.CellStyle = stringCenterStyle;
                                        #endregion
                                    }
                                }

                                colIndex++;
                            }
                        }
                    }
                    catch
                    {
                        Trace.WriteLine("表格无数据");
                    }

                    return theTable;
                }

                //6,3 合并单元格
                //6,3 .1 合并单元格 电缆结构技术参数
                void MergeCellConstruction()
                {
                    if (isDoubleCable)
                    {
                        if (isArmoured)
                        {
                            MergeCells(tableConstructionPara, 0, 0, 0, 5);
                            MergeCells(tableConstructionPara, 1, 1, 0, 1);
                            MergeCells(tableConstructionPara, 1, 1, 1, 3);
                            MergeCells(tableConstructionPara, 2, 2, 0, 1);
                            MergeCells(tableConstructionPara, 3, 10, 0, 0);
                            MergeCells(tableConstructionPara, 6, 7, 1, 1);
                            MergeCells(tableConstructionPara, 6, 7, 2, 2);
                            MergeCells(tableConstructionPara, 8, 9, 1, 1);
                            MergeCells(tableConstructionPara, 8, 9, 2, 2);
                            MergeCells(tableConstructionPara, 8, 9, 3, 3);
                            MergeCells(tableConstructionPara, 11, 15, 0, 0);
                            MergeCells(tableConstructionPara, 12, 13, 1, 1);
                            MergeCells(tableConstructionPara, 12, 13, 2, 2);
                            MergeCells(tableConstructionPara, 12, 13, 3, 3);
                            MergeCells(tableConstructionPara, 17, 18, 0, 0);
                            MergeCells(tableConstructionPara, 19, 22, 0, 0);

                            MergeCells(tableConstructionPara, 23, 26, 0, 0);
                            MergeCells(tableConstructionPara, 27, 27, 0, 1);
                            MergeCells(tableConstructionPara, 28, 29, 0, 1);
                            MergeCells(tableConstructionPara, 28, 29, 1, 1);
                            MergeCells(tableConstructionPara, 30, 31, 0, 1);
                            MergeCells(tableConstructionPara, 30, 31, 1, 1);
                            MergeCells(tableConstructionPara, 30, 31, 2, 2);
                            MergeCells(tableConstructionPara, 32, 32, 0, 1);

                            MergeCells(tableConstructionPara, 33, 33, 0, 1);
                            MergeCells(tableConstructionPara, 34, 34, 0, 1);
                            MergeCells(tableConstructionPara, 35, 35, 0, 1);
                            MergeCells(tableConstructionPara, 36, 36, 0, 1);
                            MergeCells(tableConstructionPara, 37, 37, 0, 1);
                            MergeCells(tableConstructionPara, 38, 38, 0, 1);
                            MergeCells(tableConstructionPara, 39, 39, 0, 1);
                            MergeCells(tableConstructionPara, 40, 40, 0, 1);
                            MergeCells(tableConstructionPara, 41, 41, 0, 1);
                            MergeCells(tableConstructionPara, 42, 42, 0, 1);
                        }
                        else
                        {
                            MergeCells(tableConstructionPara, 0, 0, 0, 5);
                            MergeCells(tableConstructionPara, 1, 1, 0, 1);
                            MergeCells(tableConstructionPara, 1, 1, 1, 3);
                            MergeCells(tableConstructionPara, 2, 2, 0, 1);
                            MergeCells(tableConstructionPara, 3, 10, 0, 0);
                            MergeCells(tableConstructionPara, 6, 7, 1, 1);
                            MergeCells(tableConstructionPara, 6, 7, 2, 2);
                            MergeCells(tableConstructionPara, 8, 9, 1, 1);
                            MergeCells(tableConstructionPara, 8, 9, 2, 2);
                            MergeCells(tableConstructionPara, 8, 9, 3, 3);
                            MergeCells(tableConstructionPara, 11, 15, 0, 0);
                            MergeCells(tableConstructionPara, 12, 13, 1, 1);
                            MergeCells(tableConstructionPara, 12, 13, 2, 2);
                            MergeCells(tableConstructionPara, 12, 13, 3, 3);
                            MergeCells(tableConstructionPara, 17, 20, 0, 0);

                            // MergeCells(theTable, 18, 21, 0, 0);
                            //MergeCells(theTable, 22, 25, 0, 0);
                            MergeCells(tableConstructionPara, 21, 21, 0, 1);
                            MergeCells(tableConstructionPara, 22, 23, 0, 1);
                            MergeCells(tableConstructionPara, 22, 23, 1, 1);
                            // MergeCells(theTable, 23, 28, 1, 1);
                            MergeCells(tableConstructionPara, 24, 25, 0, 1);
                            MergeCells(tableConstructionPara, 24, 25, 1, 1);
                            MergeCells(tableConstructionPara, 24, 25, 2, 2);
                            // MergeCells(theTable, 29, 30, 1, 1);
                            //MergeCells(theTable, 29, 30, 2, 2);

                            MergeCells(tableConstructionPara, 26, 26, 0, 1);
                            MergeCells(tableConstructionPara, 27, 27, 0, 1);
                            MergeCells(tableConstructionPara, 28, 28, 0, 1);
                            MergeCells(tableConstructionPara, 29, 29, 0, 1);
                            MergeCells(tableConstructionPara, 30, 30, 0, 1);
                            MergeCells(tableConstructionPara, 31, 31, 0, 1);
                            MergeCells(tableConstructionPara, 32, 32, 0, 1);
                            MergeCells(tableConstructionPara, 33, 33, 0, 1);
                            MergeCells(tableConstructionPara, 34, 34, 0, 1);
                            MergeCells(tableConstructionPara, 35, 35, 0, 1);
                            MergeCells(tableConstructionPara, 36, 36, 0, 1);


                        }
                    }
                    else if (isMultiCore)
                    {

                        if (isArmoured)
                        {
                            MergeCells(tableConstructionPara, 0, 0, 0, 5);
                            MergeCells(tableConstructionPara, 1, 1, 0, 1);
                            MergeCells(tableConstructionPara, 1, 1, 1, 3);
                            MergeCells(tableConstructionPara, 2, 2, 0, 1);
                            MergeCells(tableConstructionPara, 3, 8, 0, 0);

                            //MergeCells(theTable, 5, 6, 1, 1);
                            //MergeCells(theTable, 5, 6, 2, 2);
                            //MergeCells(theTable, 7, 8, 1, 1);
                            //MergeCells(theTable, 7, 8, 2, 2);
                            //MergeCells(theTable, 7, 8, 3, 3);
                            MergeCells(tableConstructionPara, 9, 12, 0, 0);
                            //MergeCells(theTable, 11, 12, 1, 1);
                            //MergeCells(theTable, 11, 12, 2, 2);
                            //MergeCells(theTable, 11, 12, 3, 3);
                            MergeCells(tableConstructionPara, 14, 15, 0, 0);
                            MergeCells(tableConstructionPara, 16, 19, 0, 0);
                            MergeCells(tableConstructionPara, 20, 23, 0, 0);
                            // MergeCells(tableConstructionPara, 24, 24, 0, 1);
                            MergeCells(tableConstructionPara, 24, 24, 0, 1);
                            MergeCells(tableConstructionPara, 25, 25, 0, 1);
                            MergeCells(tableConstructionPara, 26, 26, 0, 1);
                            MergeCells(tableConstructionPara, 27, 27, 0, 1);
                            MergeCells(tableConstructionPara, 28, 28, 0, 1);
                            MergeCells(tableConstructionPara, 29, 29, 0, 1);
                            MergeCells(tableConstructionPara, 30, 30, 0, 1);
                            MergeCells(tableConstructionPara, 31, 31, 0, 1);
                            MergeCells(tableConstructionPara, 32, 32, 0, 1);
                            MergeCells(tableConstructionPara, 33, 33, 0, 1);
                            MergeCells(tableConstructionPara, 34, 34, 0, 1);
                            MergeCells(tableConstructionPara, 35, 35, 0, 1);
                            MergeCells(tableConstructionPara, 36, 36, 0, 1);
                            MergeCells(tableConstructionPara, 37, 37, 0, 1);


                            //MergeCells(theTable, 38, 38, 0, 1);
                            //MergeCells(theTable, 39, 39, 0, 1);
                        }
                        else
                        {
                            Trace.WriteLine("merge: singe, not armoured");
                            MergeCells(tableConstructionPara, 0, 0, 0, 5);
                            MergeCells(tableConstructionPara, 1, 1, 0, 1);
                            MergeCells(tableConstructionPara, 1, 1, 1, 3);
                            MergeCells(tableConstructionPara, 2, 2, 0, 1);
                            MergeCells(tableConstructionPara, 3, 8, 0, 0);
                            //MergeCells(theTable, 5, 6, 1, 1);
                            //MergeCells(theTable, 5, 6, 2, 2);
                            //MergeCells(theTable, 7, 8, 1, 1);
                            //MergeCells(theTable, 7, 8, 2, 2);
                            //MergeCells(theTable, 7, 8, 3, 3);
                            MergeCells(tableConstructionPara, 9, 12, 0, 0);
                            //MergeCells(theTable, 11, 12, 1, 1);
                            //MergeCells(theTable, 11, 12, 2, 2);
                            //MergeCells(theTable, 11, 12, 3, 3);
                            MergeCells(tableConstructionPara, 14, 17, 0, 0);
                            Trace.WriteLine("merge: singe, not armoured here");
                            //MergeCells(theTable, 18, 21, 0, 0);
                            //MergeCells(theTable, 22, 25, 0, 0);

                            MergeCells(tableConstructionPara, 18, 18, 0, 1);
                            MergeCells(tableConstructionPara, 19, 19, 0, 1);
                            MergeCells(tableConstructionPara, 20, 20, 0, 1);
                            MergeCells(tableConstructionPara, 21, 21, 0, 1);
                            MergeCells(tableConstructionPara, 22, 22, 0, 1);
                            MergeCells(tableConstructionPara, 23, 23, 0, 1);
                            MergeCells(tableConstructionPara, 24, 24, 0, 1);
                            MergeCells(tableConstructionPara, 25, 25, 0, 1);
                            MergeCells(tableConstructionPara, 26, 26, 0, 1);
                            MergeCells(tableConstructionPara, 27, 27, 0, 1);
                            MergeCells(tableConstructionPara, 28, 28, 0, 1);
                            MergeCells(tableConstructionPara, 29, 29, 0, 1);
                            MergeCells(tableConstructionPara, 30, 30, 0, 1);
                            MergeCells(tableConstructionPara, 31, 31, 0, 1);

                            //MergeCells(theTable, 32, 32, 0, 1);
                            //MergeCells(theTable, 33, 33, 0, 1);
                            //MergeCells(theTable, 34, 34, 0, 1);
                            //MergeCells(theTable, 35, 35, 0, 1);
                            //MergeCells(theTable, 36, 36, 0, 1);
                            //MergeCells(theTable, 37, 37, 0, 1);
                            //MergeCells(theTable, 38, 38, 0, 1);
                            //MergeCells(theTable, 39, 39, 0, 1);
                            //MergeCells(theTable, 40, 40, 0, 1);
                            //MergeCells(theTable, 41, 41, 0, 1);

                        }

                    }
                    else
                    {
                        if (isArmoured)
                        {
                            MergeCells(tableConstructionPara, 0, 0, 0, 5);
                            MergeCells(tableConstructionPara, 1, 1, 0, 1);
                            MergeCells(tableConstructionPara, 1, 1, 1, 3);
                            MergeCells(tableConstructionPara, 2, 2, 0, 1);
                            MergeCells(tableConstructionPara, 3, 8, 0, 0);

                            //MergeCells(theTable, 5, 6, 1, 1);
                            //MergeCells(theTable, 5, 6, 2, 2);
                            //MergeCells(theTable, 7, 8, 1, 1);
                            //MergeCells(theTable, 7, 8, 2, 2);
                            //MergeCells(theTable, 7, 8, 3, 3);
                            MergeCells(tableConstructionPara, 9, 12, 0, 0);
                            //MergeCells(theTable, 11, 12, 1, 1);
                            //MergeCells(theTable, 11, 12, 2, 2);
                            //MergeCells(theTable, 11, 12, 3, 3);
                            MergeCells(tableConstructionPara, 13, 14, 0, 0);
                            MergeCells(tableConstructionPara, 15, 18, 0, 0);
                            MergeCells(tableConstructionPara, 19, 22, 0, 0);
                            MergeCells(tableConstructionPara, 23, 23, 0, 1);
                            MergeCells(tableConstructionPara, 24, 24, 0, 1);
                            MergeCells(tableConstructionPara, 25, 25, 0, 1);
                            MergeCells(tableConstructionPara, 26, 26, 0, 1);
                            MergeCells(tableConstructionPara, 27, 27, 0, 1);
                            MergeCells(tableConstructionPara, 28, 28, 0, 1);
                            MergeCells(tableConstructionPara, 29, 29, 0, 1);
                            MergeCells(tableConstructionPara, 30, 30, 0, 1);
                            MergeCells(tableConstructionPara, 31, 31, 0, 1);
                            MergeCells(tableConstructionPara, 32, 32, 0, 1);
                            MergeCells(tableConstructionPara, 33, 33, 0, 1);
                            MergeCells(tableConstructionPara, 34, 34, 0, 1);
                            MergeCells(tableConstructionPara, 35, 35, 0, 1);
                            MergeCells(tableConstructionPara, 36, 36, 0, 1);


                            //MergeCells(theTable, 38, 38, 0, 1);
                            //MergeCells(theTable, 39, 39, 0, 1);
                        }
                        else
                        {
                            Trace.WriteLine("merge: singe, not armoured");
                            MergeCells(tableConstructionPara, 0, 0, 0, 5);
                            MergeCells(tableConstructionPara, 1, 1, 0, 1);
                            MergeCells(tableConstructionPara, 1, 1, 1, 3);
                            MergeCells(tableConstructionPara, 2, 2, 0, 1);
                            MergeCells(tableConstructionPara, 3, 8, 0, 0);
                            //MergeCells(theTable, 5, 6, 1, 1);
                            //MergeCells(theTable, 5, 6, 2, 2);
                            //MergeCells(theTable, 7, 8, 1, 1);
                            //MergeCells(theTable, 7, 8, 2, 2);
                            //MergeCells(theTable, 7, 8, 3, 3);
                            MergeCells(tableConstructionPara, 9, 12, 0, 0);
                            //MergeCells(theTable, 11, 12, 1, 1);
                            //MergeCells(theTable, 11, 12, 2, 2);
                            //MergeCells(theTable, 11, 12, 3, 3);
                            MergeCells(tableConstructionPara, 13, 16, 0, 0);
                            Trace.WriteLine("merge: singe, not armoured here");
                            //MergeCells(theTable, 18, 21, 0, 0);
                            //MergeCells(theTable, 22, 25, 0, 0);

                            MergeCells(tableConstructionPara, 17, 17, 0, 1);
                            MergeCells(tableConstructionPara, 18, 18, 0, 1);
                            MergeCells(tableConstructionPara, 19, 19, 0, 1);
                            MergeCells(tableConstructionPara, 20, 20, 0, 1);
                            MergeCells(tableConstructionPara, 21, 21, 0, 1);
                            MergeCells(tableConstructionPara, 22, 22, 0, 1);
                            MergeCells(tableConstructionPara, 23, 23, 0, 1);
                            MergeCells(tableConstructionPara, 24, 24, 0, 1);
                            MergeCells(tableConstructionPara, 25, 25, 0, 1);
                            MergeCells(tableConstructionPara, 26, 26, 0, 1);
                            MergeCells(tableConstructionPara, 27, 27, 0, 1);
                            MergeCells(tableConstructionPara, 28, 28, 0, 1);
                            MergeCells(tableConstructionPara, 29, 29, 0, 1);
                            MergeCells(tableConstructionPara, 30, 30, 0, 1);

                            //MergeCells(theTable, 32, 32, 0, 1);
                            //MergeCells(theTable, 33, 33, 0, 1);
                            //MergeCells(theTable, 34, 34, 0, 1);
                            //MergeCells(theTable, 35, 35, 0, 1);
                            //MergeCells(theTable, 36, 36, 0, 1);
                            //MergeCells(theTable, 37, 37, 0, 1);
                            //MergeCells(theTable, 38, 38, 0, 1);
                            //MergeCells(theTable, 39, 39, 0, 1);
                            //MergeCells(theTable, 40, 40, 0, 1);
                            //MergeCells(theTable, 41, 41, 0, 1);

                        }

                    }
                }

                //6,3 .2 合并单元格 电缆非电气技术参数
                void MergeCellNonElectric()
                {
                    Trace.WriteLine("合并0");
                    MergeCells(tableNonElecticPara, 0, 0, 0, 8);
                    MergeCells(tableNonElecticPara, 1, 1, 0, 2);
                    MergeCells(tableNonElecticPara, 1, 1, 2, 3);
                    MergeCells(tableNonElecticPara, 1, 1, 3, 4);
                    MergeCells(tableNonElecticPara, 2, 7, 0, 0);
                    MergeCells(tableNonElecticPara, 2, 2, 1, 2);
                    MergeCells(tableNonElecticPara, 2, 2, 3, 4);
                    MergeCells(tableNonElecticPara, 2, 2, 4, 5);
                    Trace.WriteLine("合并1");
                    MergeCells(tableNonElecticPara, 3, 3, 1, 2);
                    MergeCells(tableNonElecticPara, 3, 3, 3, 4);
                    MergeCells(tableNonElecticPara, 3, 3, 4, 5);
                    MergeCells(tableNonElecticPara, 4, 4, 1, 2);
                    MergeCells(tableNonElecticPara, 4, 4, 3, 4);
                    MergeCells(tableNonElecticPara, 4, 4, 4, 5);
                    MergeCells(tableNonElecticPara, 5, 5, 1, 2);
                    MergeCells(tableNonElecticPara, 5, 5, 3, 4);
                    MergeCells(tableNonElecticPara, 5, 5, 4, 5);
                    Trace.WriteLine("合并2");
                    MergeCells(tableNonElecticPara, 6, 6, 1, 2);
                    MergeCells(tableNonElecticPara, 6, 6, 3, 4);
                    MergeCells(tableNonElecticPara, 6, 6, 4, 5);
                    MergeCells(tableNonElecticPara, 7, 7, 1, 2);
                    MergeCells(tableNonElecticPara, 7, 7, 3, 4);
                    MergeCells(tableNonElecticPara, 7, 7, 4, 5);
                    Trace.WriteLine("合并3");
                    MergeCells(tableNonElecticPara, 8, 9, 0, 0);
                    MergeCells(tableNonElecticPara, 8, 9, 1, 1);
                    MergeCells(tableNonElecticPara, 8, 8, 4, 5);
                    MergeCells(tableNonElecticPara, 8, 8, 5, 6);
                    MergeCells(tableNonElecticPara, 8, 8, 6, 7);
                    MergeCells(tableNonElecticPara, 9, 9, 4, 5);
                    MergeCells(tableNonElecticPara, 9, 9, 5, 6);
                    MergeCells(tableNonElecticPara, 9, 9, 6, 7);
                    Trace.WriteLine("合并4");
                    MergeCells(tableNonElecticPara, 10, 31, 0, 0);
                    MergeCells(tableNonElecticPara, 10, 10, 1, 3);
                    MergeCells(tableNonElecticPara, 11, 11, 1, 2);
                    MergeCells(tableNonElecticPara, 12, 12, 1, 2);
                    MergeCells(tableNonElecticPara, 13, 13, 1, 2);
                    MergeCells(tableNonElecticPara, 14, 14, 1, 2);
                    MergeCells(tableNonElecticPara, 15, 15, 1, 2);
                    MergeCells(tableNonElecticPara, 16, 16, 1, 2);
                    MergeCells(tableNonElecticPara, 17, 17, 1, 2);
                    MergeCells(tableNonElecticPara, 18, 18, 1, 2);
                    MergeCells(tableNonElecticPara, 19, 19, 1, 2);
                    MergeCells(tableNonElecticPara, 20, 20, 1, 2);
                    MergeCells(tableNonElecticPara, 21, 21, 1, 2);
                    MergeCells(tableNonElecticPara, 22, 22, 1, 2);
                    MergeCells(tableNonElecticPara, 23, 23, 1, 2);
                    MergeCells(tableNonElecticPara, 24, 24, 1, 2);
                    MergeCells(tableNonElecticPara, 25, 25, 1, 2);
                    MergeCells(tableNonElecticPara, 26, 26, 1, 2);
                    MergeCells(tableNonElecticPara, 27, 27, 1, 2);
                    MergeCells(tableNonElecticPara, 28, 28, 1, 2);
                    MergeCells(tableNonElecticPara, 29, 29, 1, 2);
                    MergeCells(tableNonElecticPara, 30, 30, 1, 2);
                    MergeCells(tableNonElecticPara, 31, 31, 1, 2);
                }

                //6,3 .2 合并单元格 材料配置
                void MergeCellMaterial()
                {
                    Trace.WriteLine("合并0");
                    //  MergeCells(tableMaterial, 0, 0, 0, 8);
                    MergeCells(tableMaterial, 1, 2, 0, 0);
                    MergeCells(tableMaterial, 1, 2, 1, 1);
                    MergeCells(tableMaterial, 8, 10, 0, 0);
                    MergeCells(tableMaterial, 8, 8, 1, 5);
                    MergeCells(tableMaterial, 9, 9, 1, 5);
                    MergeCells(tableMaterial, 10, 10, 1, 5);
                }

                //6,3 .3 合并单元格 环境条件
                void MergeCellEnvironment()
                {
                    MergeCells(tableEnvironment, 0, 0, 0, 1);
                    MergeCells(tableEnvironment, 1, 1, 0, 1);
                    MergeCells(tableEnvironment, 2, 2, 0, 1);
                    MergeCells(tableEnvironment, 3, 3, 0, 1);
                    MergeCells(tableEnvironment, 4, 4, 0, 1);
                    MergeCells(tableEnvironment, 5, 5, 0, 1);
                    MergeCells(tableEnvironment, 6, 6, 0, 1);
                    MergeCells(tableEnvironment, 7, 8, 0, 0);
                    MergeCells(tableEnvironment, 9, 9, 0, 1);
                    MergeCells(tableEnvironment, 10, 10, 0, 1);
                }

            }

            public void CreateWordQuoteKeyValuePair()
            {
                //1 创建word文档
                var aWord = new XWPFDocument();

                //2 添加内容

                //2.1 ----第一页----
                //2.1.1 页面布局
                ulong pageWidth = theMargin(aWord, isLandscape: true); //页边距
                //2.1.2 插入文字
                var XWPFParagraph2 = aWord.CreateParagraph();
                SetRunStyle(XWPFParagraph2, " ");//插入空格

                //2.1.3 创建表格: 电缆结构技术参数
                XWPFTable tableQuote = CreateTable(TableDataInputQuote(), aWord);
                //2.1.3.1 设置表格宽度 
                ColumnWidthProportion(tableQuote, pageWidth, 3.79, 1.59, 2.33, 2.62, 1.85, 2.33, 2.54, 2.65, 4.64);
                //2.1.3.2 设置单元格对齐方式， 默认左上对齐
                // positioning(tableQuote);
                //2.1.3.3 合并单元格
                MergeCellQuote();


                //5  保存文档

                saveWord(aWord, $"报价定额{type_specInFileName} ");
                //saveWord(aWord, "Output报价Bom");


                //6   内部方法
                //6.1 输入数据至tableConstructionPara
                //6.1.1  至tableConstructionPara
                List<List<KeyValuePair<int, object>>> TableDataInputQuote()
                {
                    var tableData = new List<List<KeyValuePair<int, object>>>();   //table数据

                    // 添加元素		
                    var rowList1 = new List<KeyValuePair<int, object>>();  //行数据			
                    rowList1.Add(new KeyValuePair<int, object>(0, "营销报价定额\r\n(kg/km)"));
                    rowList1.Add(new KeyValuePair<int, object>(0, "——"));
                    rowList1.Add(new KeyValuePair<int, object>(0, "产品型号"));
                    rowList1.Add(new KeyValuePair<int, object>(1, theType));//待输入
                    rowList1.Add(new KeyValuePair<int, object>(0, "——"));
                    rowList1.Add(new KeyValuePair<int, object>(0, "电压等级"));
                    rowList1.Add(new KeyValuePair<int, object>(1, voltage + " kV")); //"0.6/1kV"));//待输入
                    rowList1.Add(new KeyValuePair<int, object>(0, "文件编号"));
                    rowList1.Add(new KeyValuePair<int, object>(1, file_code)); //"Q/LN1 05 003-2024"));//待输入
                    tableData.Add(rowList1);

                    var rowList2 = new List<KeyValuePair<int, object>>();  //行数据
                    rowList2.Add(new KeyValuePair<int, object>(0, "——"));
                    rowList2.Add(new KeyValuePair<int, object>(0, "——"));
                    rowList2.Add(new KeyValuePair<int, object>(0, "产品名称"));
                    rowList2.Add(new KeyValuePair<int, object>(1, production_name));  // "阻燃交联聚乙烯绝缘电力电缆"));//待输入
                    rowList2.Add(new KeyValuePair<int, object>(0, "——"));
                    rowList2.Add(new KeyValuePair<int, object>(0, "——"));
                    rowList2.Add(new KeyValuePair<int, object>(0, "——"));
                    rowList2.Add(new KeyValuePair<int, object>(0, "编制依据"));
                    rowList2.Add(new KeyValuePair<int, object>(1, source));  // "GB/T 12706.1-2020"));//待输入
                    tableData.Add(rowList2);

                    var rowList3 = new List<KeyValuePair<int, object>>();  //行数据
                    rowList3.Add(new KeyValuePair<int, object>(0, "规格"));
                    rowList3.Add(new KeyValuePair<int, object>(0, "导体"));
                    rowList3.Add(new KeyValuePair<int, object>(0, "绕包"));
                    rowList3.Add(new KeyValuePair<int, object>(0, "绝缘"));
                    rowList3.Add(new KeyValuePair<int, object>(0, "成缆"));
                    rowList3.Add(new KeyValuePair<int, object>(0, "——"));
                    rowList3.Add(new KeyValuePair<int, object>(0, "内衬层"));
                    rowList3.Add(new KeyValuePair<int, object>(0, "铠装"));
                    rowList3.Add(new KeyValuePair<int, object>(0, "外护套"));
                    tableData.Add(rowList3);

                    var rowList4 = new List<KeyValuePair<int, object>>();  //行数据
                    rowList4.Add(new KeyValuePair<int, object>(0, "——"));
                    rowList4.Add(new KeyValuePair<int, object>(1, conductorMaterial));  // "铜\r\n(kg/km)"));
                    rowList4.Add(new KeyValuePair<int, object>(1, mica_tapeMaterial));  // "云母带"));//输入
                    rowList4.Add(new KeyValuePair<int, object>(1, insulationMaterialSelected));  // "二步法硅烷交联绝缘料")); //输入    
                    rowList4.Add(new KeyValuePair<int, object>(1, bufferMaterialSelected));  // "填充绳"));////输入
                    rowList4.Add(new KeyValuePair<int, object>(1, tapeMaterialSelected));  // "三合一金云母带"));//输入
                    rowList4.Add(new KeyValuePair<int, object>(1, inner_sheathMaterial));  // "H-90 PVC\r\n护套料"));//输入
                    rowList4.Add(new KeyValuePair<int, object>(1, armourMaterialSelected));  // (isDoubleCable || isMultiCore) ? "多芯采用镀锌钢带" : "单芯采用不锈钢带"));  // 待输入(1, "单芯采用不锈钢带，\r\n多芯采用镀锌钢带"));//待输入
                    rowList4.Add(new KeyValuePair<int, object>(1, outer_sheathMaterialSelected));  // "ZH-90 PVC护套料（氧指数大于等于36%）"));//输入
                    tableData.Add(rowList4);

                    var rowList5 = new List<KeyValuePair<int, object>>();  //行数据
                    rowList5.Add(new KeyValuePair<int, object>(1, spec));//待输入
                    rowList5.Add(new KeyValuePair<int, object>(1, conductorWeight));//待输入
                    rowList5.Add(new KeyValuePair<int, object>(1, mica_tapeWeight));//待输入
                    rowList5.Add(new KeyValuePair<int, object>(1, insulationWeightSelected)); //"insulation1Weight"));//待输入
                    rowList5.Add(new KeyValuePair<int, object>(1, bufferWeightSelected));//待输入
                    rowList5.Add(new KeyValuePair<int, object>(1, tapeWeightSelected));//待输入
                    rowList5.Add(new KeyValuePair<int, object>(1, inner_sheathWeight));//待输入
                    rowList5.Add(new KeyValuePair<int, object>(1, armourWeight));//待输入
                    rowList5.Add(new KeyValuePair<int, object>(1, outer_sheathWeight));//待输入
                    tableData.Add(rowList5);

                    return tableData;
                }

                //6.2 创建table for word
                XWPFTable CreateTable(List<List<KeyValuePair<int, object>>> inputDataList, XWPFDocument theWord, string tableMark = "Normal")
                {
                    // 创建表格
                    var theTable = theWord.CreateTable(inputDataList.Count, inputDataList[0].Count); //（行数，列数） 行数+1用于表头
                                                                                                     // Trace.WriteLine("读取Word 结构参数 表格");
                    try
                    {
                        // 添加数据行
                        for (int rowIndex = 0; rowIndex < inputDataList.Count; rowIndex++)
                        {
                            //var row = theTable.GetRow(rowIndex + 1);
                            var row = theTable.GetRow(rowIndex);
                            //row.Height = 800;//高度.默认为自动高度
                            int colIndex = 0;
                            Trace.WriteLine("行循环");
                            foreach (var keyValuePair in inputDataList[rowIndex])
                            {
                                Trace.WriteLine("列循环");
                                int styleInt = keyValuePair.Key;//方法1
                                                                // object value = inputDataList[colIdx];
                                object value = keyValuePair.Value;

                                //  row.GetCell(colIndex).SetText(value);
                                var cell = row.GetCell(colIndex);
                                cell.SetVerticalAlignment(XWPFTableCell.XWPFVertAlign.CENTER); // //单元格内文字，垂直位置，默认靠上
                                cell.RemoveParagraph(0);//去掉段落,否则单元格前面会多一个空行
                                                        // var cellPara = cell.AddParagraph();
                                                        //cellPara.Alignment = ParagraphAlignment.CENTER;//水平靠中
                                                        //cellPara.SpacingBefore = 0; // 设置段落间距：段前为0
                                                        // cellPara.SpacingAfter = 0; // 设置段落间距：段后为0
                                                        // cellPara.SpacingBetween=0; // 设置行距单倍
                                                        // cellPara.SetSpacingBetween (0); // 设置行距1倍
                                                        // var run = cellPara.CreateRun();                                

                                if (Convert.ToString(value).Contains("\r\n"))
                                {
                                    string valueString = Convert.ToString(value);
                                    string value1 = valueString.Split("\r\n")[0];
                                    string value2 = valueString.Split("\r\n")[1];

                                    dataInput(value1);
                                    dataInput(value2);

                                }
                                else dataInput(value);

                                void dataInput(object value)
                                {
                                    var aParagraph = cell.AddParagraph();
                                    //Trace.WriteLine("run this ");
                                    switch (tableMark)
                                    {
                                        case "Normal":
                                            aParagraph.Alignment = ParagraphAlignment.CENTER; //单元格内文字，水平位置，默认靠左
                                            break;
                                        case "Material":
                                            if (rowIndex < 8 || colIndex == 0) aParagraph.Alignment = ParagraphAlignment.CENTER;//单元格内文字，水平位置，默认靠左
                                                                                                                                // else aParagraph.Alignment = ParagraphAlignment.LEFT;
                                            break;
                                        case "Environment":
                                            if (rowIndex == 0 || colIndex > 0) aParagraph.Alignment = ParagraphAlignment.CENTER;//单元格内文字，水平位置，默认靠左
                                                                                                                                // else aParagraph.Alignment = ParagraphAlignment.LEFT;
                                            break;
                                    }
                                    var run = aParagraph.CreateRun();
                                    run.FontSize = 9;
                                    run.FontFamily = "宋体";
                                    //run.SetFontFamily("宋体", FontCharRange.CS); //这代码只改变汉字字体，英语自动

                                    // 根据数据类型应用样式
                                    if (value == DBNull.Value || value == "")
                                    {
                                        //cell.SetText("数据未提供");
                                        run.SetText("数据未提供");//要设置文字颜色，文字输入也须用run设置。
                                        run.SetColor(ColorConverterWord.ToHexColor("Red")); //文字颜色
                                                                                            // cell.SetColor(ColorConverterWord.ToHexColor("Red"));//填充色
                                        Trace.WriteLine("数据未提供");

                                    }
                                    else
                                    {
                                        //cell.SetText(Convert.ToString(value));//这个只能设置文字，无法设置文字颜色      
                                        run.SetText(Convert.ToString(value));//用Run可以改变文字颜色，字体，大小 
                                        Trace.WriteLine($"Value: {Convert.ToString(value)}");
                                        if (faultMarkArray.Any(term => value.ToString().Contains(term, StringComparison.OrdinalIgnoreCase)))
                                        {
                                            run.SetColor(ColorConverterWord.ToHexColor("Red")); //文字颜色
                                            //cell.SetColor(ColorConverterWord.ToHexColor("Yellow")); //填充色
                                        }
                                        else if (styleInt == 1)
                                        {
                                            run.SetColor(ColorConverterWord.ToHexColor("Blue")); //文字颜色
                                                                                                 //cell.SetColor(ColorConverterWord.ToHexColor("Yellow")); //填充色
                                        }

                                        #region 类型判断
                                        //if (rowIndex == 0) cell.CellStyle = titleStyle;
                                        //else if (faultMarkArray.Any(term => value.ToString().Contains(term, StringComparison.OrdinalIgnoreCase)))
                                        //    cell.CellStyle = warnStyle;
                                        //else if (styleInt == 1)//方法2
                                        //    cell.CellStyle = stringBlueStyle;
                                        //else if (value.ToString().Contains("电缆长期允许载流量", StringComparison.OrdinalIgnoreCase))
                                        //{
                                        //    //合并单元格，文字超出单元格范围，行高不会自动变化
                                        //    iRow.Height = 576;// 14.4*2 *20;1/20个点为最小单位
                                        //    cell.CellStyle = stringCenterStyle;
                                        //}
                                        //else
                                        //    cell.CellStyle = stringCenterStyle;
                                        #endregion
                                    }
                                }

                                colIndex++;
                            }
                        }
                    }
                    catch
                    {
                        Trace.WriteLine("表格无数据");
                    }

                    return theTable;
                }

                //7 合并单元格
                //7.1 合并单元格 报价Bom
                void MergeCellQuote()
                {
                    //Trace.WriteLine("here 123");
                    MergeCells(tableQuote, 0, 1, 0, 1);
                    MergeCells(tableQuote, 0, 0, 2, 3);
                    MergeCells(tableQuote, 1, 1, 2, 5);
                    MergeCells(tableQuote, 2, 3, 0, 0);
                    MergeCells(tableQuote, 2, 2, 4, 5);
                    // Trace.WriteLine("here 234");
                }
            }

            //2.2 方法2：原有word文档，替代文字
            public void ReplaceWordTech(string templatePath)  // 替换原有文档
                                                              // public void ReplaceWordTech(string templatePath, string outputName)  // 替换原有文档
                                                              //public void ReplaceWordTech(string templatePath, string outputName, Dictionary<string, string> theDictionary)  // 替换原有文档

            {
                //string voltageLevel, armourString, coreString;
                //if (type_spec.Contains("0.6/1")) voltageLevel = "低压"; else voltageLevel = "高压？中压？低压？";
                //if (isArmoured) armourString = "铠装"; else armourString = "无铠装";
                //if (isDoubleCable) coreString = "双缆"; else if (isMultiCore) coreString = "多芯"; else coreString = "单芯";

                inPutTemplateFullName = templatePath + "替换模版电力电缆" + voltageLevel + armourString + coreString + ".docx"; //ZB-YJV 0.6/1 3×300＋1×150
                Trace.WriteLine($"inPutTemplateFullName：{inPutTemplateFullName}");
                //string outputName = $"工艺参数替代{type_specInFileName}";

                // Dictionary<string, string> theDictionary = InputDictionary();

                var dictionaryReplacement = new Dictionary<string, string>
                        {
                            { "【高低压】", voltageLevel },
                            { "【单多芯】", coreString},
                            { "【规格型号】", type_spec.ToString() },
                            { "【导体材料】", conductorMaterial},
                            { "【规格】",spec },
                            { "【最小单线根数】", pieces_1.ToString()},
                            { "【最小单线根数2】", pieces_2.ToString()},
                            { "【对应截面】", areaConductor[0]},
                            { "【对应截面2】", areaConductor[1]},
                            { "【导体外径】",conductDiameter_1.ToString() },
                            { "【导体外径2】", conductDiameter_2.ToString()},
                            { "【绝缘材料】", insulationMaterialSelected},
                            { "【绝缘厚度】", insulationThick_1.ToString()},
                            { "【绝缘厚度2】", insulationThick_2.ToString()},
                            { "【填充材料】", bufferMaterialSelected},
                            { "【内衬材料】", inner_sheathMaterial},
                            { "【内衬厚度】", inner_thick.ToString()},
                            { "【铠装材料】", armourMaterialSelected},
                            { "【铠装厚度】", steel_thick.ToString()},
                            { "【钢带层数】", "2"},
                            { "【钢带宽度】", steel_width.ToString()},
                            { "【外护材料】", outer_sheathMaterialFront},
                            { "【颜色】",  "黑色"},
                            { "【标称厚度有无铠装】", isArmoured ? "标称厚度t(有铠装)" : "标称厚度t(无铠装)"},
                            { "【外护厚度】", sheathThick.ToString()},
                            { "【外护薄点厚度】", isArmoured ? "铠装80%" : "无铠装85%"},
                            { "【电缆外径】", cableDiameter.ToString()},
                            { "【20度电阻】", resistant20_1.ToString()},
                            { "【20度电阻2】", resistant20_2.ToString()},
                            { "【90度电阻】", resistant90_1.ToString()},
                            { "【90度电阻2】", resistant90_2.ToString()},
                            { "【载流量】", current40.ToString()},
                            { "【敷设最大牵引力】", "70"},
                            { "【电缆质量】",cableWeight.ToString()},
                            { "【阻燃级别】", flameRedartant},
                            { "【无卤性能】", halogenFree},
                            { "【低烟性能】",smokeFree }
                        };



                //using (var templateStream = new FileStream(templatePath, FileMode.Open))
                using (var templateStream = new FileStream(inPutTemplateFullName, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                using (var doc = new XWPFDocument(templateStream))
                {
                    // 替换段落中的文本
                    foreach (var para in doc.Paragraphs)
                    {
                        string text = para.Text;
                        foreach (var replacement in dictionaryReplacement)
                        {
                            if (text.Contains(replacement.Key))
                            {
                                // text = text.Replace(replacement.Key, replacement.Value);
                                para.ReplaceText(replacement.Key, replacement.Value);
                            }
                        }
                    }

                    // 替换表格中的文本
                    foreach (var table in doc.Tables)
                    {
                        foreach (var row in table.Rows)
                        {
                            foreach (var cell in row.GetTableCells())
                            {

                                foreach (var para in cell.Paragraphs)
                                {
                                    string text = para.Text;
                                    foreach (var replacement in dictionaryReplacement)
                                    {
                                        if (text.Contains(replacement.Key))
                                        {
                                            para.ReplaceText(replacement.Key, replacement.Value);

                                            // 替换后更新text变量
                                            text = para.Text;
                                            if (faultMarkArray.Any(term => text.Contains(term, StringComparison.OrdinalIgnoreCase)))
                                            {
                                                // 先移除所有现有的runs
                                                while (para.Runs.Count > 0)
                                                {
                                                    para.RemoveRun(0);
                                                }

                                                XWPFRun redRun = para.CreateRun();
                                                redRun.FontSize = 9;
                                                redRun.SetText(text);
                                                redRun.SetColor(ColorConverterWord.ToHexColor("Red")); //文字颜色
                                            }
                                            else if (text == "")
                                            {
                                                while (para.Runs.Count > 0)
                                                {
                                                    para.RemoveRun(0);
                                                }

                                                XWPFRun redRun = para.CreateRun();
                                                //cell.SetText("数据未提供");
                                                redRun.FontSize = 9;
                                                redRun.SetText("数据未提供");//要设置文字颜色，文字输入也须用run设置。
                                                redRun.SetColor(ColorConverterWord.ToHexColor("Red")); //文字颜色
                                                                                                       // cell.SetColor(ColorConverterWord.ToHexColor("Red"));//填充色  
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }

                    // 保存生成的文档
                    saveWord(doc, $"工艺参数替代{type_specInFileName} ");

                }

                Dictionary<string, string> InputDictionaryXXX()
                {
                    Dictionary<string, string> dictionaryReplacement;

                    if (isDoubleCable)
                        if (isArmoured)
                            dictionaryReplacement = new Dictionary<string, string>
                        {
                            { "【高低压】", voltageLevel },
                            { "【单多芯】", coreString},
                            { "【规格型号】", type_spec.ToString() },
                            { "【导体材料】", conductorMaterial},
                            { "【规格】",spec },
                            { "【最小单线根数】", pieces_1.ToString()},
                            { "【最小单线根数2】", pieces_2.ToString()},
                            { "【对应截面】", areaConductor[0]},
                            { "【对应截面2】", areaConductor[1]},
                            { "【导体外径】",conductDiameter_1.ToString() },
                            { "【导体外径2】", conductDiameter_2.ToString()},
                            { "【绝缘材料】", insulationMaterialSelected},
                            { "【绝缘厚度】", insulationThick_1.ToString()},
                            { "【绝缘厚度2】", insulationThick_2.ToString()},
                            { "【填充材料】", bufferMaterialSelected},
                            { "【内衬材料】", inner_sheathMaterial},
                            { "【内衬厚度】", inner_thick.ToString()},
                            { "【铠装材料】", armourMaterialSelected},
                            { "【铠装厚度】", steel_thick.ToString()},
                            { "【钢带层数】", "2"},
                            { "【钢带宽度】", steel_width.ToString()},
                            { "【外护材料】", outer_sheathMaterialFront},
                            { "【颜色】",  "黑色"},
                            { "【标称厚度有无铠装】", isArmoured ? "标称厚度t(有铠装)" : "标称厚度t(无铠装)"},
                            { "【外护厚度】", sheathThick.ToString()},
                            { "【外护薄点厚度】", isArmoured ? "铠装80%" : "无铠装85%"},
                            { "【电缆外径】", cableDiameter.ToString()},
                            { "【20度电阻】", resistant20_1.ToString()},
                            { "【20度电阻2】", resistant20_2.ToString()},
                            { "【90度电阻】", resistant90_1.ToString()},
                            { "【90度电阻2】", resistant90_2.ToString()},
                            { "【载流量】", current40.ToString()},
                            { "【敷设最大牵引力】", "70"},
                            { "【电缆质量】",cableWeight.ToString()},
                            { "【阻燃级别】", flameRedartant},
                            { "【无卤性能】", halogenFree},
                            { "【低烟性能】",smokeFree }
                        };
                        else
                            dictionaryReplacement = new Dictionary<string, string>
                        {
                            { "【高低压】", voltageLevel },
                            { "【单多芯】", coreString},
                            { "【规格型号】", type_spec.ToString() },
                            { "【导体材料】", conductorMaterial},
                            { "【规格】",spec },
                            { "【最小单线根数】", pieces_1.ToString()},
                            { "【最小单线根数2】", pieces_2.ToString()},
                            { "【对应截面】", areaConductor[0]},
                            { "【对应截面2】", areaConductor[1]},
                            { "【导体外径】",conductDiameter_1.ToString() },
                            { "【导体外径2】", conductDiameter_2.ToString()},
                            { "【绝缘材料】", insulationMaterialSelected},
                            { "【绝缘厚度】", insulationThick_1.ToString()},
                            { "【绝缘厚度2】", insulationThick_2.ToString()},
                            { "【填充材料】", bufferMaterialSelected},
                            //{ "【内衬材料】", inner_sheathMaterial},
                            //{ "【内衬厚度】", inner_thick.ToString()},
                            //{ "【铠装材料】", armourMaterialSelected},
                            //{ "【铠装厚度】", steel_thick.ToString()},
                            //{ "【钢带层数】", "2"},
                            //{ "【钢带宽度】", steel_width.ToString()},
                            { "【外护材料】", outer_sheathMaterialFront},
                            { "【颜色】",  "黑色"},
                            { "【标称厚度有无铠装】", isArmoured ? "标称厚度t(有铠装)" : "标称厚度t(无铠装)"},
                            { "【外护厚度】", sheathThick.ToString()},
                            { "【外护薄点厚度】", isArmoured ? "铠装80%" : "无铠装85%"},
                            { "【电缆外径】", cableDiameter.ToString()},
                            { "【20度电阻】", resistant20_1.ToString()},
                            { "【20度电阻2】", resistant20_2.ToString()},
                            { "【90度电阻】", resistant90_1.ToString()},
                            { "【90度电阻2】", resistant90_2.ToString()},
                            { "【载流量】", current40.ToString()},
                            { "【敷设最大牵引力】", "70"},
                            { "【电缆质量】",cableWeight.ToString()},
                            { "【阻燃级别】", flameRedartant},
                            { "【无卤性能】", halogenFree},
                            { "【低烟性能】",smokeFree }
                        };
                    else if (isMultiCore)
                        if (isArmoured)
                            dictionaryReplacement = new Dictionary<string, string>
                        {
                            { "【高低压】", voltageLevel },
                            { "【单多芯】", coreString},
                            { "【规格型号】", type_spec.ToString() },
                            { "【导体材料】", conductorMaterial},
                            { "【规格】",spec },
                            { "【最小单线根数】", pieces_1.ToString()},
                            //{ "【最小单线根数2】", pieces_2.ToString()},
                            { "【对应截面】", areaConductor[0]},
                            //{ "【对应截面2】", areaConductor[1]},
                            { "【导体外径】",conductDiameter_1.ToString() },
                            //{ "【导体外径2】", conductDiameter_2.ToString()},
                            { "【绝缘材料】", insulationMaterialSelected},
                            { "【绝缘厚度】", insulationThick_1.ToString()},
                            //{ "【绝缘厚度2】", insulationThick_2.ToString()},
                            { "【填充材料】", bufferMaterialSelected},
                            { "【内衬材料】", inner_sheathMaterial},
                            { "【内衬厚度】", inner_thick.ToString()},
                            { "【铠装材料】", armourMaterialSelected},
                            { "【铠装厚度】", steel_thick.ToString()},
                            { "【钢带层数】", "2"},
                            { "【钢带宽度】", steel_width.ToString()},
                            { "【外护材料】", outer_sheathMaterialFront},
                            { "【颜色】",  "黑色"},
                            { "【标称厚度有无铠装】", isArmoured ? "标称厚度t(有铠装)" : "标称厚度t(无铠装)"},
                            { "【外护厚度】", sheathThick.ToString()},
                            { "【外护薄点厚度】", isArmoured ? "铠装80%" : "无铠装85%"},
                            { "【电缆外径】", cableDiameter.ToString()},
                            { "【20度电阻】", resistant20_1.ToString()},
                            //{ "【20度电阻2】", resistant20_2.ToString()},
                            { "【90度电阻】", resistant90_1.ToString()},
                            //{ "【90度电阻2】", resistant90_2.ToString()},
                            { "【载流量】", current40.ToString()},
                            { "【敷设最大牵引力】", "70"},
                            { "【电缆质量】",cableWeight.ToString()},
                            { "【阻燃级别】", flameRedartant},
                            { "【无卤性能】", halogenFree},
                            { "【低烟性能】",smokeFree }
                        };
                        else
                            dictionaryReplacement = new Dictionary<string, string>
                        {
                            { "【高低压】", voltageLevel },
                            { "【单多芯】", coreString},
                            { "【规格型号】", type_spec.ToString() },
                            { "【导体材料】", conductorMaterial},
                            { "【规格】",spec },
                            { "【最小单线根数】", pieces_1.ToString()},
                            //{ "【最小单线根数2】", pieces_2.ToString()},
                            { "【对应截面】", areaConductor[0]},
                            //{ "【对应截面2】", areaConductor[1]},
                            { "【导体外径】",conductDiameter_1.ToString() },
                            //{ "【导体外径2】", conductDiameter_2.ToString()},
                            { "【绝缘材料】", insulationMaterialSelected},
                            { "【绝缘厚度】", insulationThick_1.ToString()},
                            //{ "【绝缘厚度2】", insulationThick_2.ToString()},
                            { "【填充材料】", bufferMaterialSelected},
                            //{ "【内衬材料】", inner_sheathMaterial},
                            //{ "【内衬厚度】", inner_thick.ToString()},
                            //{ "【铠装材料】", armourMaterialSelected},
                            //{ "【铠装厚度】", steel_thick.ToString()},
                            //{ "【钢带层数】", "2"},
                            //{ "【钢带宽度】", steel_width.ToString()},
                            { "【外护材料】", outer_sheathMaterialFront},
                            { "【颜色】",  "黑色"},
                            { "【标称厚度有无铠装】", isArmoured ? "标称厚度t(有铠装)" : "标称厚度t(无铠装)"},
                            { "【外护厚度】", sheathThick.ToString()},
                            { "【外护薄点厚度】", isArmoured ? "铠装80%" : "无铠装85%"},
                            { "【电缆外径】", cableDiameter.ToString()},
                            { "【20度电阻】", resistant20_1.ToString()},
                            //{ "【20度电阻2】", resistant20_2.ToString()},
                            { "【90度电阻】", resistant90_1.ToString()},
                            //{ "【90度电阻2】", resistant90_2.ToString()},
                            { "【载流量】", current40.ToString()},
                            { "【敷设最大牵引力】", "70"},
                            { "【电缆质量】",cableWeight.ToString()},
                            { "【阻燃级别】", flameRedartant},
                            { "【无卤性能】", halogenFree},
                            { "【低烟性能】",smokeFree }
                        };
                    else
                        if (isArmoured)
                        dictionaryReplacement = new Dictionary<string, string>
                        {
                            { "【高低压】", voltageLevel },
                            { "【单多芯】", coreString},
                            { "【规格型号】", type_spec.ToString() },
                            { "【导体材料】", conductorMaterial},
                            { "【规格】",spec },
                            { "【最小单线根数】", pieces_1.ToString()},
                            //{ "【最小单线根数2】", pieces_2.ToString()},
                            { "【对应截面】", areaConductor[0]},
                            //{ "【对应截面2】", areaConductor[1]},
                            { "【导体外径】",conductDiameter_1.ToString() },
                            //{ "【导体外径2】", conductDiameter_2.ToString()},
                            { "【绝缘材料】", insulationMaterialSelected},
                            { "【绝缘厚度】", insulationThick_1.ToString()},
                            //{ "【绝缘厚度2】", insulationThick_2.ToString()},
                            //{ "【填充材料】", bufferMaterialSelected},
                            { "【内衬材料】", inner_sheathMaterial},
                            { "【内衬厚度】", inner_thick.ToString()},
                            { "【铠装材料】", armourMaterialSelected},
                            { "【铠装厚度】", steel_thick.ToString()},
                            { "【钢带层数】", "2"},
                            { "【钢带宽度】", steel_width.ToString()},
                            { "【外护材料】", outer_sheathMaterialFront},
                            { "【颜色】",  "黑色"},
                            { "【标称厚度有无铠装】", isArmoured ? "标称厚度t(有铠装)" : "标称厚度t(无铠装)"},
                            { "【外护厚度】", sheathThick.ToString()},
                            { "【外护薄点厚度】", isArmoured ? "铠装80%" : "无铠装85%"},
                            { "【电缆外径】", cableDiameter.ToString()},
                            { "【20度电阻】", resistant20_1.ToString()},
                            //{ "【20度电阻2】", resistant20_2.ToString()},
                            { "【90度电阻】", resistant90_1.ToString()},
                            //{ "【90度电阻2】", resistant90_2.ToString()},
                            { "【载流量】", current40.ToString()},
                            { "【敷设最大牵引力】", "70"},
                            { "【电缆质量】",cableWeight.ToString()},
                            { "【阻燃级别】", flameRedartant},
                            { "【无卤性能】", halogenFree},
                            { "【低烟性能】",smokeFree }
                        };
                    else
                        dictionaryReplacement = new Dictionary<string, string>
                        {
                             { "【高低压】", voltageLevel },
                            { "【单多芯】", coreString},
                            { "【规格型号】", type_spec.ToString() },
                            { "【导体材料】", conductorMaterial},
                            { "【规格】",spec },
                            { "【最小单线根数】", pieces_1.ToString()},
                            //{ "【最小单线根数2】", pieces_2.ToString()},
                            { "【对应截面】", areaConductor[0]},
                            //{ "【对应截面2】", areaConductor[1]},
                            { "【导体外径】",conductDiameter_1.ToString() },
                            //{ "【导体外径2】", conductDiameter_2.ToString()},
                            { "【绝缘材料】", insulationMaterialSelected},
                            { "【绝缘厚度】", insulationThick_1.ToString()},
                            //{ "【绝缘厚度2】", insulationThick_2.ToString()},
                            //{ "【填充材料】", bufferMaterialSelected},
                            //{ "【内衬材料】", inner_sheathMaterial},
                            //{ "【内衬厚度】", inner_thick.ToString()},
                            //{ "【铠装材料】", armourMaterialSelected},
                            //{ "【铠装厚度】", steel_thick.ToString()},
                            //{ "【钢带层数】", "2"},
                            //{ "【钢带宽度】", steel_width.ToString()},
                            { "【外护材料】", outer_sheathMaterialFront},
                            { "【颜色】",  "黑色"},
                            { "【标称厚度有无铠装】", isArmoured ? "标称厚度t(有铠装)" : "标称厚度t(无铠装)"},
                            { "【外护厚度】", sheathThick.ToString()},
                            { "【外护薄点厚度】", isArmoured ? "铠装80%" : "无铠装85%"},
                            { "【电缆外径】", cableDiameter.ToString()},
                            { "【20度电阻】", resistant20_1.ToString()},
                            //{ "【20度电阻2】", resistant20_2.ToString()},
                            { "【90度电阻】", resistant90_1.ToString()},
                            //{ "【90度电阻2】", resistant90_2.ToString()},
                            { "【载流量】", current40.ToString()},
                            { "【敷设最大牵引力】", "70"},
                            { "【电缆质量】",cableWeight.ToString()},
                            { "【阻燃级别】", flameRedartant},
                            { "【无卤性能】", halogenFree},
                            { "【低烟性能】",smokeFree }
                        };

                    return dictionaryReplacement;
                }

            }

            public void ReplaceWordQuote(string templatePath)  // 替换原有文档
            {
                //string voltageLevel, armourString, coreString;
                //if (type_spec.Contains("0.6/1")) voltageLevel = "低压"; else voltageLevel = "高压？中压？低压？";
                //if (isArmoured) armourString = "铠装"; else armourString = "无铠装";
                //if (isDoubleCable) coreString = "双缆"; else if (isMultiCore) coreString = "多芯"; else coreString = "单芯";

                inPutTemplateFullName = templatePath + "替换模板报价Bom.docx"; //ZB-YJV 0.6/1 3×300＋1×150
                Trace.WriteLine($"inPutTemplateFullName：{inPutTemplateFullName}");
                //string outputName = $"报价定额替代{type_specInFileName}";

                // Dictionary<string, string> theDictionary = InputDictionary();
                var dictionaryReplacement = new Dictionary<string, string> 
                {
                    { "【型号】", theType },
                    { "【电压大小】", voltage + " kV"},
                    { "【文件编号】",  file_code},
                    { "【产品名称】",  production_name},
                    { "【编制依据】",  source},
                    { "【导体材料】",  conductorMaterial},
                    { "【绕包材料】",  mica_tapeMaterial},
                    { "【绝缘材料】",  insulationMaterialSelected},
                    { "【填充材料】",  bufferMaterialSelected},
                    { "【成缆绕包材料】",  tapeMaterialSelected},
                    { "【内衬材料】",  inner_sheathMaterial},
                    { "【铠装材料】",  armourMaterialSelected},
                    { "【外护材料】", outer_sheathMaterialSelected},
                    { "【规格】", spec },
                    { "【导体质量】",  conductorWeight.ToString()},
                    { "【绕包质量】",  mica_tapeWeight.ToString()},
                    { "【绝缘质量】",  insulationWeightSelected.ToString()},
                    { "【填充质量】",  bufferWeightSelected.ToString()},
                    { "【成缆绕包质量】",  tapeWeightSelected.ToString()},
                    { "【内衬质量】", inner_sheathWeight.ToString() },
                    { "【铠装质量】",  armourWeight.ToString()},
                    { "【外护质量】",  outer_sheathWeight.ToString()}
                };

                using (var templateStream = new FileStream(inPutTemplateFullName, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                using (var doc = new XWPFDocument(templateStream))
                {
                    // 替换段落中的文本
                    foreach (var para in doc.Paragraphs)
                    {
                        string text = para.Text;
                        foreach (var replacement in dictionaryReplacement)
                        {
                            if (text.Contains(replacement.Key))
                            {
                                text = text.Replace(replacement.Key, replacement.Value);
                                para.ReplaceText(replacement.Key, replacement.Value);
                            }
                        }
                    }

                    // 替换表格中的文本
                    foreach (var table in doc.Tables)
                    {
                        foreach (var row in table.Rows)
                        {
                            foreach (var cell in row.GetTableCells())
                            {

                                foreach (var para in cell.Paragraphs)
                                {
                                    string text = para.Text;
                                    foreach (var replacement in dictionaryReplacement)
                                    {
                                        if (text.Contains(replacement.Key))
                                        {
                                            para.ReplaceText(replacement.Key, replacement.Value);

                                            // 替换后更新text变量
                                            text = para.Text;
                                            if (faultMarkArray.Any(term => text.Contains(term, StringComparison.OrdinalIgnoreCase)))
                                            {
                                                // 先移除所有现有的runs
                                                while (para.Runs.Count > 0)
                                                {
                                                    para.RemoveRun(0);
                                                }

                                                XWPFRun redRun = para.CreateRun();
                                                redRun.FontSize = 9;
                                                redRun.SetText(text);
                                                redRun.SetColor(ColorConverterWord.ToHexColor("Red")); //文字颜色
                                            }
                                            else if (text == "")
                                            {
                                                while (para.Runs.Count > 0)
                                                {
                                                    para.RemoveRun(0);
                                                }

                                                XWPFRun redRun = para.CreateRun();
                                                //cell.SetText("数据未提供");
                                                redRun.FontSize = 9;
                                                redRun.SetText("数据未提供");//要设置文字颜色，文字输入也须用run设置。
                                                redRun.SetColor(ColorConverterWord.ToHexColor("Red")); //文字颜色
                                                                                                       // cell.SetColor(ColorConverterWord.ToHexColor("Red"));//填充色  
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }

                    // 保存生成的文档
                    saveWord(doc, $"报价定额替代{type_specInFileName} ");
                }
            }

            //3 --------------------------------分页分节--------------------------------------------------
            void addBreak(XWPFParagraph XWPFParagraph1, bool isPageBreak = false) //isPageBreak = false 一般分页 isPageBreak = true 分节，页码页边距会变化
            {
                var run = XWPFParagraph1.CreateRun();
                if (isPageBreak)
                //分节，这个插在某个Run的前面，会在该run的前,及该run所在的段后，各插入一个分节符号；
                //如果插在段的结尾，会在段后插入一个分节符
                {
                    run.AddBreak(BreakType.PAGE); //这个插在段文字后面
                    var ctp = XWPFParagraph1.GetCTP();
                    // 添加段落属性（如果不存在）
                    if (ctp.pPr == null) ctp.pPr = new CT_PPr();
                    // 添加节属性
                    ctp.pPr.sectPr = new CT_SectPr();
                }
                else run.AddBreak();   //分页，这插在其他各个Run的前面 或后面，就会在相应位置插入一个分页
            }

            //4 --------------------------------Word 页边距--------------------------------------------------
            ulong theMargin(XWPFDocument doc, bool isLandscape = false, ulong theLeft = 1440, ulong theRight = 1440, ulong theTop = 1440, ulong theBottom = 1130, ulong theHeader = 720, ulong theFooter = 720)
            //  void theMargin(XWPFDocument doc, CT_SectPr sectPr, ulong theLeft = 1440, ulong theRight = 1440, ulong theTop = 1440, ulong theBottom = 1130, ulong theHeader = 720, ulong theFooter = 720)
            {
                if (doc.Document.body == null) doc.Document.body = new CT_Body();
                var sectPr = doc.Document.body.AddNewSectPr();

                // 设置页边距
                sectPr.pgMar = new CT_PageMar
                {
                    left = theLeft,   // 左页边距 （1英寸=1440 twips, 1厘米≈567 twips, 1 磅 = 20 twips）
                    right = theRight,  // 右页边距 1英寸
                    top = theTop,    // 上页边距 1.25英寸
                    bottom = theBottom, // 下页边距 1.25英寸
                    header = theHeader,  // 页眉边距 0.5英寸
                    footer = theFooter   // 页脚边距 0.5英寸
                };
                ulong pageWidth;
                if (isLandscape)
                {
                    var pgSz = sectPr.pgSz;
                    if (pgSz == null)
                    {
                        pgSz = new NPOI.OpenXmlFormats.Wordprocessing.CT_PageSz();
                        sectPr.pgSz = pgSz;
                    }
                    pgSz.orient = NPOI.OpenXmlFormats.Wordprocessing.ST_PageOrientation.landscape;
                    pageWidth = pgSz.w = 16838;  // A4宽度(297mm) in twentieths of a point
                    pgSz.h = 11906;  // A4高度(210mm) in twentieths of a point
                    pgSz.orient = NPOI.OpenXmlFormats.Wordprocessing.ST_PageOrientation.landscape; //.portrait 
                    pageWidth = pageWidth - theLeft - theRight;
                }
                else pageWidth = sectPr.pgSz.w - theLeft - theRight;
                return pageWidth;
            }


            //5 --------------------------------文字格式----------------------------------------------------
            void SetRunStyle(XWPFParagraph XWPFParagraph1, string theText, string fontFamily = "宋体", int fontSize = 0, string color = "Black", bool isBold = false, bool isItalic = false, UnderlinePatterns underline = UnderlinePatterns.None, bool breakChangeRow = true)
            {
                var run = XWPFParagraph1.CreateRun();
                run.SetText(theText);
                run.FontFamily = fontFamily;
                // run.SetFontFamily(fontFamily, FontCharRange.CS); //这代码只改变汉字字体，英语自动
                if (fontSize > 0) run.FontSize = fontSize; //默认为10.5，这个数字无法设定，只好设为0避开设置。
                run.CharacterSpacing = 0;// 设置文字间距离
                //if (color != null) run.SetColor(color);
                if (color != null) run.SetColor(ColorConverterWord.ToHexColor(color));

                run.IsBold = isBold;//加粗
                run.IsItalic = isItalic;//斜体
                run.Underline = underline;
                if (breakChangeRow) run.AddBreak(BreakType.TEXTWRAPPING);//空白行
            }

            //6 --------------------------------文字行距--------------------------------------------------
            void rowDistance(XWPFParagraph XWPFParagraph1, double size, int setType, int spacingBefore = 0, int SpacingAfter = 0)
            {
                if (spacingBefore > 0) XWPFParagraph1.SpacingBefore = 100; // 100=10磅 //段落前间距
                if (SpacingAfter > 0) XWPFParagraph1.SpacingAfter = 200; // 200=20磅   //段落后间距
                                                                         //设置段落内行距规则
                switch (setType)
                {
                    case 1:
                        XWPFParagraph1.SetSpacingBetween(size, LineSpacingRule.AUTO); // 倍行距  ？值 = 倍数(double) × 240. 行距会根据文字大小调节
                        break;
                    case 2:
                        XWPFParagraph1.SetSpacingBetween(size, LineSpacingRule.EXACT); // 常用值：12磅、16磅、20磅等  1磅 = 1 / 72英寸
                        break;
                    case 3:
                        XWPFParagraph1.SetSpacingBetween(size, LineSpacingRule.ATLEAST); //15 设置最小磅值，实际行距可能大于此值, 确保文本不会重叠
                        break;
                }
            }


            //7 --------------------------------Word 页码----------------------------------------------------
            void thePageNumber(XWPFDocument doc, bool hasFirstPage = false, short textForPage = 0)
            {
                // 1. 本节封面(即首页)页脚- 留空.只有首页不要页码时用，此时不要分节符。
                if (hasFirstPage)
                {
                    XWPFFooter footerFirst = doc.CreateFooter(HeaderFooterType.FIRST);
                    XWPFParagraph footerFirstPara = footerFirst.CreateParagraph();
                    footerFirstPara.CreateRun().SetText(""); // 空白页脚
                }


                // 2. 正文页脚（默认页脚） - 添加页码
                XWPFFooter footerDefault = doc.CreateFooter(HeaderFooterType.DEFAULT);
                XWPFParagraph footerPara = footerDefault.CreateParagraph();
                footerPara.Alignment = ParagraphAlignment.CENTER;

                // 使用多个Run确保所有文本正确显示
                if (textForPage == 1 || textForPage == 2)
                {
                    XWPFRun textRun1 = footerPara.CreateRun();
                    textRun1.SetText("第 ");
                }
                // 当前页码字段（从1开始）
                XWPFRun pageRun = footerPara.CreateRun();
                pageRun.GetCTR().AddNewFldChar().fldCharType = ST_FldCharType.begin;
                pageRun.GetCTR().AddNewInstrText().Value = "PAGE";
                pageRun.GetCTR().AddNewFldChar().fldCharType = ST_FldCharType.end;

                if (textForPage == 2)
                {
                    XWPFRun textRun2 = footerPara.CreateRun();
                    textRun2.SetText(" 页，共 ");

                    // 总页数字段（不包括封面,也不包括其他节）
                    XWPFRun numPagesRun = footerPara.CreateRun();
                    numPagesRun.GetCTR().AddNewFldChar().fldCharType = ST_FldCharType.begin;
                    numPagesRun.GetCTR().AddNewInstrText().Value = "SECTIONPAGES";  // "SECTIONPAGES" 本节计数页码，不包括封面页。  "NUMPAGES" 所有计数页码
                    numPagesRun.GetCTR().AddNewFldChar().fldCharType = ST_FldCharType.end;
                }

                if (textForPage == 1 || textForPage == 2)
                {
                    XWPFRun textRun3 = footerPara.CreateRun();
                    textRun3.SetText(" 页");
                }
                // ---------------------------- 设置页码起始值----------------------------
                // 获取文档的节属性（正文部分）
                var sectPr = doc.Document.body.sectPr; //页边距里，这一句doc.Document.body.AddNewSectPr();已经设定了doc.Document.body.sectPr，所以页边距一定要在页面前面
                if (sectPr == null)
                {
                    sectPr = new CT_SectPr();
                    doc.Document.body.sectPr = sectPr;
                }

                // 设置页码从1开始（封面计入）
                sectPr.pgNumType = new CT_PageNumber();
                sectPr.pgNumType.start = "1"; // 设置起始页码为1。有分节符时，封面页为起始页，无分节符呢？
            }



            //8 --------------------------------表格字符位置------------------------------------------------
            void positioning(XWPFTable table)
            {
                //这是整体位置设置，个别单元格，在程序中设置
                foreach (var row in table.Rows)
                {
                    foreach (var cell in row.GetTableCells())
                    {
                        cell.SetVerticalAlignment(XWPFTableCell.XWPFVertAlign.CENTER); //.TOP);   // 垂直居中
                        var paragraphs = cell.Paragraphs;
                        if (paragraphs.Count > 0)
                        {
                            int i = 0;
                            foreach (var theText in paragraphs)
                            {
                                paragraphs[i].Alignment = ParagraphAlignment.CENTER; // 水平居中，
                            }
                        }
                    }
                }
            }

            //9 --------------------------------表格行高----------------------------------------------------       
            void height(XWPFTable table)
            {
                // 第一行：自动行高（默认行为）
                var row0 = table.GetRow(0);
                row0.GetCell(0).SetText("自动行高（根据内容自动调整）");
                row0.GetCell(4).SetText("长文本示例：NPOI 是一个强大的 .NET 库，用于操作 Office 文档");

                // 第二行：固定行高（精确值）
                var row1 = table.GetRow(1);
                row1.Height = 800;  // 高度值（单位：缇twip，1/20点）
                row1.GetCell(0).SetText("固定行高 (800 twips ≈ 0.56厘米)");

                // 第三行：最小行高（自动撑高但保证最小高度）
                var row2 = table.GetRow(2);
                row2.Height = 400;   // 最小高度
                row2.GetCell(0).SetText("最小行高 (400 twips)，内容超出行高时会自动增加行高");
                row2.GetCell(4).SetText("当单元格内容超过最小高度时，行高会自动增加以适应内容。这是通过设置固定高度实现的自动扩展效果。");

                // 禁止行跨页断开
                //row3.IsBreakAcrossPages = false;

            }


            //10 --------------------------------表格列宽----------------------------------------------------        

            void ColumnWidth(XWPFTable table, params ulong[] widths)
            {

                // 获取并重置网格
                var grid = table.GetCTTbl().tblGrid;
                grid.gridCol.Clear();

                // 添加新列定义
                foreach (ulong width in widths)
                {
                    grid.AddNewGridCol().w = width;//（1英寸=1440 twips, 1厘米≈567 twips, 1 磅 = 20 twips）
                }
                table.GetCTTbl().tblPr.AddNewTblLayout().type = ST_TblLayoutType.@fixed;
            }
            void ColumnWidthProportion(XWPFTable table, double tableWidth, params double[] columnRatios)
            {
                // 1 cm = 567 twips
                // 1 inch = 1440 twips
                // 示例：`8000 twips ≈ 14.1cm`。

                double totalRatio = columnRatios.Sum();
                //double tableWidth = 567* tableWidthInCm; // 总宽度（twips）

                var grid = new CT_TblGrid();
                foreach (var ratio in columnRatios)
                {
                    double colWidth = (tableWidth * (ratio / totalRatio));
                    //Trace.WriteLine($"colWidth: {colWidth}");
                    grid.AddNewGridCol().w = (ulong)colWidth;
                }
                // Trace.WriteLine("here mmm")
                table.GetCTTbl().tblGrid = grid;
                table.GetCTTbl().tblPr.AddNewTblLayout().type = ST_TblLayoutType.@fixed; // ST_TblLayoutType.autofit;//autofit会根据文字大小自动宽度
            }

            //11 --------------------------------合并单元格-------------------------------------------------
            #region
            // 跨列非合并，单元格后移动
            void span(XWPFTable table)
            {
                table.GetRow(2).GetCell(2).GetCTTc().AddNewTcPr().AddNewGridspan().val = "2"; // 跨2列，单元格会向右移动
                table.GetRow(2).RemoveCell(3);  // 删除列2的单元格3 
            }


            //// 合并第 2 列和第 3 列
            void mergeColumns(XWPFTable table)
            {
                XWPFTableCell firstColumnCell = table.GetRow(2).GetCell(2);
                firstColumnCell.GetCTTc().AddNewTcPr().AddNewHMerge().val = ST_Merge.restart; // 开始合并列

                XWPFTableCell secondColumnCell = table.GetRow(2).GetCell(3);
                secondColumnCell.GetCTTc().AddNewTcPr().AddNewHMerge().val = ST_Merge.@continue; // 继续合并列
            }

            //合并行
            void mergeRows(XWPFTable table)
            {
                XWPFTableCell firstRowCell = table.GetRow(0).GetCell(3);
                firstRowCell.GetCTTc().AddNewTcPr().AddNewVMerge().val = ST_Merge.restart; // 开始合并行	

                XWPFTableCell secondRowCell = table.GetRow(1).GetCell(3);
                secondRowCell.GetCTTc().AddNewTcPr().AddNewVMerge().val = ST_Merge.@continue; // 继续合并行	
            }

            //合并2行2列单元格
            void mergeRow2Column2(XWPFTable table)
            {
                // 设置水平合并属性（跨2列）
                var tcPr = table.GetRow(0).GetCell(0).GetCTTc().tcPr ?? table.GetRow(0).GetCell(0).GetCTTc().AddNewTcPr();

                tcPr.AddNewGridspan().val = "2";
                //   tcPr.gridSpan = new CT_DecimalNumber { val = "2" }; // 跨2列

                tcPr.vMerge = new NPOI.OpenXmlFormats.Wordprocessing.CT_VMerge();
                tcPr.vMerge.val = ST_Merge.restart; // 合并开始



                var belowTcPr = table.GetRow(1).GetCell(0).GetCTTc().tcPr ?? table.GetRow(1).GetCell(0).GetCTTc().AddNewTcPr();
                belowTcPr.AddNewGridspan().val = "2";// 跨2列

                belowTcPr.vMerge = new NPOI.OpenXmlFormats.Wordprocessing.CT_VMerge();
                belowTcPr.vMerge.val = ST_Merge.@continue; // 合并延续

                //// 4. 删除被合并的冗余单元格
                table.GetRow(0).RemoveCell(1); // 删除第0行第1列
                table.GetRow(1).RemoveCell(1); // 删除第1行第1列
            }

            //合并多行多列的单元格
            public void MergeCells(XWPFTable table, int rowStart, int rowEnd, int colStart, int colEnd)
            {
                // 1. 设置起始单元格的合并属性
                var startCell = table.GetRow(rowStart).GetCell(colStart);
                var tcPr = startCell.GetCTTc().tcPr ?? startCell.GetCTTc().AddNewTcPr();

                // 跨列设置
                //tcPr.gridSpan = new CT_DecimalNumber { val = colEnd - colStart + 1 };
                tcPr.AddNewGridspan().val = $"{colEnd - colStart + 1}";
                // 跨行设置
                tcPr.vMerge = new CT_VMerge { val = ST_Merge.restart };

                // 2. 设置后续行的合并属性
                for (int row = rowStart + 1; row <= rowEnd; row++)
                {
                    var cell = table.GetRow(row).GetCell(colStart);
                    if (cell == null) continue;

                    var cellTcPr = cell.GetCTTc().tcPr ?? cell.GetCTTc().AddNewTcPr();
                    // cellTcPr.gridSpan = new CT_DecimalNumber { val = colEnd - colStart + 1 };
                    cellTcPr.AddNewGridspan().val = $"{colEnd - colStart + 1}";
                    cellTcPr.vMerge = new CT_VMerge { val = ST_Merge.@continue };
                }

                // 3. 删除被合并的冗余单元格
                for (int row = rowStart; row <= rowEnd; row++)
                {
                    for (int col = colEnd; col > colStart; col--)
                    {
                        var rowObj = table.GetRow(row);
                        if (rowObj.GetCell(col) != null)
                        {
                            rowObj.RemoveCell(col);
                        }
                    }
                }
            }
            #endregion


            //12 --------------------------------保存Word----------------------------------------------------      
            void saveWord(XWPFDocument doc, string fileName)
            {
                try
                {
                    using (SaveFileDialog saveFileDialog = new SaveFileDialog())//弹出对话框，可指定Excel 存储路径，文件名
                    {
                        saveFileDialog.Filter = "Word文件|*.docx";
                        saveFileDialog.Title = "保存Word文件";
                        saveFileDialog.FileName = fileName + DateTime.Now.ToString("yyyy-MM-dd HH-mm-ss-fff") + ".docx";

                        // string filePath = "output.xlsx";//默认在...bin\里面
                        //string filePath = @"C:\Temp\MySqlOutput.xlsx";

                        //string tempPath = Path.GetTempPath();//C:\Users\Admin\AppData\Local\Temp
                        //string filePath = Path.Combine(tempPath, "MySQLOutput.xlsx");

                        if (saveFileDialog.ShowDialog() == DialogResult.OK)//选择保存路径
                        {
                            // string filePath = "output.docx";//默认在...bin\里面

                            //string filePath = @"C:\Temp\MySqlOutput.docx";

                            //string tempPath = Path.GetTempPath();//C:\Users\Admin\AppData\Local\Temp
                            //string filePath = Path.Combine(tempPath, "MySQLOutput.docx");

                            string filePath = saveFileDialog.FileName;

                            Directory.CreateDirectory(Path.GetDirectoryName(filePath)); //创建目录
                                                                                        // 4.1 （FileStream）
                            /*   using (var fs = new FileStream(filePath, FileMode.Create))
                             // FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                              {
                                  workbook.Write(fs);
                              }
                              */
                            #region 核心代码
                            using (FileStream fs = new FileStream(filePath, FileMode.Create))
                            {
                                doc.Write(fs);
                            }
                            #endregion
                        }
                    }
                    Trace.WriteLine("文件保存成功");
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"创建Word文件失败: {ex.ToString()}");
                    Trace.WriteLine($"创建Word文件失败: {ex.ToString()}");
                }
                finally
                {
                    // 5. 手动清理资源（根据NPOI版本可能需要）
                    if (doc is IDisposable disposable)
                    {
                        disposable.Dispose();
                    }
                }
            }

        }

        //private void label1_Click(object sender, EventArgs e)
        //{

        //}
    }


    public static class ColorConverterWord   //文字颜色，NPOI为16进制，用这个方便阅读
    {
        // 常用颜色映射表
        private static readonly Dictionary<string, string> ColorMap =
            new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
        { "Black", "000000" },
        { "White", "FFFFFF" },
        { "Red", "FF0000" },
        { "Green", "00FF00" },
        { "Blue", "0000FF" },
        { "Yellow", "FFFF00" },
        { "Cyan", "00FFFF" },
        { "Magenta", "FF00FF" },
        { "Gray", "808080" },
        { "DarkRed", "8B0000" },
        { "DarkGreen", "006400" },
        { "DarkBlue", "00008B" },
        { "Orange", "FFA500" },
        { "Purple", "800080" },
        { "Pink", "FFC0CB" }
        };

        /// <summary>
        /// 将任意格式颜色转换为NPOI所需的6位十六进制格式
        /// </summary>
        /// <param name="colorInput">颜色输入（名称、十六进制、RGB值）</param>
        /// <returns>6位十六进制颜色字符串</returns>
        public static string ToHexColor(string colorInput)
        {
            if (string.IsNullOrWhiteSpace(colorInput))
                return "000000"; // 默认黑色

            // 如果已经是6位十六进制格式
            if (System.Text.RegularExpressions.Regex.IsMatch(colorInput, "^[0-9A-Fa-f]{6}$"))
            {
                return colorInput.ToUpper();
            }

            // 处理带#前缀的十六进制
            if (colorInput.StartsWith("#"))
            {
                string hex = colorInput.Substring(1).ToUpper();
                return hex.Length switch
                {
                    6 => hex, // #RRGGBB
                    3 => $"{hex[0]}{hex[0]}{hex[1]}{hex[1]}{hex[2]}{hex[2]}", // #RGB => RRGGBB
                    _ => "000000" // 无效格式
                };
            }

            // 检查颜色名称
            if (ColorMap.TryGetValue(colorInput, out string hexValue))
            {
                return hexValue;
            }

            // 处理RGB格式 (255,0,0)
            if (colorInput.Contains(","))
            {
                string[] parts = colorInput.Split(',');
                if (parts.Length == 3 &&
                    int.TryParse(parts[0], out int r) &&
                    int.TryParse(parts[1], out int g) &&
                    int.TryParse(parts[2], out int b))
                {
                    return $"{r:X2}{g:X2}{b:X2}";
                }
            }

            // 尝试解析为系统颜色
            try
            {
                Color color = Color.FromName(colorInput);
                if (color.IsKnownColor)
                {
                    return $"{color.R:X2}{color.G:X2}{color.B:X2}";
                }
            }
            catch
            {
                // 忽略解析错误
            }

            // 默认返回黑色
            return "000000";
        }
        private static int Clamp(int value) => Math.Clamp(value, 0, 255);
    }
}
