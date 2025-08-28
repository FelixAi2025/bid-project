using System.Diagnostics;
using System.Text.RegularExpressions;
using static MysqlToExcelWord.FormStart;
using MatchRegex = System.Text.RegularExpressions.Match;

namespace MysqlToExcelWord
{
    public partial class CnLv : UserControl
    {
        internal static FormSpec? specForm;
        internal ControlOutput controlOutput;
        bool specSelected;
        public CnLv()
        {
            InitializeComponent();
            theType = type_spec = spec = voltage = type_specInFileName = "";
            tableNameFromButton = typeInMaterialTable = "";
            controlOutput = new ControlOutput();
            this.Controls.Add(controlOutput);
        }

        private void CnLv_Load(object sender, EventArgs e)
        {
            checkBoxList = new List<CheckBox>();
            foreach (Control control in this.Controls)
            {
                // if (control is CheckBox checkBox && checkBox.Tag?.ToString() == "batch")
                if (control is CheckBox checkBox)
                {
                    checkBoxList.Add(checkBox);
                    checkBox.CheckedChanged += CheckBox_CheckedChanged;
                }
            }
        }

        public void ToCancel()
        {
            foreach (CheckBox theCheckBox in checkBoxList)
            {
                theCheckBox.Checked = false;
            }

            theType = type_spec = spec = voltage = type_specInFileName = "";
            tableNameFromButton = typeInMaterialTable = "";
        }

        private void CheckBox_CheckedChanged(object sender, EventArgs e)
        {
            // 获取被点击的CheckBox
            CheckBox checkBoxClicked = sender as CheckBox;

            if (checkBoxClicked != null)
            {
                InspectCheckBox();
                if (checkBoxClicked.Checked)
                {

                    // tableNameFromButton = "Z_YJVx2";

                    specForm = new FormSpec();
                    specForm.EventCableSpec += (sender, value) =>
                    {

                        specSelected = true;
                        //  var theSender = sender as SpecForm;
                        type_spec = value;
                        type_specInFileName = type_spec.Replace('/', '_');
                        Trace.WriteLine($"\n  type_spec.Replace： {type_specInFileName}");
                        //textBox1.Text = typeFromButton +" "+ type_spec;
                        //    Trace.WriteLine($"theSender get ： {theSender.Type_specFromMysql}");
                        //   Trace.WriteLine($"specForm get ： {specForm.Type_specFromMysql}");
                        Trace.WriteLine($"\nvalue get ： {value}");
                        spec = specForm.SpecFromButton;
                        theType = type_spec.Split(' ')[0];
                        voltage = type_spec.Split(' ')[1];
                        // textBox1.Text = value;
                        MatchCollection matchCollection = Regex.Matches(spec, @"(\d×)|(\dX)|(\dx)");
                        foreach (MatchRegex match1 in matchCollection)
                        {
                            MatchRegex theMatch = Regex.Match(match1.Value, @"\d");
                            //   Trace.WriteLine($"数字：{theMatch.Value}");
                            isMultiCore = (Convert.ToInt16(theMatch.Value) > 1);
                            Trace.WriteLine($"theMatch.Value: {Convert.ToInt16(theMatch.Value)}");
                            // specMini.Add(match1.Value);
                        }
                        controlOutput.button1.Enabled = controlOutput.button2.Enabled = controlOutput.button3.Enabled = controlOutput.button4.Enabled = true;
                        controlOutput.textBox1.Text = type_spec;
                        controlOutput.EventCancel += (sender, e) => { ToCancel(); };
                        getMysqlData();  //MySql数据查询

                    };

                    specForm.FormClosed += (sender, e) =>
                    {
                        if (!specSelected) ToCancel();
                        else specSelected = false;
                        // Trace.WriteLine($"SpecForm_FormClosed receive value {value}");
                    };


                    tableNameFromButton = Regex.Replace(checkBoxClicked.Name, "cBox_", "", RegexOptions.IgnoreCase); //不区分大小写
                    typeInMaterialTable = Regex.Replace(tableNameFromButton, "_", "-");
                    specForm.TableNameFromButton = tableNameFromButton;
                    specForm.Show();
                    //specForm.ShowDialog();//无法监听操作
                }
            }
        }

        void InspectCheckBox()
        {

            numChecked = 0;
            foreach (CheckBox theCheckBox in checkBoxList)
            {
                numChecked += theCheckBox.Checked ? 1 : 0;
            }

            Trace.WriteLine($"numChecked:   {numChecked}");

            if (numChecked >= 1)
            {

                foreach (CheckBox theCheckBox in checkBoxList)
                {
                    theCheckBox.Enabled = false;
                }
                label2.ForeColor = label3.ForeColor = label6.ForeColor = label7.ForeColor = label8.ForeColor = Color.Gray;
                // button1.Enabled= button3.Enabled = true;
            }
            else
            {
                /*  cBox_Z_VVx2.Enabled = true; cBox_ZA_VV.Enabled = true; cBox_ZB_VV.Enabled = true; cBox_ZC_VV.Enabled = true;
                  cBox_ZC_YJV.Enabled = true; cBox_ZB_YJV.Enabled = true; cBox_ZA_YJV.Enabled = true; cBox_Z_YJVx2.Enabled = true; cBox_ZC_YJY.Enabled = true;
                  cBox_ZB_YJY.Enabled = true; cBox_ZA_YJY.Enabled = true; cBox_Z_YJYx3.Enabled = true; cBox_N_VVX2.Enabled = true; cBox_N_VV.Enabled = true;
                  cBox_N_YJV.Enabled = true; cBox_N_YJVx2.Enabled = true; cBox_N_YJY.Enabled = true; cBox_N_YJYx3.Enabled = true; cBox_WDZAN_YJY.Enabled = true;
                  cBox_WDZBN_YJY.Enabled = true; cBox_WDZCN_YJY.Enabled = true; cBox_WDZN_YJYx3.Enabled = true; cBox_WDZ_YJYx3.Enabled = true; cBox_WDZC_YJY.Enabled = true;
                  cBox_WDZB_YJY.Enabled = true; cBox_WDZA_YJY.Enabled = true; cBox_ZN_YJVx2.Enabled = true; cBox_ZCN_YJV.Enabled = true; cBox_ZBN_YJV.Enabled = true;
                  cBox_ZAN_YJV.Enabled = true; cBox_ZN_VVx2.Enabled = true; cBox_ZCN_VV.Enabled = true; cBox_ZBN_VV.Enabled = true; cBox_ZAN_VV.Enabled = true;
                  cBox_ZN_YJYx3.Enabled = true; cBox_ZCN_YJY.Enabled = true; cBox_ZBN_YJY.Enabled = true; cBox_ZAN_YJY.Enabled = true; cBox_YJY.Enabled = true;
                  cBox_YJYx3.Enabled = true; cBox_YJV.Enabled = true; cBox_YJVx2.Enabled = true; cBox_VV.Enabled = true; cBox_VVx2.Enabled = true;
                */
                foreach (CheckBox theCheckBox in checkBoxList)
                {
                    theCheckBox.Enabled = true;
                }
                label2.ForeColor = label3.ForeColor = label6.ForeColor = label7.ForeColor = label8.ForeColor = Color.Black;
                //button1.Enabled = button3.Enabled = button4.Enabled = false;
                //textBox1.Text = "";
            }

        }





    }
}
