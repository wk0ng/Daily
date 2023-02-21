using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Text.RegularExpressions;
using MSWord = Microsoft.Office.Interop.Excel;

namespace Daily
{
    public partial class Form1 : Form
    {
        private string intoFilename = "";
        private int maxDay = 0;
        struct Info
        {
            public string name;
            public List<string> days;
            public List<string> other;
            public List<string> latter;
        }
        public Form1()
        {
            InitializeComponent();
            this.StartPosition = FormStartPosition.CenterScreen;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            OpenFileDialog dialog = new OpenFileDialog();
            dialog.Multiselect = false;//该值确定是否可以选择多个文件
            dialog.Title = "请选择文件";
            dialog.Filter = "Excel文件(*.xls,*.xlsx)|*.xls;*.xlsx";
            if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                this.intoFilename = dialog.FileName;
                this.textBox1.Text = this.intoFilename;
                this.textBox2.Text = this.textBox1.Text + "____导出.xlsx";
            }
        }

        private void textBox3_TextChanged(object sender, EventArgs e)
        {
            textBox3.Text = Regex.Replace(textBox3.Text, @"\D", ",");
            textBox3.Text = Regex.Replace(textBox3.Text, @",{2,}", ",");
            textBox3.Select(textBox3.TextLength, 0);
        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.button2.Text = "正在统计";
            this.button2.Enabled = false;
            this.button1.Enabled = false;
            this.textBox2.ReadOnly = true;
            this.textBox3.ReadOnly = true;
            this.textBox4.ReadOnly = true;

            // 加载不需要填写日期，给1~9前面加0
            String[] notNeed = textBox3.Text.Split(',');

            for(int i=0; i<notNeed.Length; i++)
            {
                if (notNeed[i].Length == 1)
                {
                    notNeed[i] = '0' + notNeed[i];
                }
            }

            // 加载需要填写日期，给1~9前面加0
            List<String> needSub = new List<string>();
            for(int i=1; i<maxDay; i++)
            {
                string tmpday = i.ToString();
                if (i < 10)
                {
                    tmpday = "0" + tmpday;
                }

                if (!notNeed.Contains(tmpday))
                {
                    needSub.Add(tmpday);
                }
            }

            MSWord.Application excelInput = new MSWord.Application();
            excelInput.Visible = false;
            MSWord.Workbooks wbs = excelInput.Workbooks;

            MSWord._Workbook wb = wbs.Add(textBox1.Text);

            Dictionary<string, Info> infoList = new Dictionary<string, Info>();
            for(int i=1;i < wb.Sheets.Count+1; i++)
            {
                MSWord._Worksheet inpuSheet = wb.Sheets[i];
                int rowCount = inpuSheet.Rows.Count;

                for(int r = 2; r < rowCount + 1; r++)
                {
                    String name = ((MSWord.Range)inpuSheet.Cells[r, 2]).Text;
                    String myDate = ((MSWord.Range)inpuSheet.Cells[r, 4]).Text;
                    if(name != "" && myDate != "")
                    {
                        String day = myDate.Split(new string[] { "年" }, StringSplitOptions.RemoveEmptyEntries)[1].Split(new string[] { "月" }, StringSplitOptions.RemoveEmptyEntries)[1].Split(new string[] { "日" }, StringSplitOptions.RemoveEmptyEntries)[0];
                        String timeHour = myDate.Split(' ')[1].Split(':')[0];
                        String timeMin = myDate.Split(' ')[1].Split(':')[1];

                        if (!infoList.ContainsKey(name))
                        {
                            Info newUser = new Info();
                            newUser.name = name;
                            newUser.days = new List<string>();
                            newUser.other = new List<string>();
                            newUser.latter = new List<string>();
                            infoList.Add(name, newUser);
                        }

                        infoList[name].days.Add(day);

                        if (notNeed.Contains(day))
                        {
                            infoList[name].other.Add(day);
                        }
                        else
                        {
                            if(timeHour=="23" && timeMin != "00")
                            {

                                infoList[name].latter.Add(day+"日"+" "+timeHour+":"+timeMin);
                            }
                        }
                    }
                    else
                    {
                        r = rowCount;
                    }
                }
            }
            //关闭excel
            wb.Close();
            wbs.Close();
            excelInput.Quit();
            System.Runtime.InteropServices.Marshal.ReleaseComObject(excelInput);

            //开始写入excel
            object Nothing = System.Reflection.Missing.Value;

            MSWord.Application excelOutput = new MSWord.Application();
            excelOutput.Visible = false;
            MSWord.Workbooks wbsOut = excelOutput.Workbooks;

            MSWord._Workbook wbOut = wbsOut.Add(Nothing);

            MSWord.Worksheet wbSheet = (MSWord.Worksheet)wbOut.Sheets[1];
            wbSheet.Name = "日报统计";

            wbSheet.Cells[1, 1] = "序号";
            wbSheet.Cells[1, 2] = "钉钉昵称";
            wbSheet.Cells[1, 3] = "提交次数";
            wbSheet.Cells[1, 4] = "未提交";
            wbSheet.Cells[1, 5] = "晚交";
            wbSheet.Cells[1, 6] = "加班确认";

            Dictionary<string, Info>.KeyCollection allKeys = infoList.Keys;

            int tNum = 1;
            
            foreach(string k in allKeys)
            {
                Info tmp = infoList[k];
                int total = tmp.days.Count;
                string notSub = "";
                string latterSub = string.Join("、", tmp.latter.ToArray());
                string otherSub = string.Join("、", tmp.other.ToArray());

                for(int i =0;i<needSub.Count; i++)
                {
                    if (!tmp.days.Contains(needSub[i]))
                    {
                        if (notSub != "")
                        {
                            notSub = notSub + "、";
                        }
                        notSub = notSub + needSub[i];
                    }
                }

                wbSheet.Cells[tNum + 1, 1] = tNum.ToString();
                wbSheet.Cells[tNum + 1, 2] = tmp.name;
                wbSheet.Cells[tNum + 1, 3] = total.ToString();
                wbSheet.Cells[tNum + 1, 4] = notSub;
                wbSheet.Cells[tNum + 1, 5] = latterSub;
                wbSheet.Cells[tNum + 1, 6] = otherSub;

                tNum = tNum + 1;
            }

            wbSheet.SaveAs(textBox2.Text, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, MSWord.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing);


            //关闭excel
            wbOut.Close(false, Type.Missing, Type.Missing);
            wbsOut.Close();
            excelOutput.Quit();
            System.Runtime.InteropServices.Marshal.ReleaseComObject(excelOutput);

            this.button2.Text = "导出";
            this.button2.Enabled = true;
            this.button1.Enabled = true;
            this.textBox2.ReadOnly = false;
            this.textBox3.ReadOnly = false;
            this.textBox4.ReadOnly = false;
            MessageBox.Show("统计完成！");
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            this.maxDay = System.Threading.Thread.CurrentThread.CurrentUICulture.Calendar.GetDaysInMonth(DateTime.Now.Year, DateTime.Now.Month);
            textBox4.Text = maxDay.ToString();

            var startMoth = DateTime.Now.AddDays(-DateTime.Now.Day + 1).Date;
            var stopMoth = startMoth.AddMonths(1).AddDays(-1).Date;

            var tempDate = startMoth;
            var today = DateTime.Now.Date;

            while (tempDate <= stopMoth)
            {
                // Monday 从 国内星期天开始
                if(tempDate.DayOfWeek == DayOfWeek.Sunday || tempDate.DayOfWeek == DayOfWeek.Saturday || tempDate>= today)
                {
                    if (textBox3.Text.Length != 0)
                    {
                        textBox3.Text = textBox3.Text + ",";
                    }
                    textBox3.Text = textBox3.Text + tempDate.Day.ToString();
                }
                tempDate = tempDate.AddDays(1);
            }
        }

        private void textBox4_TextChanged(object sender, EventArgs e)
        {
            this.maxDay = int.Parse(textBox4.Text);
        }
    }
}
