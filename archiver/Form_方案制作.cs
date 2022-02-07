using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Windows.Forms;
using archiver.ConsoleColorWriter;
using archiver.Properties;
using Xceed.Document.NET;

namespace archiver
{
    public partial class Form_方案制作 : Form
    {
        myutil doc;//等保报告
        myutil tempo;//出报告模板
        string str_公司;
        string str_系统;
        string str_P号;
        bool tryoldVer = false;
        public Form_方案制作()
        {
            
            InitializeComponent();
            AllocConsole();
        }
        [DllImport("kernel32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        static extern bool AllocConsole();
        private void Form_方案制作_Load(object sender, EventArgs e)
        {
            loadSampleList();
            //加载保存的人名
            textBox_zuozhe.Text = Settings.Default.author_str;
            textBox_pingshen.Text = Settings.Default.auditor_Str;
        }

        #region textDrop
        private void textBox1_DragDrop(object sender, DragEventArgs e)
        {
            textBox1.Text = ((System.Array)e.Data.GetData(DataFormats.FileDrop)).GetValue(0).ToString();
            string path = textBox1.Text;
            this.Hide();
            ConsoleWriter.WriteColoredText("加载报告...",ConsoleColor.Green);
            doc = new myutil(path);
            ConsoleWriter.WriteColoredText("加载完成", ConsoleColor.Green);
            this.Show();
        }

        private void textBox1_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                e.Effect = DragDropEffects.Link;
            }
            else
            {
                e.Effect = DragDropEffects.None;
            }
        }
        //加载报告函数
        void textbox_do()
        {
            string path = textBox1.Text;
            this.Hide();
            ConsoleWriter.WriteColoredText("加载报告...", ConsoleColor.Green);
            doc = new myutil(path);
            ConsoleWriter.WriteColoredText("加载完成", ConsoleColor.Green);
            this.Show();
        }


        #endregion

        #region 加载模板

        private void loadSample()
        {
            if (listBox1.SelectedItem != null)
            {
                string tempo_path = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + @$"\Sample\方案";
                if (!System.IO.Directory.Exists(tempo_path))
                {
                    toolStripStatusLabel1.ForeColor = System.Drawing.Color.Red;
                    toolStripStatusLabel1.Text = "获取模板 失败";
                }
                else
                {
                    //拼接字符串，加载项目模板
                    string  a = tempo_path + "\\"+ listBox1.SelectedItem.ToString();
                    tempo = new myutil(a);
                    ConsoleWriter.WriteYEllow("模板已加载" + a);
                }
            }
            else
            {
                toolStripStatusLabel1.ForeColor = System.Drawing.Color.Red;
                toolStripStatusLabel1.Text = "先选择一个模板";
            }
        }

        private void loadSampleList()
        {
            string tempo_path = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + @$"\Sample\方案";
            if (!System.IO.Directory.Exists(tempo_path))
            {
                toolStripStatusLabel1.ForeColor = System.Drawing.Color.Red;
                toolStripStatusLabel1.Text = "获取模板 失败，检查桌面的Sample\\出报告申请_模板.docx";
            }
            else
            {
                toolStripStatusLabel1.Text = "获取模板 成功";

                string[] fangan_list = Directory.GetFiles(tempo_path,"*.docx");
                listBox1.Items.Clear();
                foreach (var item in fangan_list)
                {
                    listBox1.Items.Add(item.Replace(tempo_path+"\\", ""));
                }
            }
        }
        #endregion

        /// <summary>
        /// 刷新按钮
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button1_Click(object sender, EventArgs e)
        {
            string tempo_path = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + @$"\Sample\方案";

            string[] fangan_list = Directory.GetFiles(tempo_path, "*.docx");
            listBox1.Items.Clear();
            foreach (var item in fangan_list)
            {
                listBox1.Items.Add(item.Replace(tempo_path + "\\", ""));
            }
        }



        #region 通过报告制作（自动）

        private void button2_Click(object sender, EventArgs e)
        {
            if (listBox1.SelectedItem == null)
            {
                toolStripStatusLabel1.Text = "请先选择一个模板文件。";
                return;
            }
            loadSample();
            // 设定预设临时变量
            Xceed.Document.NET.Cell cell;
            Xceed.Document.NET.Table table;
            List<int> ilist = new List<int>();
            List<Xceed.Document.NET.Table> tlist;
            string a;
            string tmpstr;
            this.Hide();

            ConsoleWriter.WriteYEllow("项目编号↓");
            try
            {
                tmpstr = doc.Find_Paragraph_for_text("本报告记录编号：");
                tmpstr = myutil.get_string_after(tmpstr, "本报告记录编号：", "P202107109".Length);
                str_P号 = tmpstr;

            }
            catch (Exception)
            {

                try
                {
                    tmpstr = doc.Find_Paragraph_for_text("本报告记录号：");
                    tmpstr = myutil.get_string_after(tmpstr, "本报告记录号：", "P202107109".Length);
                    str_P号 = tmpstr;
                }
                catch (Exception)
                {

                    ConsoleWriter.WriteCyan("无法自动获取项目编号（P2022XXXXX)，请手动输入:");
                    tmpstr = Console.ReadLine();
                }
            }
            Console.WriteLine(tmpstr);
            tempo._replacePatterns.Add("P2021xxxxx", tmpstr);//后面页数
            tempo._replacePatterns.Add("P2021XXXXX", tmpstr);//首页

            ConsoleWriter.WriteYEllow("公司名↓");
            a = doc.table_Get_cell_text(doc.tables[2], 1, 1);
            str_公司 = a;
            Console.WriteLine(a);
            tmpstr = a;
            tempo._replacePatterns.Add("AAAAA", tmpstr);

            //tempo.cw_read_dictionary("AAAAA");



            ConsoleWriter.WriteYEllow("系统名称↓");
            a = doc.table_Get_cell_text(doc.tables[1], 1, 1);
            str_系统 = a;
            Console.WriteLine(a);

            tempo._replacePatterns.Add("BBBBB", a);
            //tempo.cw_read_dictionary("BBBBB");

            ConsoleWriter.WriteYEllow("封面日期 2021年XX月XX日:");
            cell = tempo.table_index_Get_cell(0, 3, 1);
            a = doc.Find_Paragraph_for_text("方案编制过程。");
            //有的测评过程写的时间是左右都有年份的，有的不是。
            if (a.Length == "1、2021年10月08日～2021年10月09日，测评准备过程。".Length)
            {
                a = a.Substring("2、20xx年xx月xx日～".Length, "20xx年xx月xx日".Length);
            }
            else
            {
                ConsoleWriter.WriteGray("gantaniangde");
                string a1;
                string a2;
                a1 = a.Substring("2、".Length, "20xx年".Length);
                a2 = a.Substring("2、20xx年xx月xx日～".Length, "xx月xx日".Length);
                a = a1 + a2;
            }
            Console.WriteLine(a);
            tempo.cell_settext_Big(cell, a);


            ConsoleWriter.WriteYEllow("方案作者:");
            cell = tempo.table_index_Get_cell(1, 1, 1);
            tmpstr = "黄瀚国";
            if (textBox_zuozhe.Text == "")
            {
                tmpstr = Console.ReadLine();
            }
            else
            {
                tmpstr = textBox_zuozhe.Text;
            }
            Console.WriteLine(tmpstr);
            tempo.cell_settext(cell, tmpstr);


            ConsoleWriter.WriteYEllow("创作日期 2021-XX-XX:");
            cell = tempo.table_index_Get_cell(1, 1, 3);
            tmpstr = "2021-12-31";
            a = doc.Find_Paragraph_for_text("方案编制过程。");
            //有的测评过程写的时间是左右都有年份的，有的不是。
            if (a.Length == "1、2021年10月08日～2021年10月09日，方案编制过程。".Length)
            {
                a = a.Substring("2、".Length, "20xx年xx月xx日".Length);
            }
            else
            {
                string a1;
                string a2;
                a1 = a.Substring("2、".Length, "20xx年".Length);
                a2 = a.Substring("2、20xx年".Length, "xx月xx日".Length);
                a = a1 + a2;
            }
            tmpstr = a.Replace("年", "-").Replace("日", "").Replace("月", "-");
            Console.WriteLine(tmpstr);
            tempo.cell_settext(cell, tmpstr);


            ConsoleWriter.WriteYEllow("审核日期 2021-XX-XX:");
            a = doc.Find_Paragraph_for_text("方案编制过程。");
            //有的测评过程写的时间是左右都有年份的，有的不是。
            if (a.Length == "1、2021年10月08日～2021年10月09日，方案编制过程。".Length)
            {
                a = a.Substring("2、20xx年xx月xx日～".Length, "20xx年xx月xx日".Length);
            }
            else
            {
                string a1;
                string a2;
                a1 = a.Substring("2、".Length, "20xx年".Length);
                a2 = a.Substring("2、20xx年xx月xx日～".Length, "xx月xx日".Length);
                a = a1 + a2;
            }
            tmpstr = a.Replace("年", "-").Replace("日", "").Replace("月", "-");
            Console.WriteLine(tmpstr);
            cell = tempo.table_index_Get_cell(2, 1, 3);
            tempo.cell_settext(cell, tmpstr);


            ConsoleWriter.WriteYEllow("审核人:");
            cell = tempo.table_index_Get_cell(2, 1, 1);
            tmpstr = "陈家琦";
            if (textBox_pingshen.Text == "")
            {
                tmpstr = Console.ReadLine();
            }
            else
            {
                tmpstr = textBox_pingshen.Text;
                Console.WriteLine(tmpstr);
            }
            tempo.cell_settext(cell, tmpstr);


            ConsoleWriter.WriteYEllow("系统责任描述↓");
            int i = doc.Find_Paragraph_for_ilist("总体评价")[0] + 1;
            a = doc.document.Paragraphs[i].Text;
            Console.WriteLine(a);
            if (a == "")
            {
                ConsoleWriter.WriteErrorMessage("注意，未找到系统责任描述↓，请手动填写：");
                a = Console.ReadLine();
            }
            tempo._replacePatterns.Add("被测对象情况描述（必须包括被测对象责任主体、业务描述、网络拓扑描述）。", a);


            ConsoleWriter.WriteYEllow("系统业务描述↓");
            i = doc.Find_Paragraph_for_ilist("业务和采用的技术")[1] + 1;
            a = doc.document.Paragraphs[i].Text;

            Console.WriteLine(a);
            
            if (a == "")
            {
                ConsoleWriter.WriteErrorMessage("注意，未找到系统责任描述↓，请手动填写：");
                a = Console.ReadLine();
            }

            tempo._replacePatterns.Add("系统提供的服务介绍；系统存储的主要业务数据。", a);

            ConsoleWriter.WriteYEllow("系统功能截图↓(未实现)");
            //暂时还是先算了

            ConsoleWriter.WriteYEllow("网络架构的说明↓");
            i = doc.Find_Paragraph_for_ilist("网络结构")[1] + 1;
            a = doc.document.Paragraphs[i].Text;
            Console.WriteLine(a);
            ConsoleWriter.WriteCyan("请确认一下描述是否正确：(Enter确定，n 更换)");
            int j = 1;
            while (Console.ReadLine() == "n")
            {
                j++;
                i = doc.Find_Paragraph_for_ilist("网络结构")[j] + 1;
                a = doc.document.Paragraphs[i].Text;
                Console.WriteLine(a+ "(Enter确定，n 更换)");
                
            }
            tempo._replacePatterns.Add("网络架构图说明。", a);
            tempo.remove_p(tempo.Find_Paragraph_for_p("至少说明边界、区域划分、主要的设备等。"));


            ConsoleWriter.WriteYEllow("测评首日↓");
            string day = "2021.11.20";
            a = doc.Find_Paragraph_for_text("现场实施过程。");
            a = a.Substring("3、".Length, "20xx年xx月xx日".Length);
            Console.WriteLine("第一天："+a);
            day = a;
            DateTime day1 = DateTime.Parse(day);
            DateTime day2 = addOne_not_weekend(day1);
            DateTime day3 = addOne_not_weekend(day2);
            DateTime day4 = addOne_not_weekend(day3);
            DateTime day5 = addOne_not_weekend(day4);
            DateTime day6 = addOne_not_weekend(day5);

            tlist = tempo.findTableList("序号	日期	时间	测评内容	测评方法	配合人员");
            table = tlist[0];
            cell = tempo.table_Get_cell(table, 1, 1);
            tempo.cell_settext(cell, day1.ToString("yyyy.MM.dd"));

            cell = tempo.table_Get_cell(table, 4, 1);
            tempo.cell_settext(cell, day2.ToString("yyyy.MM.dd"));

            cell = tempo.table_Get_cell(table, 7, 1);
            tempo.cell_settext(cell, day3.ToString("yyyy.MM.dd"));

            cell = tempo.table_Get_cell(table, 9, 1);
            tempo.cell_settext(cell, day4.ToString("yyyy.MM.dd"));

            cell = tempo.table_Get_cell(table, 11, 1);
            tempo.cell_settext(cell, day5.ToString("yyyy.MM.dd"));

            cell = tempo.table_Get_cell(table, 13, 1);
            tempo.cell_settext(cell, day6.ToString("yyyy.MM.dd"));






            replace_tables_with_doc();
            //替换所有字典中的文字
            tempo.ReplaceTextWithText_all_noBracket();


            toolStripStatusLabel1.Text = "已经自动保存到桌面-out文件夹";
            //tempo.save($"{str_P号}_{str_公司}_{str_系统}_项目方案.docx");
            tempo.save($"{str_P号}_GF01_项目方案_{str_系统}.docx");

            ConsoleWriter.WriteYEllow(@"
已经自动保存到桌面-out文件夹
还有以下事项需要手工操作：
1. 刷新目录
2. 业务截图、拓扑图、工具测试图
3. 是否涉及物理环境？删除相关字。
4. 
");

            this.Show();
        }

        #endregion

        #region 测试输出按钮
        private void button3_Click(object sender, EventArgs e)
        {
            if (listBox1.SelectedItem == null)
            {
                toolStripStatusLabel1.Text = "请先选择一个模板文件。";
                return;
            }
            loadSample();
            // 设定预设临时变量
            Xceed.Document.NET.Cell cell;
            Xceed.Document.NET.Table table;
            List<int> ilist = new List<int>();
            List<Xceed.Document.NET.Table> tlist;
            string tmpstr;
            this.Hide();

            ConsoleWriter.WriteYEllow("项目编号↓");
            tmpstr = "P202199999";
            tempo._replacePatterns.Add("P2021xxxxx", tmpstr);//后面页数
            tempo._replacePatterns.Add("P2021XXXXX", tmpstr);//首页

            ConsoleWriter.WriteYEllow("公司名↓");
            tmpstr = "金元顺安基金管理有限公司";
            tempo._replacePatterns.Add("AAAAA", tmpstr);

            //tempo.cw_read_dictionary("AAAAA");



            ConsoleWriter.WriteYEllow("系统名称↓");
            tmpstr = "投资管理系统";
            tempo._replacePatterns.Add("BBBBB", tmpstr);
            //tempo.cw_read_dictionary("BBBBB");

            ConsoleWriter.WriteYEllow("封面日期 2021年XX月XX日:");
            cell = tempo.table_index_Get_cell(0, 3, 1);
            tmpstr = "2021年12月31日";
            tempo.cell_settext_Big(cell, tmpstr);


            ConsoleWriter.WriteYEllow("方案作者:");
            cell = tempo.table_index_Get_cell(1, 1, 1);
            tmpstr = "黄瀚国";
            tempo.cell_settext(cell, tmpstr);


            ConsoleWriter.WriteYEllow("创作日期 2021-XX-XX:");
            cell = tempo.table_index_Get_cell(1, 1, 3);
            tmpstr = "2021-12-31";
            tempo.cell_settext(cell, tmpstr);


            ConsoleWriter.WriteYEllow("审核日期 2021-XX-XX:");
            cell = tempo.table_index_Get_cell(2, 1, 3);
            tmpstr = "2021-12-31";
            tempo.cell_settext(cell, tmpstr);


            ConsoleWriter.WriteYEllow("审核人:");
            cell = tempo.table_index_Get_cell(2, 1, 1);
            tmpstr = "陈家琦";
            tempo.cell_settext(cell, tmpstr);

            ConsoleWriter.WriteYEllow("系统责任描述↓");
            tmpstr = "责任情况的描述";
            tempo._replacePatterns.Add("被测对象情况描述（必须包括被测对象责任主体、业务描述、网络拓扑描述）。", tmpstr);
            

            ConsoleWriter.WriteYEllow("系统业务描述↓");
            tmpstr = "业务情况的描述";
            tempo._replacePatterns.Add("系统提供的服务介绍；系统存储的主要业务数据。", tmpstr);

            ConsoleWriter.WriteYEllow("系统功能截图↓");
            //暂时还是先算了

            ConsoleWriter.WriteYEllow("系统业务描述↓");
            tmpstr = "网络架构的说明";
            tempo._replacePatterns.Add("网络架构图说明。", tmpstr);
            tempo.remove_p(tempo.Find_Paragraph_for_p("至少说明边界、区域划分、主要的设备等。"));

            var p = tempo.Find_Paragraph_for_p("安全管理文档");
            ConsoleWriter.WriteYEllow(p.Text);

            ConsoleWriter.WriteYEllow("测评首日↓");
            string day = "2021.11.20";
            DateTime day1 = DateTime.Parse(day);
            DateTime day2 = addOne_not_weekend(day1);
            DateTime day3 = addOne_not_weekend(day2);
            DateTime day4 = addOne_not_weekend(day3);
            DateTime day5 = addOne_not_weekend(day4);
            DateTime day6 = addOne_not_weekend(day5);

            tlist = tempo.findTableList("序号	日期	时间	测评内容	测评方法	配合人员");
            table = tlist[0];
            cell = tempo.table_Get_cell(table, 1, 1);
            tempo.cell_settext(cell, day1.ToString("yyyy.MM.dd"));

            cell = tempo.table_Get_cell(table, 4, 1);
            tempo.cell_settext(cell, day2.ToString("yyyy.MM.dd"));

            cell = tempo.table_Get_cell(table, 7, 1);
            tempo.cell_settext(cell, day3.ToString("yyyy.MM.dd"));

            cell = tempo.table_Get_cell(table, 9, 1);
            tempo.cell_settext(cell, day4.ToString("yyyy.MM.dd"));

            cell = tempo.table_Get_cell(table, 11, 1);
            tempo.cell_settext(cell, day5.ToString("yyyy.MM.dd"));

            cell = tempo.table_Get_cell(table, 13, 1);
            tempo.cell_settext(cell, day6.ToString("yyyy.MM.dd"));


            replace_tables_with_doc();

            //替换所有字典中的文字
            tempo.ReplaceTextWithText_all_noBracket();


            toolStripStatusLabel1.Text = "已经自动保存到桌面-out文件夹";
            tempo.save($"{str_公司}_{str_系统}_方案.docx");
            tempo.save($"{str_P号}_GF01_测评方案_{str_系统}_方案.docx");
            this.Show();

            
        }
        #endregion

        #region 图片展示按钮
        private void button4_Click(object sender, EventArgs e)
        {
            if (textBox1.Text == "")
            {
                toolStripStatusLabel1.Text = "请先加载报告";
                return;
            }
            var form_a = new form_loading(doc.Bitmaplist);
            form_a.Show();

        }
        #endregion

        #region 归档信息快查

        private void button7_Click(object sender, EventArgs e)
        {
            this.Hide();
            ConsoleWriter.WriteSeperator('^');
            var table1 = doc.findTableList("被测对象")[0];
            ConsoleWriter.WriteGray("被测对象名称:");
            Cell cell = table1.Rows[1].Cells[1];
            Console.WriteLine(doc.cell_get_text(cell));
            ConsoleWriter.WriteGray("安全保护等级:");
            cell = table1.Rows[1].Cells[3];
            Console.WriteLine(doc.cell_get_text(cell));

            table1 = doc.findTableList("被测单位")[0];

            ConsoleWriter.WriteGray("被测单位:");
            cell = table1.Rows[1].Cells[1];
            Console.WriteLine(doc.cell_get_text(cell));
            ConsoleWriter.WriteGray("联系人姓名（甲方）:");
            cell = table1.Rows[3].Cells[2];
            Console.WriteLine(doc.cell_get_text(cell));

             cell = doc.table_index_Get_cell(0, 2, 1);
            ConsoleWriter.WriteGray("报告日期：");
            Console.WriteLine(doc.cell_get_text(cell));

            ConsoleWriter.WriteGray("出任务时间：");
            int k = doc.Find_Paragraph_for_ilist("本次等级测评分为四个过程")[0];
            for (int i = k; i < k+5; i++)
            {
                Console.WriteLine(doc.Paragraphs[i].Text);
            }


            ConsoleWriter.WriteGray("测评设备：");
            try
            {
                int i;
                i = doc.Find_Paragraph_for_ilist("本次验证测试使用以下工具：")[0];
                ConsoleWriter.WriteGray("可能使用的扫描工具：");
                ConsoleWriter.WriteYEllow(doc.document.Paragraphs[i-1].Text);

                ConsoleWriter.WriteCyan(doc.document.Paragraphs[i + 1].Text);
                ConsoleWriter.WriteCyan(doc.document.Paragraphs[i + 2].Text);
                ConsoleWriter.WriteCyan(doc.document.Paragraphs[i + 3].Text);
                ConsoleWriter.WriteCyan(doc.document.Paragraphs[i + 4].Text);
            }
            catch (Exception)
            {
                ConsoleWriter.WriteGray("无法获得扫描的提示信息");
            }
            try
            {
                int i;

                i = doc.Find_Paragraph_for_ilist("附录F")[1];
                ConsoleWriter.WriteGray("渗透测试的参考信息：");
                ConsoleWriter.WriteCyan(doc.document.Paragraphs[i + 1].Text);
            }
            catch (Exception)
            {
                ConsoleWriter.WriteGray("无法获得渗透测试的提示信息");
            }

            Console.WriteLine("按Enter继续……");
            Console.ReadLine();
            this.Show();
        }
        static void Thread1(List<Bitmap> bitmaps)
        {
            //展示bitmap窗体
            Application.Run(new form_loading(bitmaps));

        }
        #endregion

        #region 命令行制作

        private void button5_Click(object sender, EventArgs e)
        // 手动制作
        {
            if (listBox1.SelectedItem == null)
            {
                toolStripStatusLabel1.Text = "请先选择一个模板文件。";
                return;
            }
            loadSample();
            // 设定预设临时变量
            Xceed.Document.NET.Cell cell;
            Xceed.Document.NET.Table table;
            List<int> ilist = new List<int>();
            List<Xceed.Document.NET.Table> tlist;
            string tmpstr;
            this.Hide();

            ConsoleWriter.WriteYEllow("项目编号↓");
            tmpstr = Console.ReadLine();
            str_P号 = tmpstr;
            tempo._replacePatterns.Add("P2021xxxxx", tmpstr);//后面页数
            tempo._replacePatterns.Add("P2021XXXXX", tmpstr);//首页

            ConsoleWriter.WriteYEllow("公司名↓");
            tmpstr = Console.ReadLine();
            str_公司 = tmpstr;
            tempo._replacePatterns.Add("AAAAA", tmpstr);

            //tempo.cw_read_dictionary("AAAAA");



            ConsoleWriter.WriteYEllow("系统名称↓");
            tmpstr = Console.ReadLine();
            str_系统 = tmpstr;
            tempo._replacePatterns.Add("BBBBB", tmpstr);
            //tempo.cw_read_dictionary("BBBBB");

            ConsoleWriter.WriteYEllow("封面日期 2021年XX月XX日:");
            cell = tempo.table_index_Get_cell(0, 3, 1);
            tmpstr = "2021年12月31日";
            tmpstr = Console.ReadLine();
            tempo.cell_settext_Big(cell, tmpstr);


            ConsoleWriter.WriteYEllow("方案作者:");
            cell = tempo.table_index_Get_cell(1, 1, 1);
            tmpstr = "黄瀚国";
            tmpstr = Console.ReadLine();
            tempo.cell_settext(cell, tmpstr);


            ConsoleWriter.WriteYEllow("创作日期 2021-XX-XX:");
            cell = tempo.table_index_Get_cell(1, 1, 3);
            tmpstr = "2021-12-31";
            tmpstr = Console.ReadLine();
            tmpstr = tmpstr.Replace("年", "-").Replace("日", "").Replace("月", "-");
            tempo.cell_settext(cell, tmpstr);


            ConsoleWriter.WriteYEllow("审核日期 2021-XX-XX:");
            cell = tempo.table_index_Get_cell(2, 1, 3);
            ConsoleWriter.WriteYEllow("自动使用创作日期同一天");
            tempo.cell_settext(cell, tmpstr);


            ConsoleWriter.WriteYEllow("审核人:");
            cell = tempo.table_index_Get_cell(2, 1, 1);
            tmpstr = "陈家琦";
            tmpstr = Console.ReadLine();
            tempo.cell_settext(cell, tmpstr);

            ConsoleWriter.WriteYEllow("系统责任描述↓");
            tmpstr = "责任情况的描述";
            tmpstr = Console.ReadLine();
            tempo._replacePatterns.Add("被测对象情况描述（必须包括被测对象责任主体、业务描述、网络拓扑描述）。", tmpstr);


            ConsoleWriter.WriteYEllow("系统业务描述↓");
            tmpstr = "业务情况的描述";
            tmpstr = Console.ReadLine();
            tempo._replacePatterns.Add("系统提供的服务介绍；系统存储的主要业务数据。", tmpstr);

            ConsoleWriter.WriteYEllow("系统功能截图↓(未实现)");
            //暂时还是先算了

            ConsoleWriter.WriteYEllow("网络架构的说明↓");
            tmpstr = "网络架构的说明";
            tmpstr = Console.ReadLine();
            tempo._replacePatterns.Add("网络架构图说明。", tmpstr);
            tempo.remove_p(tempo.Find_Paragraph_for_p("至少说明边界、区域划分、主要的设备等。"));

            //var p = tempo.Find_Paragraph_for_p("安全管理文档");
            //ConsoleWriter.WriteYEllow(p.Text);

            ConsoleWriter.WriteYEllow("测评首日↓");
            string day = "2021.11.20";
            tmpstr = Console.ReadLine();
            DateTime day1 = DateTime.Parse(day);
            DateTime day2 = addOne_not_weekend(day1);
            DateTime day3 = addOne_not_weekend(day2);
            DateTime day4 = addOne_not_weekend(day3);
            DateTime day5 = addOne_not_weekend(day4);
            DateTime day6 = addOne_not_weekend(day5);

            tlist = tempo.findTableList("序号	日期	时间	测评内容	测评方法	配合人员");
            table = tlist[0];
            cell = tempo.table_Get_cell(table, 1, 1);
            tempo.cell_settext(cell, day1.ToString("yyyy.MM.dd"));

            cell = tempo.table_Get_cell(table, 4, 1);
            tempo.cell_settext(cell, day2.ToString("yyyy.MM.dd"));

            cell = tempo.table_Get_cell(table, 7, 1);
            tempo.cell_settext(cell, day3.ToString("yyyy.MM.dd"));

            cell = tempo.table_Get_cell(table, 9, 1);
            tempo.cell_settext(cell, day4.ToString("yyyy.MM.dd"));

            cell = tempo.table_Get_cell(table, 11, 1);
            tempo.cell_settext(cell, day5.ToString("yyyy.MM.dd"));

            cell = tempo.table_Get_cell(table, 13, 1);
            tempo.cell_settext(cell, day6.ToString("yyyy.MM.dd"));


            replace_tables_with_doc();
            //替换所有字典中的文字
            tempo.ReplaceTextWithText_all_noBracket();

            ConsoleWriter.WriteYEllow(@"
还有以下事项需要手工操作：
1. 刷新目录
2. 业务截图、拓扑图、工具测试图
3. 是否涉及物理环境？删除相关字。
4. 
");

            toolStripStatusLabel1.Text = "已经自动保存到桌面-out文件夹";
            tempo.save($"{str_P号}{str_公司}_{str_系统}_测评方案.docx");
            this.Show();
        }
        #endregion


        #region 复制table\选择工具

        void replace_tables_with_doc()
        {
            try
            {
                int i;
                i = doc.Find_Paragraph_for_ilist("本次验证测试使用以下工具：")[0];
                ConsoleWriter.WriteGray("可能使用的扫描工具：");
                ConsoleWriter.WriteQuestionMessage(doc.document.Paragraphs[i + 1].Text);
                ConsoleWriter.WriteQuestionMessage(doc.document.Paragraphs[i + 2].Text);
                ConsoleWriter.WriteQuestionMessage(doc.document.Paragraphs[i + 3].Text);
                ConsoleWriter.WriteQuestionMessage(doc.document.Paragraphs[i + 4].Text);


            }
            catch (Exception)
            {

                ConsoleWriter.WriteGray("无法获得扫描的提示信息");

            }
            try
            {
                int i;

                i = doc.Find_Paragraph_for_ilist("附录F")[1];
                ConsoleWriter.WriteGray("渗透测试的参考信息：");
                ConsoleWriter.WriteQuestionMessage(doc.document.Paragraphs[i + 1].Text);
            }
            catch (Exception)
            {

                ConsoleWriter.WriteGray("无法获得渗透测试的提示信息");
            }


            //删除多余的工具行
            ConsoleWriter.WriteColoredText("↓主要测评工具 - 选择需要留下的工具序号(空格分隔)：", ConsoleColor.Green);
            Delete_selectRow("序号	工具名称	厂商	系统版本	漏洞库版本");
            ConsoleWriter.WriteColoredText("↓工具测试 - 选择需要留下的工具序号(空格分隔)：", ConsoleColor.Green);
            Delete_selectRow("序号 工具名称    测评对象 接入点 说明");

            //所有设备
            ConsoleWriter.WriteGray("copy 表格 A.机房 →系统构成机房 ");
            CopyTable("序号	机房名称	物理位置	重要程度	备注", "序号	机房名称	物理位置	重要程度");

            ConsoleWriter.WriteGray("copy 表格 A.网络设备 →系统构成 网络设备 ");
            CopyTable("序号 设备名称    是否虚拟设备 系统及版本   品牌及型号 用途  重要程度  备注",
                "序号	设备名称	虚拟设备	系统及版本	品牌型号	用途	重要程度	备注");

            ConsoleWriter.WriteGray("copy 表格 A.安全设备 →系统构成 安全设备 ");
            CopyTable(
                "序号	设备名称	是否虚拟设备	系统及版本	品牌及型号	用途	重要程度	备注",
                "序号	设备名称	虚拟设备	系统及版本 品牌型号	用途	重要程度	备注", 1, 1);
            if (tryoldVer)
            {
                ConsoleWriter.WriteQuestionMessage("自动尝试复制旧版表格");
                ConsoleWriter.WriteGray("copy 表格 A.安全设备 →系统构成 安全设备 （旧版）");
                CopyTable(
                "序号	设备名称	是否虚拟设备	系统及版本	品牌及型号	用途	重要程度	备注",
                "序号	设备名称	虚拟设备	系统及版本 品牌型号	用途	重要程度	备注", 0, 1);
                tryoldVer = false;
            }
            ConsoleWriter.WriteGray("copy 表格 A.服务器 →系统构成 服务器 ");
            CopyTable(
                "序号	设备名称	所属业务应用系统/平台名称	是否虚拟设备	操作系统及版本	数据库管理系统及版本	中间件及版本	重要程度	备注",
                "序号	设备名称	所属业务应用系统/平台名称	虚拟设备	操作系统及版本	数据库管理系统及版本	中间件及版本	重要程度	备注"
                );
            if (tryoldVer)
            {
                ConsoleWriter.WriteQuestionMessage("自动尝试复制旧版表格");
                ConsoleWriter.WriteGray("copy 表格 A.服务器 →系统构成 服务器 （旧版）");
                CopyTable(
                    "序号	设备名称	所属业务应用系统/平台	虚拟设备  操作系统及版本 数据库管理系统及版本  中间件及版本  重要程度",
                    "序号	设备名称	所属业务应用系统/平台名称	虚拟设备	操作系统及版本	数据库管理系统及版本	中间件及版本	重要程度	备注"
                    );
                tryoldVer = false;
            }
            ConsoleWriter.WriteGray("copy 表格 A.5终端设备 →系统构成 终端设备 ");
            CopyTable(
                "序号	设备名称	是否虚拟设备	操作系统/控制软件及版本	用途	重要程度	备注",
                "序号	设备名称	虚拟设备	操作系统/控制软件及版本	设备类别/用途	重要程度	备注"
                );
            //注意，模板这里的表格头部顺序和报告中不一样。
            ConsoleWriter.WriteGray("copy 表格 A.6	系统管理软件/平台 →系统构成 终端设备 ");
            CopyTable_withHead(
                "序号	系统管理软件/平台名称	主要功能	版本	所在设备名称	重要程度	备注",
                "序号	管理软件/平台名称	所在设备名称	版本	主要功能	重要程度	备注"
                );
            if (tryoldVer)
            {
                ConsoleWriter.WriteQuestionMessage("自动尝试复制旧版表格");
                ConsoleWriter.WriteGray("copy 表格 A.6	系统管理软件/平台 →系统构成 终端设备 (旧版) ");
                CopyTable_withHead(
                    "序号	设备名称	虚拟设备  系统及版本   设备类别 / 用途 重要程度",
                    "序号	管理软件/平台名称	所在设备名称	版本	主要功能	重要程度	备注"
                    );
                tryoldVer = false;
            }
            ConsoleWriter.WriteGray("copy 表格 A.7	业务应用系统软件/平台 →系统构成 终端设备 ");
            CopyTable(
                "序号	业务应用系统/平台名称	主要功能	业务应用软件及版本	开发厂商	重要程度	备注",
                "序号	业务应用系统/平台名称	主要功能	业务应用软件及版本	开发厂商	重要程度"
                );
            if (tryoldVer)
            {
                ConsoleWriter.WriteQuestionMessage("自动尝试复制旧版表格");
                ConsoleWriter.WriteGray("copy 表格 A.7	业务应用系统软件/平台 →系统构成 终端设备（旧版） ");
                CopyTable(
                    "序号	业务应用系统/平台名称	主要功能	业务应用软件及版本	开发厂商	重要程度",
                    "序号	业务应用系统/平台名称	主要功能	业务应用软件及版本	开发厂商	重要程度"
                    );
                tryoldVer = false;
            }
            ConsoleWriter.WriteGray("copy 表格 A.8 数据资源 →系统构成  数据资源");
            CopyTable(
                "序号	数据类别	所属业务应用	安全防护需求	重要程度",
                "序号	数据类别	所属业务应用	安全防护需求"
                , 1, 0);

            ConsoleWriter.WriteGray("copy 表格 A.10 安全相关人员 →系统构成  安全相关人员");
            CopyTable(
                "序号	姓名	岗位/角色	联系方式	所属单位",
                "序号	姓名	岗位/角色	联系方式"
                , 1, 0);
            ConsoleWriter.WriteGray("copy 表格 A.12 安全管理文档 →系统构成  管理文档");
            CopyTable(
                "序号	文档名称	主要内容",
                "序号	文档名称	主要内容"
                , 1, 0);

            //选中被测评的设备
            ConsoleWriter.WriteGray("Copy表格 2.3.2.选择结果机房 →选择结果 机房 ");
            CopyTable("序号	机房名称	物理位置	重要程度",
                "序号	机房名称	物理位置	重要程度",
                0, 1);

            ConsoleWriter.WriteGray("Copy表格 2.3.2.选择结果网路设备 →选择结果 网路设备 ");
            CopyTable_withHead("序号	设备名称	虚拟设备  系统及版本   品牌及型号   用途  重要程度",
                "序号	设备名称	虚拟设备	系统及版本	品牌型号  用途  重要程度  备注",
                0, 2);

            ConsoleWriter.WriteGray("Copy表格 2.3.2.选择结果安全设备 →选择结果 安全设备 ");
            CopyTable_withHead("序号	设备名称	虚拟设备  系统及版本   品牌及型号   用途  重要程度",
                "序号	设备名称	虚拟设备  系统及版本 品牌型号  用途  重要程度",
                0, 0);

            //ConsoleWriter.WriteGray("Copy表格 2.3.2.选择结果 →选择结果 服务器 ");
            //CopyTable_withHead("序号	设备名称	所属业务应用系统/平台	虚拟设备  操作系统及版本 数据库管理系统及版本  中间件及版本  重要程度",
            //    "序号	设备名称	所属业务应用系统/平台名称	虚拟设备	操作系统及版本 数据库管理系统及版本  中间件及版本  重要程度",
            //    0, 0);
            try
            {
                ConsoleWriter.WriteGray("Copy表格 2.3.2.选择结果 →选择结果 服务器 ");
                ConsoleWriter.WriteQuestionMessage("尝试自动识别表格数量");
                var tlist = tempo.findTableList("序号	设备名称	所属业务应用系统/平台名称	虚拟设备	操作系统及版本 数据库管理系统及版本  中间件及版本  重要程度");
                CopyTable_withHead("序号	设备名称	所属业务应用系统/平台	虚拟设备  操作系统及版本 数据库管理系统及版本  中间件及版本  重要程度",
                    "序号	设备名称	所属业务应用系统/平台名称	虚拟设备	操作系统及版本 数据库管理系统及版本  中间件及版本  重要程度",
                    0, tlist.Count - 1);
                tryoldVer = false;
                //ConsoleWriter.WriteCyan("是否粘贴服务器表格？（enter确定，n跳过该表）");
                //var rd = Console.ReadLine();
                //if (rd == "\n")
                //{
                //}
            }
            catch (Exception) { }
            

            ConsoleWriter.WriteGray("Copy表格 2.3.2.5 终端设备 →选择结果 终端设备 ");
            CopyTable_withHead("序号	设备名称	虚拟设备  操作系统及版本 用途  重要程度",
                "序号	设备名称	虚拟设备  操作系统 /控制软件及版本 设备类别 / 用途 重要程度",
                0, 0);


            ConsoleWriter.WriteGray("Copy表格 2.3.2 系统管理软件/平台 → 选择结果 系统管理软件/平台 ");
            CopyTable_withHead(
                "序号	系统管理软件/平台名称	主要功能	版本	所在设备名称	重要程度",
                "序号	管理软件/平台名称	所在设备名称	版本	主要功能	重要程度	备注",
                0, 0);

            ConsoleWriter.WriteGray("Copy表格 2.3.2	业务应用系统/平台 → 选择结果 业务应用软件/平台 ");
            CopyTable_withHead(
                "序号	业务应用系统/平台名称	主要功能	业务应用软件及版本	开发厂商	重要程度",
                "序号	业务应用系统/ 平台名称   主要功能    业务应用软件及版本   开发厂商    重要程度",
                0, 1);

            ConsoleWriter.WriteGray("Copy表格 2.3.2	数据资源 → 选择结果 数据资源 ");
            CopyTable(
                "序号	数据类别	所属业务应用	安全防护需求	重要程度",
                "序号	数据类别	所属业务应用	安全防护需求",
                1, 1);

            ConsoleWriter.WriteGray("Copy表格 2.3.2	人员 → 选择结果 人员 ");
            CopyTable(
                "序号	姓名	岗位/角色	联系方式	所属单位",
                "序号	姓名	岗位/角色	联系方式",
                1, 1);
            ConsoleWriter.WriteGray("Copy表格 2.3.2	安全管理文档 → 选择结果 安全管理文档 ");
            CopyTable(
                "序号	文档名称	主要内容",
                "序号	文档名称	主要内容",
                1, 1);
        }
        #endregion

        #region 表格复制函数

        /// <summary>
        /// 
        /// </summary>
        /// <param name="t1head">table1 表头</param>
        /// <param name="i1">table1 所在位数</param>
        /// <param name="t2head">table2 表头</param>
        /// <param name="i2">table 2 所在位数</param>
        void CopyTable(string t1head, string t2head, int i1 = 0, int i2 = 0)
        {
            try
            {
                bool toremove = false;
                var table1 = doc.findTableList(t1head)[i1];
                ConsoleWriter.WriteColoredText("table 报告中 ↑", ConsoleColor.Green);
                var table2 = tempo.findTableList(t2head)[i2];
                ConsoleWriter.WriteColoredText("table 模板中 ↑", ConsoleColor.Green);
                //如果t1比t2更宽，增加一列临时列
                if (table1.ColumnCount > table2.ColumnCount)
                {
                    table2.InsertColumn();
                    toremove = true;
                }
                //如果t1比t2更窄，直接给t2瘦身
                else if (table1.ColumnCount < table2.ColumnCount)
                {
                    table2.RemoveColumn(table2.ColumnCount - 1);
                }
                //删除所有空的内容行
                while (table2.RowCount > 1)
                {
                    table2.RemoveRow(table2.RowCount - 1);
                }
                //从内容行数开始复制
                for (int i = 1; i < table1.RowCount; i++)
                {
                    Xceed.Document.NET.Row row = table1.Rows[i];

                    table2.InsertRow(row);
                }
                //删除多复制过来的列
                if (toremove)
                {
                    table2.RemoveColumn(table2.ColumnCount - 1);
                }
                ConsoleWriter.WriteColoredText("复制表完毕;", ConsoleColor.Yellow);
            }
            catch (Exception)
            {
                ConsoleWriter.WriteErrorMessage("无法复制这张表，请手动复制！");
                tryoldVer = true;
            }


        }

        /// <summary>
        /// 复制表t1到表t2（包含表头）
        /// </summary>
        /// <param name="t1head">table1 表头</param>
        /// <param name="i1">table1 所在位数</param>
        /// <param name="t2head">table2 表头</param>
        /// <param name="i2">table 2 所在位数</param>
        void CopyTable_withHead(string t1head, string t2head, int i1 = 0, int i2 = 0)
        {
            try
            {
                bool toremove = false;
                var table1 = doc.findTableList(t1head)[i1];
                ConsoleWriter.WriteColoredText("table 报告中 ↑", ConsoleColor.Green);
                var table2 = tempo.findTableList(t2head)[i2];
                ConsoleWriter.WriteColoredText("table 模板中 ↑", ConsoleColor.Green);
                //如果t1比t2更宽，增加一列临时列
                if (table1.ColumnCount > table2.ColumnCount)
                {
                    table2.InsertColumn();
                    toremove = true;
                }
                //如果t1比t2更窄，直接给t2瘦身
                else if (table1.ColumnCount < table2.ColumnCount)
                {
                    table2.RemoveColumn(table2.ColumnCount - 1);
                }
                //删除所有空的内容行
                while (table2.RowCount > 1)
                {
                    table2.RemoveRow(table2.RowCount - 1);
                }
                //从内容行数开始复制
                for (int i = 0; i < table1.RowCount; i++)
                {
                    Xceed.Document.NET.Row row = table1.Rows[i];

                    table2.InsertRow(row);
                }
                //删除多复制过来的列
                if (toremove)
                {
                    table2.RemoveColumn(table2.ColumnCount - 1);
                }
                //删除顶部的原始行（表头总是有问题服了）
                table2.RemoveRow(0);
                ConsoleWriter.WriteColoredText("复制表完毕;", ConsoleColor.Yellow);
            }
            catch (Exception)
            {
                ConsoleWriter.WriteErrorMessage("无法复制这张表，请手动复制！");
                tryoldVer = true;
            }


        }

        #endregion

        #region 删除表格行数函数

        /// <summary>
        /// 
        /// </summary>
        /// <param name="t1head"></param>
        /// <param name="i1"></param>
        void Delete_selectRow(string t1head,int i1=0)
        {
            var table = tempo.findTableList(t1head)[i1];
            List<int> ilist = new List<int>();
            //输出所有的可选项
            for (int i = 1; i < table.RowCount; i++)
            {
                var cell = tempo.table_Get_cell(table, i, 1);
                ConsoleWriter.WriteYEllow(i + " " + tempo.cell_get_text(cell));
                ilist.Add(i);
            }
            //读取用户的选择（保留项）
            string v = Console.ReadLine();
            if (v == "")
            {
                ConsoleWriter.WriteGray("本次测评未涉及工具测试。\n 需要手动修改 4.2 主要测评工具、5.3 工具测试的内容");
                return;
            }

            var vlist = v.Split(" ");
            //移除要保留的项目
            foreach (var i in vlist)
            {
                ilist.Remove(Convert.ToInt32(i));
            }
            //倒序删除所有不需要的row
            ilist.Reverse();
            foreach (var i in ilist)
            {
                table.RemoveRow(i);
            }
        }
        #endregion

        #region 日期顺延函数
        /// <summary>
        /// 获取下一天且下一天不在周六、周日中
        /// </summary>
        /// <param name="day"></param>
        /// <returns></returns>
        public static DateTime addOne_not_weekend(DateTime day)
        {
           day = day.AddDays(1);
           while(day.DayOfWeek == DayOfWeek.Saturday || day.DayOfWeek == DayOfWeek.Sunday)
           {
                day = day.AddDays(1);
           }
           return day;
        }

        #endregion

        #region 帮助按钮

        private void button6_Click(object sender, EventArgs e)
        {
            this.Hide();
            Console.WriteLine(@"
本工具可以通过测评报告（2021版）生成项目方案

0. 准备好Sample文件夹并放在桌面上
1. 把报告拖动到textbox中
2. 选择【自动制作】/【命令行制作】来生成报告
2.1 在命令行输入选择要保留的扫描工具
3. 生成之后在桌面out文件夹里可以找到生成的方案
4. 按照提示进行部分手动更改后结束

- [x] 项目编号
- [x] AAAAA 公司
- [x] BBBBB 项目
- [x] 作者
- [x] 发布日期 - 首页
- [x] 发布日期 - 表格
- [x] 审核日期 - 表格
- [x] 被测对象情况描述（必须包括被测对象责任主体、业务描述、网络拓扑描述）。
- [ ] 业务截图
- [x] 业务描述
- [ ] 拓扑图
- [x] 拓扑描述
- [x] 一堆表-1
- [x] 一堆表-2
- [x] 现场时间安排
- [ ] 接入点图、接入工具修改");
            ConsoleWriter.WriteSeperator('-');
        ConsoleWriter.WriteYEllow("按Enter键继续");
            Console.ReadLine();
            this.Show();
        }

        #endregion

        #region 自动填空按钮
        private void button8_Click(object sender, EventArgs e)
        {
            var tables = doc.document.Tables;
            foreach (var table in tables)
            {
                foreach(var row in table.Rows)
                {
                    bool jump = true;
                    foreach(var cell in row.Cells)
                    {
                        if (jump)//跳过第一列（序号列）
                        {
                            jump = false;

                            continue;
                        }
                        if (doc.cell_get_text(cell)=="")//如果没有杠，自动添加
                        {
                            doc.cell_settext(cell, "--");

                        }

                    }
                }
            }
            doc.save("LD18-11 信息系统基本情况调查表.docx");
            toolStripStatusLabel1.Text = "done!";
        }
        #endregion

        #region 测试按钮

        private void button9_Click(object sender, EventArgs e)
        {
            var t = doc.findTableList("测评结论和综合得分")[0];
            var cell = t.Rows[2].Cells[1];
            Console.WriteLine(doc.cell_get_text(cell));


        }
        #endregion



        private void button10_Click(object sender, EventArgs e)
        {


            Settings.Default.author_str = textBox_zuozhe.Text;
            Settings.Default.auditor_Str = textBox_pingshen.Text;

            toolStripStatusLabel1.Text = "已保存！下次会自动加载人名";
        }

        private void button11_Click(object sender, EventArgs e)
        {
            this.Hide();

            //如何可以，顺便显示下时间安排
            if (textBox1.Text != "")
            {
                ConsoleWriter.WriteGray("出任务时间：");
                int k = doc.Find_Paragraph_for_ilist("本次等级测评分为四个过程")[0];
                for (int i = k; i < k + 5; i++)
                {
                    Console.WriteLine(doc.Paragraphs[i].Text);
                }
            }

            ConsoleWriter.WriteSeperator('-');
            Console.WriteLine("模板时间提示：");
            ConsoleWriter.WriteYEllow(@"
[A3.0] 实施日期：2021.12.01
[A2.5] 实施日期：2021-09-27
[A2.4] 实施日期：2021-08-09
[A2.3] 实施日期：2021-04-30
[A2.2] 实施日期：2020-08-05
[A2.1] 实施日期：2020-08-01");

            toolStripStatusLabel1.Text = "提示已经显示在命令行";
            this.Show();
        }
    }

}