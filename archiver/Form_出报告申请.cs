using archiver.ConsoleColorWriter;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace archiver
{
    public partial class Form_出报告申请 : Form
    {
        myutil doc ;//等保报告
        myutil tempo;//出报告模板
        string str_公司 = "xxx公司";
        string str_系统 = "xxx系统";
        string str_P项目编号 = "Pxxxxxxxxx";

        string opendoc = "";
        public Form_出报告申请()
        {
            InitializeComponent();
            button1.Enabled = false;
        }
        //version v2.0 v3.0
        private void loadSample(string version)
        {
            string tempo_path = Environment.GetFolderPath(Environment.SpecialFolder.Desktop)+ @$"\Sample\出报告申请\出报告申请_模板{version}.docx";
            if (!File.Exists(tempo_path))
            {
                MessageBox.Show("dude,获取模板 失败");
                toolStripStatusLabel1.ForeColor = System.Drawing.Color.Red;
                toolStripStatusLabel1.Text = $"获取模板 失败" + tempo_path;
            }
            else
            {
                tempo = new myutil(tempo_path);
                toolStripStatusLabel1.Text = "获取模板 成功";
            }
        }


        #region textDrop
        private void textBox1_DragDrop(object sender, DragEventArgs e)
        {
            textBox1.Text = ((System.Array)e.Data.GetData(DataFormats.FileDrop)).GetValue(0).ToString();
            string path = textBox1.Text;

            doc = new myutil(path);

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
        #endregion


        //自动生成输出目录
        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            // 更换源文件，重新生成doc
            string path = textBox1.Text;
            doc = new myutil(path);
        }
        private void gp1_button1_Click(object sender, EventArgs e)
        {
            if (textBox1.Text == "")
            {
                toolStripStatusLabel1.Text = "先吧报告拖进来，然后点击生成…";
                return;
            }
            if (radioButton1.Checked)
            {
                loadSample("v3.0");
            }
            else
            {
                loadSample("v2.0");
            }
            string a = "";
            tempo._replacePatterns.Clear();
            a = doc.Paragraphs[0].Text;
            a = myutil.get_string_after(a, "报告编号：", "31001322060-19012-21-0092-01".Length);
            tempo.write_dictionary("报告编号", a);

            a = a.Substring(0, "31001322060-19012".Length);
            tempo.write_dictionary("备案号", a);

            a = doc.table_Get_cell_text(doc.tables[4], 4, 1);
            a = myutil.get_string_bewteen(a, "综合得分为", "分");
            tempo.write_dictionary("分数", a);

            a = doc.table_Get_cell_text(doc.tables[2], 1, 1);
            str_公司 = a;
            tempo.write_dictionary("xxx公司", a);

            a = doc.table_Get_cell_text(doc.tables[1], 1, 1);
            str_系统 = a;
            tempo.write_dictionary("xxx系统", a);

            a = doc.table_Get_cell_text(doc.tables[1], 1, 3);
            if (a.Contains("3")) a = "3";
            else if (a.Contains("2")) a = "2";
            tempo.write_dictionary("级别数字", a);

            this.Hide();
            a = doc.Find_Paragraph_for_text("本报告记录编号：");
            a = myutil.get_string_after(a,"本报告记录编号：", "P202107109".Length);
            tempo.write_dictionary("项目编号",a);
            str_P项目编号 = a;

            a = doc.Find_Paragraph_for_text("测评准备过程。");
            a = date_process(a,true);
            tempo.write_dictionary("起始日期",a);

            a = doc.Find_Paragraph_for_text("分析与报告编制过程。",2);
            a = date_process(a,false);
            tempo.write_dictionary("最终日期",a);
            a = doc.table_index_Get_cell_text(3, 8, 4);
            //a=  a.Substring(5);
            tempo.write_dictionary("出报告日期",a);

            Console.WriteLine("三防信息A(如果没有请直接打回车)");
            a = doc.Find_Paragraph_for_text("在防网页篡改方面");
            if (a == "") tempo.remove_p(tempo.Find_Paragraph_for_p("防网页篡改"));
            else
            {
                a = a.Substring("在防网页篡改方面".Length ) ;
                tempo.write_dictionary("防网页篡改", a);
            }

            Console.WriteLine("三防信息B.1");
            a = doc.Find_Paragraph_for_text("在防数据泄露方面");
            if (a == "") a = doc.Find_Paragraph_for_text("在防数据泄漏方面");
            if (a == "") tempo.remove_p(tempo.Find_Paragraph_for_p("防数据泄漏"));
            else
            {
                a = a.Substring("在防数据泄露方面：".Length );
                tempo.write_dictionary("防数据泄漏", a);
            }

            Console.WriteLine("三防信息B.2");
            a = doc.Find_Paragraph_for_text("在防数据勒索方面");
            if (a == "") tempo.remove_p(tempo.Find_Paragraph_for_p("防数据勒索"));
            else
            {
                a = a.Substring("在防数据勒索方面：".Length ) ;
                tempo.write_dictionary("防数据勒索", a);

            }

            Console.WriteLine("三防信息C");
            a = doc.Find_Paragraph_for_text("在防服务中断方面");
            if (a == "") tempo.remove_p(tempo.Find_Paragraph_for_p("防服务中断"));
            else
            {
                a = a.Substring("在防服务中断方面：".Length );
                tempo.write_dictionary("防服务中断", a);
            }


            ConsoleWriter.WriteCyan("还请输入：");
            tempo.cw_read_dictionary("我方人员");
            tempo.ReplaceTextWithText_all();
            ConsoleWriter.WriteCyan("已经自动保存到桌面-out文件夹，" +
                "\n 三防记得手动改成分号");
            this.Show();
            toolStripStatusLabel1.Text = "已经自动保存到桌面-out文件夹";

            opendoc = $"出报告申请_{str_公司}_{str_系统}.docx"; //用于后面打开文件查看
            button1.Enabled = true; //解锁打开查看的按钮

            tempo.save($"出报告申请_{str_公司}_{str_系统}.docx");
        }
        //手工录入
        private void gp2_button1_Click(object sender, EventArgs e)
        {
            this.Hide();
            tempo._replacePatterns.Clear();
            tempo.cw_read_dictionary("项目编号");
            tempo.cw_read_dictionary("备案号");
            tempo.cw_read_dictionary("报告编号");
            tempo.cw_read_dictionary("分数");
            str_公司 = tempo.cw_read_dictionary("xxx公司");
            str_系统 = tempo.cw_read_dictionary("xxx系统");
            tempo.cw_read_dictionary("级别数字");
            tempo.cw_read_dictionary("起始日期");
            tempo.cw_read_dictionary("出报告日期");
            tempo.cw_read_dictionary("我方人员");
            Console.WriteLine("三防：");
            tempo.cw_read_dictionary("防网页篡改");
            tempo.cw_read_dictionary("防数据泄漏");
            tempo.cw_read_dictionary("防数据勒索");
            tempo.cw_read_dictionary("防服务中断");
            tempo.ReplaceTextWithText_all();
            Console.WriteLine("录入信息完毕，请回到GUI窗体继续操作");
            toolStripStatusLabel1.Text = "录入信息完毕，已经保存到out";
            tempo.save($"出报告申请_{str_公司}_{str_系统}.docx");
            this.Show();
        }
        private string date_process(string a ,bool givemefirst)
        {
            //提取日期（结束日期）
            //ConsoleWriter.WriteCyan("在字符串中寻找日期：" + a);

            string patternA = @"\d\d\d\d年(\d)*月(\d)*日";
            string patternB = @"(\d)*月(\d)*日";
            //如果获取的短日期个数为2，但是长日期仅有1个，那么就是如2021年12月1日～12月1日这种写法
            //如果长短日期都只有一个，那么就是2021年12月1日这种写法（只有一天之类的）

            int shortDcount = 0;
            int LongDcount = 0;

            Regex rg = new Regex(patternB);
            MatchCollection matchedShortDate = rg.Matches(a);
            shortDcount = matchedShortDate.Count;
            //Console.WriteLine("shortDcount" + shortDcount);
            rg = new Regex(patternA);
            MatchCollection matchedLongDate = rg.Matches(a);
            LongDcount = matchedLongDate.Count;
            //Console.WriteLine("LongDcount"+ LongDcount);

            if (givemefirst)
            {
                a = matchedLongDate[0].Value;
                return a;
            }

            if (shortDcount > LongDcount)
            {
                string year = matchedLongDate[LongDcount - 1].Value.Substring(0, 5);
                string MandD = matchedShortDate[shortDcount - 1].Value;
                a = year + MandD;
            }
            else if (shortDcount == LongDcount)
            {
                a = matchedLongDate[LongDcount - 1].Value;
            }
            Console.WriteLine(a);
            return a;
        }
        //private void button1_Click(object sender, EventArgs e)
        //{
        //    //出报告申请_xxx公司_xxx系统.docx
        //    string filename = $"出报告申请_{str_公司}_{str_系统}.docx";
        //    tempo.save(filename);
        //    toolStripStatusLabel1.Text = "保存到桌面out文件夹了。";
        //}

        /// <summary>
        /// 打开生成的文件
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button1_Click_1(object sender, EventArgs e)
        {
            string savepath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\" + "out" + "\\" + opendoc;
            LaunchCommandLineApp(savepath);
        }

        static void LaunchCommandLineApp(string wordpath)
        {
            Process process = new Process();
            process.StartInfo.FileName = "cmd.exe";
            process.StartInfo.Arguments = "/c" + "start " + wordpath;
            process.StartInfo.UseShellExecute = false;   //是否使用操作系统shell启动 
            process.StartInfo.CreateNoWindow = false;   //是否在新窗口中启动该进程的值 (不显示程序窗口)
            process.Start();
            process.WaitForExit();  //等待程序执行完退出进程
            process.Close();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            string rootfoldername = $"{str_P项目编号}_{str_公司}_{str_系统}_出报告申请";
            string rootpath =  Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\" + rootfoldername;
            System.IO.Directory.CreateDirectory(rootpath);
            createFile(rootpath, $"{str_P项目编号}_GB01_测试报告_{str_系统}.docx");
            createFile(rootpath, $"{str_P项目编号}云计算附件.docx");
            createFile(rootpath, $"SSTL2022{str_公司}_{str_系统}网络安全等级保护测评报告附件.docx");
            createFile(rootpath, $"差距分析_{str_公司}_{str_系统}.xlsx");
            createFile(rootpath, $"出报告申请_{str_公司}_{str_系统}.docx");
            createFile(rootpath, $"大数据安全测评补充项_{str_公司}_{str_系统}.xlsx");
            createFile(rootpath, $"等保测评委托书_唯鸣官网.pdf");
            createFile(rootpath, $"漏洞清单_{str_公司}_{str_系统}.xlsx");
            createFile(rootpath, $"{str_公司}_{str_系统}.bak");

            Console.WriteLine("end:");

        }

        private void createFile(string pathString ,string fileName = "")
        {
            // Create a file name for the file you want to create.
            if (fileName == "")
            {
                fileName = System.IO.Path.GetRandomFileName();

            }

            // This example uses a random string for the name, but you also can specify
            // a particular name.
            //string fileName = "MyNewFile.txt";

            // Use Combine again to add the file name to the path.
            pathString = System.IO.Path.Combine(pathString, fileName);

            // Verify the path that you have constructed.
            Console.WriteLine("Path to my file: {0}\n", pathString);
            Console.WriteLine("Created file: {0}\n", pathString);

            // Check that the file doesn't already exist. If it doesn't exist, create
            // the file and write integers 0 - 99 to it.
            // DANGER: System.IO.File.Create will overwrite the file if it already exists.
            // This could happen even with random file names, although it is unlikely.
            if (!System.IO.File.Exists(pathString))
            {
                using (System.IO.FileStream fs = System.IO.File.Create(pathString))
                {
                    for (byte i = 0; i < 100; i++)
                    {
                        fs.WriteByte(i);
                    }
                }
            }
            else
            {
                Console.WriteLine("File \"{0}\" already exists.", fileName);
                return;
            }
        }
    }
}
