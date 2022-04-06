using archiver.ConsoleColorWriter;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;

namespace archiver
{
    public partial class Form_ModifyFormat : Form
    {
        public myutil doc;
        public Form_ModifyFormat()
        {
            InitializeComponent();
            AllocConsole();
        }
        #region console
        [DllImport("kernel32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        static extern bool AllocConsole();
        #endregion
        #region textDrop
        private void button4_Click(object sender, EventArgs e)
        {
            string path = textBox1.Text;
            this.Hide();
            ConsoleWriter.WriteColoredText("加载报告...", ConsoleColor.Green);
            doc = new myutil(path);
            ConsoleWriter.WriteColoredText("加载完成", ConsoleColor.Green);
            this.Show();
        }
        private void textBox1_DragDrop(object sender, DragEventArgs e)
        {
            textBox1.Text = ((System.Array)e.Data.GetData(DataFormats.FileDrop)).GetValue(0).ToString();

            string path = textBox1.Text;
            this.Hide();
            ConsoleWriter.WriteColoredText("加载报告...", ConsoleColor.Green);
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
        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }
        #endregion
        #region test
        private void Form_ModifyFormat_Load(object sender, EventArgs e)
        {
            textBox1.Text = @"C:\Users\kasusa\Desktop\P202201056_基金销售系统_测评报告.docx";
            string path = textBox1.Text;
            this.Hide();
            ConsoleWriter.WriteColoredText("加载报告(有时候假死，请按回车跳出)...", ConsoleColor.Green);
            doc = new myutil(path);
            ConsoleWriter.WriteColoredText("加载完成", ConsoleColor.Green);
            this.Show();
        }

        private void reloaddoc()
        {

        }
        #endregion

        /// <summary>
        /// 开始处理文件
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button1_Click(object sender, EventArgs e)
        {
            this.Hide();
            //- [ ] 清空项目编号（同时提取到P号
            var p = doc.Find_Paragraph_for_p("项目编号：");
            string a = p.Text.Substring(5);
            Console.WriteLine(a);
            if (p != null)doc.remove_p(p);
            //- [ ] 声明在最后一段增加：本报告记录编号：P202201056-GB01。
            //- [ ] 声明四号

            int i1 = doc.Find_Paragraph_for_i("声明", 1);
            int i2 = doc.Find_Paragraph_for_i("等级测评结论", 1);
            for (int i = i1; i < i2; i++)
            {
                p = doc.Paragraphs[i];
                p.FontSize(14);
            }
            p = doc.Find_Paragraph_for_p("不得对相关内容擅自进行增加、修改和伪造或掩盖事实");
            //p = p.NextParagraph;
            //p.FontSize(14);//因为声明是4号字，所以这个也变一下
            for (int i = 0; i < 7; i++)//增加几个空行让单位跑到右下角
            {
                p.InsertParagraphAfterSelf(" ");
            }
            p.InsertParagraphAfterSelf($"本报告记录编号：{a}").FontSize(14);
            Console.WriteLine("[ok] 声明在最后一段增加：本报告记录编号，且增加了空行、字号修改为14");

            //- [ ] 首页报告时间变成 中文
            var t = doc.tables[0];
            var c = doc.table_Get_cell(t, 2, 1);
            a = doc.table_Get_cell_text(t, 2, 1);
            string fengmian_time = a;
            Console.WriteLine("获取到报告封面时间："+a);
            doc.cell_replace_text(c, myutil.datetochinese(a));
            Console.WriteLine("[ok] 首页报告时间变成中文了 但是长度变短了");

            //- [ ] 网络安全等级测评基本信息表 测评单位表格 的编制日期、审核日期、批准日期自动填写（根据首页时间-1-2）
            t = doc.findTableList("测评单位")[0];
            c = doc.table_Get_cell(t, 8, 4);
            doc.cell_settext(c, fengmian_time);
            c = doc.table_Get_cell(t, 7, 4);
            doc.cell_settext(c, myutil.datetosmalldate(fengmian_time,-1));
            c = doc.table_Get_cell(t, 6, 4);
            doc.cell_settext(c, myutil.datetosmalldate(fengmian_time,-2));
            Console.WriteLine("[ok] 测评审核日期已经在表格中自动填写（编制审核批准各一天，审核日期同封面日期）");

            //- [ ] 等级测评结论扩展表（云计算安全）、等级测评结论扩展表（大数据安全都删除）
            ConsoleWriter.WriteCyan("是否删除 等级测评结论扩展表（云计算安全)？（y/n）");
            if (Console.ReadLine() != "n") 
            {
                t = doc.findTableList("等级测评结论扩展表（云计算安全）")[0];
                t.Remove();

                i1 = doc.Find_Paragraph_for_i("等级测评结论扩展表（云计算安全）");//删除两个云计算安全标题下的空行
                for (int i = i1;; i++)// i < i1+7
                {
                    if (doc.Paragraphs[i].Text.Contains("等级测评结论扩展表（大数据安全）"))
                    {
                        break;
                    }
                    doc.remove_p(doc.Paragraphs[i]);
                }

                Console.WriteLine("[ok] 等级测评结论扩展表（云计算安全）已移除");
            }


            ConsoleWriter.WriteCyan("是否删除 等级测评结论扩展表（大数据安全)？（y/n）");
            if (Console.ReadLine() != "n")
            {
                t = doc.findTableList("等级测评结论扩展表（大数据安全）")[0];
                t.Remove();
                i1 = doc.Find_Paragraph_for_i("等级测评结论扩展表（大数据安全）");//删除两个云计算安全标题下的空行
                for (int i = i1; ; i++)//i < i1 + 6
                {
                    if (doc.Paragraphs[i].Text == "总体评价")
                    {
                        break;
                    }
                    doc.remove_p(doc.Paragraphs[i]);
                }
                Console.WriteLine("[ok] 等级测评结论扩展表（大数据安全）已移除");
            }


            //- [ ] 总体评价 两端对齐
            i1 = doc.Find_Paragraph_for_i("总体评价", 1);
            i2 = doc.Find_Paragraph_for_i("主要安全问题及整改建议", 1);
            //Console.WriteLine($"i1 {i1},i2 {i2}");
            for (int i = i1 + 1; i < i2; i++)
            {
                //Console.WriteLine(i);
                p = doc.Paragraphs[i];
                a = p.Text;
                if (a.Length > 20) a = a.Substring(0, 20) + "...";
                ConsoleWriter.Writehiddeninfo($"[两端对齐] {a}");
                p.Alignment = Xceed.Document.NET.Alignment.both;
            }
            //- [ ] 1.2 测评依据 行变为我单位样式
            //doc.remove_p_from_to("测评过程中主要依据的标准：", "测评过程",1);
            doc.remove_p(doc.Find_Paragraph_for_p("信息安全技术 网络安全等级保护基本"));
            p = doc.Find_Paragraph_for_p("测评过程中主要依据的标准");
            string putong = @"1. GB/T 22239-2019：《信息安全技术 网络安全等级保护基本要求》
2. GB/T 28448-2019：《信息安全技术 网络安全等级保护测评要求》
以下为本次测评的相关参考标准和文档：
3. GB/T 28449-2018：《信息安全技术 网络安全等级保护测评过程指南》
4. GB/T 20984-2007：《信息安全技术 信息安全风险评估规范》
5. GB/T 17859-1999：《计算机信息系统 安全保护等级划分准则》";
            
            string jinrong = @"1. GB/T 22239-2019：《信息安全技术 网络安全等级保护基本要求》
2. GB/T 28448-2019：《信息安全技术 网络安全等级保护测评要求》
3. JR/T 0071.2-2020：《金融行业网络安全等级保护实施指引 第 2 部分：基本要求》
4. JR/T 0072—2020：《金融行业网络安全等级保护测评指南》
以下为本次测评的相关参考标准和文档：
5. GB/T 17859—1999 ：《计算机信息系统 安全保护等级划分准则》
6. GB/T 28449-2018：《信息安全技术 网络安全等级保护测评过程指南》
7. GB/T 20984-2007：《信息安全技术 信息安全风险评估规范》
";
            ConsoleWriter.WriteCyan("选择测评依据，普通标（1）金融标（2）：");
            string x = Console.ReadLine();
            if (x == "1")
            {
                var strlist = putong.Replace("\r", "").Split("\n");

                for (int i = strlist.Length - 1; i >= 0; i--)
                {
                    p.InsertParagraphAfterSelf(strlist[i]);
                }
                ConsoleWriter.Writehiddeninfo(putong);
            }
            else if (x  == "2")
            {
                var strlist = jinrong.Replace("\r", "").Split("\n");

                for (int i = strlist.Length - 1; i >= 0; i--)
                {
                    p.InsertParagraphAfterSelf(strlist[i]);
                }
                ConsoleWriter.Writehiddeninfo(jinrong);
            }
            else
            {
                Console.WriteLine("没选默认普通标");
                var strlist = putong.Replace("\r", "").Split("\n");

                for (int i = strlist.Length - 1; i >= 0; i--)
                {
                    p.InsertParagraphAfterSelf(strlist[i]);
                }
                ConsoleWriter.Writehiddeninfo(putong);

            }
            Console.WriteLine("[ok] 1.2 测评依据 行变为我单位样式");

            //- [ ] 1.4 报告分发范围 改成我单位样式(数字变成汉字)
            p = doc.Find_Paragraph_for_p("等级测评报告正本一式");
            a = p.Text.Replace("2", "二").Replace("1", "一").Replace(" ","");
            p.InsertParagraphAfterSelf(a);
            doc.remove_p(p);
            Console.WriteLine("[ok] 1.4 报告分发范围 改成我单位样式(数字变成汉字)");

            //- [ ] 把业务截图增加编号 图2-x xxxx 华文仿宋 5号加粗
            //Console.WriteLine($"i1 {i1},i2 {i2}");
            i1 = doc.Find_Paragraph_for_i("业务和采用的技术",2);
            i2 = doc.Find_Paragraph_for_i("网络结构",2);

            int tj = 0;//图2-j
            for (int i = i1; i < i2; i++)
            {
                p = doc.Paragraphs[i];
                a = doc.Paragraphs[i].Text;
                if (a.Contains("图")&&a.Length<30)
                {
                    tj++;
                    string str = a.Replace("图", "图2-" + tj+" ");
                    var p2 = p.InsertParagraphAfterSelf(str);
                    p2.FontSize(10.5);//五号
                    p2.Bold();
                    p2.Alignment = Xceed.Document.NET.Alignment.center;
                    doc.remove_p(p);
                    ConsoleWriter.Writehiddeninfo(str);
                }
            }
            Console.WriteLine("[ok] 把业务截图增加编号 图2-x ");

            p = doc.Find_Paragraph_for_p("网络结构拓扑图");
            p.ReplaceText(p.Text, p.Text.Replace("图2-1 ", "图2-" + tj + " "));
            p = doc.Find_Paragraph_for_p("如图2.1");
            p.ReplaceText(p.Text, p.Text.Replace("如图2.1 ", "如图2-" + tj ));

            Console.WriteLine("[ok] 修改网络结果描述第一段”如图2.1“以及图片拓扑图图片序号修正 ");

            //- [ ] 2.2.2 2.2.3 内容改成：本次测评未涉及xxx指标
            ConsoleWriter.WriteCyan("2.2.2 【安全扩展要求指标】表格 更改成 “本次测评未涉及... ”？(y/n）");
            x =  Console.ReadLine();
            if (x != "n")
            {
                t = doc.findTableList("扩展类型安全类控制点测评项数")[0];
                c = t.Rows[1].Cells[0];
                doc.cell_settext_default(c, "本次测评未涉及安全扩展要求指标");
                c.Paragraphs[0].Bold(false);
            }
            t = doc.findTableList("安全类控制点测评项数")[1];
            c = t.Rows[1].Cells[0];
            doc.cell_settext_default(c, "本次测评未涉及其他要求指标");
            c.Paragraphs[0].Bold(false);
            Console.WriteLine("[ok] 2.2.2 2.2.3 内容改成：本次测评未涉及xxx指标");
            //- [ ] 2.3.2.6 其他设备，改成本次测评未涉及xxx
            p = doc.Find_Paragraph_for_plist("其他设备")[1];
            t = p.FollowingTables[0];
            c = t.Rows[1].Cells[0];
            doc.cell_settext_default(c, "本次测评未涉及其他设备");

            //- [ ] 3.4.7 其他系统或设备： 本次测评未涉及其他系统或设备。，删除3.4.7.1\3.4.7.2
            i1 = doc.Find_Paragraph_for_i("其他系统或设备", 2);
            p = doc.Paragraphs[i1].InsertParagraphAfterSelf("   本次测评未涉及其他系统或设备");
            for (int i = i1+2; ; i++)
            {
                p = doc.Paragraphs[i];
                a = p.Text;
                if (a.Contains("安全管理中心"))
                {
                    break;
                }
                doc.remove_p(p);
            }
            Console.WriteLine("[ok] 3.4.7 其他系统或设备： 本次测评未涉及其他系统或设备。并删除3.4.7.1 ， 3.4.7.2");
            //- [ ] 3.11 其他安全指标增加：本次测评未涉及其他安全要求指标。删除3.11.1\3.4.11.2(汇总分析)
            //- [ ] 2.3.2.10安全相关人员/所属单位都变成系统的单位
            //- [ ] 总体评价 → 分离为10个层面（要包括在xxx层面字段） → 分别填充到3章节 3.1-3.10（安全计算环境要用脑子写 = 就简单粘贴了）
            //- [ ] 3.1-3.10 总体评价 → 10段文字提取
            //- [ ] 4.3整体测评结果汇总 增加严重程度变化中的勾选框 框升高 勾降低
            //- [ ] 附录A中的设备 备注一律杠杠
            //- [ ] A.10 检查有没有密码产品行决定添加不添加新行
            //- [ ] A.11安全相关人员/所属单位都变成系统的单位
            //- [ ]
            //- [ ] C.11 本次测评未涉及其它安全要求指标,并删除两个表头
            //- [ ] 删除附录H,I

            //## 1.2 参照标准
            try
            {
                doc.save();
                ConsoleWriter.WriteGreen("已保存到out文件夹");
                button2.Enabled = true;
            }
            catch (Exception)
            {
                ConsoleWriter.WriteRed("无法保存，原因未知，请重新尝试并关闭所有word");
                this.Show();

            }

            this.Show();
        }

        #region opendoctocheck
        /// <summary>
        /// open the docx
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button2_Click(object sender, EventArgs e)
        {
            string opendoc = textBox1.Text.Split("\\")[textBox1.Text.Split("\\").Length-1];
            label2.Text = opendoc;
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
        #endregion

        private void button3_Click(object sender, EventArgs e)
        {
            this.Hide();
            int i1, i2;
            var p = doc.Paragraphs[0];
            string a;

            //- [ ] 3.4.7 其他系统或设备： 本次测评未涉及其他系统或设备。，删除3.4.7.1\3.4.7.2
            i1 = doc.Find_Paragraph_for_i("其他系统或设备", 2);
            for (int i = i1; ; i++)
            {
                p = doc.Paragraphs[i];
                a = p.Text;
                Console.WriteLine(a);
                if (a.Contains("安全管理中心"))
                {
                    break;
                }
            }

            try
            {
                doc.save();
                ConsoleWriter.WriteGreen("已保存到out文件夹");
                button2.Enabled = true;
            }
            catch (Exception)
            {
                ConsoleWriter.WriteRed("无法保存，原因位置，请重新尝试并关闭所有word");
                this.Show();
            }
            button2.Enabled = true;
            this.Show();
        }


    }
}
