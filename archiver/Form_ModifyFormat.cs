using archiver.ConsoleColorWriter;
using archiver.Properties;
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
        string targetname="";
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
            //如果已经选择了文件（textbox中有字
            if (textBox1.Text!="")
            {
                string path = textBox1.Text;
                this.Hide();
                ConsoleWriter.WriteColoredText("加载报告...", ConsoleColor.Green);
                doc = new myutil(path);
                ConsoleWriter.WriteColoredText("加载完成", ConsoleColor.Green);
                this.Show();
            }
            else
            {
                //如果未选择文件，加载上次的文件。
                textBox1.Text = Settings.Default.MF_path;
                if (textBox1.Text == "")
                {
                    MessageBox.Show("未打开过任何文件，缓存中没有记录");
                    return;
                }

                string path = textBox1.Text;
                this.Hide();
                ConsoleWriter.WriteColoredText("加载报告...", ConsoleColor.Green);
                doc = new myutil(path);
                ConsoleWriter.WriteColoredText("加载完成", ConsoleColor.Green);
                button5.Enabled = true;//允许打开源文件
                this.Show();
            }
        }
        private void textBox1_DragDrop(object sender, DragEventArgs e)
        {
            textBox1.Text = ((System.Array)e.Data.GetData(DataFormats.FileDrop)).GetValue(0).ToString();

            string path = textBox1.Text;
            this.Hide();
            ConsoleWriter.WriteColoredText("加载报告...", ConsoleColor.Green);
            doc = new myutil(path);
            ConsoleWriter.WriteColoredText("加载完成", ConsoleColor.Green);
            button5.Enabled = true;//允许打开源文件
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
            //保存位置
             Settings.Default.MF_path = textBox1.Text;
        }

        #endregion
        #region test
        private void Form_ModifyFormat_Load(object sender, EventArgs e)
        {
            //textBox1.Text = @"C:\Users\kasusa\Desktop\P202204046_维鸣官网_测评报告20220531144620（康玉婷）.docx";
            //string path = textBox1.Text;
            //this.Hide();
            //ConsoleWriter.WriteColoredText("加载报告(有时候假死，请按回车跳出)...", ConsoleColor.Green);
            //doc = new myutil(path);
            //ConsoleWriter.WriteColoredText("加载完成", ConsoleColor.Green);
            //this.Show();
        }

        #endregion

        /// <summary>
        /// 处理文件按钮
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

            //机构代码全称 SC202127130010092
            var t = doc.findTableList("测评单位")[0];
            var cell = doc.table_Get_cell(t, 1, 3);
            doc.cell_settext(cell, "SC202127130010092");


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
            for (int i = 0; i < 6; i++)//增加几个空行让单位跑到右下角
            {
                p.InsertParagraphAfterSelf(" ");
            }
            p.InsertParagraphAfterSelf($"本报告记录编号：{a}-GB01").FontSize(14);
            Console.WriteLine("[ok] 声明在最后一段增加：本报告记录编号，且增加了空行、字号修改为14");

            //- [ ] 首页报告时间变成 中文
             t = doc.tables[0];
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
            ConsoleWriter.WriteCyan("删除 等级测评结论扩展表（云计算安全)？（y/n）");
            //if (Console.ReadLine() != "n") 
            if (!checkBox1.Checked)
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


            ConsoleWriter.WriteCyan("删除 等级测评结论扩展表（大数据安全)？（y/n）");
            //if (Console.ReadLine() != "n")
            if (!checkBox2.Checked)
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
5. GB 17859-1999：《计算机信息系统 安全保护等级划分准则》";
            
            string jinrong = @"1. GB/T 22239-2019：《信息安全技术 网络安全等级保护基本要求》
2. GB/T 28448-2019：《信息安全技术 网络安全等级保护测评要求》
3. JR/T 0071.2-2020：《金融行业网络安全等级保护实施指引 第 2 部分：基本要求》
4. JR/T 0072—2020：《金融行业网络安全等级保护测评指南》
以下为本次测评的相关参考标准和文档：
5. GB 17859—1999 ：《计算机信息系统 安全保护等级划分准则》
6. GB/T 28449-2018：《信息安全技术 网络安全等级保护测评过程指南》
7. GB/T 20984-2007：《信息安全技术 信息安全风险评估规范》
";
            ConsoleWriter.WriteCyan("选择测评依据，普通标（1）金融标（2）(没选默认普通标)：");
            //string x = Console.ReadLine();
            //if (x == "1")
            if (radioButton1国标.Checked)
            {
                var strlist = putong.Replace("\r", "").Split("\n");

                for (int i = strlist.Length - 1; i >= 0; i--)
                {
                    p.InsertParagraphAfterSelf(strlist[i]);
                }
                ConsoleWriter.Writehiddeninfo(putong);
            }
            //else if (x  == "2")
            else if (radioButton2金融标.Checked)
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

            int tj = 0;//图2-tj
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
            p.ReplaceText(p.Text, p.Text.Replace("图2-1 ", "图2-" + ++tj + " "));


            Console.WriteLine("[ok] 修改网络结果描述第一段”如图2.1“以及图片拓扑图图片序号修正 ");

            //- [ ] 2.2.2 2.2.3 内容改成：本次测评未涉及xxx指标
            ConsoleWriter.WriteCyan("本项目是否包括扩展指标？（2.2.2 【安全扩展要求指标】表格 更改成 “本次测评未涉及... ”）(y/n）");
            //x =  Console.ReadLine();
            //if (x != "n")            //x =  Console.ReadLine();
            //if (x != "n")
            if(checkBox1.Checked || checkBox2.Checked)
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
            for (int i = i1+1; ; i++)
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
            i1 = doc.Find_Paragraph_for_i("其他安全要求指标", 7);
            i2 = doc.Find_Paragraph_for_i("验证测试", 2);
            for (int i = i1+1; i < i2-1; i++)
            {
                p = doc.Paragraphs[i];
                doc.remove_p(p);
            }
            p = doc.Paragraphs[i2 - 1];
            p.ReplaceText(p.Text, "本次测评未涉及其他安全要求指标。");
            ConsoleWriter.WriteText(" [ok] 3.11 其他安全指标增加：本次测评未涉及其他安全要求指标");
            //- [ ] 2.3.2.10安全相关人员/所属单位都变成系统的单位
            //- [ ] 总体评价 → 分离为10个层面（要包括在xxx层面字段） → 分别填充到3章节 3.1-3.10（安全计算环境要用脑子写 = 就简单粘贴了）
            if (checkBox3.Checked)
            {
                ConsoleWriter.WriteGreen("在总体评价章节中寻找在xxx方面、然后逐个填写到单项测评结果分析中，安全计算环境可能需要后续润色*因为安全计算环境是一套话放到多个位置");
                //存放数据的结构体
                Dictionary<string, string> myDictionary总体评价 = new Dictionary<string, string>();
                //index列表
                string[] wordlistChap3 = "安全物理环境、安全通信网络、安全区域边界、安全计算环境、安全管理中心、安全管理制度、安全管理机构、安全管理人员、安全建设管理、安全运维管理".Split('、');
                //获取数据 总体评价中每一段起前面写 在xxx方面, 就能够被识别
                foreach (var item in wordlistChap3)
                {
                    p = doc.Find_Paragraph_for_p("在" + item + "方面");
                    myDictionary总体评价.Add(item, p.Text);
                }
                //安全计算环境额外index列表
                string[] wordlistChap3_4 = "网络设备、安全设备、服务器和终端、系统管理软件/平台、业务应用系统/平台、数据资源".Split('、');
                //逐个查找并填入总体评价中收集到的语句
                i1 = doc.Find_Paragraph_for_i("单项测评结果分析", 2);
                while (p.Text != "整体测评")
                {
                    ++i1;
                    p = doc.Paragraphs[i1];

                    foreach (var item in wordlistChap3)
                    {
                        if (p.Text.Contains(item))
                        {
                            ++i1;
                            p = doc.Paragraphs[i1];
                            if (p.Text.Contains("已有安全控制措施汇总分析"))
                            {
                                ++i1;
                                p = doc.Paragraphs[i1];
                                p.ReplaceText(p.Text, myDictionary总体评价[item]);
                            }
                        }
                    }
                    foreach (var item in wordlistChap3_4)
                    {
                        if (p.Text.Contains(item))
                        {
                            ++i1;
                            p = doc.Paragraphs[i1];
                            if (p.Text.Contains("已有安全控制措施汇总分析"))
                            {
                                ++i1;
                                p = doc.Paragraphs[i1];
                                p.ReplaceText(p.Text, myDictionary总体评价["安全计算环境"].Replace("在安全计算环境方面，", "在" + item + "方面，"));
                            }
                        }
                    }

                }
                ConsoleWriter.WriteGreen("[ok] 总体评价 → 分离为10个层面（要包括在xxx层面字段） → 分别填充到3章节 3.1-3.10");
            }


            //- [ ] 4.3整体测评结果汇总 增加严重程度变化中的勾选框 框升高 勾降低
            //- [ ] 4.3整体测评结果汇总 变成无
            if (checkBox4.Checked)
            {
                t = doc.findTableList("问题编号安全问题测评对象整体测评描述严重程度变化")[0];
                while (t.Rows.Count > 2)
                {
                    t.RemoveRow(t.Rows.Count - 1);
                }
                c = doc.table_Get_cell(t, 1, 0);
                doc.cell_settext(c, "无");
            }

            //- [ ] 附录A中的设备 备注一律杠杠
            //机房
            t = doc.findTableList("序号机房名称物理位置重要程度备注")[0];
            doc.table_lastcell_ganggang(t);
            //网络设备
            t = doc.findTableList("序号设备名称是否虚拟设备系统及版本品牌及型号用途重要程度备注")[0];
            doc.table_lastcell_ganggang(t);     
            //安全设备
            t = doc.findTableList("序号设备名称是否虚拟设备系统及版本品牌及型号用途重要程度备注")[1];
            doc.table_lastcell_ganggang(t);
            //A.4	服务器/存储设备
            t = doc.findTableList("序号设备名称所属业务应用系统/平台名称是否虚拟设备操作系统及版本数据库管理系统及版本中间件及版本重要程度备注")[0];
            doc.table_lastcell_ganggang(t);            
            //终端设备
            t = doc.findTableList("序号设备名称是否虚拟设备操作系统/控制软件及版本用途重要程度备注")[0];
            doc.table_lastcell_ganggang(t);                
            //跳过其他
            //系统管理软件、平台
            t = doc.findTableList("序号系统管理软件/平台名称主要功能版本所在设备名称重要程度备注")[0];
            doc.table_lastcell_ganggang(t);    
            //业务应用系统、平台
            t = doc.findTableList("序号业务应用系统/平台名称主要功能业务应用软件及版本开发厂商重要程度备注")[0];
            doc.table_lastcell_ganggang(t);
            ConsoleWriter.WriteText(" [ok] 附录A中的设备 备注一律杠杠");

            //- [ ] A.10 检查有没有密码产品行决定添加不添加新行
            //- [ ] A.11安全相关人员/所属单位都变成系统的单位
            //- [ ]
            //- [ ] C.11 本次测评未涉及其它安全要求指标,并删除两个表头
            i1 = doc.Find_Paragraph_for_i("其他安全要求指标", 9);
            //i2 = doc.Find_Paragraph_for_i("单项测评结果记录", 3);
            for (int i = i1 + 1; i <= i1+4; i++)
            {
                p = doc.Paragraphs[i];
                doc.remove_p(p);
            }
            p = doc.Paragraphs[i1];
            p = p.InsertParagraphAfterSelf("\t本次测评未涉及其他安全要求指标。");
            //p.ReplaceText(p.Text, );
            //- [ ] 删除附录H,I
            //删除文档中红字提示
            delred();

            //全文替换
            //其他要求指标 → 其他安全要求指标
            doc._replacePatterns.Add("本次测评不涉及", "本次测评未涉及");
            doc._replacePatterns.Add("其他要求指标", "其他安全要求指标");
            //doc._replacePatterns.Add("其他设备", "其他系统或设备");
            doc.ReplaceTextWithText_all_noBracket();

            //保存
            try
            {
                //格式
                //默认导出名称：P202204046_维鸣官网_测评报告20220531144620（康玉婷）
                //目标名称：P202111003_GB01_测试报告_TOMONICHN
                string defaultname = textBox1.Text.Split("\\")[textBox1.Text.Split("\\").Length - 1];
                string Pnum = defaultname.Split("_")[0];
                string xitongname = defaultname.Split("_")[1];
                targetname = Pnum + "_GB01_测试报告_" + xitongname + ".docx";

                doc.save(targetname);
                ConsoleWriter.WriteGreen("已保存到out文件夹");
                button2.Enabled = true;
            }
            catch (Exception)
            {
                ConsoleWriter.WriteRed("无法保存，原因unknow，请重新尝试并关闭所有word,然后“重新加载”报告后重试");
                this.Show();
            }

            this.Show();
        }

        // 删除红字
        private void delred()
        {
            var p = doc.Find_Paragraph_for_p("【如果手动修订过报告，请手动刷新目录域】");
            if (p != null) doc.remove_p(p);
            
            p = doc.Find_Paragraph_for_p("【填写说明：验证测试包括漏洞扫描、渗透测试等。");
            if (p != null) doc.remove_p(p);
            
            p = doc.Find_Paragraph_for_p("【填写说明：描述漏洞扫描工具的名称及其系统版本和规则库版本，");
            if (p != null) doc.remove_p(p);
            
            p = doc.Find_Paragraph_for_p("【填写说明：按照下表对漏洞扫描结果进行汇总，详细漏洞扫描结果记录描述参见报告附录");
            if (p != null) doc.remove_p(p);
                        
            p = doc.Find_Paragraph_for_p("【填写说明：针对系统漏洞扫描或Web漏洞扫描结果进行分析，汇总被测对象存在的安全漏洞。严重程度结果为“高”、“中”或“低”。");
            if (p != null) doc.remove_p(p);
                                    
            p = doc.Find_Paragraph_for_p("【填写说明：本次测评若果未对网络设备、安全设备、服务器操作系统和应用系统等进行渗透测试，请提供特殊说明材料，以图片方式提供，说明文件需要有签字、盖章和日期。");
            if (p != null) doc.remove_p(p);
                                    
            p = doc.Find_Paragraph_for_p("【填写说明：简要描述渗透测试的工具、方法和过程等。");
            if (p != null) doc.remove_p(p);
                                    
            p = doc.Find_Paragraph_for_p("【填写说明：针对渗透测试发现的安全问题进行汇总描述，详细渗透测试过程记录描述参见报告附录。严重程度结果为“高”、“中”或“低”。");
            if (p != null) doc.remove_p(p);
                                    
            p = doc.Find_Paragraph_for_p("【特别说明：已整改的安全问题，不输出到表5-1中");
            if (p != null) doc.remove_p(p);
                                    
            p = doc.Find_Paragraph_for_p("【提示：CTRL+A 全选后，按F9，刷新 附录C 表");
            if (p != null) doc.remove_p(p);

            ConsoleWriter.WriteGreen("已经删除所有红字。");

        }

        #region 打开检查按钮
        /// <summary>
        /// open the docx
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button2_Click(object sender, EventArgs e)
        {
            //string opendoc = textBox1.Text.Split("\\")[textBox1.Text.Split("\\").Length-1];
            label2.Text = targetname;
            string savepath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\" + "out" + "\\" + targetname;
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
        private void button5_Click(object sender, EventArgs e)
        {
            string savepath = textBox1.Text;
            LaunchCommandLineApp(savepath);
        }
        #endregion

        //测试按钮
        private void button3_Click(object sender, EventArgs e)
        {
            this.Hide();
            int i1, i2;
            var p = doc.Paragraphs[0];
            string a;
            //机构代码全称
            var t = doc.findTableList("测评单位")[0];
            var cell = doc.table_Get_cell(t, 1, 3);
            doc.cell_settext(cell, "SC202127130010092");


            Console.WriteLine();
            try
            {
                doc.save("(test）"+ textBox1.Text.Split("\\")[textBox1.Text.Split("\\").Length-1] );
                ConsoleWriter.WriteGreen("已保存到out文件夹");
                button2.Enabled = true;
            }
            catch (Exception)
            {
                ConsoleWriter.WriteRed("无法保存，原因unknow，请重新尝试并关闭所有word，然后“重新加载”报告后重试");
                this.Show();
            }
            button2.Enabled = true;
            this.Show();
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {

        }


    }
}
