using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;
using archiver.ConsoleColorWriter;

namespace archiver
{
    public partial class Form_replaceForAll : Form
    {
        public string path_stub;
        public Form_replaceForAll()
        {
            InitializeComponent(); AllocConsole();
            Console.WriteLine(@"
提示：
仅支持docx文件
把目录拖入/粘贴textbox1.然后可以自动识别目录下的所有docx文件
通过shift/ctrl来多选自己想要替换的word(未实现)
输入替换内容后点击替换即可批量替换
文件导出到桌面 out 文件夹
");
        }
        [DllImport("kernel32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        static extern bool AllocConsole();  
        #region textDrop
        private void textBox1_DragDrop(object sender, DragEventArgs e)
        {
            textBox1.Text = ((System.Array)e.Data.GetData(DataFormats.FileDrop)).GetValue(0).ToString();
            string path = textBox1.Text;

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
            button2.PerformClick();
        }
        #endregion



        //一键替换
        private void button1_Click(object sender, EventArgs e)
        {

            if (textBox2 .Text == "")
            {
                toolStripStatusLabel1.Text = "被替换文字不能为空";
            }
            this.Hide();
            foreach (var item in listBox1.Items)
            {
                string filenam = item.ToString();
                myutil doc = new myutil(path_stub + filenam);
                doc._replacePatterns.Add(textBox2.Text , textBox3 .Text);
                doc.ReplaceTextWithText_all_noBracket();
                Console.WriteLine("已经处理："+filenam);
                if (radioButton1.Checked)
                {
                    doc.save();

                }
                else
                {
                    doc.saveUrl(path_stub+filenam);
                }
            }
            Console.WriteLine("处理完毕,保存到文件夹 out/replace");
            ConsoleWriter.WriteSeperator('#');
            this.Show();
        }

        //命令行模式
        private void button3_Click(object sender, EventArgs e)
        {
            if (textBox1.Text == "")
            {
                toolStripStatusLabel1.Text = "请先选择文件。";
            }
            this.Hide();
            Dictionary<string,string> map = new Dictionary<string,string>();
            ConsoleWriter.WriteSeperator('-');
            ConsoleWriter.WriteColoredText("请输入替换内容，支持多行，|（竖杠）分隔：",ConsoleColor.Green);
            while (true)
            {
                var line = Console.ReadLine().Split("|");
                if (line[0] == "") break;
                if (line.Length==1)
                {
                    ConsoleWriter.WriteYEllow("未检查到竖杠，请正确地重新输入");
                    continue;
                }
                map.Add(line[0], line[1]);
                ConsoleWriter.WriteCyan("成功录入,请继续输入（回车离开）");
            }

            Console.WriteLine("获取到字典：");

            foreach (var kvp in map)
            {
                ConsoleWriter.WriteCyan(kvp.Key + " → " + kvp.Value);
            }
            Console.WriteLine("将会替换:" + map.Count);


            foreach (var item in listBox1.Items)
            {
                string filenam = item.ToString();
                myutil doc = new myutil(path_stub + filenam);
                doc._replacePatterns = map;

                doc.ReplaceTextWithText_all_noBracket();
                Console.WriteLine("已经处理：" + filenam);
                if (radioButton1.Checked)
                {
                    doc.save();
                }
                else
                {
                    doc.saveUrl(path_stub + filenam);
                }
            }
            Console.WriteLine("处理完毕,保存到文件夹 out/replace");
            ConsoleWriter.WriteSeperator('#');
            this.Show();
        }

        //刷新列表
        private void button2_Click(object sender, EventArgs e)
        {
            string path = textBox1.Text;
            if (!Directory.Exists(path))
            {
                Console.WriteLine("目录不存在");
                return;
            }
            else
            {
                path_stub = path + "\\";
                string[] fangan_list = Directory.GetFiles(path, "*.docx");
                
                listBox1.Items.Clear();
                foreach (var item in fangan_list)
                {
                    if (!item.StartsWith("~"))
                    {
                        listBox1.Items.Add(item.Replace(path + "\\", ""));
                    }
                }
            }
            //string text = textBox1.Text;
            //string filename = text.Split('\\')[text.Split('\\').Length - 1];
            toolStripStatusLabel1.Text = "已刷新";
        }

        #region savelocation

        private void radioButton1_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton1.Checked)
            {
                radioButton2.Checked = false;
            }
            else 
            {
                radioButton2.Checked = true;
            }
        }

        private void radioButton2_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton2.Checked)
            {
                radioButton1.Checked = false;
            }
            else
            {
                radioButton1.Checked = true;
            }
        }
        #endregion


        //移除选中项目
        private void button4_Click(object sender, EventArgs e)
        {
            List<int> f = new List<int>();
            for (int i = listBox1.Items.Count - 1; i >=0; i--)
            {
                if (listBox1.SelectedIndices.Contains(i))
                {
                    listBox1.Items.RemoveAt(i);
                }
            }
            f.Reverse();

            toolStripStatusLabel1.Text = "已删除";
            
        }

        //仅保留选中项目
        private void button5_Click(object sender, EventArgs e)
        {
            List<int> f = new List<int>();
            for (int i = listBox1.Items.Count - 1; i >= 0; i--)
            {
                if (!listBox1.SelectedIndices.Contains(i))
                {
                    listBox1.Items.RemoveAt(i);
                }
            }
            f.Reverse();

            toolStripStatusLabel1.Text = "已删除";
        }
    }
}
