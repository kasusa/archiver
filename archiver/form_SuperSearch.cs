using archiver.ConsoleColorWriter;
using archiver.Properties;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Media;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;

namespace archiver
{
    public partial class form_SuperSearch : Form
    {
        string fullguanlipath = "";
        string currentrootpath = "";
        LinkedList<string> filepathlist = new LinkedList<string>();
        public form_SuperSearch()
        {
            InitializeComponent();
            AllocConsole();
            treeView1.ExpandAll();
        }
        [DllImport("kernel32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        static extern bool AllocConsole();

        #region textDrop
        private void textBox1_DragDrop(object sender, DragEventArgs e)
        {
            fullguanlipath = ((System.Array)e.Data.GetData(DataFormats.FileDrop)).GetValue(0).ToString();
            //debug
            //Console.WriteLine(fullguanlipath);
            //Console.ReadLine();
            //MessageBox.Show(fullguanlipath);
            PaintTreeView(treeView1, fullguanlipath);
            Settings.Default.lastguanliPath = fullguanlipath;
            Settings.Default.Save();
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

        #region treeviewRender
        private void PaintTreeView(TreeView treeView, string fullPath)
        {
            try
            {
                treeView.Nodes.Clear(); //清空TreeView
                treeView.Nodes.Add("[Root]");
                //currentrootpath = fullguanlipath + "\\";
                filepathlist.Clear();
                GetMultiNode(treeView.Nodes[0], fullPath);
                treeView1.ExpandAll();

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + "\r\n出错的位置为：Form1.PaintTreeView()");
            }
        }


        private bool GetMultiNode(TreeNode treeNode, string path)
        {
            currentrootpath = path;
            if (Directory.Exists(path) == false)
            { return false; }

            DirectoryInfo dirs = new DirectoryInfo(path); //获得程序所在路径的目录对象
            DirectoryInfo[] dir = dirs.GetDirectories();//获得目录下文件夹对象
            FileInfo[] file = dirs.GetFiles();//获得目录下文件对象
            int dircount = dir.Count();//获得文件夹对象数量
            int filecount = file.Count();//获得文件对象数量
            int sumcount = dircount + filecount;

            if (sumcount == 0)
            { return false; }

            //循环文件夹
            for (int j = 0; j < dircount; j++)
            {
                ConsoleWriter.WriteCyan("floder:" + dir[j].Name);
                treeNode.Nodes.Add("[F]  "+dir[j].Name);
                string pathNodeB = path + "\\" + dir[j].Name;
                GetMultiNode(treeNode.Nodes[j], pathNodeB);
            }

            //循环文件
            for (int j = 0; j < filecount; j++)
            {
                string filepath = path + "\\" + file[j].Name;
                //Console.WriteLine(file[j].Name);
                //如果名字前面~或者后面不是docx跳过
                if (file[j].Name.Contains("~")) continue;
                if (file[j].Name.EndsWith(".doc")) {
                    ConsoleWriter.WriteRed("[Reject] " + file[j].Name );
                    continue;
                };

                    if (!file[j].Name.EndsWith(".docx")) continue;
                treeNode.Nodes.Add(file[j].Name);
                ConsoleWriter.WriteGreen("[Accept] "+ file[j].Name );//+ filepath
                filepathlist.AddLast(filepath);
                
            }
            return true;
        }
        #endregion

        //刷新按钮
        private void button2_Click(object sender, EventArgs e)
        {
            PaintTreeView(treeView1, Settings.Default.lastguanliPath);
        }

        private void treeView1_AfterSelect(object sender, TreeViewEventArgs e)
        {
            //MessageBox.Show(treeView1.SelectedNode.Text + "文本已复制");
        }

        private void button3_Click(object sender, EventArgs e)
        {
            treeView1.Font = new System.Drawing.Font(treeView1.Font.FontFamily, treeView1.Font.Size - 1); 
        }

        private void button4_Click(object sender, EventArgs e)
        {
            treeView1.Font = new System.Drawing.Font(treeView1.Font.FontFamily, treeView1.Font.Size + 1); ;
        }

        private void treeView1_NodeMouseDoubleClick(object sender, TreeNodeMouseClickEventArgs e)
        {
            //和enter功能一样
            string filename = treeView1.SelectedNode.Text;
            //Console.WriteLine("filename " + filename);
            string filepath = "";
            foreach (var ipath in filepathlist)
            {
                if (ipath.Contains(filename))
                {
                    filepath = ipath;
                    break;
                }
            }
            //Console.WriteLine("尝试打开： " + filepath);
            LaunchCommandLineApp(filepath);
        }

        private void treeView1_KeyDown(object sender, KeyEventArgs e)
        {
            //打开word
            if (e.KeyCode == Keys.Enter)
            {
                string filename = treeView1.SelectedNode.Text;
                //Console.WriteLine("filename "+filename);
                string filepath = "";
                foreach (var ipath in filepathlist)
                {
                    if (ipath.Contains(filename))
                    {
                        filepath = ipath;
                        break;
                    }
                }
                //Console.WriteLine("尝试打开： "+filepath);
                LaunchCommandLineApp(filepath);

            }
            if (e.KeyCode == Keys.F2)
            {
                string filename = treeView1.SelectedNode.Text;
                filename = filename.Replace(".docx","");
                //Console.WriteLine("复制到剪切板：" +"《"+filename +"》");
                SystemSounds.Exclamation.Play();
                Clipboard.SetText("《" + filename + "》");
            }
        }
        static void LaunchCommandLineApp(string wordpath)
        {
            Process process = new Process();
            process.StartInfo.FileName = "cmd.exe";
            //process.StartInfo.Arguments = "/c" + "start "+wordpath;
            process.StartInfo.Arguments = "/c" + "Explorer " + $"\"{wordpath}\"";
            process.StartInfo.UseShellExecute = false;   //是否使用操作系统shell启动 
            process.StartInfo.CreateNoWindow = false;   //是否在新窗口中启动该进程的值 (不显示程序窗口)
            process.Start();
            process.WaitForExit();  //等待程序执行完退出进程
            process.Close();
        }

        //黑白切换按钮
        private void button5_Click(object sender, EventArgs e)
        {
            if (button5.Text == "黑")
            {
                treeView1.ForeColor = Color.White;
                treeView1.BackColor = Color.Black;
                button5.Text = "白";
            }
            else
            {
                treeView1.ForeColor = Color.Black;
                treeView1.BackColor = Color.White;
                button5.Text = "黑";
            }
 
        }


        private void SearchForkeyWord(string v,string filepath)
        {
            string a = filepath.Substring(filepath.LastIndexOf("\\")+1);

            myutil doc = new myutil(filepath);
            var plist = doc.Find_Paragraph_for_plist(v);
            if (plist.Count == 0)
            {
                if (checkBox3.Checked == false)//未勾选“不显示未包含关键字的文件”
                {
                    ConsoleWriter.WriteGray(a);
                    Console.WriteLine();
                    ConsoleWriter.WriteRed("    文件内未查询到相应关键字。\n");
                }
            }
            else
            {
                ConsoleWriter.WriteGray(a);
                Console.WriteLine();
                //一些显示缓冲，慢慢的显示每个文件的结果
                if (checkBox2.Checked == true)
                {

                    ConsoleWriter.WriteGreen("匹配段落数：" + plist.Count + "\n");
                    return;
                }
                else if (checkBox1.Checked == false && plist.Count >= 6)
                {
                    ConsoleWriter.WriteGreen("匹配段落数(enter显示，n跳过)：" + plist.Count);
                    string abc = Console.ReadLine();
                    if (abc == "n")
                    {
                        return;
                    }
                }
                foreach (var p in plist)
                {
                    ConsoleWriter.Writehighlight(p.Text, v);
                }
            }


        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (textBox1.Text == "")
            {
                ConsoleWriter.WriteRed("请先输入搜索内容，空的搜不了");
                return;
            }
            ConsoleWriter.WriteSeperator('#');
            ConsoleWriter.WriteCyan("开始进行深度搜索："+textBox1.Text);
            foreach (var item in filepathlist)
            {
                string filenam = item.ToString();
                SearchForkeyWord(textBox1.Text, item);
            }
            ConsoleWriter.WriteGreen("搜索完毕");
            ConsoleWriter.WriteSeperator('#');
        }

        private void textBox1_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                button1.PerformClick();
            }
        }

        private void button7_Click(object sender, EventArgs e)
        {
            Clipboard.SetText(textBox1.Text);
        }

        private void button8_Click(object sender, EventArgs e)
        {
            treeView1.CollapseAll();
        }
    }
}
