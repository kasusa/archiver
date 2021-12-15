using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Xceed.Document.NET;
using Xceed.Words.NET;
using System.Runtime.InteropServices;


namespace archiver
{
    public partial class Form1 : Form
    {
        myutil doc;
        public Form1()
        {
            InitializeComponent();
            AllocConsole();
        }
        [DllImport("kernel32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        static extern bool AllocConsole();


        private void button1_Click(object sender, EventArgs e)
        {

        }

        #region hide


        private void mylog(string v)
        {
            textBoxlog.Text += v+Environment.NewLine;
        }

        private void textBox1_DragDrop(object sender, DragEventArgs e)
        {
            textBox1.Text = ((System.Array)e.Data.GetData(DataFormats.FileDrop)).GetValue(0).ToString();

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

        //自动生成输出目录
        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            // 更换源文件，重新生成doc
            string path = textBox1.Text;
            doc = new myutil(path);
        }
        #endregion

        private void button2_Click(object sender, EventArgs e)
        {
            mylog(myutil. get_string_after(textBox_input.Text, "离开家", "；拉伸看解法，".Length));

        }


        private void label1_Click(object sender, EventArgs e)
        {

        }


    }
}
