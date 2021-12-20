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
            MyData.a_出报告日期 = doc.table_index_Get_cell_text(0, 2, 1).Trim();
            MyData.a_甲方单位 = doc.table_index_Get_cell_text(0, 0, 1).Trim();
            MyData.a_系统名称 = doc.table_index_Get_cell_text(1, 1, 1).Trim();
            MyData.a_下标数量 = 3;
            //todo 网络结构的描述要设定，用于填入报告-2.1.3网络描述中
            MyData.a_网络结构 = "";
            MyData.a_被测对象描述_1 = doc.table_index_Get_cell(4,3,1).Paragraphs[0].Text;
            MyData.a_被测对象描述_2 = doc.table_index_Get_cell(4,3,1).Paragraphs[0].Text;
            
            //MyData.a_xitong;
            mylog(MyData.a_系统名称);
            mylog(MyData.a_被测对象描述_1);

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

        }


        private void label1_Click(object sender, EventArgs e)
        {

        }


    }
}
