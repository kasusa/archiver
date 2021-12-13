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
        }



        private void button1_Click(object sender, EventArgs e)
        {
           doc._replacePatterns.Clear();
           doc.cw_read_dictionary("项目编号");
           doc.cw_read_dictionary("备案号");
           doc.cw_read_dictionary("报告编号");
           doc.cw_read_dictionary("分数");
           doc.cw_read_dictionary("xxx公司");
           doc.cw_read_dictionary("xxx系统");
           doc.cw_read_dictionary("级别数字");
           doc.cw_read_dictionary("起始日期");
           doc.cw_read_dictionary("出报告日期");
           doc.cw_read_dictionary("我方人员");
           doc.ReplaceTextWithText_all();
           doc.save();
        }


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

        private void textBox2_TextChanged(object sender, EventArgs e)
        {
            
        }
    }
}
