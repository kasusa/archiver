using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace archiver
{
    public partial class Form_出报告申请 : Form
    {
        myutil doc ;//等保报告
        myutil tempo;//出报告模板
        public Form_出报告申请()
        {
            InitializeComponent();
            string tempo_path = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + '\\' + "Sample\\出报告申请_模板.docx";
            tempo = new myutil(tempo_path);

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

        private void button1_Click(object sender, EventArgs e)
        {

        }
    }
}
