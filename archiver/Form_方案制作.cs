using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace archiver
{
    public partial class Form_方案制作 : Form
    {
        myutil doc;//等保报告
        myutil tempo;//出报告模板
        public Form_方案制作()
        {
            InitializeComponent();
        }

        private void Form_方案制作_Load(object sender, EventArgs e)
        {
            loadSample();

        }

        private void loadSample()
        {
            string tempo_path = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + @$"\Sample\方案";
            if (!System.IO.Directory.Exists(tempo_path))
            {
                toolStripStatusLabel1.ForeColor = System.Drawing.Color.Red;
                toolStripStatusLabel1.Text = "获取模板 失败，检查桌面的Sample\\出报告申请_模板.docx";
            }
            else
            {
                //tempo = new myutil(tempo_path);
                toolStripStatusLabel1.Text = "获取模板 成功";
            }
        }

    }
}
