using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Text;
using System.Windows.Forms;

namespace archiver
{
    public partial class Form_出报告申请 : Form
    {
        myutil doc ;//等保报告
        myutil tempo;//出报告模板
        string str_公司 = "xxx公司";
        string str_系统 = "xxx系统";
        public Form_出报告申请()
        {
            InitializeComponent();
            try
            {
                string tempo_path = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + '\\' + "Sample\\出报告申请_模板.docx";
                tempo = new myutil(tempo_path);
            }
            catch (Exception)
            {
                toolStripStatusLabel1.Text = "获取模板 失败，检查桌面的Sample\\出报告申请_模板.docx";
                throw;
            }
            toolStripStatusLabel1.Text = "获取模板 成功";

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
            string a = "";
            tempo._replacePatterns.Clear();
            a = doc.Paragraphs[0].Text;
            a = myutil.get_string_after(a, "报告编号", "31001322060-19012-21-0092-01".Length);
            tempo.write_dictionary("报告编号", a);
            a = myutil.get_string_before(a, "-0092", "31001322060-19012".Length);
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
            Console.WriteLine("还请输入：");
            tempo.cw_read_dictionary("项目编号");
            tempo.cw_read_dictionary("起始日期");
            tempo.cw_read_dictionary("出报告日期");
            tempo.cw_read_dictionary("我方人员");
            tempo.ReplaceTextWithText_all();
            this.Show();
            doc.save($"出报告申请_{str_公司}_{str_系统}.docx");
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
            tempo.cw_read_dictionary("xxx公司");
            tempo.cw_read_dictionary("xxx系统");
            tempo.cw_read_dictionary("级别数字");
            tempo.cw_read_dictionary("起始日期");
            tempo.cw_read_dictionary("出报告日期");
            tempo.cw_read_dictionary("我方人员");
            tempo.ReplaceTextWithText_all();
            Console.WriteLine("录入信息完毕，请回到GUI窗体继续操作。（请勿关闭本窗）");
            toolStripStatusLabel1.Text = "录入信息完毕";
            this.Show();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //出报告申请_xxx公司_xxx系统.docx
            string filename = $"出报告申请_{str_公司}_{str_系统}.docx";
            tempo.save(filename);
            toolStripStatusLabel1.Text = "保存到桌面out文件夹了。";
        }
    }
}
