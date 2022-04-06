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
using archiver.ConsoleColorWriter;
using System.Text.RegularExpressions;

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
            //List<string> a =  doc.document.FindUniqueByPattern("[^ -~]*hello", RegexOptions.None);
            Table t = doc.findTableList("被测对象")[0];
            string a = doc.table_Get_cell_text(t,0,1);
            Paragraph p = doc.Paragraphs[0];

            textBox_input.Text = a;
        }


        private void label1_Click(object sender, EventArgs e)
        {
            
        }
        /// <summary>
        /// 提取和处理日期
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button3_Click(object sender, EventArgs e)
        {
            //string  a = doc.Find_Paragraph_for_text("方案编制过程。");
            string a = textBox_input.Text = "1、2021年12月28日～2021年12月29日，测评准备过程。";
            textBoxlog.Text = date_process(a);
        }

        private string date_process(string a)
        {
            //提取日期（结束日期）
            ConsoleWriter.WriteCyan("在字符串中寻找日期："+a);

            string patternA = @"\d\d\d\d年(\d)*月(\d)*日";
            string patternB = @"(\d)*月(\d)*日";
            //如果获取的短日期个数为2，但是长日期仅有1个，那么就是如2021年12月1日～12月1日这种写法
            //如果长短日期都只有一个，那么就是2021年12月1日这种写法（只有一天之类的）

            int shortDcount = 0;
            int LongDcount = 0;

            Regex rg = new Regex(patternB);
            MatchCollection matchedShortDate = rg.Matches(a);
            shortDcount = matchedShortDate.Count;
            //Console.WriteLine("shortDcount" + shortDcount);
             rg = new Regex(patternA);
            MatchCollection matchedLongDate = rg.Matches(a);
            LongDcount = matchedLongDate.Count;
            //Console.WriteLine("LongDcount"+ LongDcount);

            if (shortDcount > LongDcount)
            {
                string year = matchedLongDate[LongDcount - 1].Value.Substring(0, 5);
                string MandD = matchedShortDate[shortDcount - 1].Value;
                 a = year+ MandD;
            }else if (shortDcount == LongDcount)
            {
                a = matchedLongDate[LongDcount - 1].Value;
            }
            Console.WriteLine(a);
            return a;
        }

        private void button4_Click(object sender, EventArgs e)
        {
            //string  a = doc.Find_Paragraph_for_text("方案编制过程。");
            string a = textBox_input.Text;
            textBoxlog.Text = date_process(a);
        }
    }
}
