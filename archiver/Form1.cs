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

namespace archiver
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Console.WriteLine("Hello, World!");
            string path = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\1.docx";
            label1.Text = path;
            var document = DocX.Load(path);


            var p = Find_Paragraph_for_i(document, "我的后面是");
            if (p == null )
            {
                mylog("啥也找不到");
            }
            //List<int> a = document.FindAll("我的听啊俺");
            //foreach (var index in a)
            //{
            //    textBoxlog.Text += a.t;
            //    textBoxlog.Text +=  document.Paragraphs[index].Text;

            //}
        }

        public Paragraph Find_Paragraph_for_p(DocX document, string v)
        {
            foreach (var p in document.Paragraphs)
            {
                if (p.Text.Contains(v))
                {
                    mylog("【找到:】"+p.Text + Environment.NewLine);
                    return p;
                }
            }
            return null;
        }
        public int Find_Paragraph_for_i(DocX document, string v)
        {
            for (int i = 0; i < document.Paragraphs.Count; i++)
            {
                var p = document.Paragraphs[i];
                if (p.Text.Contains(v))
                {
                    mylog("【找到:】" + p.Text + Environment.NewLine+"在"+i);
                    return i;
                }
            }
            return -1;
        }

        private void mylog(string v)
        {
            textBoxlog.Text += v+Environment.NewLine;
        }
    }
}
