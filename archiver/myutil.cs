using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using Xceed.Document.NET;
using Xceed.Words.NET;

namespace archiver
{
    public class myutil
    {
       public  DocX document = null;
       public string path = "";
       public string savepath = "";

        public myutil(string str_path)
        {

            this.path = str_path;
            var document = DocX.Load(str_path);
            this.document = document;

            //获取savepath （同目录下加上-out）
            string text = path;
            string filename = text.Split('\\')[text.Split('\\').Count() - 1];
            savepath = text.Replace(filename, "");
            //string newfilename = filename.Split(".")[0] + "-out." + filename.Split(".")[1];
            //savepath = savepath + newfilename;

            // 把文件拷贝到temp文件夹防止被占用
            string sourceFile = str_path;
            string destinationFile = @"c:\temp\temp.docx";
            bool isrewrite = true; // true=覆盖已存在的同名文件,false则反之
            System.IO.File.Copy(sourceFile, destinationFile, isrewrite);
            this.path = destinationFile;
        }

        public Paragraph Find_Paragraph_for_p(string v)
        {
            foreach (var p in document.Paragraphs)
            {
                if (p.Text.Contains(v))
                {
                    //mylog("【找到:】" + p.Text + Environment.NewLine);
                    return p;
                }
            }
            return null;
        }

        public List<Paragraph> Find_Paragraph_for_plist( string v)
        {
            List < Paragraph > plist = new List < Paragraph >();
            foreach (var p in document.Paragraphs)
            {
                if (p.Text.Contains(v))
                {
                    //mylog("【找到:】" + p.Text + Environment.NewLine);
                    plist.Add( p);
                }
            }
            return plist;
        }
        public int Find_Paragraph_for_i( string v)
        {
            for (int i = 0; i < document.Paragraphs.Count; i++)
            {
                var p = document.Paragraphs[i];
                if (p.Text.Contains(v))
                {
                    //mylog("【找到:】" + p.Text + Environment.NewLine + "在" + i);
                    return i;
                }
            }
            return -1;
        }

        public List<int> Find_Paragraph_for_ilist( string v)
        {
            List<int> ilist = new List<int>();
            for (int i = 0; i < document.Paragraphs.Count; i++)
            {
                var p = document.Paragraphs[i];
                if (p.Text.Contains(v))
                {
                    //mylog("【找到:】" + p.Text + Environment.NewLine + "在" + i);
                    ilist.Add (i);
                }
            }
            return ilist;
        }

        //进行换字功能前请先设置好字典。
        //可以同时替换多个位置的文字。
        //前提是替换用变量前写成【xxx】这种格式，两边都是中文方括号即可
        //字典里面左边写【xxx】中的xxx 右边写上需要替换为的内容，后面会把方括号给替换掉的
        public Dictionary<string, string> _replacePatterns = new Dictionary<string, string>()
        {
            //例子
            //{ "报告编号", "P123445555" },
        };

        // 通过命令行向dict中写入数据
        // doc.cw_read_dictionary("报告编号");
        //输出 报告编号：
        // 获取输入 P123456
        //字典加入： 报告编号 ， P123456
        public string cw_read_dictionary(string a)
        {
            Console.WriteLine(a + ":");
            string x = Console.ReadLine();
            this._replacePatterns.Add(a, x);
            return x;
        }
        public void ReplaceTextWithText_all()
        {
            // Check if some of the replace patterns are used in the loaded document.
            if (document.FindUniqueByPattern(@"【(.*?)】", RegexOptions.IgnoreCase).Count > 0)
            {
                // Do the replacement of all the found tags and with green bold strings.
                document.ReplaceText("【(.*?)】", ReplaceFunc, false, RegexOptions.IgnoreCase);

                // Save this document to disk.
                Console.WriteLine("\tCreated: ReplacedTextWithText.docx\n");
            }
        }
        //这个是上一个函数会利用的子函数，原理我不懂我是照着官方文档copy的
        private string ReplaceFunc(string findStr)
        {
            if (_replacePatterns.ContainsKey(findStr))
            {
                return _replacePatterns[findStr];
            }
            return findStr;
        }



        public bool save()
        {
            document.SaveAs(savepath);
            return true;
        }
    }
}
