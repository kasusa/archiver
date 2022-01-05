using archiver.ConsoleColorWriter;
using System;
using System.Collections.Generic;
using System.IO;
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
        public DocX document = null;
        public string path = "";
        public List<Table> tables;
        public System.Collections.ObjectModel.ReadOnlyCollection<Paragraph> Paragraphs;
        
        public myutil(string str_path)
        {
            //如果没有创建temp文件夹（用于暂存文件）则自动创建
            string tempath = @$"c:\temp";
            if (!Directory.Exists(tempath))
                Directory.CreateDirectory(tempath);

            this.path = str_path;
            //复制到这里c、temp 防被占用
            string sourceFile = str_path;
            string filename = str_path.Split('\\')[str_path.Split('\\').Count() - 1];
            string destinationFile = @$"c:\temp\{filename}";
            bool isrewrite = true; // true=覆盖已存在的同名文件,false则反之
            System.IO.File.Copy(sourceFile, destinationFile, isrewrite);
            this.path = destinationFile;

            var document = DocX.Load(path);
            this.document = document;
            this.tables = document.Tables;
            this.Paragraphs = document.Paragraphs;
            
        }

        public void remove_p(Paragraph paragraph)
        {
            if (paragraph != null)
            {
                document.RemoveParagraph(paragraph);

            }
        }



        #region string操作    

        /// <summary>
        /// 获取str中指定搜索字段后面的字符串
        /// </summary>
        /// <param name="str"></param>
        /// <param name="str_search"></param>
        /// <param name="len"></param>
        /// <returns></returns>
        public static string get_string_after(string str, string str_search, int len)
        {
            string x = str.Substring(str.LastIndexOf(str_search) + str_search.Length, len);
            return x;
        }

        public static string get_string_before(string str, string str_search, int len)
        {
            string x = str.Substring(0,str.LastIndexOf(str_search) - len);
            return x;
        }

        /// <summary>
        /// 获取字符串-中间部分
        /// </summary>
        /// <param name="str"></param>
        /// <param name="str_search"></param>
        /// <param name="str_search2"></param>
        /// <returns></returns>
        public static string get_string_bewteen(string str, string str_search, string str_search2)
        {
            //截取前面一段
            string x = str.Substring(0, str.LastIndexOf(str_search2));
            //找到str_search末位置
            int startindex = str.LastIndexOf(str_search) + str_search.Length;
            //计算所需str长度
            int len = x.Length - startindex;
            //获取所需字符串
            x = str.Substring(startindex, len);
            return x;
        }
        #endregion


        #region table操作

        public List<Table> findTableList(string v1)
        {
            v1 = v1.Replace(" ", "").Replace("\t", "");
            List<Table> tlist = new List<Table>();
            //Console.WriteLine("开始寻找表头是 :"+v1+ "的表格");
            for (int i = 0; i < tables.Count; i++)
            {
                string rowstring = "";
                for (int j = 0; j < tables[i].ColumnCount; j++)
                {
                    if (rowstring.Contains("安全通用要求测评指标类"))
                    {
                        break;
                    }
                    rowstring += cell_get_text(table_Get_cell(tables[i], 0, j));
                }
                Console.WriteLine("look@me:" + i + ":" + rowstring);

                if (rowstring== v1)
                {
                    //Console.WriteLine("找到了table"+i);
                    tlist.Add(tables[i]);
                }
            }
                Console.WriteLine("找到table个数：" + tlist.Count);

            return tlist;
        }

        /// <summary>
        /// 获取table中的指定cell
        /// </summary>
        /// <param name="table"></param>
        /// <param name="rowindex"></param>
        /// <param name="cellindex"></param>
        /// <returns></returns>
        public Cell table_Get_cell(Table table, int rowindex, int cellindex)
        {
            try
            {
                return table.Rows[rowindex].Cells[cellindex];
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                Console.WriteLine("cell is not exist!");
                return null;
            }
        }

        public Cell table_index_Get_cell(int i, int rowindex, int cellindex)
        {
            Table table = tables[i];
            try
            {
                return table.Rows[rowindex].Cells[cellindex];
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                Console.WriteLine("cell is not exist!");
                return null;
            }

        }

        /// <summary>
        /// 获取table指定 rowindex,cellindex的文字内容
        /// </summary>
        /// <param name="table"></param>
        /// <param name="rowindex"></param>
        /// <param name="cellindex"></param>
        /// <returns></returns>
        public string table_Get_cell_text(Table table, int rowindex, int cellindex)
        {
            Cell cell = table.Rows[rowindex].Cells[cellindex];
            return cell_get_text(cell);
        }
        public string table_index_Get_cell_text(int tableindex, int rowindex, int cellindex)
        {
            Cell cell = this.tables[tableindex].Rows[rowindex].Cells[cellindex];
            return cell_get_text(cell);
        }

        /// <summary>
        /// 给table增加一行合并行
        /// </summary>
        /// <param name="table"></param>
        /// <param name="v"></param>
        public void table_add_merged_cell(Table table, string v)
        {
            table.InsertRow();
            Row lastrow = table.Rows[table.RowCount - 1];

            lastrow.MergeCells(0,lastrow.Cells.Count -1 );

            Cell cell = lastrow.Cells[0];
            cell_settext(cell,v);

        }

        /// <summary>
        /// 把table最后一列杠杠(除了第一行)
        /// </summary>
        /// <param name="table"></param>
        public void table_lastcell_ganggang(Table table)
        {
            int linelength = table.Rows[0].Cells.Count;
            int lastindex = linelength - 1;
            int rowcount = table.RowCount;
            for (int i = 1; i < rowcount; i++)
            {
                Cell cell = table.Rows[i].Cells[lastindex];
                cell_settext(cell, "--");
            }
        }
        /// <summary>
        /// 获取table第一行文字
        /// </summary>
        /// <param name="table"></param>
        /// <returns></returns>
        public string table_get_first_row_text(Table table)
        {
            string text = "";
            var row = table.Rows.First();
            int cellcount = row.Cells.Count;
            for (int i = 0; i < cellcount; i++)
            {
               text += " "+ cell_get_text(table_Get_cell(table, 0, i));
            }
            Console.WriteLine(text);
            return text;
        }
        #endregion

        #region cell操作

        /// <summary>
        /// 获取cell的段落(string列表)
        /// </summary>
        /// <param name="cell"></param>
        /// <returns></returns>
        public List<string> cell_get_textList(Cell cell)
        {
            if(cell == null) return null;
            List<string> texts = new List<string>();

            foreach (var p in cell.Paragraphs)
            {
                texts.Add(p.Text);
            }
            return texts;
        }

        /// <summary>
        /// 获取cell的内容文字（所有段落）
        /// </summary>
        /// <param name="cell"></param>
        /// <returns></returns>
        public string cell_get_text(Cell cell)
        {

            if (cell == null) return null;
            string texts = "";

            foreach (var p in cell.Paragraphs)
            {
                texts += p.Text;
            }
            return texts;
        }

        /// <summary>
        /// 清空这个cell中的文字
        /// </summary>
        /// <param name="cell"></param>
        public void cell_clear(Cell cell)
        {
            if (cell == null) return;

            var paragraphs = cell.Paragraphs;
            foreach (var paragraph in paragraphs)
            {
                cell.RemoveParagraph(paragraph);
            }
            cell.InsertParagraph("");//最后添加一个防止word报错
           
        }

        /// <summary>
        /// 为空cell添加内容（5号，居中）
        /// </summary>
        /// <param name="cell"></param>
        /// <param name="v"></param>
        public void cell_settext(Cell cell, string v)
        {
            cell_clear(cell);
            cell.Paragraphs[0].Append(v);

            //居中五号
            cell.Paragraphs[0].Alignment = Alignment.center;
            cell.Paragraphs[0].FontSize(10.5d);
        }

        public void cell_settext_Big(Cell cell, string v)
        {
            cell_clear(cell);
            cell.Paragraphs[0].Append(v);

            //居中四号
            cell.Paragraphs[0].Alignment = Alignment.center;
            cell.Paragraphs[0].FontSize(14d);
            cell.Paragraphs[0].Bold();
            cell.Paragraphs[0].Font("宋体");

        }


        #endregion

        #region 年月日大写
        /// <summary>
        /// 输入日期（2021年12月5日）输出 全大写日期 二〇二一年十二月五日
        /// </summary>
        /// <param name="date"></param>
        /// <returns></returns>
        public static string datetochinese(string date)
        {
            Dictionary<string, string> chinese_data_dict = new Dictionary<string, string>() {
            {"1", "一" },
            {"2", "二"},
            {"3", "三"},
            {"4", "四"},
            {"5", "五"},
            {"6", "六"},
            {"7", "七"},
            {"8", "八"},
            {"9", "九"},
            {"0", "〇"}
            };
            Dictionary<string, string> chinese_data2_dict = new Dictionary<string, string>()
            {
                {"一", "十"},
                {"二", "二十"},
                {"三", "三十"},
            };

            foreach(var chr in date)
            {
                string c = chr.ToString();
                if (chinese_data_dict.ContainsKey(c))
                {
                    date = date.Replace(c, chinese_data_dict[c]);
                }
            }

            string y = date.Split("年")[0];
            string m = date.Split("年")[1].Split("月")[0];
            string temp = "";
            if(m.Length == 2)
            {
                temp += chinese_data2_dict[m[0].ToString()];
                temp += m[1];
                m = temp;
            }


            string d = date.Split("月")[1].Split("日")[0];
            temp = "";
            if (d.Length == 2)
            {
                temp += chinese_data2_dict[d[0].ToString()];
                temp += d[1];
                d = temp;
            }


            date = y + "年" + m + "月" + d + "日";
            return date;
        }
        #endregion


        #region 段落查找


        public Paragraph Find_Paragraph_for_p(string v)
        {
            foreach (var p in document.Paragraphs)
            {
                if (p.Text.Contains(v))
                {
                    //Console.WriteLine("【找到:】" + p.Text + Environment.NewLine);
                    return p;
                }
            }
            return null;
        }

        public string Find_Paragraph_for_text(string v,int count = 1)
        {
            foreach (var p in document.Paragraphs)
            {
                if (p.Text.Contains(v))
                {
                    count--;
                    if (count == 0)
                    {
                        return p.Text;
                    }
                }
            }
            return "";
        }

        public List<Paragraph> Find_Paragraph_for_plist( string v)
        {
            List < Paragraph > plist = new List < Paragraph >();
            foreach (var p in document.Paragraphs)
            {
                if (p.Text.Contains(v))
                {
                    Console.WriteLine("【找到:】" + p.Text + Environment.NewLine);
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
        #endregion

        #region 段落替换    
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
        public void write_dictionary(string a,string x)
        {
            //a 被替换
            //x 替换成
            Console.WriteLine($"{a}\t:{x}");
            this._replacePatterns.Add(a, x);
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
        public void ReplaceTextWithText_all_noBracket()
        {
            foreach(var keyValuePair in _replacePatterns)
            {
                document.ReplaceText(keyValuePair.Key, keyValuePair.Value);
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

        #endregion

        #region save

        //使用原文件名保存
        public bool save()
        {
            checkoutdir();
            string text = this.path;
            string filename = text.Split('\\')[text.Split('\\').Count() - 1];
            string savepath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "//" + "out" + "//" + filename;
            document.SaveAs(savepath);
            return true;
        }
        //新建名称保存
        public bool save(string filename)
        {
            checkoutdir();
            string savepath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "//" + "out" + "//" + filename;
            document.SaveAs(savepath);
            return true;
        }
        private void checkoutdir()
        {
            string outpath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + '\\' + "out";
            if (!Directory.Exists(outpath))
                Directory.CreateDirectory(outpath);
        }

        #endregion

    }
}
