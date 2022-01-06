using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace archiver
{
    public partial class form_loading : Form
    {
        List<Bitmap> _bitmapList = new List<Bitmap>();
        public form_loading(List<Bitmap> bitmapList)
        {
            _bitmapList = bitmapList;
            InitializeComponent();
            //增加listbox中的选项
            int i = 0;
            //int i = bitmapList.Count-1;
            foreach (var bitmap in bitmapList)
            {
                listBox1.Items.Add("图片-"+i);
                i++;
            }
            listBox1.SelectedIndex = 0;
        }

        private void listBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            int i = listBox1.SelectedIndex;
            pictureBox1.Image = _bitmapList[i];
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Clipboard.SetImage(pictureBox1.Image);
        }
    }
}
