using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;

namespace archiver
{
    public partial class Form_main : Form
    {
        public Form_main()
        {
            InitializeComponent();
            AllocConsole();

        }

        [DllImport("kernel32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        static extern bool AllocConsole();
        private void button1_Click(object sender, EventArgs e)
        {

            this.Hide();
            Form a = new Form_出报告申请();
            a.ShowDialog();
            this.Show();



        }

        private void button_Click(object sender, EventArgs e)
        {

        }

        private void button0_Click(object sender, EventArgs e)
        {
            this.Hide();
            Form a = new Form1();
            a.ShowDialog();
            this.Show();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.Hide();
            Form a = new Form_方案制作();
            a.ShowDialog();
            this.Show();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            this.Hide();
            Form a = new Form_replaceForAll();
            a.ShowDialog();
            this.Show();
        }

        private void groupBox1_Enter(object sender, EventArgs e)
        {

        }
    }
}
