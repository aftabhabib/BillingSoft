using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace BillingSoft
{
    public partial class First : Form
    {
        private int v;

        public First()
        {
            InitializeComponent();
        }

        public First(int v)
        {
            this.v = v;
        }

        private void button3_Click(object sender, EventArgs e)
        {

            this.Hide();

            Form1 s = new Form1(1);

        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.Hide();

            Form1 s = new Form1(2);
        }

        private void Form4_Load(object sender, EventArgs e)
        {

        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void richTextBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {

        }
    }
}
