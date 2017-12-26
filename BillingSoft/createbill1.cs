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
   
    public partial class createbill1 : Form
    {
        int cnt = 0;
        private int v;
        private string v1;
        private DataSet data;
        private string text1;
        private string text2;
        private string text3;
        private string text4;

        public createbill1()
        {
            InitializeComponent();
            
         
        }

        public createbill1(int v)
        {
            InitializeComponent();

            this.v = v;
        }

        public createbill1(string v1, DataSet data, string text1, string text2, string text3, string text4, int v2)
        {
            InitializeComponent();
            this.v1 = v1;
            this.data = data;
            this.text1 = text1;
            this.text2 = text2;
            this.text3 = text3;
            this.text4 = text4;
            this.v = v2;
        }

        private void Form3_Load(object sender, EventArgs e)
        {

        }

        private void label3_Click(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            if(textBox2.Text=="")
            {
                MessageBox.Show("enter bill no.");
            }

            else if (textBox3.Text=="")
            {
                MessageBox.Show("enter challen no.");
            }

            else if (textBox1.Text == "")
            {
                MessageBox.Show("enter bill date");
            }

            else if (textBox4.Text == "")
            {
                MessageBox.Show("enter challen date");
            }

            else
            {

                Form2 ss = new Form2(v1, data,textBox2.Text, textBox3.Text, textBox1.Text, textBox4.Text, v);

                ss.Show();
                this.Hide();

            }


        }


        private void label7_Click(object sender, EventArgs e)
        {

        }

        private void label6_Click(object sender, EventArgs e)
        {

        }

        private void dateTimePicker1_ValueChanged(object sender, EventArgs e)
        {
            string theDate = dateTimePicker1.Value.ToShortDateString();
            int x = theDate.ToLower().IndexOf('/');
            theDate = theDate.Remove(x, 1);
            int y = theDate.ToLower().IndexOf('/');
            string str = theDate.Substring(x, y - x);
            theDate = theDate.Remove(y+1,2);
            theDate = theDate.Remove(x, y - x);
            theDate = str + '/' + theDate;
            textBox1.Text = theDate;

        }

        private void dateTimePicker2_ValueChanged(object sender, EventArgs e)
        {
            string theDate = dateTimePicker2.Value.ToShortDateString();

            int x = theDate.ToLower().IndexOf('/');
            theDate = theDate.Remove(x, 1);
            int y = theDate.ToLower().IndexOf('/');
            string str = theDate.Substring(x, y - x);
            theDate = theDate.Remove(y + 1, 2);
            theDate = theDate.Remove(x, y - x);
            theDate = str + '/' + theDate;

            textBox4.Text = theDate;

        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
        }

        private void textBox4_TextChanged(object sender, EventArgs e)
        {
        }

        private void button2_Click(object sender, EventArgs e)
        {
            First ss = new First();
            ss.Show();
            this.Hide();
        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {

        }
    }
}
