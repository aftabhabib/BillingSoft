using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;


namespace BillingSoft
{
    public partial class Form1 : Form
    {

        OleDbConnection con;
        OleDbDataAdapter da;
        private string text1;
        private string text2;
        private string text3;
        private string text4;
        private int v;

        public Form1(int s)
        {
            v = s;

            InitializeComponent();

            try
            {
                string path = " ";
                // this is the path that you are checking.
                if (v == 1)
                {
                    path = @"C:\Program Files (x86)\YatinPatel\BillingSoft. Setup\data LB.xlsx";
                }

                else
                {
                    path = @"C:\Program Files (x86)\YatinPatel\BillingSoft. Setup\data RP.xlsx";
                }



                con = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + path + ";Extended Properties=Excel 12.0");


                da = new OleDbDataAdapter("select * from [sheet1$]", con);

                da.Fill(data);

                if (v == 1)
                {
                    path = @"C:\Program Files (x86)\YatinPatel\BillingSoft. Setup\BILL LB.xlsm";
                }
                else
                {
                    path = @"C:\Program Files (x86)\YatinPatel\BillingSoft. Setup\BILL RP.xlsm";
                }

                openFileDialog2.FileName = path;

            }

            catch (Exception ex)
            {
                MessageBox.Show("can't load select meanullay");

            }

           


            this.Hide();

            createbill1 ss = new createbill1(openFileDialog2.FileName.ToString(), data, text1, text2, text3, text4, v);

            ss.Show();



        }

        public Form1(string text1, string text2, string text3, string text4,int v)
        {
            InitializeComponent();


       
            this.text1 = text1;
            this.text2 = text2;
            this.text3 = text3;
            this.text4 = text4;
            this.v = v;

        }

        public Form1()
        {
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
                 openFileDialog1.InitialDirectory = @"C:\";


                openFileDialog1.InitialDirectory = @"C:\";
                openFileDialog1.Title = "Browse excel data file";

                openFileDialog1.CheckFileExists = true;
                openFileDialog1.CheckPathExists = true;

                openFileDialog1.DefaultExt = "xlsx";
                openFileDialog1.Filter = "Excel Files(2003)|*.xls|Excel Files(2007)|*.xlsx|macro enablled excel files|*.xlsm";

                openFileDialog1.FilterIndex = 2;
                openFileDialog1.RestoreDirectory = true;




                if (openFileDialog1.ShowDialog() != DialogResult.Cancel)
                {

                    con = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + openFileDialog1.FileName.ToString() + ";Extended Properties=Excel 12.0");

                }


                else
                {

                    MessageBox.Show("opening cancel :(");
                }

            

            da = new OleDbDataAdapter("select * from [sheet1$]", con);

            da.Fill(data);
            
        }

        private void button2_Click(object sender, EventArgs e)
        {
                 openFileDialog2.InitialDirectory = @"C:\";
                openFileDialog2.Title = "Browse excel Sample Bill";

                openFileDialog2.CheckFileExists = true;
                openFileDialog2.CheckPathExists = true;

                openFileDialog2.DefaultExt = "xlsm";
                openFileDialog2.Filter = "macro enablled excel files|*.xlsm|Excel Files(2003)|*.xls|Excel Files(2007)|*.xlsx";

                openFileDialog2.FilterIndex = 2;
                openFileDialog2.RestoreDirectory = true;

                if (openFileDialog2.ShowDialog() != DialogResult.Cancel)
                {

                }
                else
                {
                    MessageBox.Show("opening cancel :(");
                    Form1 ss = new Form1();
                    ss.Show();
                    this.Hide();

                }

            

           // MessageBox.Show(openFileDialog2.FileName.ToString()+"done!!");

         }

        private void button3_Click(object sender, EventArgs e)
        {
            
            //MessageBox.Show(openFileDialog2.FileName.ToString() + "done!!");
            Form2 ss = new Form2(openFileDialog2.FileName.ToString(),data,text1,text2,text3,text4,v);
            ss.Show();
            this.Hide();
        }
    }
}
