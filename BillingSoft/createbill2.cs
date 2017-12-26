using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace BillingSoft
{
    public partial class Form2 : Form
    {
         private string str;
        private int number = 0;
        private double tamount = 0;
        private int count = 0;

        Microsoft.Office.Interop.Excel.Workbooks wrbks = null;
        Microsoft.Office.Interop.Excel.Workbook wrbk = null;
        Microsoft.Office.Interop.Excel.Worksheet wrst = null;

        Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();
        private DataSet data;
        private string text1;
        private string text2;
        private string text3;
        private string text4;
        private int v;

        public Form2()
        {
            InitializeComponent();
        }

         

        public Form2(string text1, string text2, string text3, string text4)
        {
            InitializeComponent();

            this.text1 = text1;
            this.text2 = text2;
            this.text3 = text3;
            this.text4 = text4;

        }

        public Form2(string s, DataSet data, string text1, string text2, string text3, string text4,int v)  
        {
            InitializeComponent();
            str = s;
            dta = data;
            // MessageBox.Show(str);

            // MessageBox.Show(data.Tables[0].Rows.Count.ToString());

            try
            {
                int q = dta.Tables[0].Rows.Count;
            }

            catch (Exception ex)
            {
                MessageBox.Show("can't load data please select excel fie from setting menu");
            }

                    

            excel.Application.Workbooks.Add(true);

            wrbks = excel.Workbooks;

            try
            {


                wrbk = wrbks.Open(str);

            wrst = wrbk.Worksheets[1];

                int q = wrst.Rows.Count;
            }
            catch (Exception ex)
            {
                MessageBox.Show("can't load sample bill Please select excel sample bill from setting menu :)");
            }


            this.v = v;


            this.text1 = text1;
            this.text2 = text2;
            this.text3 = text3;
            this.text4 = text4;
        }

        public Form2(string text1, string text2, string text3, string text4, int v) : this(text1, text2, text3, text4)
        {
            InitializeComponent();
             
            this.v = v;


            this.text1 = text1;
            this.text2 = text2;
            this.text3 = text3;
            this.text4 = text4;


        }

        private void Form2_Load(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {

            wrst.Range["F8"].Value = text1;
            wrst.Range["G8"].Value = text2;
            wrst.Range["F9"].Value = text3;
            wrst.Range["G9"].Value = text4;


            wrst.Range["G34"].Value = tamount;

            
            wrst.Range["G35"].Value = tamount * 0.09;

            wrst.Range["G36"].Value = tamount * 0.09;

            wrst.Range["A40"].Value = (tamount * 0.09 * 2) + tamount;
            

            //                wrst.Range[wrst.Cells[3, 1], wrst.Cells[3, 7]].Merge();

            /*
                saveFileDialog1.InitialDirectory = "C:\bills";
                saveFileDialog1.Title = "Save as excel file";
                saveFileDialog1.FileName = "1";
                saveFileDialog1.Filter = "macro enablled excel files|*.xlsm|Excel Files(2003)|*.xls|Excel Files(2007)|*.xlsx";
              */
              

                try
                {

                string str = "C:\\bills";

                if (v == 1)
                    str = str + "\\" + "LB\\";
                else
                    str = str + "\\" + "RP\\";

                    str += text1 + ".xlsm";


                if (!File.Exists(str))
                {
                    FileStream fs = File.Create(str);
                    fs.Close();
                }

                else
                {
                    MessageBox.Show("Bill is exist ");
                }

                //MessageBox.Show(v.ToString() + str);

                    excel.ActiveWorkbook.SaveCopyAs(str);
                    excel.ActiveWorkbook.Saved = true;


                    GC.Collect();
                    GC.WaitForPendingFinalizers();

                    excel.Quit();

                MessageBox.Show("SAVED!!");
            }


                catch(Exception ex)
                {
                    MessageBox.Show("saving aborted may be file c:\\bills not exist :(");
                }

            

           // MessageBox.Show("done!!");

        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void label3_Click(object sender, EventArgs e)
        {

        }

        private void monthCalendar1_DateChanged(object sender, DateRangeEventArgs e)
        {

        }

        private void button2_Click(object sender, EventArgs e)
        {
           
        }

        private void label8_Click(object sender, EventArgs e)
        {

        }

        private void button3_Click(object sender, EventArgs e)
        {
            

            int code;

            try
            {

                //MessageBox.Show(textBox4.Text+textBox1.Text);
                if (int.TryParse(textBox4.Text, out code))
                {
                    foreach (DataRow r in dta.Tables[0].Rows)

                    {
                        int temp;

                        int.TryParse(r[0].ToString(), out temp);

                        if (temp == code)
                        {
                            textBox5.Text = r[1].ToString();
                            textBox6.Text = r[2].ToString();
                        }


                    }
                }



                else
                { MessageBox.Show("You haven't attached file here or entered item is not integer!!! Please attach from setting menu."); }


            }

            catch (Exception ex)
            {
                MessageBox.Show("Can't open files please select form setting menu");


            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            
            count++;
            number++;

            string no = count.ToString();

            string code = textBox4.Text.ToString();
            string name = textBox5.Text.ToString();
            string price = textBox6.Text.ToString();
            string qty = textBox7.Text.ToString();


            int qt,cd;
            double pc;
            if(int.TryParse(textBox4.Text, out cd))
            {

            }
            else
            {
                MessageBox.Show("write code is in int form");
                return;
            }
            if(double.TryParse(textBox6.Text, out pc))
            {

            }
            else
            {
                MessageBox.Show("Wrint price in float form");
                return;
            }

            if(int.TryParse(textBox7.Text, out qt))
            {

            }
            else
            {
                MessageBox.Show("write qty is in int form");
                return;
            }

            double amount = qt * pc;

            string amt = amount.ToString();

            tamount += amount;

            
            string a = 'A' + (count+17).ToString();

            wrst.Range[a].Value = number;


            a = 'B' + (count+17).ToString();

            wrst.Range[a].Value = code;

            a = 'C' + (count+17).ToString() ;

            wrst.Range[a].Value =  name;

            a = 'E' + (count+17).ToString();

            wrst.Range[a].Value = qt;

            a = 'F' + (count+17).ToString();

            wrst.Range[a].Value = pc;

            a = 'G' + (count+17).ToString() ;

            wrst.Range[a].Value = amount;

            if(v==2)
            {
                count++;
                a = 'C' + (count + 17).ToString();
                wrst.Range[a].Value = "HSN-84669200";
            }

            textBox4.Text = null;
            textBox5.Text = null;
            textBox6.Text = null;
            textBox7.Text = null;

            textBox8.Text = tamount.ToString();



        }

        private void button5_Click(object sender, EventArgs e)
        {
            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();

            // Open the Workbook:
            /*
            openFileDialog1.InitialDirectory = @"C:\";
            openFileDialog1.Title = "Browse excel data file";

            openFileDialog1.CheckFileExists = true;
            openFileDialog1.CheckPathExists = true;

            openFileDialog1.DefaultExt = "xlsm";
            openFileDialog1.Filter = "macro enablled excel files|*.xlsm|Excel Files(2003)|*.xls|Excel Files(2007)|*.xlsx";

            openFileDialog1.FilterIndex = 2;
            openFileDialog1.RestoreDirectory = true;


            if (openFileDialog1.ShowDialog() != DialogResult.Cancel)
            {
            */

            string st;

            try {

                st = "C:\\bills";

                if (v == 1)
                    st = st + "\\" + "LB\\";
                else
                    st = st + "\\" + "RP\\";

                st += text1 + ".xlsm";


                if (!File.Exists(st))
                {
                    MessageBox.Show("Please First Save Bill");
                    return;
                }
                

                Microsoft.Office.Interop.Excel.Workbook wb = excelApp.Workbooks.Open(
                st,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing);

                // Get the first worksheet.
                // (Excel uses base 1 indexing, not base 0.)
                Microsoft.Office.Interop.Excel.Worksheet ws = (Microsoft.Office.Interop.Excel.Worksheet)wb.Worksheets[1];

                // Print out 1 copy to the default printer:

                bool userDidntCancel =
        excelApp.Dialogs[Microsoft.Office.Interop.Excel.XlBuiltInDialog.xlDialogPrint].Show(
            Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
            Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
            Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
            Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
            Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
            Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

                
                
                if(!userDidntCancel)
                {
                    excelApp.Quit();
                    return;
                }

                excelApp.Quit();

            }

            catch (Exception ex)
            {
                MessageBox.Show("opening cancel :(");

            }

            excel.Application.Workbooks.Add(true);

            wrbks = excel.Workbooks;

            try
            {


                st = "C:\\bills";

                if (v == 1)
                    st = st + "\\" + "LB\\";
                else
                    st = st + "\\" + "RP\\";

                st += text1 + ".xlsm";

                
                wrbk = wrbks.Open(st);

                wrst = wrbk.Worksheets[1];

            //    MessageBox.Show(st);


                wrst.Range["G3"].Value = "ORIGINAL";

                st = st.Substring(0, st.Length - 5);
                st += "o" + ".xlsm";

                excel.ActiveWorkbook.SaveCopyAs(st);
                excel.ActiveWorkbook.Saved = true;


                GC.Collect();
                GC.WaitForPendingFinalizers();

                excel.Quit();




            }
            catch (Exception ex)
            {
                MessageBox.Show("cant open file");
            }

            try
            {

                st = "C:\\bills";

                if (v == 1)
                    st = st + "\\" + "LB\\";
                else
                    st = st + "\\" + "RP\\";

                st += text1 + "o.xlsm";


                if (!File.Exists(st))
                {
                    MessageBox.Show("Please First Save Bill");
                    return;
                }


                Microsoft.Office.Interop.Excel.Workbook wb = excelApp.Workbooks.Open(
                st,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing);

                // Get the first worksheet.
                // (Excel uses base 1 indexing, not base 0.)
                Microsoft.Office.Interop.Excel.Worksheet ws = (Microsoft.Office.Interop.Excel.Worksheet)wb.Worksheets[1];

                // Print out 1 copy to the default printer:

                bool userDidntCancel =
        excelApp.Dialogs[Microsoft.Office.Interop.Excel.XlBuiltInDialog.xlDialogPrint].Show(
            Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
            Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
            Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
            Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
            Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
            Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);



                if (!userDidntCancel)
                {
                    excelApp.Quit();
                    return;
                }

                excelApp.Quit();
            }

            catch (Exception ex)
            {
                MessageBox.Show("opening cancel :(");

            }



        }

        private void button6_Click(object sender, EventArgs e)
        {
            createbill1 ss = new createbill1(v);
            ss.Show();
            this.Hide();
        }

        private void button2_Click_1(object sender, EventArgs e)
        {
            Form1 ss = new Form1(text1,text2,text3,text4,v);
            ss.Show();
            this.Hide();
        }

        private void textBox4_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }
    }
}



