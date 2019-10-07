using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Excel;
using System.Data.OleDb;

namespace Knn_EksikVeriTamamlama
{
    public partial class Form1 : Form
    {
        Test test = new Test();
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string adres = @"C:\Users\HP PC\Desktop\Dersler2019\4.Sınıf\Meta-Sezgisel\egitim.xls";
            Excel.Application xlOrn = new Microsoft.Office.Interop.Excel.Application();
            if (xlOrn == null)
            {
                MessageBox.Show("Excel yüklü değil!!");
                return;
            }

            Excel.Workbook xlWorkBook;
            Excel.Workbook xWorkBookTest;
            Excel.Worksheet xlWorkSheet;
            object misValue = System.Reflection.Missing.Value;

            xlWorkBook = xlOrn.Workbooks.Open(adres, 0, true, 5, "", "", true, Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
           
        

            Excel.Range xlRange = xlWorkSheet.UsedRange; // get the entire used range

            int numberOfRows = xlRange.Rows.Count;
            int numberOfCols = xlRange.Columns.Count;
          
            double[,] veri_dizi = new double[numberOfRows-1, numberOfCols];
            int indis1, indis2;
      
            
            MessageBox.Show(numberOfRows+" "+numberOfCols);
            for (int i = 2; i <= numberOfRows; i++)//5
            {
                for(int p=1; p<= numberOfCols; p++)//2
                {
                    if (i == 2)
                    {
                        indis1 = 0;
                        indis2 = p - 1;
                        veri_dizi[indis1, indis2] = (double)(xlRange.Cells[i, p].Value2());
                    }
                    else
                    {
                        indis1 = i-2;
                        indis2 = p-1;
                        veri_dizi[indis1, indis2] = (double)(xlRange.Cells[i, p].Value2());
                    }
                }
       
            }

          //  double[,] test_dizi = test.TestIslemleri();
   
            int k = Convert.ToInt32(textBox1.Text);
            test.kontrol(veri_dizi,k, numberOfRows, numberOfCols);
        

        }
    }
}
