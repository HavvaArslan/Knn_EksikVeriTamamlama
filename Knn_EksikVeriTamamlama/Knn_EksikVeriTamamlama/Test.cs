using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Excel;
using System.Data.OleDb;
using System.Windows.Forms;


namespace Knn_EksikVeriTamamlama
{
     public class Test
    {
        public int row,col;
        int index = 0,count=0;
        public static double[] temp;
        int[] index_table;
        int[] sutun_table;
        double[] siralanmis_dizi;
        double[] tut_dizi;
        double[] indisler_dizi;
        double[] point;
        double gercek = 0;
        double ortalama_hata=0;
        Excel.Application ExcelUygulama;
        Excel.Workbook ExcelProje;
        Excel.Worksheet ExcelSayfa;
        object Missing = System.Reflection.Missing.Value;
        Excel.Range ExcelRange;


        public double[,] TestIslemleri(int egitim_row, int egitim_col)
        {
            string adres = @"C:\Users\HP PC\Desktop\Dersler2019\4.Sınıf\Meta-Sezgisel\test2.xls";
            Excel.Application xlOrn = new Microsoft.Office.Interop.Excel.Application();
          
            Excel.Workbook xWorkBookTest;
            Excel.Worksheet xlWorkSheet;
            object misValue = System.Reflection.Missing.Value;

            xWorkBookTest = xlOrn.Workbooks.Open(adres, 0, true, 5, "", "", true, Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
            xlWorkSheet = (Excel.Worksheet)xWorkBookTest.Worksheets.get_Item(1);
            Excel.Range xlRange = xlWorkSheet.UsedRange; 
            

            int numberOfRows = xlRange.Rows.Count;
            int numberOfCols = xlRange.Columns.Count;
         
            int indis1 = 0, indis2 = 0;
            row = numberOfRows; col = numberOfCols;
            double[,] test_dizi = new double[numberOfRows - 1, numberOfCols];
            indis1 = 0; indis2 = 0;

            for (int i = 2; i <= numberOfRows; i++)//5
            {
                for (int p = 1; p <= numberOfCols; p++)//2
                {
                    if (i == 2)
                    {
                        if (xlRange.Cells[i, p].Value2() == null)
                        {
                            indis1 = 0;
                            indis2 = p - 1;
                            test_dizi[indis1, indis2] =-1;
                        }
                        else
                        {
                            indis1 = 0;
                            indis2 = p - 1;
                            test_dizi[indis1, indis2] = (double)(xlRange.Cells[i, p].Value2());
                        }
                    }
                    else
                    {
                        if (xlRange.Cells[i, p].Value2() == null)
                        {
                            indis1 = i - 2;
                            indis2 = p - 1;
                            test_dizi[indis1, indis2] = -1;
                        }
                        else
                        {
                            indis1 = i - 2;
                            indis2 = p - 1;
                            test_dizi[indis1, indis2] = (double)(xlRange.Cells[i, p].Value2());
                        }    
                    }
                }
            }

            temp = new double[(egitim_row - 2)];
            index_table = new int[(egitim_row - 2)];
            sutun_table = new int[(egitim_row - 2)];
            indisler_dizi = new double[(egitim_row - 2)];
            return test_dizi;

        }
        public double[] oklit(int satir,int sutun,double[,] egitim_seti,int egitim_row)
        {
           
            double[,] satirlik_veri = new double[1, col];
            for (int k=0; k < 1; k++)
            {
                for (int m = 0; m < col; m++)
                {
                    satirlik_veri[k, m] = egitim_seti[satir, m];
                }
            }
           
            double toplam = 0;
            index = 0;
            Array.Clear(temp, 0, temp.Length);
            for (int i = 0; i < egitim_row - 1; i++)
            {
                    for (int j = 0; j < col; j++)
                    {
                        if (i != satir)
                        {
                            if (j != sutun)
                            {
                                toplam += (egitim_seti[i, j] - satirlik_veri[0, j]) * (egitim_seti[i, j] - satirlik_veri[0, j]);

                            }
                        }
                    }
                
                if(i != satir) {
                    temp[index] = Math.Sqrt(toplam);
                    index_table[index] = i;
                    sutun_table[index] = sutun;
                     index++;
                }
                toplam = 0;
            }
            return temp;
        }

        public void kontrol(double[,] egitim_seti,int k,int egitim_row,int egitim_col)
        {
            double[] deg;
            double sum=0;
            int deger = 0, deger2 = 0;
            double ort;
            ExcelOlustur();
            double[,] test_dizi = TestIslemleri(egitim_row, egitim_col);
            for (int i = 0; i < row - 1; i++)
            {
                sum = 0;
                for (int j = 0; j < col; j++)
                {

                    if (test_dizi[i, j] == -1)
                    {
                        oklit(i,j,egitim_seti, egitim_row);
                        tut_dizi = temp;
                        deg = DiziYedekle(temp);

                        siralanmis_dizi = sirala(deg, k);

                        for (int m = 0; m < k; m++)
                        {
                            int idx = IndisBul(siralanmis_dizi[m]);
                            deger = index_table[idx];
                            deger2 = sutun_table[idx];
                            sum += egitim_seti[deger, deger2];
                        }

                        ort = sum / k;
                        MessageBox.Show(ort + "");
                        gercek = egitim_seti[i , j ];
                        Ortalama_Hata(gercek, ort);
                       
                        ExceleYaz(i+1 , j+1, ort);
                        ort = 0.0;
                    }
                    
                }
                ExceleYaz(2, 6, ortalama_hata / k);
            }
           
            kaydet();
        }
     

        public double[] sirala(double[] dizi,int k)
        {
            double gecici;
            double[] yeni= DiziYedekle(dizi);
            double[] son=new double[k];

            int indx=0;

            for (int i = 0; i < yeni.Length - 1; i++)
            {
                for (int j = i; j < yeni.Length; j++)
                {
                    if (yeni[i] > yeni[j])
                    {
                        gecici = yeni[j];
                        yeni[j] = yeni[i];
                        yeni[i] = gecici;
                        indx++;
                    }
                }
            }
            for (int i = 0; i < k; i++)
            {
                son[i] = yeni[i];
            }
        
            return son;
        }

        public double[] kTaneAl(int k,double[] dizi)
        {
            double[] new_array = new double[k];
            for (int i = 0; i < k; i++)
            {
                new_array[i]= dizi[i];
            }
            return new_array;
        }

        public int IndisBul(double eleman)
        {
            int idx=0;
            for (int i = 0; i < tut_dizi.Length; i++)
            {
                if (tut_dizi[i] == eleman) {
                    idx = i;
                    break;
                }
            }
            return idx;
        }

        public double[] DiziYedekle(double[] dizi)
        {
            point = new double[dizi.Length];
            for (int i = 0; i < dizi.Length; ++i)
            {
                point[i] = dizi[i];
            }
            return point;
        }


        public void ExcelOlustur()
        {
            ExcelUygulama = new Excel.Application();
            ExcelProje = ExcelUygulama.Workbooks.Add(Missing);
            ExcelSayfa = (Excel.Worksheet)ExcelProje.Worksheets.get_Item(1);
            ExcelRange = ExcelSayfa.UsedRange;
            ExcelSayfa = (Excel.Worksheet)ExcelUygulama.ActiveSheet;

            ExcelUygulama.Visible = false;
            ExcelUygulama.AlertBeforeOverwriting = false;
        }

        public void ExceleYaz(int row,int col,double deger)
        {
            Excel.Range bolge = (Excel.Range)ExcelSayfa.Cells[row, col];
            bolge.Value2 = deger;

        }

        public void kaydet()
        {

            ExcelProje.SaveAs(System.Windows.Forms.Application.StartupPath + @"\eksiksonuclar.xlsx", Excel.XlFileFormat.xlWorkbookDefault, Missing, Missing, false, Missing, Excel.XlSaveAsAccessMode.xlNoChange);
            ExcelProje.Close(true, Missing, Missing);
            ExcelUygulama.Quit();
        }

        public void Ortalama_Hata(double hedef, double tahmin)
        {
            ortalama_hata += hedef - tahmin;
        }

    }
}
