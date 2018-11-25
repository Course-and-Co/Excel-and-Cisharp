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
using System.IO;
using System.Text.RegularExpressions;

namespace WindowsFormsApplication1
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        string way = "";
        
        

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

       
        private void textBox2_TextChanged(object sender, EventArgs e)
        {

        }

        private void button2_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog FBD = new FolderBrowserDialog();
            FBD.ShowNewFolderButton = false;
            if (FBD.ShowDialog() == DialogResult.OK)
            {
                way = FBD.SelectedPath;
            }
            DirectoryInfo dir = new DirectoryInfo(way);
            //textBox2.Text += "\n\r" + way + "\r\n";
            string t = "*.xlsx";
            FileInfo[] Files = dir.GetFiles(t, SearchOption.TopDirectoryOnly);
           
            foreach (FileInfo fail in Files)
            {
                textBox2.Text += "\n\r" + fail.Name + "\r\n";
               
            }
            foreach (FileInfo fail in Files)
            {
                Excel.Application excel = new Excel.Application();
                Workbook workbookb = excel.Workbooks.Open(way + @"\" + fail.Name);
                try
                {
                    
                    Worksheet excelSheet = workbookb.ActiveSheet;
                    //Read the first cell
                    int count_i = 5;
                    int count_j = 5;
                    //while (true)
                    //{
                    //    if (excelSheet.Cells[count_i, count_j].Value != null)
                    //    {
                    //        count_i++;
                    //    }
                    //    else
                    //    {
                    //        MessageBox.Show(Convert.ToString(count_i - 1));
                    //        break;
                    //    }
                    //}
                    //while (true)
                    //{
                    //    if (excelSheet.Cells[count_i, count_j].Value != null)
                    //    {
                    //        count_j++;
                    //    }
                    //    else
                    //    {
                    //        MessageBox.Show(Convert.ToString(count_j - 1));
                    //        break;
                    //    }
                    //}
                    for (int i = 1; i <= count_i ; i++)
                    {
                        for (int j = 2; j <= count_j; j++)
                        {
                            if (excelSheet.Cells[j, i].Value != null)
                            {
                                switch (i)
                                {
                                    
                                    //номер
                                    case 1:
                                        string test = excelSheet.Cells[j, i].Value.ToString();
                                        int num;
                                        bool isNum = int.TryParse(test, out num);
                                        if (isNum)
                                            MessageBox.Show(test,"Это число");
                                        else
                                            MessageBox.Show(test, "Это строка");                                        
                                        break;

                                        //ФИО
                                    case 2:
                                        test = excelSheet.Cells[j, i].Value.ToString();
                                       
                                         isNum = int.TryParse(test, out num);
                                        if (isNum ==false)
                                            MessageBox.Show(test, "Это строка"); 
                                        else
                                            MessageBox.Show(test, "Это не строка");
                                        break;


                                        //адрес
                                    case 3:
                                        test = excelSheet.Cells[j, i].Value.ToString();
                                        //Regex reg = new Regex(@"(\d+)(\.|\,)(\d+)");
                                        //MatchCollection mc = reg.Matches(s);
                                        //if (mc.Count > 0) return true;
                                        //return false;
                                        isNum = int.TryParse(test, out num);
                                        if (isNum == false)

                                            MessageBox.Show(test, "Это строка");
                                        else
                                            MessageBox.Show(test, "Это не строка");
                                        break;

                                        //назначение 
                                    case 4:
                                        test = excelSheet.Cells[j, i].Value.ToString();
                                        isNum = int.TryParse(test, out num);
                                        if (isNum == false)
                                           
                                                MessageBox.Show(test, "Это строка");
                                                                                        
                                        else
                                            MessageBox.Show(test, "Это не строка");
                                        break;

                                        //показания 1
                                    case 5:
                                        test = excelSheet.Cells[j, i].Value.ToString();
                                        //Regex reg = new Regex(@"(\d+)(\.|\,)(\d+)");
                                        //MatchCollection mc = reg.Matches(s);
                                        //if (mc.Count > 0) return true;
                                        //return false;
                                        double num2;
                                        isNum = double.TryParse(test, out num2);
                                        if (isNum)
                                            MessageBox.Show(test, "Это  число");
                                        else
                                            MessageBox.Show(test, "Это ошибка");
                                        break;

                                    //показания 2
                                    case 6:
                                    
                                        test = excelSheet.Cells[j, i].Value.ToString();
                                        //Regex reg = new Regex(@"(\d+)(\.|\,)(\d+)");
                                        //MatchCollection mc = reg.Matches(s);
                                        //if (mc.Count > 0) return true;
                                        //return false;
                                       
                                        isNum = double.TryParse(test, out num2);
                                        if (isNum)
                                            MessageBox.Show(test, "Это  число");
                                        else
                                            MessageBox.Show(test, "Это ошибка");
                                        break;

                                        //расход
                                    case 7:
                                        test = excelSheet.Cells[j, i].Value.ToString();
                                        isNum = double.TryParse(test, out num2);
                                        if (isNum)
                                            MessageBox.Show(test, "Это  число");
                                        else
                                            MessageBox.Show(test, "Это ошибка");

                                        break;

                                        //сумма
                                    case 8:
                                        test = excelSheet.Cells[j, i].Value.ToString();
                                        isNum = double.TryParse(test, out num2);
                                        if (isNum)
                                            MessageBox.Show(test, "Это  число");
                                        else
                                            MessageBox.Show(test, "Это ошибка");

                                        break;

                                    //дата
                                    case 9:
                                        DateTime date;
                                        test = excelSheet.Cells[j, i].Value.ToString();
                                        isNum = DateTime.TryParse(test, out date);
                                        if (isNum)
                                            MessageBox.Show(test, "Это  дата");
                                        else
                                            MessageBox.Show(test, "Это ошибка");

                                        break;


                                    default:
                                        MessageBox.Show("Default case");
                                        break;
                                }

                            }
                            else
                            {
                                //MessageBox.Show(Convert.ToString(count_i - 1));
                                break;
                            }

                        }
                    }



                }
                catch(Exception error)
                {
                    MessageBox.Show("ошибка1" + error);
                }
                finally
                {
                    workbookb.SaveAs(way + @"\2.xlsx" );
                    workbookb.Close();
                }




            }
        }

        
    }
}
