﻿using System;
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

        private void button2_Click(object sender, EventArgs e)
        {
            string oshibke = "";
            int yui= 1;
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
                    
                    int count_i = 1;//кол-во столбцов
                    int count_j = 1;//кол-во строк
                    while (true)
                    {
                        if (excelSheet.Cells[count_i, 1].Value != null)
                        {
                            count_i++;
                        }
                        else
                        {
                            MessageBox.Show(Convert.ToString(count_i - 1));
                            break;
                        }
                    }
                    
                    while (true)
                    {
                        if (excelSheet.Cells[1, count_j].Value != null)
                        {
                            count_j++;
                        }
                        else
                        {
                            MessageBox.Show(Convert.ToString(count_j - 1));
                            break;
                        }
                    }
                    for (int i = 1; i <= count_i-1 ; i++)
                    {
                        for (int j = 2; j <= count_j-1; j++)
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
                                            //MessageBox.Show(test, "Это число");
                                            yui = 1;
                                        else
                                            /*MessageBox.Show(test, "Это строка")*/
                                            oshibke += "\n\r Ошибка в столбце '№', строка " + j + " столбец" + i + "\n\r";
                                        break;

                                        //ФИО
                                    case 2:
                                        test = excelSheet.Cells[j, i].Value.ToString();
                                       
                                         isNum = int.TryParse(test, out num);
                                        if (isNum ==false)
                                            //MessageBox.Show(test, "Это строка"); 
                                            yui = 1;
                                        else
                                            //MessageBox.Show(test, "Это не строка");
                                            oshibke += "\n\rОшибка в столбце 'ФИО', строка " + j + " столбец" + i + "\n\r";
                                        break;


                                        //адрес
                                    case 3:
                                        test = excelSheet.Cells[j, i].Value.ToString();
                                       
                                        isNum = int.TryParse(test, out num);
                                        if (isNum == false)
                                            yui = 1;
                                            //MessageBox.Show(test, "Это строка");
                                        else
                                            //MessageBox.Show(test, "Это не строка");
                                            oshibke += "\n\rОшибка в столбце 'Адрес', строка " + j + " столбец" + i + "\n\r";
                                        break;

                                        //назначение 
                                    case 4:
                                        test = excelSheet.Cells[j, i].Value.ToString();
                                        isNum = int.TryParse(test, out num);
                                        if (isNum == false)
                                            yui = 1;
                                            //MessageBox.Show(test, "Это строка");
                                        
                                        else
                                            //MessageBox.Show(test, "Это не строка");
                                            oshibke += "\n\r Ошибка в столбце 'Назначении платежа', строка " + j + " столбец" + i + "\n\r";
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
                                            yui = 1;
                                            //MessageBox.Show(test, "Это  число");
                                        else
                                            //MessageBox.Show(test, "Это ошибка");
                                            oshibke += "\n\r Ошибка в столбце 'Показания 1', строка " + j+" столбец " + i + "\n\r";
                                        break;

                                    //показания 2
                                    case 6:
                                    
                                        test = excelSheet.Cells[j, i].Value.ToString();
                                       
                                       
                                        isNum = double.TryParse(test, out num2);
                                        if (isNum)
                                            yui = 1;
                                            //MessageBox.Show(test, "Это  число");
                                        else
                                            //MessageBox.Show(test, "Это ошибка");
                                            oshibke += "\n\r Ошибка в столбце 'Показания 2', строка " + j + " столбец " + i + "\n\r";
                                        break;

                                        //расход
                                    case 7:
                                        test = excelSheet.Cells[j, i].Value.ToString();
                                        isNum = double.TryParse(test, out num2);
                                        if (isNum)
                                            yui = 1;
                                            //MessageBox.Show(test, "Это  число");
                                        else
                                            //MessageBox.Show(test, "Это ошибка");
                                            oshibke += "\n\rОшибка в столбце 'Расход кВт*ч', строка " + j + " столбец" + i + "\n\r";

                                        break;

                                        //сумма
                                    case 8:
                                        test = excelSheet.Cells[j, i].Value.ToString();
                                        isNum = double.TryParse(test, out num2);
                                        if (isNum)
                                            yui = 1;
                                            //MessageBox.Show(test, "Это  число");
                                        else
                                            //MessageBox.Show(test, "Это ошибка");
                                            oshibke += "\n\rОшибка в столбце 'Сумма', строка " + j + " столбец" + i + "\n\r";

                                        break;

                                    //дата
                                    case 9:
                                        DateTime date;
                                        test = excelSheet.Cells[j, i].Value.ToString();
                                        isNum = DateTime.TryParse(test, out date);
                                        if (isNum)
                                            yui = 1;
                                            //MessageBox.Show(test, "Это дата");
                                        else
                                            //MessageBox.Show(test, "Это ошибка");
                                            oshibke += "\n\rОшибка в столбце 'Дата', строка " + j + " столбец " + i + "\n\r";

                                        break;


                                    default:
                                        MessageBox.Show("Default case");
                                        break;
                                }

                            }
                            else
                            {
                                oshibke += "\n\rОшибка в строке " + j + " столбец " + i +  " Отсутствует значение" + "\n\r";
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
                    //workbookb.SaveAs(way + @"\2.xlsx" );
                    workbookb.Close();
                }
                if (oshibke == "")
                {
                    oshibke = "Ошибок не обнаружено";
                }

                textBox2.Text += "\n\r" + oshibke + "\r\n";


            }
        }

        
    }
}
