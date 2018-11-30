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
using Столбец;


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


        
        Workbook workbookb;
        private void button2_Click(object sender, EventArgs e)
        {
            //Form2 f = new Form2();
            //f.ShowDialog();
            //string t = f.Data;

            FileInfo[] Files = Failname();
                foreach (FileInfo fail in Files)
                {

                    Excel.Application excel = new Excel.Application();
                    
                    
                    
                    try
                    {
                        workbookb = excel.Workbooks.Open(way + @"\" + fail.Name);
                        string oshibke = "";
                        int yui = 1;

                        Worksheet excelSheet = workbookb.ActiveSheet;

                        textBox2.Text += "\n\r" + fail.Name + "\n\r";
                        int count_i = 1;
                        int count_j = 1;

                    //строки
                        while (true)
                        {
                            if (excelSheet.Cells[count_i, 1].Value != null)
                            {
                                count_i++;
                            }
                            else
                            {
                                //MessageBox.Show(Convert.ToString(count_i - 1));
                                break;
                            }
                        }
                        //столбцы
                        while (true)
                        {
                            if (excelSheet.Cells[1, count_j].Value != null)
                            {
                                count_j++;
                            }
                            else
                            {
                                //MessageBox.Show(Convert.ToString(count_j - 1));
                                break;
                            }
                        }


                        //возможно сделать выбор типа столбца


                        for (int j = 1; j <= count_j - 1; j++)
                        {
                            for (int i = 2; i <= count_i - 1; i++)
                            {
                                if (excelSheet.Cells[i, j].Value != null)
                                {
                                    double [] pokaz1= new double [count_i-1];
                                    double [] pokaz2 = new double[count_i - 1];
                                    double [] rezult = new double[count_i - 1];
                                    string test = excelSheet.Cells[i, j].Value.ToString();
                                    switch (j)
                                    {

                                        //номер
                                        case 1:


                                            if (chislo(test) == true)
                                            {
                                                //MessageBox.Show(test, "Это число");
                                                yui = 1;
                                            }
                                            else
                                            {
                                                //MessageBox.Show(test, "Это строка");
                                                oshibke += "\n\r Ошибка в столбце '№', строка " + j + " столбец" + i + "\n\r";
                                            }

                                            break;

                                        //ФИО
                                        case 2:
                                            test = excelSheet.Cells[i, j].Value.ToString();
                                            int num;
                                            bool isNum = int.TryParse(test, out num);
                                            if (isNum == false)
                                            {
                                                //MessageBox.Show(test, "Это строка");
                                                yui = 1;
                                            }

                                            else
                                            {
                                                //MessageBox.Show(test, "Это не строка");
                                                oshibke += "\n\rОшибка в столбце 'ФИО', строка " + j + " столбец" + i + "\n\r";
                                            }

                                            break;


                                        //адрес
                                        case 3:
                                            test = excelSheet.Cells[i, j].Value.ToString();

                                            isNum = int.TryParse(test, out num);
                                            if (isNum == false)
                                            {
                                                yui = 1;
                                                //MessageBox.Show(test, "Это строка");
                                            }

                                            else
                                            {
                                                //MessageBox.Show(test, "Это не строка");
                                                oshibke += "\n\rОшибка в столбце 'Адрес', строка " + j + " столбец" + i + "\n\r";
                                            }

                                            break;

                                        //назначение 
                                        case 4:
                                            test = excelSheet.Cells[i, j].Value.ToString();
                                            isNum = int.TryParse(test, out num);
                                            if (isNum == false)
                                            {
                                                yui = 1;
                                                //MessageBox.Show(test, "Это строка");
                                            }


                                            else
                                            {
                                                //MessageBox.Show(test, "Это не строка");
                                                oshibke += "\n\r Ошибка в столбце 'Назначении платежа', строка " + j + " столбец" + i + "\n\r";
                                            }

                                            break;

                                        //показания 1
                                        case 5:
                                            test = excelSheet.Cells[i, j].Value.ToString();
                                            if (Dubl(test) == true)
                                            {
                                                pokaz1[i] =Convert.ToDouble(test);
                                            if (pokaz1[i] > 0)
                                            {

                                            }
                                            else
                                            {
                                                oshibke += "\n\r Ошибка в столбце 'Показания 1', строка " + j + " столбец " + i + " неверное значение\n\r";
                                            }
                                            //MessageBox.Show(test, "Это  число");
                                        }

                                            else
                                            {
                                                //MessageBox.Show(test, "Это ошибка");
                                                oshibke += "\n\r Ошибка в столбце 'Показания 1', строка " + j + " столбец " + i + "\n\r";
                                            }

                                            break;

                                        //показания 2(старые)
                                        case 6:
                                            test = excelSheet.Cells[i, j].Value.ToString();
                                            if (Dubl(test) == true)
                                            {
                                            pokaz2[i] = Convert.ToDouble(test);
                                            if (pokaz2[i] > 0 || pokaz1[i]>pokaz2[i])
                                            {

                                            }
                                            else
                                            {
                                                oshibke += "\n\r Ошибка в столбце 'Показания 2', строка " + j + " столбец " + i + " неверное значение\n\r";
                                            }
                                            //MessageBox.Show(test, "Это  число");
                                        }
                                            else
                                            {
                                                //MessageBox.Show(test, "Это ошибка");
                                                oshibke += "\n\r Ошибка в столбце 'Показания 2', строка " + j + " столбец " + i + "\n\r";
                                            }
                                            break;

                                        //расход
                                        case 7:
                                            test = excelSheet.Cells[i, j].Value.ToString();

                                            if (Dubl(test) == true)
                                            {
                                            rezult[i] = Convert.ToDouble(test);
                                            if (rezult[i] > 0)
                                            {

                                            }
                                            else
                                            {
                                                oshibke += "\n\r Ошибка в столбце 'Расход кВт*ч', строка " + j + " столбец " + i + " неверное значение\n\r";
                                            }
                                                //MessageBox.Show(test, "Это  число");
                                            }

                                            else
                                            {
                                                //MessageBox.Show(test, "Это ошибка");
                                                oshibke += "\n\rОшибка в столбце 'Расход кВт*ч', строка " + j + " столбец" + i + "\n\r";
                                            }


                                            break;

                                        //сумма
                                        case 8:
                                            test = excelSheet.Cells[i, j].Value.ToString();

                                            if (Dubl(test) == true)
                                            {
                                                yui = 1;
                                                //MessageBox.Show(test, "Это  число");
                                            }

                                            else
                                            {
                                                //MessageBox.Show(test, "Это ошибка");
                                                oshibke += "\n\rОшибка в столбце 'Сумма', строка " + j + " столбец" + i + "\n\r";
                                            }


                                            break;

                                        //дата
                                        case 9:
                                            test = excelSheet.Cells[i, j].Value.ToString();
                                            if (Date(test) == true)
                                            {
                                                //MessageBox.Show(test, "Это  число");
                                                yui = 1;
                                            }

                                            else
                                            {
                                                //MessageBox.Show(test, "Это ошибка");
                                                oshibke += "\n\rОшибка в столбце 'Дата', строка " + j + " столбец " + i + "\n\r"; ;
                                            }

                                            break;


                                        default:
                                            //MessageBox.Show("Default case");
                                            break;
                                    }

                                }
                                else
                                {
                                    oshibke += "\n\rОшибка в строке " + j + " столбец " + i + " Отсутствует значение" + "\n\r";
                                    //MessageBox.Show("Это ошибка");
                                    break;
                                }

                            }
                        }


                    //сделать проверку на существлование папки для сохранения файлов

                        string put= way + @"\Правильно\" + Path.GetFileNameWithoutExtension(fail.Name);
                        string putosh = way + @"\Ошибки\" + Path.GetFileNameWithoutExtension(fail.Name) + @"_osibka.xlsx";
                        if (oshibke == "")
                        {
                            oshibke = "\n\r Ошибок не обнаружено";
                        if (Directory.Exists(put))
                        {
                            workbookb.SaveAs(put);
                           
                        }
                        else
                        {
                            DirectoryInfo di = Directory.CreateDirectory(put);
                            workbookb.SaveAs(put);
                        }
                        
                        }
                        else
                        {
                        if (Directory.Exists(putosh))
                        {
                            workbookb.SaveAs(putosh);
                        }
                        else
                        {
                            DirectoryInfo di = Directory.CreateDirectory(putosh);
                            workbookb.SaveAs(putosh);
                        }

                    }
                        textBox2.Text += "\n\r" + oshibke + "\r\n";
                        textBox2.Text += "\n\r";
                    }
                    catch (Exception error)
                    {
                        MessageBox.Show("ошибка " + error);
                    }
                    finally
                    {

                        workbookb.Close();
                    }
            

            
                


            }
        }

        public FileInfo[] Failname()
        {
            FileInfo[] Files;
            FolderBrowserDialog FBD = new FolderBrowserDialog();
            FBD.ShowNewFolderButton = false;

            if (FBD.ShowDialog() == DialogResult.OK)
            {
                way = FBD.SelectedPath;
            }
            else
            {

                way = @"C:";
                MessageBox.Show("ошибка не указан путь ");

            }
            DirectoryInfo dir = new DirectoryInfo(way);
            string t = " *.xlsx";
            Files = dir.GetFiles(t, SearchOption.TopDirectoryOnly);
            return Files;







        }
        public bool Date(string test)
        {
            DateTime date;
            bool isNum = DateTime.TryParse(test, out date);
            if (isNum)
                return true;
            else
                return false;
        }

        public bool chislo(string test)
        {
            int num;
            bool isNum = int.TryParse(test, out num);
            if (isNum)
                return true;
            else
                return false;
        }

        public bool Dubl(string test)
        {
            double num;
            bool isNum = double.TryParse(test, out num);
            if (isNum)
                return true;
            else
                return false;
        }
    }
}
