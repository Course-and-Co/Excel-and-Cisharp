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
            //вызов функции выбора папки с файлами
            FileInfo[] Files = Failname();

            //цикл для проверки каждой книги по порядку
                foreach (FileInfo fail in Files)
                {
                //подключение Экземпляра excel
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

                    //вычисление количества строк
                        while (true)
                        {
                            if (excelSheet.Cells[count_i, 1].Value != null)
                                count_i++;
                            else                               
                                break;
                        }
                        //Вычисление количества столбцов
                        while (true)
                        {
                            if (excelSheet.Cells[1, count_j].Value != null)
                                count_j++;
                            else
                                break;
                        }

                        //цикл проверки данных по строчно в каждом столбце
                        //по очередно проверяется каждая строка
                        for (int j = 1; j <= count_j - 1; j++)
                        {
                            for (int i = 2; i <= count_i - 1; i++)
                            {
                                if (excelSheet.Cells[i, j].Value != null)//проверка на пустое значение ячейки
                                {
                                    double [] pokaz1= new double [count_i];//переменная для записи показаний новых
                                    double [] pokaz2 = new double[count_i];//переменная для записи показаний старых
                                    double [] rezult = new double[count_i];//переменная для записи результата разности показаний
                                    string test = excelSheet.Cells[i, j].Value.ToString();//переменная для записи значения из ячейки
                                    switch (j)
                                    {

                                        //номер
                                        case 1:
                                            if (chislo(test) == true)
                                            {
                                            }
                                            else
                                            {
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
                                            }
                                            else
                                            {
                                                oshibke += "\n\rОшибка в столбце 'ФИО', строка " + j + " столбец" + i + "\n\r";
                                            }
                                            break;


                                        //адрес
                                        case 3:
                                            test = excelSheet.Cells[i, j].Value.ToString();

                                            isNum = int.TryParse(test, out num);
                                            if (isNum == false)
                                            {
                                            }
                                            else
                                            {
                                                oshibke += "\n\rОшибка в столбце 'Адрес', строка " + j + " столбец" + i + "\n\r";
                                            }
                                            break;


                                        //назначение 
                                        case 4:
                                            test = excelSheet.Cells[i, j].Value.ToString();
                                            isNum = int.TryParse(test, out num);
                                            if (isNum == false)
                                            {
                                            }
                                            else
                                            {
                                                oshibke += "\n\r Ошибка в столбце 'Назначении платежа', строка " + j + " столбец" + i + "\n\r";
                                            }
                                            break;

                                        //показания новые
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
                                                    oshibke += "\n\r Ошибка в столбце 'Показания новые', строка " + 
                                                        j + " столбец " + i + " неверное значение\n\r";
                                                }
                                            }
                                            else
                                            {
                                                oshibke += "\n\r Ошибка в столбце 'Показания новые', строка " + 
                                                    j + " столбец " + i + "\n\r";
                                            }
                                            break;

                                        //показания старые
                                        case 6:
                                            test = excelSheet.Cells[i, j].Value.ToString();
                                            if (Dubl(test) == true)
                                            {
                                            pokaz2[i] = Convert.ToDouble(test);
                                            if (pokaz2[i] > 0 || pokaz1[i]>pokaz2[i])//старые показания не могут быть больше чем новые
                                            {

                                            }
                                            else
                                            {
                                                oshibke += "\n\r Ошибка в столбце 'Показания старые', строка " +
                                                    j + " столбец " + i + " неверное значение\n\r";
                                            }
                                        }
                                            else
                                            {
                                                oshibke += "\n\r Ошибка в столбце 'Показания старые', строка " +
                                                    j + " столбец " + i + "\n\r";
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
                                            }
                                            else
                                            {
                                                oshibke += "\n\rОшибка в столбце 'Расход кВт*ч', строка " + j + " столбец" + i + "\n\r";
                                            }
                                            break;

                                        //сумма
                                        case 8:
                                            test = excelSheet.Cells[i, j].Value.ToString();
                                            if (Dubl(test) == true)
                                            {
                                            }
                                            else
                                            {
                                                oshibke += "\n\rОшибка в столбце 'Сумма', строка " + j + " столбец" + i + "\n\r";
                                            }
                                            break;


                                        //дата
                                        case 9:
                                            test = excelSheet.Cells[i, j].Value.ToString();
                                            if (Date(test) == true)
                                            {
                                            }
                                            else
                                            {
                                                oshibke += "\n\rОшибка в столбце 'Дата', строка " + j + " столбец " + i + "\n\r"; ;
                                            }
                                            break;


                                        default:
                                        oshibke += "\n\r Ошибка что то не так \n\r"; ;
                                        break;
                                    }

                                }
                                else
                                {
                                    oshibke += "\n\rОшибка в строке " + j + " столбец " + i + " Отсутствует значение" + "\n\r";
                                    break;
                                }

                            }
                        }

                        

                    string put = way + @"\Правильно\";
                    string putosh = way + @"\Ошибки\";
                        if (oshibke == "")
                        {
                            oshibke = "\n\r Ошибок не обнаружено \n\r";
                            if (Directory.Exists(put))
                            {
                                workbookb.SaveAs(put + Path.GetFileNameWithoutExtension(fail.Name));//сохранение
                           
                            }
                            else
                            {
                                DirectoryInfo di = Directory.CreateDirectory(put);//создание папки
                                workbookb.SaveAs(put + Path.GetFileNameWithoutExtension(fail.Name));//сохранение
                            }                       
                        }
                        else
                        {
                            if (Directory.Exists(putosh))
                            {
                                workbookb.SaveAs(putosh + Path.GetFileNameWithoutExtension(fail.Name) + @"_osibka.xlsx");//сохранение
                        }
                            else
                            {
                                DirectoryInfo di = Directory.CreateDirectory(putosh);//создание папки
                                workbookb.SaveAs(putosh + Path.GetFileNameWithoutExtension(fail.Name) + @"_osibka.xlsx");//сохранение
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
                        workbookb.Close(); //закрытие сесии работы с эксель
                    }          
                }       
        }
        //функция для выбора папки с файлами
        public FileInfo[] Failname()
        {
            FileInfo[] Files;
            FolderBrowserDialog FBD = new FolderBrowserDialog();
            FBD.ShowNewFolderButton = false;

            if (FBD.ShowDialog() == DialogResult.OK)
            {
                way = FBD.SelectedPath;
            }
            else//запись базового пути для исключения ошибки по отмене выбора пути
            {
                way = @"C:";
                MessageBox.Show("ошибка не указан путь ");

            }
            DirectoryInfo dir = new DirectoryInfo(way);

            string t = "*.xlsx";//тип отбираемых файлов
            Files = dir.GetFiles(t, SearchOption.TopDirectoryOnly);
            return Files;
        }

        //функция для проверки строки на дату
        public bool Date(string test)
        {
            bool isNum = DateTime.TryParse(test, out DateTime date);
            if (isNum)
                return true;
            else
                return false;
        }

        //функция проверки строки на целое число
        public bool chislo(string test)
        {
            bool isNum = int.TryParse(test, out int num);
            if (isNum)
                return true;
            else
                return false;
        }

        //функция проверки строки на число с запятой
        public bool Dubl(string test)
        {
            bool isNum = double.TryParse(test, out double num);
            if (isNum)
                return true;
            else
                return false;
        }
    }
}
