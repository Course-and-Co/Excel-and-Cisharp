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


        
        Workbook workbookb;
        private void button2_Click(object sender, EventArgs e)
        {
            FileInfo[] Files = Failname();
            foreach (FileInfo fail in Files)
            {
                if( fail.Name[0] == '~')
                {

                MessageBox.Show("Обнаружен открытый процесс! После работы рекомендуется закрыть все процессы MS EXCEL ");//ошибка возникает из за прерваной работы програмы
            }
            else
                {
                    Excel.Application excel = new Excel.Application();

                    try
                    {
                        workbookb = excel.Workbooks.Open(way + @"\" + fail.Name);//открытие  файла
                        string oshibke = "";
                        Worksheet excelSheet = workbookb.ActiveSheet;

                        textBox2.Text += "\n\r" + fail.Name + "\n\r";//вывод имени файла
                        int count_i = 1;
                        int count_j = 1;

                        //кол-во строк
                        while (true)
                        {
                            if (excelSheet.Cells[count_i, 1].Value != null)
                                count_i++;
                            else
                                break;
                        }
                        //кол-во столбцов
                        while (true)
                        {
                            if (excelSheet.Cells[1, count_j].Value != null)
                                count_j++;
                            else
                                break;
                        }





                        for (int j = 1; j <= count_j - 1; j++)
                        {
                            for (int i = 2; i <= count_i - 1; i++)
                            {
                                if (excelSheet.Cells[i, j].Value != null)
                                {
                                    string FIO = @"[А-Я|а-я]{2,}\ [А-Я|а-я]{2,}\ [А-Я|а-я]{2,}";//формула для проверки правильности ФИО
                                    string adres = @"г\.[А-Я|а-я]{2,}\, ул\.[А-Я|а-я]{2,}\, [0-9]{1,4}";//Формула для проверки правельности адреса
                                    double[] pokaz1 = new double[count_i];
                                    double[] pokaz2 = new double[count_i];
                                    double[] rezult = new double[count_i];
                                    string test = excelSheet.Cells[i, j].Value.ToString();//значение ячейки
                                    switch (j)
                                    {

                                        //номер
                                        case 1:
                                            if (chislo(test) == true)
                                            {
                                            }
                                            else
                                            {
                                                oshibke += "\n\r Ошибка в столбце '№', строка " +
                                                    i + " столбец" + j + "\n\r";
                                            }

                                            break;

                                        //ФИО
                                        case 2:
                                            test = excelSheet.Cells[i, j].Value.ToString();
                                            if (Regex.IsMatch(test, FIO, RegexOptions.IgnoreCase))
                                            {
                                            }
                                            else
                                            {
                                                oshibke += "\n\rОшибка в столбце 'ФИО', строка " + 
                                                    i + " столбец" + j + "\n\r";
                                            }

                                                break;


                                            //адрес
                                            case 3:
                                                test = excelSheet.Cells[i, j].Value.ToString();

                                            if (Regex.IsMatch(test, adres))
                                            {
                                            }
                                            else
                                            {
                                                oshibke += "\n\rОшибка в столбце 'Адрес', строка " + 
                                                    i + " столбец" + j + "\n\r";
                                            }

                                            break;

                                        //назначение 
                                        case 4:
                                            test = excelSheet.Cells[i, j].Value.ToString();
                                                
                                            if (test=="электроэнергия" || test == "отопление")
                                            {
                                            }
                                            else
                                            {
                                                oshibke += "\n\r Ошибка в столбце 'Назначении платежа', строка " +
                                                    i + " столбец" + j + "\n\r";
                                            }

                                            break;

                                        //показания 1
                                        case 5:
                                            test = excelSheet.Cells[i, j].Value.ToString();
                                            if (Dubl(test) == true)
                                            {
                                                pokaz1[i] = Convert.ToDouble(test);
                                                if (pokaz1[i] > 0)//проверка на то что бы показания небыли меньше нуля 
                                                {

                                                }
                                                else
                                                {
                                                    oshibke += "\n\r Ошибка в столбце 'Показания 1', строка " + 
                                                        i + " столбец " + j + " неверное значение\n\r";
                                                }
                                                //MessageBox.Show(test, "Это  число");
                                            }

                                            else
                                            {
                                                //MessageBox.Show(test, "Это ошибка");
                                                oshibke += "\n\r Ошибка в столбце 'Показания 1', строка " +
                                                    i + " столбец " + j + "\n\r";
                                            }

                                            break;

                                        //показания 2(старые)
                                        case 6:
                                            test = excelSheet.Cells[i, j].Value.ToString();
                                            if (Dubl(test) == true)
                                            {
                                                pokaz2[i] = Convert.ToDouble(test);
                                                if (pokaz2[i] > 0 && pokaz1[i] > pokaz2[i])//проверка на то что бы показания небыли меньше нуля 
                                                                                           //и старые показания не превышали новые
                                                {

                                                }
                                                else
                                                {
                                                    oshibke += "\n\r Ошибка в столбце 'Показания 2', строка " +
                                                        i + " столбец " + j + " неверное значение\n\r";
                                                }
                                                
                                            }
                                            else
                                            {
                                                
                                                oshibke += "\n\r Ошибка в столбце 'Показания 2', строка " + 
                                                    i + " столбец " + j + "\n\r";
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
                                                    oshibke += "\n\r Ошибка в столбце 'Расход кВт*ч', строка " +
                                                        i + " столбец " + j + " неверное значение\n\r";
                                                }
                                                //MessageBox.Show(test, "Это  число");
                                            }

                                            else
                                            {
                                                //MessageBox.Show(test, "Это ошибка");
                                                oshibke += "\n\rОшибка в столбце 'Расход кВт*ч', строка " +
                                                    i + " столбец" + j + "\n\r";
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
                                                
                                                oshibke += "\n\rОшибка в столбце 'Сумма', строка " +
                                                    i + " столбец" + j + "\n\r";
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
                                                
                                                oshibke += "\n\rОшибка в столбце 'Дата', строка " +
                                                    i + " столбец " + j + "\n\r"; ;
                                            }

                                            break;


                                        default:
                                            oshibke += "\n\rОшибка \n\r"; ;

                                            break;
                                    }

                                }
                                else
                                {
                                    oshibke += "\n\rОшибка в строке " + i + " столбец " +
                                        j + " Отсутствует значение" + "\n\r";
                                   
                                    break;
                                }

                            }
                        }

                           

                        string put = way + @"\Правильно\";
                        string putosh = way + @"\Ошибки\";
                        if (oshibke == "")//проверка на точность файла
                        {
                            oshibke = "\n\r Ошибок не обнаружено";
                            textBox2.Text += "\n\r" + oshibke + "\r\n";
                            textBox2.Text += "\n\r";
                            if (Directory.Exists(put))//проверка на существование папки
                            {
                                    workbookb.SaveAs(put + Path.GetFileNameWithoutExtension(fail.Name));//сохранение провереного файла в папку
                            }
                            else
                            {
                                DirectoryInfo di = Directory.CreateDirectory(put);//создание папки
                                workbookb.SaveAs(put + Path.GetFileNameWithoutExtension(fail.Name));//сохранение провереного файла в папку
                            }
                        }
                        else
                        {
                            textBox2.Text += "\n\r" + oshibke + "\r\n";
                            textBox2.Text += "\n\r";
                            if (Directory.Exists(putosh))//проверка на существование папки
                            {
                               
                                workbookb.SaveAs(putosh + Path.GetFileNameWithoutExtension(fail.Name) + @"_osibka.xlsx");//сохранение провереного файла в папку
                            }
                            else
                            {
                                DirectoryInfo di = Directory.CreateDirectory(putosh);//создание папки
                                workbookb.SaveAs(putosh + Path.GetFileNameWithoutExtension(fail.Name) + @"_osibka.xlsx");//сохранение провереного файла в папку
                            }

                        }
                        
                    }
                    catch (Exception)
                    {
                        MessageBox.Show("Предупреждение! Вы не сохранили провереный файл!" );//исключение при отказе от перезаписи файлов
                    }
                    finally
                    {
                        workbookb.Close();
                    }
                    }

                    
   
            }
        }

        //функция для выбора пути к файлам
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

            string t = "*.xls*";
            Files = dir.GetFiles(t, SearchOption.TopDirectoryOnly);
            return Files;

           
        }

        //функция для определения даты
        public bool Date(string test)
        {
            DateTime date;
            bool isNum = DateTime.TryParse(test, out date);
            if (isNum)
                return true;
            else
                return false;
        }


     //функция для определения целого числа
        public bool chislo(string test)
        {
            int num;
            bool isNum = int.TryParse(test, out num);
            if (isNum)
                return true;
            else
                return false;
        }

        //функция для определения числа с запятой
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
