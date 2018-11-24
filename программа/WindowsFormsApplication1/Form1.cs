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
            textBox2.Text += "\n\r" + way + "\r\n";
            FileInfo[] Files = dir.GetFiles("*.xlsx", SearchOption.TopDirectoryOnly);
           
            foreach (FileInfo fail in Files)
            {

                textBox2.Text += "\n\r" + fail.Name + "\r\n";

               
            }
            foreach (FileInfo fail in Files)
            {
                Excel.Application excel = new Excel.Application();
                Workbook workbookb = excel.Workbooks.Open(way + @"\" + fail.Name);
                Worksheet excelSheet = workbookb.ActiveSheet;
                workbookb.Close();





            }
        }

        
    }
}
