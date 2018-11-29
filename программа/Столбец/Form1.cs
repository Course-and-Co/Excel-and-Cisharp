using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Столбец
{
    public partial class Form2 : Form
    {
        public Form2()
        {
            InitializeComponent();
        }
        public string Data
        {
            get
            {
                return comboBox1.Text;
            }
        }
        private void button1_Click(object sender, EventArgs e)
        {

        }
    }
}
