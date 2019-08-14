using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ForMyExcel
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            ExcelClass ex = new ExcelClass(@"C:\Users\vanka\source\repos\ForMyExcel\ForMyExcel\bin\Debug\Форма для ответа — копия1.xlsx", 1);

            double[,] read = ex.ReadRange(7, 6, 38, 6);
            ex.Close();
            ExcelClass ex1 = new ExcelClass(@"C:\Users\vanka\source\repos\ForMyExcel\ForMyExcel\bin\Debug\Форма для ответа — копия1 — копия.xlsx", 1);
            ex1.WriteRange(7, 7, 38, 7, read);
            ex1.ClearRange();
            ex1.Save();
            ex1.Close();
        }
    }
}
