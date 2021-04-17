using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using FileProcessor;

namespace DVRP1
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            
        }
        private string OpenDialog()
        {
            var opd = new OpenFileDialog();
            opd.Filter = "*.xlsx | *.xlsx";
            opd.ShowDialog();
            return opd.FileName;
        }

        private uint[] Courses(string c)
        {
            
            int j = 0;
            uint[] a = new uint[6];
            if (c == "" || c == null)
            {
                return new uint[1] { 0};
            }
            for (int i = 0; i < c.Length; i++)
            {
                if (c[i] > 47 && c[i] < 58)
                {
                    a[j] = Convert.ToUInt32(c[i])-48;
                    j++;
                }
                if (c[i] == '-')
                {
                    uint x = a[j-1];
                    x++;
                    while (x != Convert.ToUInt32(c[i+1])-48)
                    {
                        a[j] = x;
                        x++;
                        j++;
                    }
                }
            }
            uint[] res = new uint[j];
            for (int i = 0; i < j; i++)
            {
                res[i] = a[i];
            }
            return res;
        }

        private uint MinMaxShitHappens(string shitstring)
        {
            if (shitstring == "")
            {
                return 0;
            }
            if (shitstring == "Не обмежена" || shitstring == "Необмежена" || shitstring == "необмежена")
            {
                return 4294967294;
            }
            else
                return Convert.ToUInt32(shitstring);
        }

        public uint[] Func(string str)
        {


            str = str.Replace(",-;", " ");

            char[] array = str.ToCharArray();
            uint[] Uint = new uint[array.Length];
            for (uint i = 0; i < array.Length; i++)
            {
                Uint[i] = Convert.ToUInt32(array[i]);

            }
            return Uint;
        }

        private void StartButton_click(object sender, EventArgs e)
        {
            FileReader Dop = new FileReader(OpenDialog());

            string un_magistr_select = Dop.GetCellValue("B", 4, "Вхід норматив");
            string un_bacalavr_select = Dop.GetCellValue("B", 5, "Вхід норматив");
            string un_magistr_condit_select = Dop.GetCellValue("E", 4, "Вхід норматив");
            string un_bacalavr_condit_select = Dop.GetCellValue("E", 5, "Вхід норматив");

            string fac_magistr_select = Dop.GetCellValue("B", 8, "Вхід норматив");
            string fac_bacalavr_select = Dop.GetCellValue("B", 9, "Вхід норматив");
            string fac_magistr_condit_select = Dop.GetCellValue("E", 8, "Вхід норматив");
            string fac_bacalavr_condit_select = Dop.GetCellValue("E", 9, "Вхід норматив");
            List <Discipline> disciplines = new List<Discipline>();
            uint i = 2;
            while (Dop.GetCellValue("A", i, "Вхід УВК бак+маг") != "" && Dop.GetCellValue("A", i, "Вхід УВК бак+маг") != null)
            {
                disciplines.Add(new Discipline(
                    Dop.GetCellValue("A", i, "Вхід УВК бак+маг"),
                    Dop.GetCellValue("B", i, "Вхід УВК бак+маг"),
                    Dop.GetCellValue("C", i, "Вхід УВК бак+маг"),
                    Dop.GetCellValue("D", i, "Вхід УВК бак+маг"),
                    MinMaxShitHappens(Dop.GetCellValue("E", i, "Вхід УВК бак+маг")),
                    MinMaxShitHappens(Dop.GetCellValue("F", i, "Вхід УВК бак+маг")),
                    Courses(Dop.GetCellValue("G", i, "Вхід УВК бак+маг"))));
                i++;
            }
          

            Dop.Close();
        }

        private void button1_Click(object sender, EventArgs e)
        {

        }
    }
}
