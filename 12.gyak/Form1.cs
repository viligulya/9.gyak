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
using System.Reflection;
using Microsoft.Office.Interop.Excel;

namespace _12.gyak
{
    public partial class Form1 : Form
    {

        Excel.Application xlApp;
        Excel.Workbook xlWB;
        Excel.Worksheet xlSheet;
        public Form1()
        {
            InitializeComponent();
            //CreateExcel();
            CreateExcel();
        }

        private void CreateExcel()
        {
            try
            {
                // Excel elindítása és az applikáció objektum betöltése
                xlApp = new Excel.Application();
                // Új munkafüzet
                xlWB = xlApp.Workbooks.Add(Missing.Value);
                // Új munkalap
                xlSheet = xlWB.ActiveSheet;
                // Tábla létrehozása
                CreateTable(); // Ennek megírása a következő feladatrészben következik
                // Control átadása a felhasználónak
                xlApp.Visible = true;
                xlApp.UserControl = true;


            }
            catch (Exception ex)
            {
                string errMsg = string.Format("Error: {0}\nLine: {1}", ex.Message, ex.Source);
                MessageBox.Show(errMsg, "Error");

                xlWB.Close(false, Type.Missing, Type.Missing);
                xlApp.Quit();
                xlWB = null;
                xlApp = null;
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            
        }

        


    }
}
