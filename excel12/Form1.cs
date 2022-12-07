using Excel = Microsoft.Office.Interop.Excel;
using System.Reflection;

namespace excel12
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            try
            {
                
                xlApp = new Excel.Application();

                
                xlWB = xlApp.Workbooks.Add(Missing.Value);

                
                xlSheet = xlWB.ActiveSheet;

                
                CreateTable(); 

                
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

        Excel.Application xlApp; 
        Excel.Workbook xlWB;     
        Excel.Worksheet xlSheet;

        public void CreateTable()
        {

            string[] fejl�cek = new string[] {
        "K�rd�s",
        "1. v�lasz",
        "2. v�laszl",
        "3. v�lasz",
        "Helyes v�lasz",
        "k�p"};

            for (int i = 0; i < fejl�cek.Length; i++)
            {
                xlSheet.Cells[1, 1] = fejl�cek[0];
                xlSheet.Cells[1, 2] = fejl�cek[1];
                xlSheet.Cells[1, 3] = fejl�cek[2];
                xlSheet.Cells[1, 4] = fejl�cek[3];
                xlSheet.Cells[1, 5] = fejl�cek[4];
                xlSheet.Cells[1, 6] = fejl�cek[5];
            }

            Models.HajosContext hajosContext = new Models.HajosContext();
            var mindenK�rd�s = hajosContext.Questions.ToList();

            object[,] adatT�mb = new object[mindenK�rd�s.Count(), fejl�cek.Count()];

            for (int i = 0; i < mindenK�rd�s.Count(); i++)
            {
                adatT�mb[i, 0] = mindenK�rd�s[i].Question1;
                adatT�mb[i, 1] = mindenK�rd�s[i].Answer1;
                adatT�mb[i, 2] = mindenK�rd�s[i].Answer2;
                adatT�mb[i, 3] = mindenK�rd�s[i].Answer3;
                adatT�mb[i, 4] = mindenK�rd�s[i].CorrectAnswer;
                adatT�mb[i, 5] = mindenK�rd�s[i].Image;
            }

            int sorokSz�ma = adatT�mb.GetLength(0);
            int oszlopokSz�ma = adatT�mb.GetLength(1);

            Excel.Range adatRange = xlSheet.get_Range("A2", Type.Missing).get_Resize(sorokSz�ma, oszlopokSz�ma);
            adatRange.Value2 = adatT�mb;

            adatRange.Columns.AutoFit();

            Excel.Range fejll�cRange = xlSheet.get_Range("A1", Type.Missing).get_Resize(1, 6);
            fejll�cRange.Font.Bold = true;
            fejll�cRange.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
            fejll�cRange.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            fejll�cRange.EntireColumn.AutoFit();
            fejll�cRange.RowHeight = 40;
            fejll�cRange.Interior.Color = Color.Fuchsia;
            fejll�cRange.BorderAround2(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThick);

            adatRange.BorderAround2(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThick);

            Excel.Range elsooszlop = xlSheet.get_Range("A1", Type.Missing).get_Resize(sorokSz�ma, 1);
            elsooszlop.Font.Bold = true;


            int lastRowID = xlSheet.UsedRange.Rows.Count;
            Excel.Range utolsooszlop = xlSheet.get_Range("F2", Type.Missing).get_Resize(lastRowID,1);
            utolsooszlop.Interior.Color = Color.LightGreen;
            

        }
    }
}