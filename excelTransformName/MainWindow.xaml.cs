using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using Excel = Microsoft.Office.Interop.Excel;

namespace excelTransformName
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private GridView gridView = new GridView();
        private int colNumber;
        private Excel.Worksheet wSheet;

        public MainWindow()
        {
            InitializeComponent();

            
            //this.lbAutori.View = gridView;

            //gridView.Columns.Add(new GridViewColumn
            //{
            //    Header = "Rb.",
            //    Width = 10,
            //    DisplayMemberBinding = new Binding("Rb")
            //});
            //gridView.Columns.Add(new GridViewColumn
            //{
            //    Header = "Autori",
            //    DisplayMemberBinding = new Binding("Autor")
            //});

            
        }

        private void btUlaz_Click(object sender, RoutedEventArgs e)
        {
            Microsoft.Win32.OpenFileDialog ofd = new Microsoft.Win32.OpenFileDialog();
            if (ofd.ShowDialog() == true)
            {
                txtUlaz.Text = ofd.FileName;
                obrada(ofd.FileName);
            }
                
        }

        private void obrada(string fName)
        {
            lbAutori.Items.Clear();
            var excelApp = new Excel.Application();

            excelApp.Visible = true;

            Excel.Workbook wBook = excelApp.Workbooks.Open(fName);

            wSheet = wBook.Worksheets["Sheet1"];

            Excel.Range row = wSheet.Rows[1];
            Excel.Range find = row.Find("Autori");

            colNumber = find.Column;

            var range = wSheet.UsedRange;//wSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell);
            int lastR = range.Rows.Count;//wSheet.Range["A1", range].Row;
            
            string cellData;

            List<string> autori = new List<string>();

            for(int rowC = 2; rowC <= lastR; rowC++)
            {
                range = wSheet.Cells[rowC, colNumber];
                cellData = (string)range.Text;
                //autori.Add(cellData);
                // lbAutori.Items.Add(new myItm { Rb = rowC - 1, Autor = cellData });
                lbAutori.Items.Add(cellData);
            }
            //wSheet.Columns.Insert(Type.Missing,1);
            //object SaveChanges = (object) false;
            //wBook.Close(SaveChanges,Type.Missing,Type.Missing);
            //excelApp.Quit();
            //excelApp = null;
        }

        private void btnObr_Click(object sender, RoutedEventArgs e)
        {

            for (var i = 0; i < lbAutori.Items.Count; i++)
            {
                string[] imena = lbAutori.Items[i].ToString().Split(',');
                string[] autor = new string[imena.Length];
                for (var j = 0; j < imena.Length; j++)
                {
                    string[] imeprezime = imena[j].Trim().Split(' ');
                    string ime = imeprezime[imeprezime.Length - 1];
                    imeprezime[imeprezime.Length - 1] = ime[0] + ".";
                    autor[j] = string.Join(" ", imeprezime);//prezime i ime sa skracenjem

                   // gridView.Columns.Add(new GridViewColumn { Header = "1", DisplayMemberBinding = new Binding("a" + j), Width = 20 });
                    //lbAutori.Items[i].
                }
                string autori = string.Join(", ", autor);
                wSheet.Cells[i+2, colNumber] = autori;

            }
            lbAutori.Items.Clear();
        }
    }
}
