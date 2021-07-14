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
using Word = Microsoft.Office.Interop.Word;
using Excel = Microsoft.Office.Interop.Excel;
using System.Threading;
using System.IO;
using Microsoft.Win32;
using System.Windows.Forms;

namespace DocFromTableData
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : System.Windows.Window
    {

        private string pathSrcFile;
        private string pathSrcTemplate;
        private string pathOutputFolder;

        public Dictionary<int, String[]> listUniversity;
        Word.Application oWordApp;
        Word.Document oWordDoc;
        //Excel.Application oExcelApp;
        //Excel.Workbook oExcelWorkbook;

        public MainWindow()
        {
            oWordApp = new Word.Application();
            //oExcelApp = new Excel.Application();
            listUniversity = new Dictionary<int, String[]>();
            InitializeComponent();
            
        }

        public bool readFromWordSrc()
        {
            oWordDoc = oWordApp.Documents.Open(pathSrcFile);
            //Получаем таблицу с именами Ректоров и названия университетов
            string nameEstablishment;
            string nameDirector;
            foreach (Word.Table table in oWordDoc.Tables)
            {
                for (int i = 2; i < table.Rows.Count; i++)
                {
                    nameEstablishment = table.Cell(i, 2).Range.Text.Trim();
                    nameDirector = table.Cell(i, 3).Range.Text.Trim(new char[]{'\r','\a'});
                    listUniversity.Add(i -1, new string[] { nameEstablishment, nameDirector });
                    Console.WriteLine($"{i-1}:\n Наименование вуза: {nameEstablishment} \n Имя Ректора/Директора: {nameDirector} \n");
                }
            }
            oWordDoc.Close();
            return true;
        }

        public bool writeToTemplate()
        {
            foreach (KeyValuePair<int,String[]> infoUniversity in listUniversity)
            {
                //TODO переделать считываниеы
                oWordDoc =  oWordApp.Documents.Open(pathSrcTemplate);
                oWordDoc.Bookmarks["FullNameFirst"].Range.Text= infoUniversity.Value[1];
                oWordDoc.Bookmarks["FullNameSecond"].Range.Text = infoUniversity.Value[1];
                oWordDoc.Bookmarks["UniversityTitle"].Range.Text = infoUniversity.Value[0];
                oWordDoc.SaveAs2($"{pathOutputFolder}\\{infoUniversity.Key}.docx");
                Console.WriteLine($"Запись {infoUniversity.Key}.docx - Завершена!");
                lblStatusWork.Content = $"Запись {infoUniversity.Key}.docx - Завершена!";


            }
            oWordDoc.Close();
            return true;
        }


        private void btnFileDialogSrc_Click(object sender, RoutedEventArgs e)
        {
            var selectFilePicker = new Microsoft.Win32.OpenFileDialog();
            if (selectFilePicker.ShowDialog() == true)
            {
                pathSrcFile = selectFilePicker.FileName;
                txtBoxPathFileSrcData.Text = pathSrcFile;


            }
        }

        private void btnFileSelectTemplate_Click(object sender, RoutedEventArgs e)
        {
            var selectFilePicker = new Microsoft.Win32.OpenFileDialog();
            if (selectFilePicker.ShowDialog() == true)
            {
                pathSrcTemplate = selectFilePicker.FileName;
                txtBoxPathSelectTemplate.Text = pathSrcTemplate;
            }
        }

        private void btnSelectFolderOnSave_Click(object sender, RoutedEventArgs e)
        {
            var selectFolderPicker = new FolderBrowserDialog();
            selectFolderPicker.ShowDialog();
            pathOutputFolder = selectFolderPicker.SelectedPath;
            txtBoxPathSelectOutputFolder.Text = pathOutputFolder;
        }

        private void btnStartGenerateFiles_Click(object sender, RoutedEventArgs e)
        {
            lblStatusWork.Content = "В процессе.";
            Thread.Sleep(500);
            readFromWordSrc();
            writeToTemplate();
            lblStatusWork.Content = "Завершено!";
            oWordApp.Quit();
        }
    }
}
