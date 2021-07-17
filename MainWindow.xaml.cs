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

namespace DocFromTableData
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : System.Windows.Window
    {

        private string pathSrcFile = "";
        private string pathSrcTemplate = "";
        private string pathOutputFolder = "";

        private List<object> tablesSrcData;
        //Словарь совместимости. На один столбец несколько закладок
        private Dictionary<int, List<int>> dictCompatibility;
        private Dictionary<int,string> dictBookmark;


        Word.Application oWordApp;
        Word.Document oWordDoc;
        //Excel.Application oExcelApp;
        //Excel.Workbook oExcelWorkbook;

        public MainWindow()
        {
            oWordApp = new Word.Application();
            InitializeComponent();

        }



        public bool readFromWordSrcDoc()
        {
            oWordDoc = oWordApp.Documents.Open(pathSrcFile);
            dictCompatibility = new Dictionary<int, List<int>>();
            tablesSrcData = new List<object>();
            Dictionary<int, string> dictTitleColumn;
            List<Dictionary<int, string>> dictDataSrc;
            Dictionary<string, object> tableData;
            List<int> listIndex;
            listTitleColumn.Items.Clear();
            //Получаем таблицу с именами Ректоров и названия университетов
            //КАК-ТО РАЗДЕЛИТЬ ИНФУ ПО ТАБЛИЦАМ
            foreach (Word.Table table in oWordDoc.Tables)
            {
                tableData = new Dictionary<string, object>();
                dictTitleColumn = new Dictionary<int, string>();
                dictDataSrc = new List<Dictionary<int, string>>();
                listIndex = new List<int>();
                string bufText;
                
                
                for (int i = 0; i < table.Columns.Count; i++)
                {
                    //Подумать над заменой магического числа 1
                    bufText = table.Cell(1, i).Range.Text.Replace("\r\a","");
                    if(bufText != "" && bufText != "№")
                    {
                        listIndex.Add(i);
                        dictTitleColumn[i] = bufText;
                        listTitleColumn.Items.Add(
                            new TextBlock()
                            {
                                Text = bufText,
                                TextWrapping = TextWrapping.Wrap
                            }
                        ); ;
                    }
                }
                //Получаем данные только с тех столбцов, что были получены с заголовков
                for (int i = 2; i <= table.Rows.Count; i++)
                {
                    dictDataSrc.Add(new Dictionary<int,string>());
                    foreach (int index in listIndex)
                    {
                        dictDataSrc[i-2][index] = table.Rows[i].Cells[index].Range.Text.Replace("\r\a", "");
                    }
                    
                }

                tableData.Add("title",dictTitleColumn);
                tableData.Add("data", dictDataSrc);

                tablesSrcData.Add(tableData);
            }
            lblStatusWork.Content = "Данные источника считаны!";
            oWordDoc.Close();
            return true;

        }


        public bool readFromWordTemplateDoc()
        {
            oWordDoc = oWordApp.Documents.Open(pathSrcTemplate);
            dictBookmark = new Dictionary<int, string>();
            dictCompatibility = new Dictionary<int, List<int>>();
            listChkBoxBookmarks.Items.Clear();
            int i = 0;
            foreach (Word.Bookmark item in oWordDoc.Bookmarks)
            {
                dictBookmark.Add(i, item.Range.Text);
                listChkBoxBookmarks.Items.Add(new CheckBox()
                {
                    Tag = i,
                    Content = item.Range.Text
                });
                i++;
            }
            oWordDoc.Close();
            
            return true;
        }


        public bool writeToTemplate()
        {
            foreach (KeyValuePair<int,String[]> infoUniversity in tablesSrcData)
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
                readFromWordSrcDoc();

            }
        }

        private void btnFileSelectTemplate_Click(object sender, RoutedEventArgs e)
        {
            var selectFilePicker = new Microsoft.Win32.OpenFileDialog();
            if (selectFilePicker.ShowDialog() == true)
            {
                pathSrcTemplate = selectFilePicker.FileName;
                txtBoxPathSelectTemplate.Text = pathSrcTemplate;
                readFromWordTemplateDoc();
            }
        }

        private void btnSelectFolderOnSave_Click(object sender, RoutedEventArgs e)
        {
            //var selectFolderPicker = new FolderBrowserDialog();
            //selectFolderPicker.ShowDialog();
            //pathOutputFolder = selectFolderPicker.SelectedPath;
            //txtBoxPathSelectOutputFolder.Text = pathOutputFolder;
        }

        private void btnStartGenerateFiles_Click(object sender, RoutedEventArgs e)
        {
            lblStatusWork.Content = "";
            if (pathSrcFile != "" && pathSrcTemplate != "" && pathOutputFolder != "")
            {
                lblStatusWork.Content = "В процессе.";
                Thread.Sleep(500);
                writeToTemplate();
                lblStatusWork.Content = "Завершено!";
                oWordApp.Quit();
            }
            else
            {
                lblStatusWork.Content = "Не все пути указаны!";
            }
        }
    }
}
