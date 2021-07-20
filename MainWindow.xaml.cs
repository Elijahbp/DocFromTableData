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

        private List<Dictionary<string,object>> tablesSrcData;
        //Словарь совместимости. На один столбец несколько закладок
        private Dictionary<int, List<int>> dictCompatibility;
        private Dictionary<int,string> dictBookmark;
        List<int> listSelectedIndex;


        Word.Application oWordApp;
        //Excel.Application oExcelApp;
        //Excel.Workbook oExcelWorkbook;
        int selectedColumnBlockTag = int.MinValue;
        int selectedColumnToTitle = 0;
        //CheckBox selectedCheckBox;


        public MainWindow()
        {
            oWordApp = new Word.Application();
            dictCompatibility = new Dictionary<int, List<int>>();
            tablesSrcData = new List<Dictionary<string,object>>();
            listSelectedIndex = new List<int>();
            InitializeComponent();
        }



        public async void readFromWordSrcDoc()
        {
            Word.Document oWordDoc = oWordApp.Documents.Open(pathSrcFile);
            
            Dictionary<int, string> dictTitleColumn;
            List<Dictionary<int, string>> dictDataSrc;
            Dictionary<string, object> tableData;
            List<int> listIndex;
            listTitleColumn.Items.Clear();
            //Получаем таблицу с именами Ректоров и названия университетов
            //КАК-ТО РАЗДЕЛИТЬ ИНФУ ПО ТАБЛИЦАМ
            await Task.Run(() => {
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
                        bufText = table.Cell(1, i).Range.Text.Replace("\r", "").Replace("\a", "");
                        if (bufText != "" && bufText != "№")
                        {
                            listIndex.Add(i);
                            dictTitleColumn[i] = bufText;
                        }
                    }
                    //Получаем данные только с тех столбцов, что были получены с заголовков
                    for (int i = 2; i <= table.Rows.Count; i++)
                    {
                        dictDataSrc.Add(new Dictionary<int, string>());
                        foreach (int index in listIndex)
                        {
                            //TODO - придумать что-то с волшебным числом 2!!!
                            dictDataSrc[i - 2][index] = table.Rows[i].Cells[index].Range.Text.Replace("\r", "").Replace("\a","");
                        }

                    }

                    tableData.Add("title", dictTitleColumn);
                    tableData.Add("data", dictDataSrc);
                    tablesSrcData.Add(tableData);
                }
            });
            oWordDoc.Close();
            lblComboBox.Visibility = Visibility.Visible;
            comboBoxTitles.Visibility = Visibility.Visible;
            foreach (Dictionary<string,object> table in tablesSrcData)
            {
                Dictionary<int,string> titleData = (Dictionary<int, string>) table["title"];
                foreach (KeyValuePair<int,string> content in titleData)
                {
                    listTitleColumn.Items.Add(getTextBlockColumnData(content.Value,content.Key));
                    comboBoxTitles.Items.Add(getTextBlockColumnData(content.Value, content.Key));
                }
            }
            lblComboBox.IsEnabled = true;
            comboBoxTitles.IsEnabled = true;
            lblStatusWork.Content = "Данные источника считаны!";
        }



        public async void readFromWordTemplateDoc()
        {
            Word.Document oWordDoc = oWordApp.Documents.Open(pathSrcTemplate);
            dictBookmark = new Dictionary<int, string>();
            dictCompatibility = new Dictionary<int, List<int>>();
            listChkBoxBookmarks.Items.Clear();
            int i = 0;
            await Task.Run(()=>
            {
                foreach (Word.Bookmark item in oWordDoc.Bookmarks)
                {
                    dictBookmark.Add(i, item.Range.Text.Replace("\r", "").Replace("\a", ""));
                    i++;
                }
            });
            oWordDoc.Close();
            foreach (KeyValuePair<int,string> kvPair in dictBookmark)
            {
                listChkBoxBookmarks.Items.Add(getCheckBoxBookmarks(kvPair.Value, kvPair.Key));
            }



        }

        public void generateDocuments()
        {
            int i = 1;
            string nameBookmarks;
            string dataColumn;
            string titleDocument;
            foreach (Dictionary<string,object> table in tablesSrcData)
            {
                   
                foreach (Dictionary<int,string> rowData in (List<Dictionary<int, string>>)table["data"])
                {
                    Word.Document oWordDoc = oWordApp.Documents.Open(pathSrcTemplate);
                    foreach (KeyValuePair<int, List<int>> kvPair in dictCompatibility)
                    {
                        foreach (int indexBookmarks in kvPair.Value)
                        {
                            nameBookmarks = dictBookmark[indexBookmarks];
                            dataColumn = rowData[kvPair.Key];
                            oWordDoc.Bookmarks[nameBookmarks].Range.Text = dataColumn;
                        }
                    }
                    titleDocument = rowData[selectedColumnToTitle].Replace(" ", "_").Replace("\"","");
                    if (titleDocument == "")
                    {
                        titleDocument = "empty_title_" + i;
                        i++;
                    }
                    oWordDoc.SaveAs2($"{pathOutputFolder}\\{titleDocument}.docx");//TODO ПОМЕНЯТЬ
                    oWordDoc.Close();
                }
            }
            lblStatusWork.Content = $"Запись документов - Завершена!";
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
           var selectFolderPicker = new System.Windows.Forms.FolderBrowserDialog();
           selectFolderPicker.ShowDialog();
           pathOutputFolder = selectFolderPicker.SelectedPath;
           txtBoxPathSelectOutputFolder.Text = pathOutputFolder;
        }

        private void btnStartGenerateFiles_Click(object sender, RoutedEventArgs e)
        {
            lblStatusWork.Content = "";
            if (pathSrcFile != "" && pathSrcTemplate != "" && pathOutputFolder != "")
            {
                lblStatusWork.Content = "В процессе.";
                Thread.Sleep(500);
                generateDocuments();
                lblStatusWork.Content = "Завершено!";
                oWordApp.Quit();
            }
            else
            {
                lblStatusWork.Content = "Не все пути указаны!";
            }
        }

        private ListBoxItem getTextBlockColumnData(string content,int tag)
        {
            ListBoxItem listBoxItem = new ListBoxItem()
            {
                Name = "listBoxItem_" + tag,
                Tag = tag,
                Content = content,
                //TextWrapping = TextWrapping.Wrap,
            };
            listBoxItem.Selected += selectItemListBoxColumData;
            return listBoxItem;
        }

        private CheckBox getCheckBoxBookmarks(string content, int tag)
        {
            CheckBox checkBox = new CheckBox()
            {
                Name = "checkBox_" + tag,
                Tag = tag,
                Content = content,
            };
            checkBox.Checked += selectCheckBoxBookmarks;
            checkBox.Unchecked += unselectCheckBoxBookmarks;
            return checkBox;
        }

        private void selectItemListBoxColumData(object sender, RoutedEventArgs e)
        {
            ListBoxItem selectedColumnBlock = (ListBoxItem)sender;
            if (selectedColumnBlockTag != (int)selectedColumnBlock.Tag)
            {
                selectedColumnBlockTag = (int)selectedColumnBlock.Tag;
                //Корректно представляем выбранные заголовки у других столбцов (делаем недоступными)
                List<int> selectedBookmarks = new List<int>();
                foreach (KeyValuePair<int,List<int>> kvPair in dictCompatibility)
                {
                    // Если выбран столбец, у которого нет выбранных заголовков, все ранее выбраные заголовки скрываем от выбора, а те что привязаны к нему, и свободны - выводим
                    if (kvPair.Key != selectedColumnBlockTag)
                    {
                        selectedBookmarks.AddRange(kvPair.Value);
                    }
                    
                }
                foreach (CheckBox checkBox in listChkBoxBookmarks.Items)
                {
                    if (selectedBookmarks.Contains((int)checkBox.Tag))
                    {
                        checkBox.IsEnabled = false;
                    }
                    else
                    {
                        checkBox.IsEnabled = true;
                    }
                }
            }
        }

        private void selectedCombBoxColumnToTitle(object sender, RoutedEventArgs e)
        {
            selectedColumnToTitle = (int)((ListBoxItem)((ComboBox)sender).SelectedItem).Tag;
        }

        private void selectCheckBoxBookmarks(object sender, RoutedEventArgs e)
        {
            CheckBox selectedCheckBox = (CheckBox)sender;
            if (selectedColumnBlockTag != int.MinValue) {
                List<int> selectedBookmarks;
                if (dictCompatibility.TryGetValue(selectedColumnBlockTag, out selectedBookmarks))
                {
                    selectedBookmarks.Add((int)selectedCheckBox.Tag);
                }
                else
                {
                    dictCompatibility.Add(selectedColumnBlockTag,new List<int>() { (int)selectedCheckBox.Tag});
                }
            }
            else
            {
                selectedCheckBox.IsChecked = false;
                lblStatusWork.Content = "Перед выбором загаловков, выберите столбец!";
            }
        }
        private void unselectCheckBoxBookmarks(object sender, RoutedEventArgs e)
        {
            CheckBox selectedCheckBox = (CheckBox)sender;
            if (selectedColumnBlockTag != int.MinValue)
            {
                List<int> selectedBookmarks;
                if (dictCompatibility.TryGetValue(selectedColumnBlockTag, out selectedBookmarks))
                {
                    selectedBookmarks.Remove((int)selectedCheckBox.Tag);
                    if(selectedBookmarks.Count == 0)
                    {
                        dictCompatibility.Remove(selectedColumnBlockTag);
                        selectedColumnBlockTag = int.MinValue;
                    }
                }
            }
        }
    }
}
