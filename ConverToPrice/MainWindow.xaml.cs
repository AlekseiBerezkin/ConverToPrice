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

namespace ConverToPrice
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }
        private List<InputItem> list = new List<InputItem>();
        private async void ButtonOpenXlsx_Click(object sender, RoutedEventArgs e)
        {
            List<string> fn = OpenSaveDialogs.OpenDialog("sf", false, DialogFilterType.Xlsx);

            if (fn == null) { }
            else
            {
                ButtonOpenXlsx.IsEnabled = false;
                PrintEvent("Чтение файла");
                try
                {
                    await Task.Run((Action)(() => list = ExcelReader.ReadXlsxOzon(fn.First<string>())));
                }
                catch (Exception ex)
                {
                    PrintEvent("Данные не были прочитаны. Ошибка:"+ex);
                    return;
                }
                
                PrintEvent("Прочитано: " + list.Count.ToString() + " строк.");
                ButtonRun.IsEnabled = true;
                ButtonOpenXlsx.IsEnabled = false;
            }
        }

        private void PrintEvent(string e)
        {
            Jurnal.Items.Add(DateTime.Now.ToString("hh:mm:ss") +" "+e);
        }

        private async void ButtonRun_Click(object sender, RoutedEventArgs e)
        {
            PrintEvent("Старт преобразования");
            string fn = "ИП Вельмога. Прайс лист " + DateTime.Now.ToString("dd-MM-yy") + ".xlsx";
            string file = OpenSaveDialogs.SaveDialog("sf2", DialogFilterType.Xlsx, fn);
            if (file == null)
            {
                fn = (string)null;
            }
            else
            {
                await Task.Run((Action)(() =>
                {
                    Dictionary<string, List<InputItem>> dictionary = new Dictionary<string, List<InputItem>>();
                    Dispatcher.Invoke(() => { PrintEvent("Вычисление количества"); });
                    foreach (InputItem inputItem in this.list)
                    {
                        if (!dictionary.ContainsKey(inputItem._1_Id))
                            dictionary.Add(inputItem._1_Id, new List<InputItem>());
                        dictionary[inputItem._1_Id].Add(inputItem);
                    }
                    Dispatcher.Invoke(() => { PrintEvent("Сборка файла"); });
                    List<InputItem> items_list = new List<InputItem>();
                    foreach (KeyValuePair<string, List<InputItem>> keyValuePair in dictionary)
                    {
                        InputItem inputItem = new InputItem()
                        {
                            _1_Id = keyValuePair.Value.First<InputItem>()._1_Id,
                            _2_Name = keyValuePair.Value.First<InputItem>()._2_Name,
                            _3_Year = keyValuePair.Value.First<InputItem>()._3_Year,
                            _4_Mrk = keyValuePair.Value.First<InputItem>()._4_Mrk,
                            _5_Rozn = keyValuePair.Value.First<InputItem>()._5_Rozn,
                            _6_CountOZ = keyValuePair.Value.Where(p => p.isWB == false).Count(),
                            _8_Count = (double)keyValuePair.Value.Count,
                            _7_CountWB =keyValuePair.Value.Where(p => p.isWB == true).Count()
                        };
                        items_list.Add(inputItem);
                    }
                    
                    try
                    {
                        Dispatcher.Invoke(() => { PrintEvent("Запись в файл"); });
                        Exporter.WriteFile(file, items_list);
                        Dispatcher.Invoke(() => { PrintEvent("Данные успешно записаны"); });
                    }
                    catch (Exception ex)
                    {
                        Dispatcher.Invoke(() => { PrintEvent("ОШИБКА при записи файла"+ex); });
                    }

                }));
                fn = (string)null;
            }
            list.Clear();
            ButtonRun.IsEnabled = false;
            ButtonOpenXlsx.IsEnabled = true;
        }
    }
}
