using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
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
using word = Microsoft.Office.Interop.Word;
using System.IO;

namespace WpfApp1
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
      
        public static string Tarif;
        public static double priceitog;
        public static int ch = 1;
        public static int minut = 0;
        public MainWindow()
        {
            InitializeComponent();
           
        }
        /// <summary>
        /// Для ввода в поле только чисел
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void min_PreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            Regex regex = new Regex("[^0-9,]+");
            e.Handled = regex.IsMatch(e.Text);
        }
        /// <summary>
        /// Расчёт стоимости 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Button_Click(object sender, RoutedEventArgs e)
        {
           
            try
            {
                Tarif = tarif.Text;
                if (Tarif == "1 тариф")
                {
                    int minutes = Convert.ToInt32(min.Text);
                    if (minutes > 200)
                    {
                        int dif = minutes - 200;
                        double pricedif = dif * 1.6;
                        priceitog = (0.7 * 200) + pricedif;
                        oplata.Content = "К оплате: " + priceitog;
                        overmin.Content = "Минут сверхнормы: " + dif;
                    }
                    else
                    {
                        priceitog = (0.7 * Convert.ToInt32(min.Text));
                        oplata.Content = "К оплате: " + priceitog;
                        overmin.Content = "Минут сверхнормы: 0";
                    }
                }
                else if (Tarif == "2 тариф")
                {
                    int minutes = Convert.ToInt32(min.Text);
                    if (minutes > 100)
                    {
                        int dif = minutes - 100;
                        double pricedif = dif * 1.6;
                        priceitog = (0.3 * 100) + pricedif;
                        oplata.Content = "К оплате: " + priceitog;
                        overmin.Content = "Минут сверхнормы: " + dif;
                    }
                    else
                    {
                        priceitog = (0.3 * Convert.ToInt32(min.Text));
                        oplata.Content = "К оплате: " + priceitog;
                        overmin.Content = "Минут сверхнормы: 0";
                    }
                }
                else MessageBox.Show("Выберите ваш тариф!");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Заполните все поля корректно" + ex.Message);
            }
        }
        /// <summary>
        /// Выод квитанции в ворд
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Button_Click_1(object sender, RoutedEventArgs e)
        {
            minut = Convert.ToInt32(min.Text);
            word.Document document = null;

            word.Application app = new word.Application();

            string putword = Environment.CurrentDirectory.ToString() + @"\шаблон.docx";

            document = app.Documents.Add(putword);

            document.Activate();

            word.Bookmarks bookm = document.Bookmarks;

            word.Range range;

            string[] data = new string[4] { DateTime.Now.ToString("dd.MM.yyyy HH:mm"), Tarif, minut.ToString(), priceitog.ToString("F2") };

            int i = 0;

            foreach (word.Bookmark mark in bookm)

            {

                range = mark.Range;

                range.Text = data[i];

                i++;

            }

            document.SaveAs2(Environment.CurrentDirectory.ToString() + @"\Документ.docx");   

            document.Close();
            
            document = null;
            ch++;
        }
    }
}
    
