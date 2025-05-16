using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using gestion_concrets.Services;
using gestion_concrets.ViewModels;
using gestion_concrets.Views;

namespace gestion_concrets
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

     

        private void dashBt_Click(object sender, RoutedEventArgs e)
        {

        }

        private void gestBtn(object sender, RoutedEventArgs e)
        {
           
        }

        private void addBtn_Click(object sender, RoutedEventArgs e)
        {
            MainFrame.Navigate(new PersonView());
        }
    }
}