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
    
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
            try
            {
                MainFrame.Navigate(new DashboardView());
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Erreur lors de la navigation vers le tableau de bord : {ex.Message}\n{ex.StackTrace}", "Erreur", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

     

        private void dashBt_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                MainFrame.Navigate(new DashboardView());
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Erreur lors de la navigation vers le tableau de bord : {ex.Message}\n{ex.StackTrace}", "Erreur", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void gestBtn(object sender, RoutedEventArgs e)
        {
            MainFrame.Navigate(new PersonListPage());
        }

        private void addBtn_Click(object sender, RoutedEventArgs e)
        {
            
            MainFrame.Navigate(new PersonView());
        }

        private void dataBaseBtn_Click(object sender, RoutedEventArgs e)
        {
            MainFrame.Navigate(new DatabaseView());
        }
    }
}