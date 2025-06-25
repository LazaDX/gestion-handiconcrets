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
using gestion_concrets.Services;

namespace gestion_concrets.Views
{
    /// <summary>
    /// Logique d'interaction pour DashboardView.xaml
    /// </summary>
    public partial class DashboardView : Page
    {
        private readonly DatabaseService _databaseService;
        public DashboardView()
        {
            InitializeComponent();
            _databaseService = new DatabaseService();
            try
            {
                LoadStatistics();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Erreur lors du chargement des statistiques : {ex.Message}\n{ex.StackTrace}", "Erreur", MessageBoxButton.OK, MessageBoxImage.Error);
                TotalPersonsText.Text = "Erreur";
                TotalMenText.Text = "Erreur";
                TotalWomenText.Text = "Erreur";
            }
        }

        private void LoadStatistics()
        {
            TotalPersonsText.Text = _databaseService.GetTotalPersonsCount().ToString();
            TotalMenText.Text = _databaseService.GetTotalMenCount().ToString();
            TotalWomenText.Text = _databaseService.GetTotalWomenCount().ToString();
        }

        private void ExportTotal_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                _databaseService.ExportToExcel("Total", "SELECT * FROM BPerson");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Erreur lors de l'exportation : {ex.Message}\n{ex.StackTrace}", "Erreur", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void ExportMen_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                _databaseService.ExportToExcel("Hommes", "SELECT * FROM BPerson WHERE B4 = @Gender", ("@Gender", "Lahy"));
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Erreur lors de l'exportation : {ex.Message}\n{ex.StackTrace}", "Erreur", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void ExportWomen_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                _databaseService.ExportToExcel("Femmes", "SELECT * FROM BPerson WHERE B4 = @Gender", ("@Gender", "Vavy"));
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Erreur lors de l'exportation : {ex.Message}\n{ex.StackTrace}", "Erreur", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
    }
}
