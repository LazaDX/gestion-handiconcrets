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
using gestion_concrets.Models;
using gestion_concrets.ViewModels;

namespace gestion_concrets.Views
{
    
    public partial class PersonView : Page
    {
        public PersonView()
        {
            try
            {
                InitializeComponent();
                DataContext = new PersonViewModel();
                System.Diagnostics.Debug.WriteLine("PersonView initialisé avec succès");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Erreur dans PersonView : {ex.Message}");
                MessageBox.Show($"Erreur lors du chargement de PersonView : {ex.Message}");
            }
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {

        }
    }
}
