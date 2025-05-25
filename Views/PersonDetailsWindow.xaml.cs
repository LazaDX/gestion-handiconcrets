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
using System.Windows.Shapes;
using gestion_concrets.ViewModels;

namespace gestion_concrets.Views
{
    /// <summary>
    /// Logique d'interaction pour PersonDetailsWindow.xaml
    /// </summary>
    public partial class PersonDetailsWindow : Window
    {
        public PersonDetailsWindow(PersonViewModel viewModel)
        {
            InitializeComponent();
            DataContext = viewModel;
            Content = new PersonDetailsView
            {
                DataContext = viewModel
            };

        }
    }
}
