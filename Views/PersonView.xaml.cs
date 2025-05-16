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
    /// <summary>
    /// Logique d'interaction pour PersonView.xaml
    /// </summary>
    public partial class PersonView : Page
    {
        public PersonView()
        {
            InitializeComponent();
            this.DataContext = new PersonViewModel();
        }
    }
}
