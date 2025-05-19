using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Runtime.CompilerServices;
using gestion_concrets.Models;
using gestion_concrets.Services;
using Wpf.Ui.Controls;
using System.Windows;
using System.Diagnostics;

namespace gestion_concrets.ViewModels
{
    public class PersonListViewModel : INotifyPropertyChanged
    {
        private readonly DatabaseService _databaseService;
        private ObservableCollection<BPerson> _persons;

        public ObservableCollection<BPerson> Persons
        {
            get => _persons;
            set
            {
                _persons = value;
                OnPropertyChanged();
            }
        }

        public PersonListViewModel(DatabaseService databaseService)
        {
            _databaseService = databaseService;
            LoadPersons();
        }

        private void LoadPersons()
        {
            try
            {
                Persons = new ObservableCollection<BPerson>(_databaseService.GetAllPersons());
            }
            catch (Exception ex)
            {
                // À adapter selon ton gestionnaire d'erreurs
                Debug.WriteLine($"[ERREUR LORS DE L'AFFICHAGE] : {ex.Message}");
                
            }
        }

        public event PropertyChangedEventHandler PropertyChanged;
        protected void OnPropertyChanged([CallerMemberName] string propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}
