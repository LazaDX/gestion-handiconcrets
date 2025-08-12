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
using gestion_concrets.Views;
using System.Windows.Input;
using System.Windows.Data;

namespace gestion_concrets.ViewModels
{
    public class PersonListViewModel : INotifyPropertyChanged
    {
        private readonly DatabaseService _databaseService;
        private ObservableCollection<BPerson> _persons;

        private string _searchText;
        private ICollectionView _filteredPersons;

        public ICollectionView FilteredPersons
        {
            get
            {
                if (_filteredPersons == null)
                {
                    _filteredPersons = CollectionViewSource.GetDefaultView(Persons);
                    _filteredPersons.Filter = PersonFilter;
                }
                return _filteredPersons;
            }
        }

        public string SearchText
        {
            get => _searchText;
            set
            {
                _searchText = value;
                OnPropertyChanged(nameof(SearchText));
                FilteredPersons.Refresh(); // Actualise le filtre quand le texte change
            }
        }

        public ICommand SearchCommand { get; }
        public ICommand ResetSearchCommand { get; }

        public ObservableCollection<BPerson> Persons
        {
            get => _persons;
            set
            {
                _persons = value;
                OnPropertyChanged();

                _filteredPersons = null; // Réinitialiser pour forcer la recréation
                OnPropertyChanged(nameof(FilteredPersons));
            }
        }



        public ICommand ViewDetailsCommand { get; }

        public PersonListViewModel(DatabaseService databaseService)
        {
            _databaseService = databaseService;
            ViewDetailsCommand = new RelayCommand(ViewDetails);
            SearchCommand = new RelayCommand(ExecuteSearch);
            ResetSearchCommand = new RelayCommand(ExecuteResetSearch);


            LoadPersons();

            DatabaseService.DataChanged += (s, e) => LoadPersons();
        }


        private bool PersonFilter(object item)
        {
            if (string.IsNullOrWhiteSpace(SearchText))
                return true;

            var person = item as BPerson;
            if (person == null)
                return false;

            // Recherche dans plusieurs propriétés
            return person.B2?.IndexOf(SearchText, StringComparison.OrdinalIgnoreCase) >= 0 ||
                   person.B3?.IndexOf(SearchText, StringComparison.OrdinalIgnoreCase) >= 0 ||
                   person.B4?.IndexOf(SearchText, StringComparison.OrdinalIgnoreCase) >= 0 ||
                   person.Adress?.IndexOf(SearchText, StringComparison.OrdinalIgnoreCase) >= 0 ||
                   person.Email?.IndexOf(SearchText, StringComparison.OrdinalIgnoreCase) >= 0 ||
                   person.A1?.IndexOf(SearchText, StringComparison.OrdinalIgnoreCase) >= 0 ||
                   person.A2?.IndexOf(SearchText, StringComparison.OrdinalIgnoreCase) >= 0 ||
                   person.Phone?.IndexOf(SearchText, StringComparison.OrdinalIgnoreCase) >= 0;
        }

        private void ExecuteViewDetails(object parameter)
        {
            if (parameter is BPerson selectedPerson)
            {
                // Logique pour afficher les détails
            }
        }

        private void ExecuteSearch()
        {
            FilteredPersons.Refresh();
        }

        private void ExecuteResetSearch()
        {
            SearchText = string.Empty;
            FilteredPersons.Refresh();
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



        private void ViewDetails(object parameter)
        {
            System.Diagnostics.Debug.WriteLine($"ViewDetails appelé avec parameter: {parameter}");
            if (parameter is BPerson selectedPerson)
            {
                System.Diagnostics.Debug.WriteLine($"SelectedPerson Id: {selectedPerson.Id}");
                try
                {

                    var personViewModel = new PersonViewModel
                    {
                        BPerson = _databaseService.GetBPersonById(selectedPerson.Id) ?? new BPerson(),
                        Alocalisation = _databaseService.GetAlocalisationById(selectedPerson.Id) ?? new Alocalisation(),
                        Ddescription = _databaseService.GetDdescriptionById(selectedPerson.Id) ?? new Ddescription(),
                        Eclimat = _databaseService.GetEclimatById(selectedPerson.Id) ?? new Eclimat(),
                        IsReadOnly = true
                    };

                    var personDetailsWindow = new PersonDetailsWindow(personViewModel);
                    personDetailsWindow.Show();
                    Debug.WriteLine("PersonDetailsWindow affichée");
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"Erreur dans ViewDetails : {ex.Message}\n{ex.StackTrace}");
                }
            }
            else
            {
                System.Diagnostics.Debug.WriteLine("Parameter n'est pas un BPerson");
            }
        }
        public event PropertyChangedEventHandler PropertyChanged;
        protected void OnPropertyChanged([CallerMemberName] string propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }



    }
}
