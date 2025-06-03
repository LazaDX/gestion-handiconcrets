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
       
        public ObservableCollection<BPerson> Persons
        {
            get => _persons;
            set
            {
                _persons = value;
                OnPropertyChanged();
                
            }
        }

       

        public ICommand ViewDetailsCommand { get; }

        public PersonListViewModel(DatabaseService databaseService)
        {
            _databaseService = databaseService;
            ViewDetailsCommand = new RelayCommand(ViewDetails);
           
            LoadPersons();
            DataChangedNotifier.DataChanged += (s, e) => LoadPersons();
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
                        IIapplicationCDPH = _databaseService.GetIIapplicationCDPHById(selectedPerson.Id) ?? new IIapplicationCDPH(),
                        IIIright = _databaseService.GetIIIrightById(selectedPerson.Id) ?? new IIIright(),
                        Itransmission = _databaseService.GetItransmissionById(selectedPerson.Id) ?? new Itransmission(),
                        IVdutyGov = _databaseService.GetIVdutyGovById(selectedPerson.Id) ?? new IVdutyGov(),
                        VdevSupport = _databaseService.GetVdevSupportById(selectedPerson.Id) ?? new VdevSupport(),
                        VIpartnerCollab = _databaseService.GetVIpartnerCollabById(selectedPerson.Id) ?? new VIpartnerCollab(),
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
