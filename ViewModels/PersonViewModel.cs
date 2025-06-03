using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.Reflection;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Input;
using gestion_concrets.Models;
using gestion_concrets.Services;
using gestion_concrets.Views;

namespace gestion_concrets.ViewModels
{
    public class PersonViewModel : BaseViewModel
    {
        private readonly DatabaseService _databaseService;
        private BPerson _bPerson = new BPerson();
        private Alocalisation _alocalisation = new Alocalisation();
        private IIapplicationCDPH _iiApplicationCDPH = new IIapplicationCDPH();
        private IIIright _iiiright = new IIIright();
        private Itransmission _itransmission = new Itransmission();
        private IVdutyGov _ivDutyGov = new IVdutyGov();
        private VdevSupport _vdevSupport = new VdevSupport();
        private VIpartnerCollab _viPartnerCollab = new VIpartnerCollab();
        private int _selectedTabIndex;
        private bool _isReadOnly;
        private DateTime _currentDate = DateTime.Today; // Date
        public BPerson BPerson
        {
            get => _bPerson;
            set => SetProperty(ref _bPerson, value);
        }

        public Alocalisation Alocalisation
        {
            get => _alocalisation;
            set => SetProperty(ref _alocalisation, value);
        }

        public IIapplicationCDPH IIapplicationCDPH
        {
            get => _iiApplicationCDPH;
            set => SetProperty(ref _iiApplicationCDPH, value);
        }

        public IIIright IIIright
        {
            get => _iiiright;
            set => SetProperty(ref _iiiright, value);
        }

        public Itransmission Itransmission
        {
            get => _itransmission;
            set => SetProperty(ref _itransmission, value);
        }

        public IVdutyGov IVdutyGov
        {
            get => _ivDutyGov;
            set => SetProperty(ref _ivDutyGov, value);
        }

        public VdevSupport VdevSupport
        {
            get => _vdevSupport;
            set => SetProperty(ref _vdevSupport, value);
        }

        public VIpartnerCollab VIpartnerCollab
        {
            get => _viPartnerCollab;
            set => SetProperty(ref _viPartnerCollab, value);
        }

        public bool IsReadOnly
        {
            get => _isReadOnly;
            set
            {
                _isReadOnly = value;
                OnPropertyChanged();
            }
        }

        public int SelectedTabIndex
        {
            get => _selectedTabIndex;
            set
            {
                if (SetProperty(ref _selectedTabIndex, value))
                {
                    OnPropertyChanged(nameof(CanGoNext));
                    OnPropertyChanged(nameof(CanGoPrevious));
                }
            }
        }

        public DateTime CurrentDate // Date
        {
            get => _currentDate;
            set => SetProperty(ref _currentDate, value);
        }


        public bool CanGoNext => SelectedTabIndex < 7;
        public bool CanGoPrevious => SelectedTabIndex > 0;

        public ICommand ToggleEditCommand { get; }
        public ICommand SaveCommand { get; }

        public ICommand DeleteCommand { get; }
        public ICommand UpdateCommand { get; }
        public ICommand NextTabCommand { get; }
        public ICommand PreviousTabCommand { get; }

        public ICommand RefreshCommand { get; }

        public event EventHandler CloseWindowRequested;


        // BPerson
        public ObservableCollection<string> SexeOptions { get; } = new ObservableCollection<string> { "Lahy", "Vavy" };
        public ObservableCollection<string> Skills { get; } = new ObservableCollection<string> { "1. Ambaratonga fototra", "2. Ambaratonga", "3. Ambaratonga ambony", "4. Hafa", "5. Nahavita ambaratonga fototra" };
        public ObservableCollection<string> SkillsSubOptions { get; } = new ObservableCollection<string> { "2.1 Collège", "2.2. lycée" };
        public ObservableCollection<string> Working { get; } = new ObservableCollection<string> { "Eny", "Tsia" };
        public ObservableCollection<string> LivingJob { get; } = new ObservableCollection<string> { "Mpiasam-panjakàna", "Mpiasa tsy miankina", "Miasatena" };
        public ObservableCollection<string> Disability { get; } = new ObservableCollection<string> { "Ara-batana", "ara-tsaian", "ara-pahitana", "ara-pihainoana", "ara-paharanitan-tsaina", "aretina tsy sitrana" };
        public ObservableCollection<string> DisabilityEquipment { get; } = new ObservableCollection<string> { "Canne blanche", "appareil auditif", "bequille", "Chaise roulante", "Chaussure orthopédique", "Attèle" };
        public ObservableCollection<string> Worship { get; } = new ObservableCollection<string> { "Mpitovo", "Tokatena", "Mananjanaka", "Tsy manan-janaka", "Manam-bady" };
        
        // Itransmission
        public ObservableCollection<string> Marginalization { get; } = new ObservableCollection<string> { "Eny", "Tsia" };
        public ObservableCollection<string> I51Options { get; } = new ObservableCollection<string> { "Eny", "Tsia" };
        public ObservableCollection<string> I52Options { get; } = new ObservableCollection<string> { "Eny", "Tsia" };
        public ObservableCollection<string> I53Options { get; } = new ObservableCollection<string> { "Eny", "Tsia" };
        public ObservableCollection<string> I54Options { get; } = new ObservableCollection<string> { "Eny", "Tsia" };
        public ObservableCollection<string> I55Options { get; } = new ObservableCollection<string> { "Eny", "Tsia" };
        public ObservableCollection<string> I56Options { get; } = new ObservableCollection<string> { "Eny", "Tsia" };
        public ObservableCollection<string> I57Options { get; } = new ObservableCollection<string> { "Eny", "Tsia" };
        public ObservableCollection<string> I58Options { get; } = new ObservableCollection<string> { "Eny", "Tsia" };
        public ObservableCollection<string> I59Options { get; } = new ObservableCollection<string> { "Eny", "Tsia" };
        public ObservableCollection<string> I510Options { get; } = new ObservableCollection<string> { "Eny", "Tsia" };
        
        // IIapplicationCDPH
        public ObservableCollection<string> II1Options { get; } = new ObservableCollection<string> { "Eny", "Tsia" };
        
        // IVdutyGov
        public ObservableCollection<string> IV11Options { get; } = new ObservableCollection<string> { "Eny", "Tsia" };
        public ObservableCollection<string> IV51Options { get; } = new ObservableCollection<string> { "Eny", "Tsia" };
        
        // VdevSupport
        public ObservableCollection<string> V1Options { get; } = new ObservableCollection<string> { "Eny", "Tsia" };
        public ObservableCollection<string> V41Options { get; } = new ObservableCollection<string> { "Eny", "Tsia" };
        public ObservableCollection<string> V51Options { get; } = new ObservableCollection<string> { "Eny", "Tsia" };
        
        // VIpartnerCollab
        public ObservableCollection<string> VI1Options { get; } = new ObservableCollection<string> { "Eny", "Tsia" };
        public ObservableCollection<string> VI3Options { get; } = new ObservableCollection<string> { "Eny", "Tsia" };
        
        
        public PersonViewModel(DatabaseService databaseService = null, int? idBPerson = null)
        {
            _databaseService = databaseService ?? new DatabaseService();
            IsReadOnly = false; 
            ToggleEditCommand = new RelayCommand(ToggleEdit);
            SaveCommand = new RelayCommand(SaveData);
            UpdateCommand = new RelayCommand(UpdateData);
            DeleteCommand = new RelayCommand(DeleteData, CanDeleteData);
            NextTabCommand = new RelayCommand(NextTab, () => CanGoNext);
            PreviousTabCommand = new RelayCommand(PreviousTab, () => CanGoPrevious);
            RefreshCommand = new RelayCommand(ReinitializeData);

            SelectedTabIndex = 0;

            if (idBPerson.HasValue)
            {
                LoadData(idBPerson.Value);
            }
            else
            {
                BPerson = new BPerson();
                Alocalisation = new Alocalisation();
                IIapplicationCDPH = new IIapplicationCDPH();
                IIIright = new IIIright();
                Itransmission = new Itransmission();
                IVdutyGov = new IVdutyGov();
                VdevSupport = new VdevSupport();
                VIpartnerCollab = new VIpartnerCollab();
            }
        }

       


        private void ToggleEdit(object parameter)
        {
            IsReadOnly = !IsReadOnly;
            if (IsReadOnly && BPerson.Id != 0)
            {
                LoadData(BPerson.Id);
                SetNonDefinedValues(BPerson);
                SetNonDefinedValues(Alocalisation);
                SetNonDefinedValues(IIapplicationCDPH);
                SetNonDefinedValues(IIIright);
                SetNonDefinedValues(Itransmission);
                SetNonDefinedValues(IVdutyGov);
                SetNonDefinedValues(VdevSupport);
                SetNonDefinedValues(VIpartnerCollab);
            }
            else if (!IsReadOnly && BPerson.Id != 0)
            {
                LoadData(BPerson.Id); 
            }
        }

        private bool ValidateFields()
        {
            // Valider B2 (Date de naissance)
            if (!BPerson.DateTime.HasValue ||
                BPerson.DateTime.Value < new DateTime(1900, 1, 1) ||
                BPerson.DateTime.Value > DateTime.Today)
            {
                MessageBox.Show("La date de naissance est invalide. Veuillez entrer une date entre 1900 et aujourd'hui.",
                    "Erreur", MessageBoxButton.OK, MessageBoxImage.Error);
                return false;
            }

            // Valider Email
            if (string.IsNullOrEmpty(BPerson.Email) ||
                !Regex.IsMatch(BPerson.Email, @"^[^@\s]+@[^@\s]+\.com$"))
            {
                MessageBox.Show("L'email est invalide. Veuillez entrer un email au format exemple@domaine.com.",
                    "Erreur", MessageBoxButton.OK, MessageBoxImage.Error);
                return false;
            }

            // Valider Phone
            if (string.IsNullOrEmpty(BPerson.Phone) ||
                !Regex.IsMatch(BPerson.Phone, @"^(038|034|032|037|033)\d{7}$"))
            {
                MessageBox.Show("Le numéro de téléphone est invalide. Il doit comporter 10 chiffres et commencer par 038, 034, 032, 037 ou 033.",
                    "Erreur", MessageBoxButton.OK, MessageBoxImage.Error);
                return false;
            }

            return true;
        }

        private bool validateDateAlocalisation()
        {
            // Valider A1 (Date de resencement)
            if (!Alocalisation.DateTime.HasValue ||
                Alocalisation.DateTime.Value < new DateTime(1900, 1, 1) ||
                Alocalisation.DateTime.Value > DateTime.Today)
            {
                MessageBox.Show("La date saisie est invalide. Veuillez entrer une date entre 1900 et aujourd'hui.",
                    "Erreur", MessageBoxButton.OK, MessageBoxImage.Error);
                return false;
            }

            return true;
        }

        private void ReinitializeData()
        {
            var result = MessageBox.Show("Attention ! Toutes les données saisies seront perdues. Voulez-vous continuer ?", "Confirmation", MessageBoxButton.YesNo, MessageBoxImage.Warning);
            if (result == MessageBoxResult.No)
            {
                return; // Ne pas réinitialiser si l'utilisateur a choisi "Non"
            }
            if (result == MessageBoxResult.Yes)
            {
                BPerson = new BPerson();
                Alocalisation = new Alocalisation();
                IIapplicationCDPH = new IIapplicationCDPH();
                IIIright = new IIIright();
                Itransmission = new Itransmission();
                IVdutyGov = new IVdutyGov();
                VdevSupport = new VdevSupport();
                VIpartnerCollab = new VIpartnerCollab();
                SelectedTabIndex = 0;
            }        
        }

        private void SaveData()
        {
            //if (!ValidateFields() && !validateDateAlocalisation())
            //{
            //    return;
            //}

            try
            {
                var result = MessageBox.Show($"Voulez-vous ajouter ces informations dans la base de donées ?", "Confirmation", MessageBoxButton.YesNo, MessageBoxImage.Question);
                if (result == MessageBoxResult.Yes)
                {
                    SetNonDefinedValues(BPerson);
                    SetNonDefinedValues(Alocalisation);
                    SetNonDefinedValues(IIapplicationCDPH);
                    SetNonDefinedValues(IIIright);
                    SetNonDefinedValues(Itransmission);
                    SetNonDefinedValues(IVdutyGov);
                    SetNonDefinedValues(VdevSupport);
                    SetNonDefinedValues(VIpartnerCollab);


                    _databaseService.AddFullPerson(
                        BPerson,
                        Alocalisation,
                        IIapplicationCDPH,
                        IIIright,
                        Itransmission,
                        IVdutyGov,
                        VdevSupport,
                        VIpartnerCollab
                    );
                    Debug.WriteLine($"[ DONNEES ENREGISTREES AVEC SUCCES ]");
                    MessageBox.Show("Données enregistrées avec succès !", "Information", MessageBoxButton.OK, MessageBoxImage.Exclamation);
                    
                    // Réinitialisation du formulaire
                    BPerson = new BPerson();
                    Alocalisation = new Alocalisation();
                    IIapplicationCDPH = new IIapplicationCDPH();
                    IIIright = new IIIright();
                    Itransmission = new Itransmission();
                    IVdutyGov = new IVdutyGov();
                    VdevSupport = new VdevSupport();
                    VIpartnerCollab = new VIpartnerCollab();
                    SelectedTabIndex = 0;

                    DataChangedNotifier.NotifyDataChanged();
                }
                    
            }
            catch (Exception ex)
            {
                MessageBox.Show("Erreur lors de l'enregistrement : " + ex.Message);
                Debug.WriteLine($"[ERREUR DE L'AJOUT] : {ex.Message}");
            }
        }

        private void UpdateData()
        {
            try
            {
                
                _databaseService.UpdateFullPerson(
                    BPerson,
                    Alocalisation,
                    IIapplicationCDPH,
                    IIIright,
                    Itransmission,
                    IVdutyGov,
                    VdevSupport,
                    VIpartnerCollab
                );
                Debug.WriteLine($"[DONNEES MISES A JOUR AVEC SUCCES]");
                MessageBox.Show("Données mises à jour avec succès !", "Succès",
                    MessageBoxButton.OK, MessageBoxImage.Information);

                DataChangedNotifier.NotifyDataChanged();

                CloseWindowRequested?.Invoke(this, EventArgs.Empty);

            }
            catch (Exception ex)
            {
                MessageBox.Show($"Erreur lors de la mise à jour : {ex.Message}", "Erreur",
                    MessageBoxButton.OK, MessageBoxImage.Error);
                Debug.WriteLine($"[ERREUR DE MISE A JOUR] : {ex.Message}");
            }
        }

        private void SetNonDefinedValues(object model)
        {
            foreach (PropertyInfo prop in model.GetType().GetProperties())
            {
                if (prop.PropertyType == typeof(string))
                {
                    string value = (string)prop.GetValue(model);
                    if (string.IsNullOrWhiteSpace(value))
                    {
                        prop.SetValue(model, "Non définie");
                    }
                }
            }
        }

        private void DeleteData(object parameter)
        {
            try
            {
                if (BPerson.Id == 0)
                {
                    MessageBox.Show("Aucune personne sélectionnée pour la suppression.", "Erreur",
                        MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                var result = MessageBox.Show(
                    $"Voulez-vous vraiment supprimer {BPerson.B1} ?",
                    "Confirmation de suppression",
                    MessageBoxButton.YesNo,
                    MessageBoxImage.Warning);

                if (result == MessageBoxResult.Yes)
                {
                    _databaseService.DeleteFullPerson(BPerson.Id);
                    Debug.WriteLine($"[DONNEES SUPPRIMEES AVEC SUCCES] : Personne ID {BPerson.Id}");
                    MessageBox.Show("Données supprimées avec succès !", "Succès",
                        MessageBoxButton.OK, MessageBoxImage.Information);

                    DataChangedNotifier.NotifyDataChanged();

                    CloseWindowRequested?.Invoke(this, EventArgs.Empty);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Erreur lors de la suppression : {ex.Message}", "Erreur",
                    MessageBoxButton.OK, MessageBoxImage.Error);
                Debug.WriteLine($"[ERREUR DE SUPPRESSION] : {ex.Message}");
            }
        }

        private bool CanDeleteData(object parameter)
        {
            return BPerson.Id != 0; // Activer la commande si une personne existante est sélectionnée
        }

        private void LoadData(int idBPerson)
        {
            //BPerson = _databaseService.GetBPersonById(idBPerson) ?? new BPerson();

            BPerson = _databaseService.GetBPersonById(idBPerson) ?? new BPerson();
            if (BPerson.B2 != "Non définie" && DateTime.TryParseExact(BPerson.B2, "dd/MM/yyyy", null, System.Globalization.DateTimeStyles.None, out var date))
            {
                BPerson.DateTime = date;
            }

            Alocalisation = _databaseService.GetAlocalisationById(idBPerson) ?? new Alocalisation();
            IIapplicationCDPH = _databaseService.GetIIapplicationCDPHById(idBPerson) ?? new IIapplicationCDPH();
            IIIright = _databaseService.GetIIIrightById(idBPerson) ?? new IIIright();
            Itransmission = _databaseService.GetItransmissionById(idBPerson) ?? new Itransmission();
            IVdutyGov = _databaseService.GetIVdutyGovById(idBPerson) ?? new IVdutyGov();
            VdevSupport = _databaseService.GetVdevSupportById(idBPerson) ?? new VdevSupport();
            VIpartnerCollab = _databaseService.GetVIpartnerCollabById(idBPerson) ?? new VIpartnerCollab();

            OnPropertyChanged(nameof(BPerson));
            OnPropertyChanged(nameof(Alocalisation));
            OnPropertyChanged(nameof(IIapplicationCDPH));
            OnPropertyChanged(nameof(IIIright));
            OnPropertyChanged(nameof(Itransmission));
            OnPropertyChanged(nameof(IVdutyGov));
            OnPropertyChanged(nameof(VdevSupport));
            OnPropertyChanged(nameof(VIpartnerCollab));
        }

        private void NextTab()
        {
            if (CanGoNext)
            {
                SelectedTabIndex++;
            }
        }

        private void PreviousTab()
        {
            if (CanGoPrevious)
            {
                SelectedTabIndex--;
            }
        }
    }
}