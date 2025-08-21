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
        private Ddescription _ddescription = new Ddescription();
        private Eclimat _eclimat = new Eclimat();

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

        public Ddescription Ddescription
        {
            get => _ddescription;
            set => SetProperty( ref _ddescription, value);
        }

        public Eclimat Eclimat
        {
            get => _eclimat;
            set => SetProperty(ref _eclimat, value);
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


        private bool _isWatterSelected;
        private bool _isFoodSelected;
        private bool _isAgricultureSelected;
        private bool _isBreedingSelected;
        private bool _isFishingSelected;
        private bool _isOtherJobSelected;
        private bool _isLivingHouseSelected;
        private bool _isHealthSelected;

        public bool IsWatterSelected
        {
            get => _isWatterSelected;
            set
            {
                _isWatterSelected = value;
                OnPropertyChanged();
                UpdateE5();
            }
        }

        public bool IsFoodSelected
        {
            get => _isFoodSelected;
            set
            {
                _isFoodSelected = value;
                OnPropertyChanged();
                UpdateE5();
            }
        }
       
        public bool IsAgricultureSelected
        {
            get => _isAgricultureSelected;
            set
            {
                _isAgricultureSelected = value;
                OnPropertyChanged();
                UpdateE5();
            }
        }
       
        public bool IsBreedingSelected
        {
            get => _isBreedingSelected;
            set
            {
                _isBreedingSelected = value;
                OnPropertyChanged();
                UpdateE5();
            }
        }
        
      
        public bool IsFishingSelected
        {
            get => _isFishingSelected;
            set
            {
                _isFishingSelected = value;
                OnPropertyChanged();
                UpdateE5();
            }
        }
       
        public bool IsOtherJobSelected
        {
            get => _isOtherJobSelected;
            set
            {
                _isOtherJobSelected = value;
                OnPropertyChanged();
                UpdateE5();
            }
        }
       
        public bool IsLivingHouseSelected
        {
            get => _isLivingHouseSelected;
            set
            {
                _isLivingHouseSelected = value;
                OnPropertyChanged();
                UpdateE5();
            }
        }

        public bool IsHealthSelected
        {
            get => _isHealthSelected;
            set
            {
                _isHealthSelected = value;
                OnPropertyChanged();
                UpdateE5();
            }
        }


        public ICommand ToggleEditCommand { get; }
        public ICommand SaveCommand { get; }

        public ICommand DeleteCommand { get; }
        public ICommand UpdateCommand { get; }
        public ICommand NextTabCommand { get; }
        public ICommand PreviousTabCommand { get; }

        public ICommand RefreshCommand { get; }

        public event EventHandler CloseWindowRequested;

        // Alocalisation
        public ObservableCollection<string> A2 { get; } = new ObservableCollection<string> { "SAVA", "DIANA", "ITASY", "ANALAMANAGA", "VAKINANANKARATRA", "BONGOLAVA", "SOFIA", "BOENY", "BETSIBOKA", "MELAKY", "ATSIMO ATSINANANA", "AMORON'I MANIA", "HAUTE MATSIATRA", "VATOVAVY", "FITOVINANY", "ATSIMO ANDREFANA", "ANOSY", "ANDROY", "MENABE", "BETSIBOKA", "ALAOTRA MANGORO", "ANALANJIROFO", "IHOROMBE"};

        // BPerson
        public ObservableCollection<string> Questions { get; } = new ObservableCollection<string> { "Eny", "Tsia" };
        public ObservableCollection<string> SexeOptions { get; } = new ObservableCollection<string> { "Lahy", "Vavy" };
        public ObservableCollection<string> Relationship { get; } = new ObservableCollection<string> { "Mpitovo: 1", "Manambady ara-dalàna: 2", "Nisara-bady: 3", "Maty vady: 4" };
        public ObservableCollection<string> LivingJob { get; } = new ObservableCollection<string> { "Mpiasam-panjakàna: 1", "Mpiasa tsy miankina: 2", "Miasa tena: 3", "Mpiompy: 4", "Tsy manana asa raikitra: 5", "Mpamboly: 6",  "Mpanjono: 7"};

        // Ddescription
        public ObservableCollection<string> Disability { get; } = new ObservableCollection<string> { "Ara-batana: 1", "ara-tsaina: 2", "ara-pahitana: 3", "ara-pihainoana: 4", "Ara-paharanitan-tsaina: 5", "aretina tsy sitrana: 6"};
        public ObservableCollection<string> Material { get; } = new ObservableCollection<string> { "tehina fotsy: 1", "appareil auditif: 2", "canne anglaise: 3", "Chaise roulante: 4", "béquille: 5"};

        // Eclimat
        public ObservableCollection<string> E1 { get; } = new ObservableCollection<string> { "Eny", "Tsia" };
        public ObservableCollection<string> E4 { get; } = new ObservableCollection<string> { "Eny", "Tsia" };
        public ObservableCollection<string> E5 { get; } = new ObservableCollection<string> { "Eny", "Tsia" };
        public ObservableCollection<string> E611 { get; } = new ObservableCollection<string> { "Eny", "Tsia" };
        public ObservableCollection<string> E621 { get; } = new ObservableCollection<string> { "Eny", "Tsia" };
        public ObservableCollection<string> E631 { get; } = new ObservableCollection<string> { "Eny", "Tsia" };
        public ObservableCollection<string> E641 { get; } = new ObservableCollection<string> { "Eny", "Tsia" };
        public ObservableCollection<string> E651 { get; } = new ObservableCollection<string> { "Eny", "Tsia" };
        public ObservableCollection<string> E661 { get; } = new ObservableCollection<string> { "Eny", "Tsia" };
        public ObservableCollection<string> E71 { get; } = new ObservableCollection<string> { "Eny", "Tsia" };
        public ObservableCollection<string> E81 { get; } = new ObservableCollection<string> { "Eny", "Tsia" };
        public ObservableCollection<string> E911 { get; } = new ObservableCollection<string> { "Eny", "Tsia" };
        public ObservableCollection<string> E921 { get; } = new ObservableCollection<string> { "Eny", "Tsia" };
        public ObservableCollection<string> E931 { get; } = new ObservableCollection<string> { "Eny", "Tsia" };
        public ObservableCollection<string> E941 { get; } = new ObservableCollection<string> { "Eny", "Tsia" };
        public ObservableCollection<string> E951 { get; } = new ObservableCollection<string> { "Eny", "Tsia" };
        public ObservableCollection<string> E961 { get; } = new ObservableCollection<string> { "Eny", "Tsia" };
        public ObservableCollection<string> E971 { get; } = new ObservableCollection<string> { "Eny", "Tsia" };
        public ObservableCollection<string> E981 { get; } = new ObservableCollection<string> { "Eny", "Tsia" };
        public ObservableCollection<string> E101 { get; } = new ObservableCollection<string> { "Eny", "Tsia" };
        public ObservableCollection<string> E111 { get; } = new ObservableCollection<string> { "Eny", "Tsia" };
        public ObservableCollection<string> E121 { get; } = new ObservableCollection<string> { "Eny", "Tsia" };
        public ObservableCollection<string> E131 { get; } = new ObservableCollection<string> { "Eny", "Tsia" };
        public ObservableCollection<string> E141 { get; } = new ObservableCollection<string> { "Eny", "Tsia" };


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
                Ddescription = new Ddescription();
                Eclimat = new Eclimat();
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
                    SetNonDefinedValues(Ddescription);
                    SetNonDefinedValues(Eclimat);
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
                MessageBox.Show("La date saisie dans la localisation ou dans le grand titre A et sous titre A1 est invalide ou ne doit pas être vide. Veuillez entrer une date entre 1900 et aujourd'hui.",
                    "Erreur", MessageBoxButton.OK, MessageBoxImage.Error);
                return false;
            }

            return true;
        }

        private bool validateBPersonFullname()
        {
            // Valider B1 (Nom et prénom)
            if (string.IsNullOrWhiteSpace(BPerson.B2))
            {
                MessageBox.Show("Le nom (Anarana) et prénom (Fanampin'anarana) ne doivent pas être vides.",
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
                Ddescription = new Ddescription();
                Eclimat = new Eclimat();
                SelectedTabIndex = 0;
                UpdateCheckBoxesFromE5();
            }        
        }

        private void SaveData()
        {
            if (!validateDateAlocalisation())
            {
                return;
            }else if (!validateBPersonFullname())
            {
                return;
            }

                try
                {
                    var result = MessageBox.Show($"Confirmer l'insertion si tout les informations que vous voulez enregistrer sont toutes saisies sinon, veuillez compléter et affirmer que tout les informations saisies sont correctes avant d'enregistrer dans la base de données !", "Confirmation", MessageBoxButton.YesNo, MessageBoxImage.Question);
                    if (result == MessageBoxResult.Yes)
                    {
                        SetNonDefinedValues(BPerson);
                        SetNonDefinedValues(Alocalisation);
                        SetNonDefinedValues(Ddescription);
                        SetNonDefinedValues(Eclimat);


                        _databaseService.AddFullPerson(
                            BPerson,
                            Alocalisation,
                            Ddescription,
                            Eclimat
                        );
                        Debug.WriteLine($"[ DONNEES ENREGISTREES AVEC SUCCES ]");
                        MessageBox.Show("Données enregistrées avec succès !", "Information", MessageBoxButton.OK, MessageBoxImage.Exclamation);

                        // Réinitialisation du formulaire
                        BPerson = new BPerson();
                        Alocalisation = new Alocalisation();
                        Ddescription = new Ddescription();
                        Eclimat = new Eclimat();
                        SelectedTabIndex = 0;
                        UpdateCheckBoxesFromE5();

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
                    Ddescription,
                    Eclimat
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
            return BPerson.Id != 0; 
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
            Ddescription = _databaseService.GetDdescriptionById(idBPerson) ?? new Ddescription();
            Eclimat = _databaseService.GetEclimatById(idBPerson) ?? new Eclimat();
            UpdateCheckBoxesFromE5();

            OnPropertyChanged(nameof(BPerson));
            OnPropertyChanged(nameof(Alocalisation));
            OnPropertyChanged(nameof(Ddescription));
            OnPropertyChanged(nameof(Eclimat));
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

        private void UpdateE5()
        {
            var selectedOptions = new List<string>();
            if (IsWatterSelected) selectedOptions.Add("Fisitrahana rano: 1");
            if (IsFoodSelected) selectedOptions.Add("Sakafo: 2");
            if (IsAgricultureSelected) selectedOptions.Add("Fambolena: 3");
            if (IsBreedingSelected) selectedOptions.Add("Fiompiana: 4");
            if (IsFishingSelected) selectedOptions.Add("Jono: 5");
            if (IsOtherJobSelected) selectedOptions.Add("Asa fivelomana hafa: 6");
            if (IsLivingHouseSelected) selectedOptions.Add("trano fonenana: 7");
            if (IsHealthSelected) selectedOptions.Add("Fahasalamana: 8");
            Eclimat.E5 = string.Join(", ", selectedOptions);
        }

        private void UpdateCheckBoxesFromE5()
        {
            if (string.IsNullOrEmpty(Eclimat.E5))
            {
                IsWatterSelected = false;
                IsFoodSelected = false;
                IsAgricultureSelected = false;
                IsBreedingSelected = false;
                IsFishingSelected = false;
                IsOtherJobSelected = false;
                IsLivingHouseSelected = false;
                IsHealthSelected = false;
                return;
            }

            var selected = Eclimat.E5.Split(',').Select(s => s.Trim()).ToList();
            IsWatterSelected = selected.Contains("Fisitrahana rano: 1");
            IsFoodSelected = selected.Contains("Sakafo: 2");
            IsAgricultureSelected = selected.Contains("Fambolena: 3");
            IsBreedingSelected = selected.Contains("Fiompiana: 4");
            IsFishingSelected = selected.Contains("Jono: 5");
            IsOtherJobSelected = selected.Contains("Asa fivelomana hafa: 6");
            IsLivingHouseSelected = selected.Contains("trano fonenana: 7");
            IsHealthSelected = selected.Contains("Fahasalamana: 8");
        }
    }
}