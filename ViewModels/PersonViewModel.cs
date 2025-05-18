using System;
using System.Collections.Generic;
using System.Windows;
using System.Windows.Input;
using gestion_concrets.Models;
using gestion_concrets.Services;

namespace gestion_concrets.ViewModels
{
    public class PersonViewModel : BaseViewModel
    {
        // Propriétés liées aux 8 modèles
        private readonly DatabaseService _databaseService;
        private BPerson _bPerson = new BPerson();
        public BPerson BPerson
        {
            get => _bPerson;
            set => SetProperty(ref _bPerson, value);
        }

        private Alocalisation _alocalisation = new Alocalisation();
        public Alocalisation Alocalisation
        {
            get => _alocalisation;
            set => SetProperty(ref _alocalisation, value);
        }

        private IIapplicationCDPH _iiApplicationCDPH = new IIapplicationCDPH();
        public IIapplicationCDPH IIapplicationCDPH
        {
            get => _iiApplicationCDPH;
            set => SetProperty(ref _iiApplicationCDPH, value);
        }

        private IIIright _iiiright = new IIIright();
        public IIIright IIIright
        {
            get => _iiiright;
            set => SetProperty(ref _iiiright, value);
        }

        private Itransmission _itransmission = new Itransmission();
        public Itransmission Itransmission
        {
            get => _itransmission;
            set => SetProperty(ref _itransmission, value);
        }

        private IVdutyGov _ivDutyGov = new IVdutyGov();
        public IVdutyGov IVdutyGov
        {
            get => _ivDutyGov;
            set => SetProperty(ref _ivDutyGov, value);
        }

        private VdevSupport _vdevSupport = new VdevSupport();
        public VdevSupport VdevSupport
        {
            get => _vdevSupport;
            set => SetProperty(ref _vdevSupport, value);
        }

        private VIpartnerCollab _viPartnerCollab = new VIpartnerCollab();
        public VIpartnerCollab VIpartnerCollab
        {
            get => _viPartnerCollab;
            set => SetProperty(ref _viPartnerCollab, value);
        }

        // Navigation properties
        private int _selectedTabIndex;
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

        public bool CanGoNext => SelectedTabIndex < 7; // 8 onglets (0 à 7)
        public bool CanGoPrevious => SelectedTabIndex > 0;

        // Commandes
        public ICommand SaveCommand { get; }
        public ICommand NextTabCommand { get; }
        public ICommand PreviousTabCommand { get; }

        public PersonViewModel()
        {
            _databaseService = new DatabaseService();
            SaveCommand = new RelayCommand(SaveData);
            NextTabCommand = new RelayCommand(NextTab, () => CanGoNext);
            PreviousTabCommand = new RelayCommand(PreviousTab, () => CanGoPrevious);
            SelectedTabIndex = 0; // Commencer au premier onglet
        }

        public List<string> SexeOptions { get; } = new List<string> { "Lahy", "Vavy" };

        private void SaveData()
        {
            try
            {
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
                MessageBox.Show("Données enregistrées avec succès !");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Erreur lors de l'enregistrement : " + ex.Message);
            }
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