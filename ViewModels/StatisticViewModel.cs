using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Collections.ObjectModel;
using System.Windows.Input;
using LiveChartsCore;
using LiveChartsCore.SkiaSharpView;
using gestion_concrets.Services;
using gestion_concrets.Models;
using DocumentFormat.OpenXml.Office2016.Drawing.ChartDrawing;
using System.Diagnostics;
using System.Windows.Controls;
using Wpf.Ui.Input;
using LiveChartsCore.Measure;
using LiveChartsCore.SkiaSharpView.Painting;
using SkiaSharp;
using System.Windows;
using System;
using System.Collections.Generic;


namespace gestion_concrets.ViewModels
{
    public class StatisticViewModel : BaseViewModel
    {
        private readonly DatabaseService _databaseService;
        private string _selectedRegion;
        private ObservableCollection<string> _regions;
        private ISeries[] _pieSeries;
        private ISeries[] _pieSeriesE101;
        private ISeries[] _pieSeriesE121;
        private ISeries[] _barSeries;
        private LiveChartsCore.SkiaSharpView.Axis[] _xAxes;
        private LiveChartsCore.SkiaSharpView.Axis[] _yAxes;


        public StatisticViewModel(DatabaseService databaseService = null)
        {
            _databaseService = databaseService ?? new DatabaseService();
            ExportCommand = new RelayCommand(ExportToWord);
            LoadRegions();
            UpdatePieChart();
            UpdatePieChartE101();
            UpdateBarChart();
            UpdatePieChartE121();
        }

        public ObservableCollection<string> Regions
        {
            get => _regions;
            set => SetProperty(ref _regions, value);
        }

        public string SelectedRegion
        {
            get => _selectedRegion;
            set
            {
                if (SetProperty(ref _selectedRegion, value))
                {
                    UpdatePieChart();
                    UpdateBarChart();
                    UpdatePieChartE101();
                    UpdatePieChartE121();
                }
            }
        }

        public ISeries[] PieSeries
        {
            get => _pieSeries;
            set => SetProperty(ref _pieSeries, value);
        }

        public ISeries[] PieSeriesE101
        {
            get => _pieSeriesE101;
            set => SetProperty(ref _pieSeriesE101, value);
        }

        public ISeries[] PieSeriesE121
        {
            get => _pieSeriesE121;
            set => SetProperty(ref _pieSeriesE121, value);
        }

        public ISeries[] BarSeries
        {
            get => _barSeries;
            set => SetProperty(ref _barSeries, value);
        }

        public LiveChartsCore.SkiaSharpView.Axis[] XAxes
        {
            get => _xAxes;
            set => SetProperty(ref _xAxes, value);
        }

        public LiveChartsCore.SkiaSharpView.Axis[] YAxes
        {
            get => _yAxes;
            set => SetProperty(ref _yAxes, value);
        }

        public ICommand ExportCommand { get; }

        private void LoadRegions()
        {
            var regions = _databaseService.GetUniqueRegions();
            Regions = new ObservableCollection<string>(regions);
            Regions.Insert(0, "Toutes");
            SelectedRegion = "Toutes"; // Sélection par défaut
        }

        private void UpdatePieChart()
        {
            var (enyCount, tsiaCount) = _databaseService.GetE1Counts(SelectedRegion == "Toutes" ? null : SelectedRegion);
            PieSeries = new ISeries[]
            {
                new PieSeries<double>
                {
                    Values = new double[] { enyCount },
                    Name = "Eny",
                    InnerRadius = 50,
                    DataLabelsPaint = new SolidColorPaint(SKColors.Black),
                    DataLabelsSize = 14,
                    DataLabelsPosition = PolarLabelsPosition.Outer
                },
                new PieSeries<double>
                {
                    Values = new double[] { tsiaCount },
                    Name = "Tsia",
                    InnerRadius = 50,
                    DataLabelsPaint = new SolidColorPaint(SKColors.Black),
                    DataLabelsSize = 14,
                    DataLabelsPosition = PolarLabelsPosition.Outer
                }
            };
        }

        private void UpdatePieChartE101()
        {
            var (enyCount, tsiaCount) = _databaseService.GetE101Counts(SelectedRegion == "Toutes" ? null : SelectedRegion);
            PieSeriesE101 = new ISeries[]
            {
                new PieSeries<double>
                {
                    Values = new double[] { enyCount },
                    Name = "Eny",
                    InnerRadius = 50,
                    DataLabelsPaint = new SolidColorPaint(SKColors.Black),
                    DataLabelsSize = 14,
                    DataLabelsPosition = PolarLabelsPosition.Outer
                },
                new PieSeries<double>
                {
                    Values = new double[] { tsiaCount },
                    Name = "Tsia",
                    InnerRadius = 50,
                    DataLabelsPaint = new SolidColorPaint(SKColors.Black),
                    DataLabelsSize = 14,
                    DataLabelsPosition = PolarLabelsPosition.Outer
                }
            };
        }

        private void UpdatePieChartE121()
        {
            var (enyCount, tsiaCount) = _databaseService.GetE121Counts(SelectedRegion == "Toutes" ? null : SelectedRegion);
            PieSeriesE121 = new ISeries[]
            {
                new PieSeries<double>
                {
                    Values = new double[] { enyCount },
                    Name = "Eny",
                    InnerRadius = 50,
                    DataLabelsPaint = new SolidColorPaint(SKColors.Black),
                    DataLabelsSize = 14,
                    DataLabelsPosition = PolarLabelsPosition.Outer
                },
                new PieSeries<double>
                {
                    Values = new double[] { tsiaCount },
                    Name = "Tsia",
                    InnerRadius = 50,
                    DataLabelsPaint = new SolidColorPaint(SKColors.Black),
                    DataLabelsSize = 14,
                    DataLabelsPosition = PolarLabelsPosition.Outer
                }
            };
        }

        private void UpdateBarChart()
        {
            var counts = _databaseService.GetE5Counts(SelectedRegion == "Toutes" ? null : SelectedRegion);
            var categories = new List<string>
            {
                "Fisitrahana rano: 1", 
                "Sakafo: 2", 
                "Fambolena: 3", 
                "Fiompiana: 4",
                "Jono: 5", 
                "Asa fivelomana hafa: 6", 
                "Trano fonenana: 7", 
                "Fahasalamana: 8"
            };
            var values = categories.Select(cat => (double)counts[cat]).ToArray();

            System.Diagnostics.Debug.WriteLine($"Values: {string.Join(", ", values)}");

            BarSeries = new ISeries[]
            {
                new ColumnSeries<double>
                {
                    Name = "Nombres",
                    Values = values,
                    DataLabelsPaint = new SolidColorPaint(SKColors.Black),
                    Fill = new SolidColorPaint(SKColors.Blue)
                }
            };

             XAxes = new LiveChartsCore.SkiaSharpView.Axis[]
             {
                new LiveChartsCore.SkiaSharpView.Axis
                {
                    Labels = categories.ToArray(),
                    LabelsRotation = 45,
                    TextSize = 12
                }
             };

             YAxes = new LiveChartsCore.SkiaSharpView.Axis[]
                        {
                new LiveChartsCore.SkiaSharpView.Axis
                {
                    Name = "Nombre de personnes",
                    MinLimit = 0,
                    MaxLimit = values.Any() ? values.Max() * 1.2 : 10,
                    TextSize = 12,
                    Labeler = value => value.ToString("N0")
                }
                        };
        }

        private void ExportToWord()
        {
            try
            {
                // Cette méthode sera appelée depuis le code-behind avec le chemin de l'image
                MessageBox.Show("Veuillez cliquer sur le bouton Exporter pour générer le document Word.",
                    "Information", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Erreur lors de l'exportation vers Word : {ex.Message}",
                    "Erreur", MessageBoxButton.OK, MessageBoxImage.Error);
                Debug.WriteLine($"[ERREUR EXPORT WORD] : {ex.Message}");
            }
        }
    }
}
