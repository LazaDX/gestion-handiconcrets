using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
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
using DocumentFormat.OpenXml.Drawing.Charts;
using gestion_concrets.Services;
using gestion_concrets.ViewModels;
using Xceed.Document.NET;
using Xceed.Words.NET;
using LiveChartsCore.SkiaSharpView.WPF;
using Microsoft.Win32;

namespace gestion_concrets.Views
{ 
    /// <summary>
    /// Logique d'interaction pour StatisticView.xaml
    /// </summary>
    public partial class StatisticView : Page
    {
        public StatisticView()
        {
            InitializeComponent();
            DataContext = new StatisticViewModel();
        }
        private void ExportToWord(object sender, RoutedEventArgs e)
        {
            var viewModel = DataContext as StatisticViewModel;
            if (viewModel == null) return;

            try
            {
                // Créer des fichiers temporaires pour les graphiques
                string pieChartFile = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "pie_chart.png");
                string pieChartFileE101 = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "pieE101_chart.png");
                string pieChartFileE121 = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "pieE121_chart.png");
                string barChartFile = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "bar_chart.png");

                // Capturer le PieChart
                var pieChart = (LiveChartsCore.SkiaSharpView.WPF.PieChart)FindName("PieChart");
                if (pieChart != null)
                {
                    var renderTargetBitmap = new RenderTargetBitmap(
                        (int)pieChart.ActualWidth, (int)pieChart.ActualHeight, 96, 96, PixelFormats.Pbgra32);
                    renderTargetBitmap.Render(pieChart);
                    var encoder = new PngBitmapEncoder();
                    encoder.Frames.Add(BitmapFrame.Create(renderTargetBitmap));
                    using (var fileStream = new FileStream(pieChartFile, FileMode.Create, FileAccess.Write))
                    {
                        encoder.Save(fileStream);
                    }
                }

                // Capturer le PieChartE101
                var pieChartE101 = (LiveChartsCore.SkiaSharpView.WPF.PieChart)FindName("PieChartE101");
                if (pieChartE101 != null)
                {
                    var renderTargetBitmap = new RenderTargetBitmap(
                        (int)pieChartE101.ActualWidth, (int)pieChartE101.ActualHeight, 96, 96, PixelFormats.Pbgra32);
                    renderTargetBitmap.Render(pieChartE101);
                    var encoder = new PngBitmapEncoder();
                    encoder.Frames.Add(BitmapFrame.Create(renderTargetBitmap));
                    using (var fileStream = new FileStream(pieChartFileE101, FileMode.Create, FileAccess.Write))
                    {
                        encoder.Save(fileStream);
                    }
                }

                var pieChartE121 = (LiveChartsCore.SkiaSharpView.WPF.PieChart)FindName("PieChartE121");
                if (pieChartE121 != null)
                {
                    var renderTargetBitmap = new RenderTargetBitmap(
                        (int)pieChartE121.ActualWidth, (int)pieChartE121.ActualHeight, 96, 96, PixelFormats.Pbgra32);
                    renderTargetBitmap.Render(pieChartE121);
                    var encoder = new PngBitmapEncoder();
                    encoder.Frames.Add(BitmapFrame.Create(renderTargetBitmap));
                    using (var fileStream = new FileStream(pieChartFileE121, FileMode.Create, FileAccess.Write))
                    {
                        encoder.Save(fileStream);
                    }
                }

                // Capturer le CartesianChart
                var barChart = (CartesianChart)FindName("BarChart");
                if (barChart != null)
                {
                    var renderTargetBitmap = new RenderTargetBitmap(
                        (int)barChart.ActualWidth, (int)barChart.ActualHeight, 96, 96, PixelFormats.Pbgra32);
                    renderTargetBitmap.Render(barChart);
                    var encoder = new PngBitmapEncoder();
                    encoder.Frames.Add(BitmapFrame.Create(renderTargetBitmap));
                    using (var fileStream = new FileStream(barChartFile, FileMode.Create, FileAccess.Write))
                    {
                        encoder.Save(fileStream);
                    }
                }

                // Afficher SaveFileDialog pour choisir le chemin
                var saveFileDialog = new SaveFileDialog
                {
                    Filter = "Word Documents (*.docx)|*.docx",
                    DefaultExt = ".docx",
                    FileName = "StatistiqueExport.docx"
                };

                if (saveFileDialog.ShowDialog() == true)
                {
                    // Créer un document Word avec DocX
                    using (var document = DocX.Create(saveFileDialog.FileName))
                    {
                        // Ajouter le titre et la région
                        document.InsertParagraph("Statistiques sur l'impact du changement climatique")
                               .FontSize(20)
                               .Bold()
                               .Alignment = Alignment.center;
                        document.InsertParagraph($"Région sélectionnée : {viewModel.SelectedRegion ?? "Toutes"}")
                               .FontSize(12);

                        // Ajouter le PieChart (E1)
                        if (File.Exists(pieChartFile))
                        {
                            document.InsertParagraph("Ny isan'ny olona mahatsapa ny fiovan'ny toetr'andro")
                                   .FontSize(14)
                                   .Bold();
                            var pieImage = document.AddImage(pieChartFile);
                            var piePicture = pieImage.CreatePicture(300, 300);
                            document.InsertParagraph().AppendPicture(piePicture);
                        }


                        // Ajouter le PieChart (E101)
                        if (File.Exists(pieChartFileE101))
                        {
                            document.InsertParagraph("Ny isan'ireo olona mahazo vaovao amin'ny fiovan'ny toetr'andro")
                                   .FontSize(14)
                                   .Bold();
                            var pieImage = document.AddImage(pieChartFileE101);
                            var piePicture = pieImage.CreatePicture(300, 300);
                            document.InsertParagraph().AppendPicture(piePicture);
                        }

                        // Ajouter le PieChart (E121)
                        if (File.Exists(pieChartFileE121))
                        {
                            document.InsertParagraph("Ny isan'ireo olona mahatsapa fa nahazo fanampiana avy amin'ny fanjakana vokatrin'ny fiovan'ny toetr'andro")
                                   .FontSize(14)
                                   .Bold();
                            var pieImage = document.AddImage(pieChartFileE121);
                            var piePicture = pieImage.CreatePicture(300, 300);
                            document.InsertParagraph().AppendPicture(piePicture);
                        }

                        // Ajouter le CartesianChart (E5)
                        if (File.Exists(barChartFile))
                        {
                            document.InsertParagraph("Fiantrakan'ny fiovan'ny toetr'andro")
                                   .FontSize(14)
                                   .Bold();
                            var barImage = document.AddImage(barChartFile);
                            var barPicture = barImage.CreatePicture(300, 300);
                            document.InsertParagraph().AppendPicture(barPicture);
                        }

                        //// Ajouter les comptes Eny/Tsia
                        //document.InsertParagraph($"Eny: {viewModel.EnyCount}")
                        //       .FontSize(12);
                        //document.InsertParagraph($"Tsia: {viewModel.TsiaCount}")
                        //       .FontSize(12);

                        document.Save();
                    }

                    MessageBox.Show("Document exporté avec succès !", "Succès", MessageBoxButton.OK, MessageBoxImage.Information);
                }

                // Nettoyer les fichiers temporaires
                if (File.Exists(pieChartFile)) File.Delete(pieChartFile);
                if (File.Exists(pieChartFileE101)) File.Delete(pieChartFileE101);
                if (File.Exists(pieChartFileE121)) File.Delete(pieChartFileE121);
                if (File.Exists(barChartFile)) File.Delete(barChartFile);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Erreur lors de l'exportation : {ex.Message}", "Erreur", MessageBoxButton.OK, MessageBoxImage.Error);
                System.Diagnostics.Debug.WriteLine($"[ERREUR EXPORT WORD] : {ex.Message}");
            }
        }
        

        private void ExportChart_Click(object sender, RoutedEventArgs e)
        {
            
        }
    }
    
    
}
