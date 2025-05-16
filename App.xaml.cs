using System;
using System.Windows;
using System.Diagnostics;
using gestion_concrets.Services;

namespace gestion_concrets
{
    public partial class App : Application
    {
        public App()
        {
            Debug.WriteLine("[APP] Constructeur App exécuté.");
            this.DispatcherUnhandledException += (s, e) =>
            {
                Debug.WriteLine($"[ERREUR NON GÉRÉE] : {e.Exception.Message}");
                MessageBox.Show($"Exception non gérée : {e.Exception.Message}", "Erreur", MessageBoxButton.OK, MessageBoxImage.Error);
                e.Handled = true;
            };
        }

        protected override void OnStartup(StartupEventArgs e)
        {
            Debug.WriteLine("[APP] OnStartup exécuté - Début");
            base.OnStartup(e);

            try
            {
                DatabaseService _databaseService = new DatabaseService();
                Debug.WriteLine("[APP] Appel de DatabaseService.InitializeTableDatabase()...");
                _databaseService.InitializeTableDatabase();
                Debug.WriteLine("[APP] OnStartup exécuté - Fin");
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[APP][ERREUR] : {ex.Message}");
                MessageBox.Show($"Erreur au démarrage : {ex.Message}", "Erreur", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
    }
}
