using System;
using System.IO;
using System.Windows;
using System.Windows.Input;
using Microsoft.Win32;
using System.Data.SQLite;

namespace gestion_concrets.ViewModels
{
    public class DatabaseViewModel : BaseViewModel
    {
        //private static readonly string projectRoot = Path.GetFullPath(
        //    Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "..", "..", "..")
        //);
        //private static readonly string dataFolder = Path.Combine(projectRoot, "Data");
        //private static readonly string dbPath = Path.Combine(dataFolder, "database.db");
        //private static readonly string connectionString = $"Data Source={dbPath};Version=3;";

        private static readonly string executableDirectory = AppContext.BaseDirectory;
        private static readonly string dataFolder = Path.Combine(executableDirectory, "Data");
        private static readonly string dbPath = Path.Combine(dataFolder, "database.db");
        private static readonly string connectionString = $"Data Source={dbPath};Version=3;";


        public ICommand CreateBackupCommand { get; }
        public ICommand ExportDatabaseCommand { get; }
        public ICommand ImportDatabaseCommand { get; }

        public DatabaseViewModel()
        {
            if (!Directory.Exists(dataFolder))
            {
                Directory.CreateDirectory(dataFolder);
            }

            CreateBackupCommand = new RelayCommand(CreateBackup);
            ExportDatabaseCommand = new RelayCommand(ExportDatabase);
            ImportDatabaseCommand = new RelayCommand(ImportDatabase);
        }

        private void CreateBackup()
        {
            try
            {
                var result = MessageBox.Show($"Voulez-vous créer une sauvegarde (backup) de votre base de données ?", "Confirmation", MessageBoxButton.YesNo, MessageBoxImage.Question);
                if (result == MessageBoxResult.Yes)
                {
                    if (!File.Exists(dbPath))
                    {
                        MessageBox.Show("La base de données actuelle n’existe pas.", "Erreur", MessageBoxButton.OK, MessageBoxImage.Error);
                        return;
                    }

                    if (!IsValidSqliteDatabase(dbPath))
                    {
                        MessageBox.Show("La base de données actuelle est corrompue ou n’a pas la structure attendue. Impossible de créer un backup.", "Erreur", MessageBoxButton.OK, MessageBoxImage.Error);
                        return;
                    }

                    string backupPath = Path.Combine(dataFolder, $"database_backup_{DateTime.Now:yyyyMMdd_HHmmss}.db");
                    File.Copy(dbPath, backupPath, false);
                    MessageBox.Show($"Backup créé avec succès : {Path.GetFileName(backupPath)}", "Succès", MessageBoxButton.OK, MessageBoxImage.Information);

                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Erreur lors de la création du backup : {ex.Message}", "Erreur", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void ExportDatabase()
        {
            var saveFileDialog = new SaveFileDialog
            {
                Filter = "SQLite Database (*.db)|*.db",
                Title = "Exporter la base de données",
                FileName = "database.db",
                OverwritePrompt = true
            };

            if (saveFileDialog.ShowDialog() == true)
            {
                try
                {
                    if (!File.Exists(dbPath))
                    {
                        MessageBox.Show("La base de données actuelle n’existe pas.", "Erreur", MessageBoxButton.OK, MessageBoxImage.Error);
                        return;
                    }

                    if (!IsValidSqliteDatabase(dbPath))
                    {
                        MessageBox.Show("La base de données actuelle est corrompue ou n’a pas la structure attendue. Impossible d’exporter.", "Erreur", MessageBoxButton.OK, MessageBoxImage.Error);
                        return;
                    }

                    File.Copy(dbPath, saveFileDialog.FileName, true);
                    MessageBox.Show("Base de données exportée avec succès !", "Succès", MessageBoxButton.OK, MessageBoxImage.Information);
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Erreur lors de l'export : {ex.Message}", "Erreur", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        }

        private void ImportDatabase()
        {
            var openFileDialog = new OpenFileDialog
            {
                Filter = "SQLite Database (*.db)|*.db",
                Title = "Sélectionner une base de données à importer"
            };

            if (openFileDialog.ShowDialog() == true)
            {
                string newDbPath = openFileDialog.FileName;
                try
                {
                    if (!IsValidSqliteDatabase(newDbPath))
                    {
                        MessageBox.Show($"Le fichier '{Path.GetFileName(newDbPath)}' est corrompu ou n’a pas la structure attendue.", "Erreur", MessageBoxButton.OK, MessageBoxImage.Error);
                        return;
                    }

                    var result = MessageBox.Show($"Voulez-vous remplacer la base de données actuelle par '{Path.GetFileName(newDbPath)}' ? Un backup sera créé automatiquement.", "Confirmation", MessageBoxButton.YesNo, MessageBoxImage.Question);
                    if (result == MessageBoxResult.Yes)
                    {
                        // Créer un backup automatique avant remplacement
                        if (File.Exists(dbPath))
                        {
                            string backupPath = Path.Combine(dataFolder, $"database_backup_{DateTime.Now:yyyyMMdd_HHmmss}.db");
                            try
                            {
                                File.Copy(dbPath, backupPath, false);
                                MessageBox.Show($"Backup créé avec succès : {Path.GetFileName(backupPath)}", "Information", MessageBoxButton.OK, MessageBoxImage.Information);
                            }
                            catch (Exception ex)
                            {
                                MessageBox.Show($"Erreur lors de la création du backup : {ex.Message}. L'importation va continuer.", "Avertissement", MessageBoxButton.OK, MessageBoxImage.Warning);
                            }
                        }

                        // Remplacer directement Data/database.db
                        File.Copy(newDbPath, dbPath, true);
                        MessageBox.Show("Base de données importée et remplacée avec succès !", "Succès", MessageBoxButton.OK, MessageBoxImage.Information);
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Erreur lors de l’import : {ex.Message}", "Erreur", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        }

        private bool IsValidSqliteDatabase(string path)
        {
            try
            {
                using (var connection = new SQLiteConnection($"Data Source={path};Version=3;"))
                {
                    connection.Open();
                    using (var command = new SQLiteCommand("SELECT name FROM sqlite_master WHERE type='table' AND name='BPerson';", connection))
                    {
                        var result = command.ExecuteScalar();
                        if (result == null || result.ToString() != "BPerson")
                        {
                            Console.WriteLine($"Erreur : La table 'BPerson' n’existe pas dans {path}");
                            return false;
                        }
                    }
                    connection.Close();
                    return true;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Erreur lors de la validation de {path} : {ex.Message}");
                return false;
            }
        }
    }
}