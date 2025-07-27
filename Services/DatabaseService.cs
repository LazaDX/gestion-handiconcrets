using System;
using System.Data.SQLite;
using System.IO;
using System.Diagnostics;
using System.Collections.Generic;
using gestion_concrets.Models;
using ClosedXML.Excel;

namespace gestion_concrets.Services
{
    public class DatabaseService
    {

        //private static readonly string projectRoot = Path.GetFullPath(
        //    Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "..", "..", "..","..")
        //);
        //private static readonly string dataFolder = Path.Combine(projectRoot, "Data");
        //private static readonly string dbPath = Path.Combine(dataFolder, "database.db");
        //private static readonly string connectionString = $"Data Source={dbPath};";


        private static readonly string projectRoot = GetProjectRoot();
        private static readonly string dataFolder = Path.Combine(projectRoot, "Data");
        private static readonly string dbPath = Path.Combine(dataFolder, "database.db");
        private static readonly string connectionString = $"Data Source={dbPath};";

        private static string GetProjectRoot()
        {
            string baseDir = AppDomain.CurrentDomain.BaseDirectory;
           

            // Vérifie si on est dans un environnement de développement (bin\Debug ou bin\Release)
            if (baseDir.Contains(@"\bin\Debug") || baseDir.Contains(@"\bin\Release"))
            {
                // Remonte jusqu'à la racine du projet
                return Path.GetFullPath(
                        Path.Combine(baseDir, "..", "..", "..","..")
                    );
            }
            
            return baseDir;
          
        }

        public static SQLiteConnection GetConnection()
        {
            Debug.WriteLine($"[INFO] Project root : {projectRoot}");
            Debug.WriteLine($"[INFO] Data folder  : {dataFolder}");

            try
            {
                if (!Directory.Exists(dataFolder))
                {
                    Debug.WriteLine($"[INFO] Création du dossier racine Data : {dataFolder}");
                    Directory.CreateDirectory(dataFolder);
                }
                else
                {
                    Debug.WriteLine($"[INFO] Dossier racine Data existe déjà : {dataFolder}");
                }

                if (!File.Exists(dbPath))
                {
                    Debug.WriteLine($"[INFO] Création du fichier database.db : {dbPath}");
                    SQLiteConnection.CreateFile(dbPath);
                }
                else
                {
                    Debug.WriteLine($"[INFO] Fichier database.db existe déjà : {dbPath}");
                }
                return new SQLiteConnection(connectionString);
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[ERREUR] Impossible d'initialiser la connexion SQLite : {ex.Message}");
                throw;
            }
        }
    
        public void InitializeTableDatabase()
        {
            Debug.WriteLine("[INFO] Initialisation des tables SQLite...");

            try
            {
                using var con = GetConnection();
                con.Open();
                Debug.WriteLine("[INFO] Connexion SQLite ouverte.");

                using (var pragmaCmd = new SQLiteCommand("PRAGMA foreign_keys = ON;", con))
                {
                    pragmaCmd.ExecuteNonQuery();
                    Debug.WriteLine("[INFO] PRAGMA foreign_keys activé.");
                }

                const string sql = @"
                    CREATE TABLE IF NOT EXISTS BPerson (
                        Id INTEGER PRIMARY KEY AUTOINCREMENT,
                        B1 TEXT NOT NULL,
                        B2 TEXT NOT NULL,
                        B3 TEXT NOT NULL,
                        B4 TEXT NOT NULL,
                        Adress TEXT NOT NULL,
                        Phone TEXT NOT NULL,
                        Email TEXT NOT NULL,
                        B5 TEXT NOT NULL,
                        B6 TEXT NOT NULL,
                        B61 TEXT NOT NULL
                    );
                    CREATE TABLE IF NOT EXISTS Alocalisation (
                        Id INTEGER PRIMARY KEY AUTOINCREMENT,
                        IdBPerson INTEGER NOT NULL,
                        A1 TEXT NOT NULL,
                        A2 TEXT NOT NULL,
                        A3 TEXT NOT NULL,
                        A4 TEXT NOT NULL,
                        A5 TEXT NOT NULL,
                        A6 TEXT NOT NULL,
                        A7 TEXT NOT NULL,
                        A8 TEXT NOT NULL,
                        FOREIGN KEY(IdBPerson) REFERENCES BPerson(Id) ON DELETE CASCADE
                    );
                    CREATE TABLE IF NOT EXISTS Ddescription (
                        Id INTEGER PRIMARY KEY AUTOINCREMENT,
                        IdBPerson INTEGER NOT NULL,
                        D1 TEXT NOT NULL,
                        D11 TEXT NOT NULL,
                        D2 TEXT NOT NULL,
                        D21 TEXT NOT NULL,
                        D22 TEXT NOT NULL,
                        FOREIGN KEY(IdBPerson) REFERENCES BPerson(Id) ON DELETE CASCADE
                    );
                    CREATE TABLE IF NOT EXISTS Eclimat (
                        Id INTEGER PRIMARY KEY AUTOINCREMENT,
                        IdBPerson INTEGER NOT NULL,
                        E1 TEXT NOT NULL,
                        E2 TEXT NOT NULL,
                        E3 TEXT NOT NULL,
                        E4 TEXT NOT NULL,
                        E5 TEXT NOT NULL,
                        E611 TEXT NOT NULL,
                        E612 TEXT NOT NULL,
                        E621 TEXT NOT NULL,
                        E622 TEXT NOT NULL,
                        E631 TEXT NOT NULL,
                        E632 TEXT NOT NULL,
                        E641 TEXT NOT NULL,
                        E642 TEXT NOT NULL,
                        E651 TEXT NOT NULL,
                        E652 TEXT NOT NULL,
                        E661 TEXT NOT NULL,
                        E662 TEXT NOT NULL,
                        E71 TEXT NOT NULL,
                        E72 TEXT NOT NULL,
                        E81 TEXT NOT NULL,
                        E82 TEXT NOT NULL,
                        E911 TEXT NOT NULL,
                        E912 TEXT NOT NULL,
                        E921 TEXT NOT NULL,
                        E922 TEXT NOT NULL,
                        E931 TEXT NOT NULL,
                        E932 TEXT NOT NULL,
                        E941 TEXT NOT NULL,
                        E942 TEXT NOT NULL,
                        E951 TEXT NOT NULL,
                        E952 TEXT NOT NULL,
                        E961 TEXT NOT NULL,
                        E962 TEXT NOT NULL,
                        E971 TEXT NOT NULL,
                        E972 TEXT NOT NULL,
                        E981 TEXT NOT NULL,
                        E982 TEXT NOT NULL,
                        E101 TEXT NOT NULL,
                        E102 TEXT NOT NULL,
                        E103 TEXT NOT NULL,
                        E104 TEXT NOT NULL,
                        E111 TEXT NOT NULL,
                        E112 TEXT NOT NULL,
                        E121 TEXT NOT NULL,
                        E122 TEXT NOT NULL,
                        E131 TEXT NOT NULL,
                        E132 TEXT NOT NULL,
                        E141 TEXT NOT NULL,
                        E142 TEXT NOT NULL,
                        E15 TEXT NOT NULL,
                        FOREIGN KEY(IdBPerson) REFERENCES BPerson(Id) ON DELETE CASCADE
                    );
                    ";

                using var cmd = new SQLiteCommand(sql, con);
                cmd.ExecuteNonQuery();
                Debug.WriteLine("[SUCCÈS] Tables créée ou déjà existante avec contraintes FK activés.");
            }
            catch (Exception ex) when (ex is SQLiteException || ex is IOException)
            {
                Debug.WriteLine($"[ERREUR] {ex.GetType().Name} : {ex.Message}");
                throw;
            }
        }

        // ADD all information of person
        public void AddFullPerson(BPerson person, Alocalisation localisation, Ddescription description, Eclimat climat)
        {
            using (var con = GetConnection())
            {
                con.Open();
                using (var transaction = con.BeginTransaction())
                {
                    try
                    {
                        var personQuery = @"
                            INSERT INTO BPerson (B1, B2, B3, B4, Adress, Phone, Email, B5, B6, B61)
                            VALUES (@B1, @B2, @B3, @B4, @Adress, @Phone, @Email, @B5, @B6, @B61); 
                            SELECT last_insert_rowid()";

                        long idPerson;

                        using (var cmd = new SQLiteCommand(personQuery, con, transaction))
                        {
                            cmd.Parameters.AddWithValue("@B1", person.B1);
                            cmd.Parameters.AddWithValue("@B2", person.B2);
                            cmd.Parameters.AddWithValue("@B3", person.B3);
                            cmd.Parameters.AddWithValue("@B4", person.B4);
                            cmd.Parameters.AddWithValue("@Adress", person.Adress);
                            cmd.Parameters.AddWithValue("@Phone", person.Phone);
                            cmd.Parameters.AddWithValue("@Email", person.Email);
                            cmd.Parameters.AddWithValue("@B5", person.B5);
                            cmd.Parameters.AddWithValue("@B6", person.B6);
                            cmd.Parameters.AddWithValue("@B61", person.B61);

                            idPerson = (long)cmd.ExecuteScalar();
                        }

                        var localisationQuery = @"
                            INSERT INTO Alocalisation (IdBPerson, A1, A2, A3, A4, A5, A6, A7, A8)
                            VALUES (@IdBPerson, @A1, @A2, @A3, @A4, @A5, @A6, @A7, @A8);
                        ";
                        using (var cmd = new SQLiteCommand(localisationQuery, con, transaction))
                        {
                            cmd.Parameters.AddWithValue("@IdBPerson", idPerson);
                            cmd.Parameters.AddWithValue("@A1", localisation.A1);
                            cmd.Parameters.AddWithValue("@A2", localisation.A2);
                            cmd.Parameters.AddWithValue("@A3", localisation.A3);
                            cmd.Parameters.AddWithValue("@A4", localisation.A4);
                            cmd.Parameters.AddWithValue("@A5", localisation.A5);
                            cmd.Parameters.AddWithValue("@A6", localisation.A6);
                            cmd.Parameters.AddWithValue("@A7", localisation.A7);
                            cmd.Parameters.AddWithValue("@A8", localisation.A8);

                            cmd.ExecuteNonQuery();
                        }

                        var descriptionQuery = @"
                            INSERT INTO Ddescription (IdBPerson, D1, D11, D2, D21, D22)
                            VALUES (@IdBPerson, @D1, @D11, @D2, @D21, @D22);
                        ";
                        using (var cmd = new SQLiteCommand(descriptionQuery, con, transaction))
                        {
                            cmd.Parameters.AddWithValue("@IdBPerson", idPerson);
                            cmd.Parameters.AddWithValue("@D1", description.D1);
                            cmd.Parameters.AddWithValue("@D11", description.D11);
                            cmd.Parameters.AddWithValue("@D2", description.D2);
                            cmd.Parameters.AddWithValue("@D21", description.D21);
                            cmd.Parameters.AddWithValue("@D22", description.D22);

                            cmd.ExecuteNonQuery();
                        }

                        var climatQuery = @"
                            INSERT INTO Eclimat (IdBPerson, E1, E2, E3, E4, E5, E611, E612, E621, E622, 
                                                 E631, E632, E641, E642, E651, E652, E661, E662, E71, E72, 
                                                 E81, E82, E911, E912, E921, E922, E931, E932, E941, E942, 
                                                 E951, E952, E961, E962, E971, E972, E981, E982, E101, E102, 
                                                 E103, E104, E111, E112, E121, E122, E131, E132, E141, E142, E15)
                            VALUES (@IdBPerson, @E1, @E2, @E3, @E4, @E5, @E611, @E612, @E621, @E622, 
                                    @E631, @E632, @E641, @E642, @E651, @E652, @E661, @E662, @E71, @E72, 
                                    @E81, @E82, @E911, @E912, @E921, @E922, @E931, @E932, @E941, @E942, 
                                    @E951, @E952, @E961, @E962, @E971, @E972, @E981, @E982, @E101, @E102, 
                                    @E103, @E104, @E111, @E112, @E121, @E122, @E131, @E132, @E141, @E142, @E15);
                        ";
                        using (var cmd = new SQLiteCommand(climatQuery, con, transaction))
                        {
                            cmd.Parameters.AddWithValue("@IdBPerson", idPerson);
                            cmd.Parameters.AddWithValue("@E1", climat.E1);
                            cmd.Parameters.AddWithValue("@E2", climat.E2);
                            cmd.Parameters.AddWithValue("@E3", climat.E3);
                            cmd.Parameters.AddWithValue("@E4", climat.E4);
                            cmd.Parameters.AddWithValue("@E5", climat.E5);
                            cmd.Parameters.AddWithValue("@E611", climat.E611);
                            cmd.Parameters.AddWithValue("@E612", climat.E612);
                            cmd.Parameters.AddWithValue("@E621", climat.E621);
                            cmd.Parameters.AddWithValue("@E622", climat.E622);
                            cmd.Parameters.AddWithValue("@E631", climat.E631);
                            cmd.Parameters.AddWithValue("@E632", climat.E632);
                            cmd.Parameters.AddWithValue("@E641", climat.E641);
                            cmd.Parameters.AddWithValue("@E642", climat.E642);
                            cmd.Parameters.AddWithValue("@E651", climat.E651);
                            cmd.Parameters.AddWithValue("@E652", climat.E652);
                            cmd.Parameters.AddWithValue("@E661", climat.E661);
                            cmd.Parameters.AddWithValue("@E662", climat.E662);
                            cmd.Parameters.AddWithValue("@E71", climat.E71);
                            cmd.Parameters.AddWithValue("@E72", climat.E72);
                            cmd.Parameters.AddWithValue("@E81", climat.E81);
                            cmd.Parameters.AddWithValue("@E82", climat.E82);
                            cmd.Parameters.AddWithValue("@E911", climat.E911);
                            cmd.Parameters.AddWithValue("@E912", climat.E912);
                            cmd.Parameters.AddWithValue("@E921", climat.E921);
                            cmd.Parameters.AddWithValue("@E922", climat.E922);
                            cmd.Parameters.AddWithValue("@E931", climat.E931);
                            cmd.Parameters.AddWithValue("@E932", climat.E932);
                            cmd.Parameters.AddWithValue("@E941", climat.E941);
                            cmd.Parameters.AddWithValue("@E942", climat.E942);
                            cmd.Parameters.AddWithValue("@E951", climat.E951);
                            cmd.Parameters.AddWithValue("@E952", climat.E952);
                            cmd.Parameters.AddWithValue("@E961", climat.E961);
                            cmd.Parameters.AddWithValue("@E962", climat.E962);
                            cmd.Parameters.AddWithValue("@E971", climat.E971);
                            cmd.Parameters.AddWithValue("@E972", climat.E972);
                            cmd.Parameters.AddWithValue("@E981", climat.E981);
                            cmd.Parameters.AddWithValue("@E982", climat.E982);
                            cmd.Parameters.AddWithValue("@E101", climat.E101);
                            cmd.Parameters.AddWithValue("@E102", climat.E102);
                            cmd.Parameters.AddWithValue("@E103", climat.E103);
                            cmd.Parameters.AddWithValue("@E104", climat.E104);
                            cmd.Parameters.AddWithValue("@E111", climat.E111);
                            cmd.Parameters.AddWithValue("@E112", climat.E112);
                            cmd.Parameters.AddWithValue("@E121", climat.E121);
                            cmd.Parameters.AddWithValue("@E122", climat.E122);
                            cmd.Parameters.AddWithValue("@E131", climat.E131);
                            cmd.Parameters.AddWithValue("@E132", climat.E132);
                            cmd.Parameters.AddWithValue("@E141", climat.E141);
                            cmd.Parameters.AddWithValue("@E142", climat.E142);
                            cmd.Parameters.AddWithValue("@E15", climat.E15);

                            cmd.ExecuteNonQuery();
                        }

                        transaction.Commit();
                    }
                    catch (Exception e)
                    {
                        transaction.Rollback();
                        throw new Exception("Erreur lors de l'ajout de la personne et de ses données associées : " + e.Message);
                    }
                }
            }
        }

        // GET All Persons
        public List<BPerson> GetAllPersons()
        {
            var persons = new List<BPerson>();
            using (var con = GetConnection())
            {
                con.Open();
                using (var cmd = new SQLiteCommand(@"
                    SELECT BPerson.Id, BPerson.B1, BPerson.B2, BPerson.B3, BPerson.B4, BPerson.Adress, 
                           BPerson.Phone, BPerson.Email, BPerson.B5, BPerson.B6, BPerson.B61, Alocalisation.A1, Alocalisation.A2
                    FROM BPerson
                    LEFT JOIN Alocalisation ON BPerson.Id = Alocalisation.IdBPerson", con))
                using (var reader = cmd.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        persons.Add(new BPerson
                        {
                            Id = reader.GetInt32(0),
                            B1 = reader.GetString(1),
                            B2 = reader.GetString(2),
                            B3 = reader.GetString(3),
                            B4 = reader.GetString(4),
                            Adress = reader.GetString(5),
                            Phone = reader.GetString(6),
                            Email = reader.GetString(7),
                            B5 = reader.GetString(8),
                            B6 = reader.GetString(9),
                            B61 = reader.GetString(10),
                            A1 = reader.IsDBNull(11) ? null : reader.GetString(11),
                            A2 = reader.GetString(12)
                        });
                    }
                }
            }
            return persons;
        }

        public BPerson GetBPersonById(int id)
        {
            using (var con = GetConnection())
            {
                con.Open();
                try
                {
                    var query = @"
                        SELECT Id, B1, B2, B3, B4, Adress, Phone, Email, B5, B6, B61
                        FROM BPerson
                        WHERE Id = @Id";

                    using (var cmd = new SQLiteCommand(query, con))
                    {
                        cmd.Parameters.AddWithValue("@Id", id);
                        using (var reader = cmd.ExecuteReader())
                        {
                            if (reader.Read())
                            {
                                return new BPerson
                                {
                                    Id = reader.GetInt32(0),
                                    B1 = reader.GetString(1),
                                    B2 = reader.GetString(2),
                                    B3 = reader.GetString(3),
                                    B4 = reader.GetString(4),
                                    Adress = reader.GetString(5),
                                    Phone = reader.GetString(6),
                                    Email = reader.GetString(7),
                                    B5 = reader.GetString(8),
                                    B6 = reader.GetString(9),
                                    B61 = reader.GetString(10)
                                };
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    throw new Exception("Erreur lors de la récupération de la personne : " + ex.Message);
                }
                finally
                {
                    con.Close();
                }
            }
            return null;
        }

        public Alocalisation GetAlocalisationById(int idBPerson)
        {
            using (var con = GetConnection())
            {
                con.Open();
                try
                {
                    var query = @"
                        SELECT Id, IdBPerson, A1, A2, A3, A4, A5, A6, A7, A8
                        FROM Alocalisation
                        WHERE IdBPerson = @IdBPerson";

                    using (var cmd = new SQLiteCommand(query, con))
                    {
                        cmd.Parameters.AddWithValue("@IdBPerson", idBPerson);
                        using (var reader = cmd.ExecuteReader())
                        {
                            if (reader.Read())
                            {
                                return new Alocalisation
                                {
                                    Id = reader.GetInt32(0),
                                    IdBPerson = reader.GetInt32(1),
                                    A1 = reader.GetString(2),
                                    A2 = reader.GetString(3),
                                    A3 = reader.GetString(4),
                                    A4 = reader.GetString(5),
                                    A5 = reader.GetString(6),
                                    A6 = reader.GetString(7),
                                    A7 = reader.GetString(8),
                                    A8 = reader.GetString(9)
                                };
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    throw new Exception("Erreur lors de la récupération de la localisation : " + ex.Message);
                }
                finally
                {
                    con.Close();
                }
            }
            return null;
        }

        public Ddescription GetDdescriptionById(int idBPerson)
        {
            using (var con = GetConnection())
            {
                con.Open();
                try
                {
                    var query = @"
                        SELECT Id, IdBPerson, D1, D11, D2, D21, D22
                        FROM Ddescription
                        WHERE IdBPerson = @IdBPerson";

                    using (var cmd = new SQLiteCommand(query, con))
                    {
                        cmd.Parameters.AddWithValue("@IdBPerson", idBPerson);
                        using (var reader = cmd.ExecuteReader())
                        {
                            if (reader.Read())
                            {
                                return new Ddescription
                                {
                                    Id = reader.GetInt32(0),
                                    IdBPerson = reader.GetInt32(1),
                                    D1 = reader.GetString(2),
                                    D11 = reader.GetString(3),
                                    D2 = reader.GetString(4),
                                    D21 = reader.GetString(5),
                                    D22 = reader.GetString(6)
                                };
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    throw new Exception("Erreur lors de la récupération de la description : " + ex.Message);
                }
                finally
                {
                    con.Close();
                }
            }
            return null;
        }

        public Eclimat GetEclimatById(int idBPerson)
        {
            using (var con = GetConnection())
            {
                con.Open();
                try
                {
                    var query = @"
                        SELECT Id, IdBPerson, E1, E2, E3, E4, E5, E611, E612, E621, E622, E631, E632, 
                               E641, E642, E651, E652, E661, E662, E71, E72, E81, E82, E911, E912, E921, 
                               E922, E931, E932, E941, E942, E951, E952, E961, E962, E971, E972, E981, 
                               E982, E101, E102, E103, E104, E111, E112, E121, E122, E131, E132, E141, 
                               E142, E15
                        FROM Eclimat
                        WHERE IdBPerson = @IdBPerson";

                    using (var cmd = new SQLiteCommand(query, con))
                    {
                        cmd.Parameters.AddWithValue("@IdBPerson", idBPerson);
                        using (var reader = cmd.ExecuteReader())
                        {
                            if (reader.Read())
                            {
                                return new Eclimat
                                {
                                    Id = reader.GetInt32(0),
                                    IdBPerson = reader.GetInt32(1),
                                    E1 = reader.GetString(2),
                                    E2 = reader.GetString(3),
                                    E3 = reader.GetString(4),
                                    E4 = reader.GetString(5),
                                    E5 = reader.GetString(6),
                                    E611 = reader.GetString(7),
                                    E612 = reader.GetString(8),
                                    E621 = reader.GetString(9),
                                    E622 = reader.GetString(10),
                                    E631 = reader.GetString(11),
                                    E632 = reader.GetString(12),
                                    E641 = reader.GetString(13),
                                    E642 = reader.GetString(14),
                                    E651 = reader.GetString(15),
                                    E652 = reader.GetString(16),
                                    E661 = reader.GetString(17),
                                    E662 = reader.GetString(18),
                                    E71 = reader.GetString(19),
                                    E72 = reader.GetString(20),
                                    E81 = reader.GetString(21),
                                    E82 = reader.GetString(22),
                                    E911 = reader.GetString(23),
                                    E912 = reader.GetString(24),
                                    E921 = reader.GetString(25),
                                    E922 = reader.GetString(26),
                                    E931 = reader.GetString(27),
                                    E932 = reader.GetString(28),
                                    E941 = reader.GetString(29),
                                    E942 = reader.GetString(30),
                                    E951 = reader.GetString(31),
                                    E952 = reader.GetString(32),
                                    E961 = reader.GetString(33),
                                    E962 = reader.GetString(34),
                                    E971 = reader.GetString(35),
                                    E972 = reader.GetString(36),
                                    E981 = reader.GetString(37),
                                    E982 = reader.GetString(38),
                                    E101 = reader.GetString(39),
                                    E102 = reader.GetString(40),
                                    E103 = reader.GetString(41),
                                    E104 = reader.GetString(42),
                                    E111 = reader.GetString(43),
                                    E112 = reader.GetString(44),
                                    E121 = reader.GetString(45),
                                    E122 = reader.GetString(46),
                                    E131 = reader.GetString(47),
                                    E132 = reader.GetString(48),
                                    E141 = reader.GetString(49),
                                    E142 = reader.GetString(50),
                                    E15 = reader.GetString(51)
                                };
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    throw new Exception("Erreur lors de la récupération du climat : " + ex.Message);
                }
                finally
                {
                    con.Close();
                }
            }
            return null;
        }

        // UPDATE All 
        public void UpdateFullPerson(BPerson person, Alocalisation localisation, Ddescription description, Eclimat climat)
        {
            using (var con = GetConnection())
            {
                con.Open();
                using (var transaction = con.BeginTransaction())
                {
                    try
                    {
                        // Mise à jour de BPerson
                        var personQuery = @"
                            UPDATE BPerson 
                            SET B1 = @B1, B2 = @B2, B3 = @B3, B4 = @B4, Adress = @Adress, Phone = @Phone, 
                                Email = @Email, B5 = @B5, B6 = @B6, B61 = @B61
                            WHERE Id = @Id";

                        using (var cmd = new SQLiteCommand(personQuery, con, transaction))
                        {
                            cmd.Parameters.AddWithValue("@Id", person.Id);
                            cmd.Parameters.AddWithValue("@B1", person.B1);
                            cmd.Parameters.AddWithValue("@B2", person.B2);
                            cmd.Parameters.AddWithValue("@B3", person.B3);
                            cmd.Parameters.AddWithValue("@B4", person.B4);
                            cmd.Parameters.AddWithValue("@Adress", person.Adress);
                            cmd.Parameters.AddWithValue("@Phone", person.Phone);
                            cmd.Parameters.AddWithValue("@Email", person.Email);
                            cmd.Parameters.AddWithValue("@B5", person.B5);
                            cmd.Parameters.AddWithValue("@B6", person.B6);
                            cmd.Parameters.AddWithValue("@B61", person.B61);

                            cmd.ExecuteNonQuery();
                        }

                        // Mise à jour ou insertion de Alocalisation
                        var localisationExistsQuery = "SELECT COUNT(*) FROM Alocalisation WHERE IdBPerson = @IdBPerson";
                        using (var cmd = new SQLiteCommand(localisationExistsQuery, con, transaction))
                        {
                            cmd.Parameters.AddWithValue("@IdBPerson", person.Id);
                            var exists = (long)cmd.ExecuteScalar() > 0;

                            var localisationQuery = exists ?
                                @"UPDATE Alocalisation 
                                  SET A1 = @A1, A2 = @A2, A3 = @A3, A4 = @A4, A5 = @A5, A6 = @A6, A7 = @A7, A8 = @A8
                                  WHERE IdBPerson = @IdBPerson" :
                                @"INSERT INTO Alocalisation (IdBPerson, A1, A2, A3, A4, A5, A6, A7, A8) 
                                  VALUES (@IdBPerson, @A1, @A2, @A3, @A4, @A5, @A6, @A7, @A8)";

                            using (var cmdLoc = new SQLiteCommand(localisationQuery, con, transaction))
                            {
                                cmdLoc.Parameters.AddWithValue("@IdBPerson", person.Id);
                                cmdLoc.Parameters.AddWithValue("@A1", localisation.A1);
                                cmdLoc.Parameters.AddWithValue("@A2", localisation.A2);
                                cmdLoc.Parameters.AddWithValue("@A3", localisation.A3);
                                cmdLoc.Parameters.AddWithValue("@A4", localisation.A4);
                                cmdLoc.Parameters.AddWithValue("@A5", localisation.A5);
                                cmdLoc.Parameters.AddWithValue("@A6", localisation.A6);
                                cmdLoc.Parameters.AddWithValue("@A7", localisation.A7);
                                cmdLoc.Parameters.AddWithValue("@A8", localisation.A8);

                                cmdLoc.ExecuteNonQuery();
                            }
                        }

                        // Mise à jour ou insertion de Ddescription
                        var descriptionExistsQuery = "SELECT COUNT(*) FROM Ddescription WHERE IdBPerson = @IdBPerson";
                        using (var cmd = new SQLiteCommand(descriptionExistsQuery, con, transaction))
                        {
                            cmd.Parameters.AddWithValue("@IdBPerson", person.Id);
                            var exists = (long)cmd.ExecuteScalar() > 0;

                            var descriptionQuery = exists ?
                                @"UPDATE Ddescription 
                                  SET D1 = @D1, D11 = @D11, D2 = @D2, D21 = @D21, D22 = @D22
                                  WHERE IdBPerson = @IdBPerson" :
                                @"INSERT INTO Ddescription (IdBPerson, D1, D11, D2, D21, D22) 
                                  VALUES (@IdBPerson, @D1, @D11, @D2, @D21, @D22)";

                            using (var cmdDesc = new SQLiteCommand(descriptionQuery, con, transaction))
                            {
                                cmdDesc.Parameters.AddWithValue("@IdBPerson", person.Id);
                                cmdDesc.Parameters.AddWithValue("@D1", description.D1);
                                cmdDesc.Parameters.AddWithValue("@D11", description.D11);
                                cmdDesc.Parameters.AddWithValue("@D2", description.D2);
                                cmdDesc.Parameters.AddWithValue("@D21", description.D21);
                                cmdDesc.Parameters.AddWithValue("@D22", description.D22);

                                cmdDesc.ExecuteNonQuery();
                            }
                        }

                        // Mise à jour ou insertion de Eclimat
                        var climatExistsQuery = "SELECT COUNT(*) FROM Eclimat WHERE IdBPerson = @IdBPerson";
                        using (var cmd = new SQLiteCommand(climatExistsQuery, con, transaction))
                        {
                            cmd.Parameters.AddWithValue("@IdBPerson", person.Id);
                            var exists = (long)cmd.ExecuteScalar() > 0;

                            var climatQuery = exists ?
                                @"UPDATE Eclimat 
                                  SET E1 = @E1, E2 = @E2, E3 = @E3, E4 = @E4, E5 = @E5, E611 = @E611, 
                                      E612 = @E612, E621 = @E621, E622 = @E622, E631 = @E631, E632 = @E632, 
                                      E641 = @E641, E642 = @E642, E651 = @E651, E652 = @E652, E661 = @E661, 
                                      E662 = @E662, E71 = @E71, E72 = @E72, E81 = @E81, E82 = @E82, E911 = @E911, 
                                      E912 = @E912, E921 = @E921, E922 = @E922, E931 = @E931, E932 = @E932, 
                                      E941 = @E941, E942 = @E942, E951 = @E951, E952 = @E952, E961 = @E961, 
                                      E962 = @E962, E971 = @E971, E972 = @E972, E981 = @E981, E982 = @E982, 
                                      E101 = @E101, E102 = @E102, E103 = @E103, E104 = @E104, E111 = @E111, 
                                      E112 = @E112, E121 = @E121, E122 = @E122, E131 = @E131, E132 = @E132, 
                                      E141 = @E141, E142 = @E142, E15 = @E15
                                  WHERE IdBPerson = @IdBPerson" :
                                @"INSERT INTO Eclimat (IdBPerson, E1, E2, E3, E4, E5, E611, E612, E621, E622, 
                                                       E631, E632, E641, E642, E651, E652, E661, E662, E71, E72, 
                                                       E81, E82, E911, E912, E921, E922, E931, E932, E941, E942, 
                                                       E951, E952, E961, E962, E971, E972, E981, E982, E101, E102, 
                                                       E103, E104, E111, E112, E121, E122, E131, E132, E141, E142, E15)
                                  VALUES (@IdBPerson, @E1, @E2, @E3, @E4, @E5, @E611, @E612, @E621, @E622, 
                                          @E631, @E632, @E641, @E642, @E651, @E652, @E661, @E662, @E71, @E72, 
                                          @E81, @E82, @E911, @E912, @E921, @E922, @E931, @E932, @E941, @E942, 
                                          @E951, @E952, @E961, @E962, @E971, @E972, @E981, @E982, @E101, @E102, 
                                          @E103, @E104, @E111, @E112, @E121, @E122, @E131, @E132, @E141, @E142, @E15)";

                            using (var cmdClim = new SQLiteCommand(climatQuery, con, transaction))
                            {
                                cmdClim.Parameters.AddWithValue("@IdBPerson", person.Id);
                                cmdClim.Parameters.AddWithValue("@E1", climat.E1);
                                cmdClim.Parameters.AddWithValue("@E2", climat.E2);
                                cmdClim.Parameters.AddWithValue("@E3", climat.E3);
                                cmdClim.Parameters.AddWithValue("@E4", climat.E4);
                                cmdClim.Parameters.AddWithValue("@E5", climat.E5);
                                cmdClim.Parameters.AddWithValue("@E611", climat.E611);
                                cmdClim.Parameters.AddWithValue("@E612", climat.E612);
                                cmdClim.Parameters.AddWithValue("@E621", climat.E621);
                                cmdClim.Parameters.AddWithValue("@E622", climat.E622);
                                cmdClim.Parameters.AddWithValue("@E631", climat.E631);
                                cmdClim.Parameters.AddWithValue("@E632", climat.E632);
                                cmdClim.Parameters.AddWithValue("@E641", climat.E641);
                                cmdClim.Parameters.AddWithValue("@E642", climat.E642);
                                cmdClim.Parameters.AddWithValue("@E651", climat.E651);
                                cmdClim.Parameters.AddWithValue("@E652", climat.E652);
                                cmdClim.Parameters.AddWithValue("@E661", climat.E661);
                                cmdClim.Parameters.AddWithValue("@E662", climat.E662);
                                cmdClim.Parameters.AddWithValue("@E71", climat.E71);
                                cmdClim.Parameters.AddWithValue("@E72", climat.E72);
                                cmdClim.Parameters.AddWithValue("@E81", climat.E81);
                                cmdClim.Parameters.AddWithValue("@E82", climat.E82);
                                cmdClim.Parameters.AddWithValue("@E911", climat.E911);
                                cmdClim.Parameters.AddWithValue("@E912", climat.E912);
                                cmdClim.Parameters.AddWithValue("@E921", climat.E921);
                                cmdClim.Parameters.AddWithValue("@E922", climat.E922);
                                cmdClim.Parameters.AddWithValue("@E931", climat.E931);
                                cmdClim.Parameters.AddWithValue("@E932", climat.E932);
                                cmdClim.Parameters.AddWithValue("@E941", climat.E941);
                                cmdClim.Parameters.AddWithValue("@E942", climat.E942);
                                cmdClim.Parameters.AddWithValue("@E951", climat.E951);
                                cmdClim.Parameters.AddWithValue("@E952", climat.E952);
                                cmdClim.Parameters.AddWithValue("@E961", climat.E961);
                                cmdClim.Parameters.AddWithValue("@E962", climat.E962);
                                cmdClim.Parameters.AddWithValue("@E971", climat.E971);
                                cmdClim.Parameters.AddWithValue("@E972", climat.E972);
                                cmdClim.Parameters.AddWithValue("@E981", climat.E981);
                                cmdClim.Parameters.AddWithValue("@E982", climat.E982);
                                cmdClim.Parameters.AddWithValue("@E101", climat.E101);
                                cmdClim.Parameters.AddWithValue("@E102", climat.E102);
                                cmdClim.Parameters.AddWithValue("@E103", climat.E103);
                                cmdClim.Parameters.AddWithValue("@E104", climat.E104);
                                cmdClim.Parameters.AddWithValue("@E111", climat.E111);
                                cmdClim.Parameters.AddWithValue("@E112", climat.E112);
                                cmdClim.Parameters.AddWithValue("@E121", climat.E121);
                                cmdClim.Parameters.AddWithValue("@E122", climat.E122);
                                cmdClim.Parameters.AddWithValue("@E131", climat.E131);
                                cmdClim.Parameters.AddWithValue("@E132", climat.E132);
                                cmdClim.Parameters.AddWithValue("@E141", climat.E141);
                                cmdClim.Parameters.AddWithValue("@E142", climat.E142);
                                cmdClim.Parameters.AddWithValue("@E15", climat.E15);

                                cmdClim.ExecuteNonQuery();
                            }
                        }

                        transaction.Commit();
                    }
                    catch (Exception e)
                    {
                        transaction.Rollback();
                        throw new Exception("Erreur lors de la mise à jour de la personne et de ses données associées : " + e.Message);
                    }
                }
            }
        }

        public void DeleteFullPerson(int personId)
        {
            using (var connection = GetConnection())
            {
                connection.Open();
                using (var transaction = connection.BeginTransaction())
                {
                    try
                    {
                        using (var command = new SQLiteCommand("DELETE FROM BPerson WHERE Id = @Id", connection, transaction))
                        {
                            command.Parameters.AddWithValue("@Id", personId);
                            command.ExecuteNonQuery();
                        }
                        transaction.Commit();
                    }
                    catch (Exception ex)
                    {
                        transaction.Rollback();
                        throw new Exception($"Erreur lors de la suppression de la personne ID {personId} : {ex.Message}");
                    }
                }
                connection.Close();
            }
        }

        public int GetTotalPersonsCount()
        {
            using var connection = GetConnection();
            try
            {
                connection.Open();
                Debug.WriteLine("[DEBUG] Connexion ouverte pour GetTotalPersonsCount");
                using var command = new SQLiteCommand("SELECT COUNT(*) FROM BPerson", connection);
                var result = command.ExecuteScalar();
                Debug.WriteLine("[DEBUG] Requête exécutée pour GetTotalPersonsCount");
                return Convert.ToInt32(result);
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[ERREUR] Erreur dans GetTotalPersonsCount : {ex.Message}\n{ex.StackTrace}");
                throw new Exception($"Erreur lors du comptage des personnes : {ex.Message}", ex);
            }
            finally
            {
                connection.Close();
                Debug.WriteLine("[DEBUG] Connexion fermée pour GetTotalPersonsCount");
            }
        }

        public int GetTotalMenCount()
        {
            using var connection = GetConnection();
            try
            {
                connection.Open();
                Debug.WriteLine("[DEBUG] Connexion ouverte pour GetTotalMenCount");
                using var command = new SQLiteCommand("SELECT COUNT(*) FROM BPerson WHERE B4 = @Gender", connection);
                command.Parameters.AddWithValue("@Gender", "Lahy");
                var result = command.ExecuteScalar();
                Debug.WriteLine("[DEBUG] Requête exécutée pour GetTotalMenCount");
                return Convert.ToInt32(result);
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[ERREUR] Erreur dans GetTotalMenCount : {ex.Message}\n{ex.StackTrace}");
                throw new Exception($"Erreur lors du comptage des hommes : {ex.Message}", ex);
            }
            finally
            {
                connection.Close();
                Debug.WriteLine("[DEBUG] Connexion fermée pour GetTotalMenCount");
            }
        }

        public int GetTotalWomenCount()
        {
            using var connection = GetConnection();
            try
            {
                connection.Open();
                Debug.WriteLine("[DEBUG] Connexion ouverte pour GetTotalWomenCount");
                using var command = new SQLiteCommand("SELECT COUNT(*) FROM BPerson WHERE B4 = @Gender", connection);
                command.Parameters.AddWithValue("@Gender", "Vavy");
                var result = command.ExecuteScalar();
                Debug.WriteLine("[DEBUG] Requête exécutée pour GetTotalWomenCount");
                return Convert.ToInt32(result);
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[ERREUR] Erreur dans GetTotalWomenCount : {ex.Message}\n{ex.StackTrace}");
                throw new Exception($"Erreur lors du comptage des femmes : {ex.Message}", ex);
            }
            finally
            {
                connection.Close();
                Debug.WriteLine("[DEBUG] Connexion fermée pour GetTotalWomenCount");
            }
        }

        public void ExportToExcel(string exportType, string bPersonQuery, (string, string)? parameter = null)
        {
            using var workbook = new XLWorkbook();
            using var connection = GetConnection();
            connection.Open();

            // Exporter BPerson
            var bPersonHeaders = new[] { "Id", "B1", "B2", "B3", "B4", "Adress", "Phone", "Email", "B5", "B6", "B61" };
            var worksheet = workbook.Worksheets.Add(exportType);
            for (int i = 0; i < bPersonHeaders.Length; i++)
            {
                worksheet.Cell(1, i + 1).Value = bPersonHeaders[i];
            }

            var ids = new List<int>();
            using (var command = new SQLiteCommand(bPersonQuery, connection))
            {
                if (parameter.HasValue)
                {
                    command.Parameters.AddWithValue(parameter.Value.Item1, parameter.Value.Item2);
                }
                using var reader = command.ExecuteReader();
                int row = 2;
                while (reader.Read())
                {
                    ids.Add(reader.GetInt32(0));
                    for (int i = 0; i < reader.FieldCount; i++)
                    {
                        worksheet.Cell(row, i + 1).Value = reader[i]?.ToString();
                    }
                    row++;
                }
            }

            // Exporter Alocalisation
            var alocHeaders = new[] { "Id", "IdBPerson", "A1", "A2", "A3", "A4", "A5", "A6", "A7", "A8" };
            var worksheetAloc = workbook.Worksheets.Add($"{exportType}_Alocalisation");
            for (int i = 0; i < alocHeaders.Length; i++)
            {
                worksheetAloc.Cell(1, i + 1).Value = alocHeaders[i];
            }
            string alocQuery = ids.Count > 0 ? $"SELECT * FROM Alocalisation WHERE IdBPerson IN ({string.Join(",", ids)})" : "SELECT * FROM Alocalisation WHERE 1=0";
            using (var commandAloc = new SQLiteCommand(alocQuery, connection))
            {
                using var readerAloc = commandAloc.ExecuteReader();
                int row = 2;
                while (readerAloc.Read())
                {
                    for (int i = 0; i < readerAloc.FieldCount; i++)
                    {
                        worksheetAloc.Cell(row, i + 1).Value = readerAloc[i]?.ToString();
                    }
                    row++;
                }
            }

            // Exporter Ddescription
            var descHeaders = new[] { "Id", "IdBPerson", "D1", "D11", "D2", "D21", "D22" };
            var worksheetDesc = workbook.Worksheets.Add($"{exportType}_Ddescription");
            for (int i = 0; i < descHeaders.Length; i++)
            {
                worksheetDesc.Cell(1, i + 1).Value = descHeaders[i];
            }
            string descQuery = ids.Count > 0 ? $"SELECT * FROM Ddescription WHERE IdBPerson IN ({string.Join(",", ids)})" : "SELECT * FROM Ddescription WHERE 1=0";
            using (var commandDesc = new SQLiteCommand(descQuery, connection))
            {
                using var readerDesc = commandDesc.ExecuteReader();
                int row = 2;
                while (readerDesc.Read())
                {
                    for (int i = 0; i < readerDesc.FieldCount; i++)
                    {
                        worksheetDesc.Cell(row, i + 1).Value = readerDesc[i]?.ToString();
                    }
                    row++;
                }
            }

            // Exporter Eclimat
            var climatHeaders = new[] { "Id", "IdBPerson", "E1", "E2", "E3", "E4", "E5", "E611", "E612", "E621", "E622",
                                       "E631", "E632", "E641", "E642", "E651", "E652", "E661", "E662", "E71", "E72",
                                       "E81", "E82", "E911", "E912", "E921", "E922", "E931", "E932", "E941", "E942",
                                       "E951", "E952", "E961", "E962", "E971", "E972", "E981", "E982", "E101", "E102",
                                       "E103", "E104", "E111", "E112", "E121", "E122", "E131", "E132", "E141", "E142", "E15" };
            var worksheetClimat = workbook.Worksheets.Add($"{exportType}_Eclimat");
            for (int i = 0; i < climatHeaders.Length; i++)
            {
                worksheetClimat.Cell(1, i + 1).Value = climatHeaders[i];
            }
            string climatQuery = ids.Count > 0 ? $"SELECT * FROM Eclimat WHERE IdBPerson IN ({string.Join(",", ids)})" : "SELECT * FROM Eclimat WHERE 1=0";
            using (var commandClimat = new SQLiteCommand(climatQuery, connection))
            {
                using var readerClimat = commandClimat.ExecuteReader();
                int row = 2;
                while (readerClimat.Read())
                {
                    for (int i = 0; i < readerClimat.FieldCount; i++)
                    {
                        worksheetClimat.Cell(row, i + 1).Value = readerClimat[i]?.ToString();
                    }
                    row++;
                }
            }

            // Sauvegarde du fichier et ouverture
            string filePath = System.IO.Path.Combine(System.AppDomain.CurrentDomain.BaseDirectory, $"export_{exportType.ToLower()}.xlsx");
            workbook.SaveAs(filePath);
            try
            {
                Process.Start(new ProcessStartInfo
                {
                    FileName = filePath,
                    UseShellExecute = true
                });
                Debug.WriteLine($"[DEBUG] Fichier Excel ouvert : {filePath}");
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[ERREUR] Impossible d'ouvrir le fichier Excel : {ex.Message}\n{ex.StackTrace}");
                throw new Exception($"Erreur lors de l'ouverture du fichier Excel : {ex.Message}", ex);
            }

            connection.Close();
        }


        /// <summary>
        /// Get a list of unique regions from the Alocalisation table.
        /// </summary>
        /// <returns></returns>

        public List<string> GetUniqueRegions()
        {
            var regions = new List<string>();
            using (var con = GetConnection())
            {
                con.Open();
                using (var cmd = new SQLiteCommand("SELECT DISTINCT A2 FROM Alocalisation", con))
                using (var reader = cmd.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        if (!reader.IsDBNull(0))
                        {
                            regions.Add(reader.GetString(0));
                        }
                    }
                }
            }
            return regions;
        }

        public (int enyCount, int tsiaCount) GetE1Counts(string region = null)
        {
            using (var con = GetConnection())
            {
                con.Open();
                string query = string.IsNullOrEmpty(region) ?
                    @"SELECT 
                        COUNT(CASE WHEN E1 = 'Eny' THEN 1 END) as enyCount,
                        COUNT(CASE WHEN E1 = 'Tsia' THEN 1 END) as tsiaCount
                      FROM Eclimat" :
                    @"SELECT 
                        COUNT(CASE WHEN E1 = 'Eny' THEN 1 END) as enyCount,
                        COUNT(CASE WHEN E1 = 'Tsia' THEN 1 END) as tsiaCount
                      FROM Eclimat
                      INNER JOIN Alocalisation ON Eclimat.IdBPerson = Alocalisation.IdBPerson
                      WHERE Alocalisation.A2 = @region";

                using (var cmd = new SQLiteCommand(query, con))
                {
                    if (!string.IsNullOrEmpty(region))
                    {
                        cmd.Parameters.AddWithValue("@region", region);
                    }
                    using (var reader = cmd.ExecuteReader())
                    {
                        if (reader.Read())
                        {
                            int enyCount = reader.IsDBNull(0) ? 0 : reader.GetInt32(0);
                            int tsiaCount = reader.IsDBNull(1) ? 0 : reader.GetInt32(1);
                            return (enyCount, tsiaCount);
                        }
                    }
                }
            }
            return (0, 0);
        }

        public (int enyCount, int tsiaCount) GetE101Counts(string region = null)
        {
            using (var con = GetConnection())
            {
                con.Open();
                string query = string.IsNullOrEmpty(region) ?
                    @"SELECT 
                        COUNT(CASE WHEN E101 = 'Eny' THEN 1 END) as enyCount,
                        COUNT(CASE WHEN E101 = 'Tsia' THEN 1 END) as tsiaCount
                      FROM Eclimat" :
                    @"SELECT 
                        COUNT(CASE WHEN E101 = 'Eny' THEN 1 END) as enyCount,
                        COUNT(CASE WHEN E101 = 'Tsia' THEN 1 END) as tsiaCount
                      FROM Eclimat
                      INNER JOIN Alocalisation ON Eclimat.IdBPerson = Alocalisation.IdBPerson
                      WHERE Alocalisation.A2 = @region";

                using (var cmd = new SQLiteCommand(query, con))
                {
                    if (!string.IsNullOrEmpty(region))
                    {
                        cmd.Parameters.AddWithValue("@region", region);
                    }
                    using (var reader = cmd.ExecuteReader())
                    {
                        if (reader.Read())
                        {
                            int enyCount = reader.IsDBNull(0) ? 0 : reader.GetInt32(0);
                            int tsiaCount = reader.IsDBNull(1) ? 0 : reader.GetInt32(1);
                            return (enyCount, tsiaCount);
                        }
                    }
                }
            }
            return (0, 0);
        }

        public (int enyCount, int tsiaCount) GetE121Counts(string region = null)
        {
            using (var con = GetConnection())
            {
                con.Open();
                string query = string.IsNullOrEmpty(region) ?
                    @"SELECT 
                        COUNT(CASE WHEN E121 = 'Eny' THEN 1 END) as enyCount,
                        COUNT(CASE WHEN E121 = 'Tsia' THEN 1 END) as tsiaCount
                      FROM Eclimat" :
                    @"SELECT 
                        COUNT(CASE WHEN E121 = 'Eny' THEN 1 END) as enyCount,
                        COUNT(CASE WHEN E121 = 'Tsia' THEN 1 END) as tsiaCount
                      FROM Eclimat
                      INNER JOIN Alocalisation ON Eclimat.IdBPerson = Alocalisation.IdBPerson
                      WHERE Alocalisation.A2 = @region";

                using (var cmd = new SQLiteCommand(query, con))
                {
                    if (!string.IsNullOrEmpty(region))
                    {
                        cmd.Parameters.AddWithValue("@region", region);
                    }
                    using (var reader = cmd.ExecuteReader())
                    {
                        if (reader.Read())
                        {
                            int enyCount = reader.IsDBNull(0) ? 0 : reader.GetInt32(0);
                            int tsiaCount = reader.IsDBNull(1) ? 0 : reader.GetInt32(1);
                            return (enyCount, tsiaCount);
                        }
                    }
                }
            }
            return (0, 0);
        }

        public Dictionary<string, int> GetE5Counts(string region = null)
        {
            var counts = new Dictionary<string, int>
    {
        { "Fisitrahana rano: 1", 0 },
        { "Sakafo: 2", 0 },
        { "Fambolena: 3", 0 },
        { "Fiompiana: 4", 0 },
        { "Jono: 5", 0 },
        { "Asa fivelomana hafa: 6", 0},
        { "Trano fonenana: 7", 0 },
        { "Fahasalamana: 8", 0}
    };

            using (var con = GetConnection())
            {
                con.Open();
                string query = string.IsNullOrEmpty(region) ?
                   @"SELECT E5 
              FROM Eclimat 
              WHERE E5 IS NOT NULL" :
            @"SELECT E5 
              FROM Eclimat 
              INNER JOIN Alocalisation ON Eclimat.IdBPerson = Alocalisation.IdBPerson 
              WHERE Alocalisation.A2 = @region 
              AND E5 IS NOT NULL";

                using (var cmd = new SQLiteCommand(query, con))
                {
                    if (!string.IsNullOrEmpty(region))
                    {
                        cmd.Parameters.AddWithValue("@region", region);
                    }
                    using (var reader = cmd.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            string e5Value = reader.GetString(0);
                            var values = e5Value.Split(',').Select(v => v.Trim());
                            foreach (var value in values)
                            {
                                if (counts.ContainsKey(value))
                                {
                                    counts[value]++;
                                    System.Diagnostics.Debug.WriteLine($"Incrémenté {value}: {counts[value]}");
                                }
                                else
                                {
                                    System.Diagnostics.Debug.WriteLine($"Valeur inattendue: {value}");
                                }
                            }

                        }
                    }
                }
            }
            return counts;
        }
    }
}