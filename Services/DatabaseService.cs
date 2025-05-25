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
  
        private static readonly string projectRoot = Path.GetFullPath(
            Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "..", "..", "..")
        );

   
        private static readonly string dataFolder = Path.Combine(projectRoot, "Data");
        private static readonly string dbPath = Path.Combine(dataFolder, "database.db");
        private static readonly string connectionString = $"Data Source={dbPath};";

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
                        B2 INTEGER NOT NULL,
                        B3 TEXT NOT NULL,
                        Adress TEXT NOT NULL,
                        Phone TEXT NOT NULL,
                        Email TEXT NOT NULL,
                        B4 TEXT NOT NULL,
                        B42 TEXT NOT NULL,
                        B51 TEXT NOT NULL,
                        B52 TEXT NOT NULL,
                        B6 TEXT NOT NULL,
                        B61 TEXT NOT NULL,
                        B7 TEXT NOT NULL,
                        B71 TEXT NOT NULL,
                        B72 TEXT NOT NULL,
                        B73 TEXT NOT NULL,
                        B74 TEXT NOT NULL,
                        B8 TEXT NOT NULL
                    );
                    CREATE TABLE IF NOT EXISTS Alocalisation (
                        Id INTEGER PRIMARY KEY AUTOINCREMENT,
                        IdBPerson INTEGER NOT NULL,
                        A1 TEXT NOT NULL,
                        A2 TEXT NOT NULL,
                        A3 TEXT NOT NULL,
                        A4 TEXT NOT NULL,
                        FOREIGN KEY(IdBPerson) REFERENCES BPerson(Id) ON DELETE CASCADE
                    );
                    CREATE TABLE IF NOT EXISTS Itransmission (
                        Id INTEGER PRIMARY KEY AUTOINCREMENT,
                        IdBPerson INTEGER NOT NULL,
                        I11 TEXT NOT NULL,
                        I12 TEXT NOT NULL,
                        I2 TEXT NOT NULL,
                        I3 TEXT NOT NULL,
                        I4 TEXT NOT NULL,
                        I51 TEXT NOT NULL,
                        I52 TEXT NOT NULL,
                        I53 TEXT NOT NULL,
                        I54 TEXT NOT NULL,
                        I55 TEXT NOT NULL,
                        I56 TEXT NOT NULL,
                        I57 TEXT NOT NULL,
                        I58 TEXT NOT NULL,
                        I59 TEXT NOT NULL,
                        I510 TEXT NOT NULL,
                        I6 TEXT NOT NULL,
                        I7 TEXT NOT NULL,
                        I8 TEXT NOT NULL,
                        I9 TEXT NOT NULL,
                        FOREIGN KEY(IdBPerson) REFERENCES BPerson(Id) ON DELETE CASCADE
                    );
                    CREATE TABLE IF NOT EXISTS IIapplicationCDPH (
                        Id INTEGER PRIMARY KEY AUTOINCREMENT,
                        IdBPerson INTEGER NOT NULL,
                        II1 TEXT NOT NULL,
                        II2 TEXT NOT NULL,
                        II3 TEXT NOT NULL,
                        II4 TEXT NOT NULL,
                        II5 TEXT NOT NULL,
                        II6 TEXT NOT NULL,
                        FOREIGN KEY(IdBPerson) REFERENCES BPerson(Id) ON DELETE CASCADE
                    );
                    CREATE TABLE IF NOT EXISTS IIIright (
                        Id INTEGER PRIMARY KEY AUTOINCREMENT,
                        IdBPerson INTEGER NOT NULL,
                        III1 TEXT NOT NULL,
                        III2 TEXT NOT NULL,
                        III3 TEXT NOT NULL,
                        III21 TEXT NOT NULL,
                        III22 TEXT NOT NULL,
                        III23 TEXT NOT NULL,
                        III24 TEXT NOT NULL,
                        III25 TEXT NOT NULL,
                        III31 TEXT NOT NULL,
                        III32 TEXT NOT NULL,
                        III33 TEXT NOT NULL,
                        III41 TEXT NOT NULL,
                        III42 TEXT NOT NULL,
                        III43 TEXT NOT NULL,
                        III51 TEXT NOT NULL,
                        III52 TEXT NOT NULL,
                        III53 TEXT NOT NULL,
                        FOREIGN KEY(IdBPerson) REFERENCES BPerson(Id) ON DELETE CASCADE
                    );
                    CREATE TABLE IF NOT EXISTS IVdutyGov (
                        Id INTEGER PRIMARY KEY AUTOINCREMENT,
                        IdBPerson INTEGER NOT NULL,
                        IV11 TEXT NOT NULL,
                        IV12 TEXT NOT NULL,
                        IV13 TEXT NOT NULL,
                        IV2 TEXT NOT NULL,
                        IV3 TEXT NOT NULL,
                        IV4 TEXT NOT NULL,
                        IV51 TEXT NOT NULL,
                        IV52 TEXT NOT NULL,
                        FOREIGN KEY(IdBPerson) REFERENCES BPerson(Id) ON DELETE CASCADE
                    );
                    CREATE TABLE IF NOT EXISTS VdevSupport (
                        Id INTEGER PRIMARY KEY AUTOINCREMENT,
                        IdBPerson INTEGER NOT NULL,
                        V1 TEXT NOT NULL,
                        V2 TEXT NOT NULL,
                        V3 TEXT NOT NULL,
                        V41 TEXT NOT NULL,
                        V42 TEXT NOT NULL,
                        V51 TEXT NOT NULL,
                        V52 TEXT NOT NULL,
                        V53 TEXT NOT NULL,
                        FOREIGN KEY(IdBPerson) REFERENCES BPerson(Id) ON DELETE CASCADE
                    );
                    CREATE TABLE IF NOT EXISTS VIpartnerCollab (
                        Id INTEGER PRIMARY KEY AUTOINCREMENT,
                        IdBPerson INTEGER NOT NULL,
                        VI1 TEXT NOT NULL,
                        VI2 TEXT NOT NULL,
                        VI3 TEXT NOT NULL,
                        VI4 TEXT NOT NULL,
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
        public void AddFullPerson(BPerson person, Alocalisation localisation, IIapplicationCDPH applicationCDPH, IIIright right, Itransmission transmission, IVdutyGov dutyGov, VdevSupport devSupport, VIpartnerCollab partnerCollab)
        {
            using (var con = GetConnection())
            {
                con.Open();
                using (var transaction = con.BeginTransaction())
                {
                    try
                    {
                        var personQuery = @"
                           INSERT INTO BPerson (B1, B2, B3, Adress, Phone, Email, B4, B42, B51, B52, B6, B61, B7, B71, B72, B73, B74, B8)
                           VALUES (@B1, @B2, @B3, @Adress, @Phone, @Email, @B4, @B42, @B51, @B52, @B6, @B61, @B7, @B71, @B72, @B73, @B74, @B8); SELECT last_insert_rowid()";

                        long idPerson;

                        using (var cmd = new SQLiteCommand(personQuery, con, transaction))
                        {
                            cmd.Parameters.AddWithValue("@B1", person.B1);
                            cmd.Parameters.AddWithValue("@B2", person.B2);
                            cmd.Parameters.AddWithValue("@B3", person.B3);
                            cmd.Parameters.AddWithValue("@Adress", person.Adress);
                            cmd.Parameters.AddWithValue("@Phone", person.Phone);
                            cmd.Parameters.AddWithValue("@Email", person.Email);
                            cmd.Parameters.AddWithValue("@B4", person.B4);
                            cmd.Parameters.AddWithValue("@B42", person.B42);
                            cmd.Parameters.AddWithValue("@B51", person.B51);
                            cmd.Parameters.AddWithValue("@B52", person.B52);
                            cmd.Parameters.AddWithValue("@B6", person.B6);
                            cmd.Parameters.AddWithValue("@B61", person.B61);
                            cmd.Parameters.AddWithValue("@B7", person.B7);
                            cmd.Parameters.AddWithValue("@B71", person.B71);
                            cmd.Parameters.AddWithValue("@B72", person.B72);
                            cmd.Parameters.AddWithValue("@B73", person.B73);
                            cmd.Parameters.AddWithValue("@B74", person.B74);
                            cmd.Parameters.AddWithValue("@B8", person.B8);

                            idPerson = (long)cmd.ExecuteScalar();
                        }

                        var localisationQuery = @"
                                INSERT INTO Alocalisation (IdBPerson, A1, A2, A3, A4)
                                VALUES (@IdBPerson, @A1, @A2, @A3, @A4);
                            ";
                        using (var cmd = new SQLiteCommand(localisationQuery, con, transaction))
                        {
                            cmd.Parameters.AddWithValue("@IdBPerson", idPerson);
                            cmd.Parameters.AddWithValue("@A1", localisation.A1);
                            cmd.Parameters.AddWithValue("@A2", localisation.A2);
                            cmd.Parameters.AddWithValue("@A3", localisation.A3);
                            cmd.Parameters.AddWithValue("@A4", localisation.A4);

                            cmd.ExecuteNonQuery();
                        }


                        var applicationQuery = @"
                                INSERT INTO IIapplicationCDPH (IdBPerson, II1, II2, II3, II4, II5, II6)
                                VALUES (@IdBPerson, @II1, @II2, @II3, @II4, @II5, @II6);
                            ";

                        using (var cmd = new SQLiteCommand(applicationQuery, con, transaction))
                        {
                            cmd.Parameters.AddWithValue("@IdBPerson", idPerson);
                            cmd.Parameters.AddWithValue("@II1", applicationCDPH.II1);
                            cmd.Parameters.AddWithValue("@II2", applicationCDPH.II2);
                            cmd.Parameters.AddWithValue("@II3", applicationCDPH.II3);
                            cmd.Parameters.AddWithValue("@II4", applicationCDPH.II4);
                            cmd.Parameters.AddWithValue("@II5", applicationCDPH.II5);
                            cmd.Parameters.AddWithValue("@II6", applicationCDPH.II6);

                            cmd.ExecuteNonQuery();
                        }

                        string rightQuery = @"
                                INSERT INTO IIIright (IdBPerson, III1, III2, III3, III21, III22, III23, III24, III25,
                                                      III31, III32, III33, III41, III42, III43, III51, III52, III53)
                                VALUES (@IdBPerson, @III1, @III2, @III3, @III21, @III22, @III23, @III24, @III25,
                                        @III31, @III32, @III33, @III41, @III42, @III43, @III51, @III52, @III53);
                            ";

                        using (var cmd = new SQLiteCommand(rightQuery, con, transaction)) { 
                             cmd.Parameters.AddWithValue("@IdBPerson", idPerson);
                            cmd.Parameters.AddWithValue("@III1", right.III1);
                            cmd.Parameters.AddWithValue("@III2", right.III2);
                            cmd.Parameters.AddWithValue("@III3", right.III3);
                            cmd.Parameters.AddWithValue("@III21", right.III21);
                            cmd.Parameters.AddWithValue("@III22", right.III22);
                            cmd.Parameters.AddWithValue("@III23", right.III23);
                            cmd.Parameters.AddWithValue("@III24", right.III24);
                            cmd.Parameters.AddWithValue("@III25", right.III25);
                            cmd.Parameters.AddWithValue("@III31", right.III31);
                            cmd.Parameters.AddWithValue("@III32", right.III32);
                            cmd.Parameters.AddWithValue("@III33", right.III33);
                            cmd.Parameters.AddWithValue("@III41", right.III41);
                            cmd.Parameters.AddWithValue("@III42", right.III42);
                            cmd.Parameters.AddWithValue("@III43", right.III43);
                            cmd.Parameters.AddWithValue("@III51", right.III51);
                            cmd.Parameters.AddWithValue("@III52", right.III52);
                            cmd.Parameters.AddWithValue("@III53", right.III53);

                            cmd.ExecuteNonQuery();
                        }

                        string transmissionQuery = @"
                            INSERT INTO Itransmission (IdBPerson, I11, I12, I2, I3, I4, I51, I52, I53, I54, I55, I56, I57, I58, I59, I510, I6, I7, I8, I9)
                            VALUES (@IdBPerson, @I11, @I12, @I2, @I3, @I4, @I51, @I52, @I53, @I54, @I55, @I56, @I57, @I58, @I59, @I510, @I6, @I7, @I8, @I9);
                        ";

                        using (var cmd = new SQLiteCommand(transmissionQuery, con, transaction))
                        {
                            cmd.Parameters.AddWithValue("@IdBPerson", idPerson);
                            cmd.Parameters.AddWithValue("@I11", transmission.I11);
                            cmd.Parameters.AddWithValue("@I12", transmission.I12);
                            cmd.Parameters.AddWithValue("@I2", transmission.I2);
                            cmd.Parameters.AddWithValue("@I3", transmission.I3);
                            cmd.Parameters.AddWithValue("@I4", transmission.I4);
                            cmd.Parameters.AddWithValue("@I51", transmission.I51);
                            cmd.Parameters.AddWithValue("@I52", transmission.I52);
                            cmd.Parameters.AddWithValue("@I53", transmission.I53);
                            cmd.Parameters.AddWithValue("@I54", transmission.I54);
                            cmd.Parameters.AddWithValue("@I55", transmission.I55);
                            cmd.Parameters.AddWithValue("@I56", transmission.I56);
                            cmd.Parameters.AddWithValue("@I57", transmission.I57);
                            cmd.Parameters.AddWithValue("@I58", transmission.I58);
                            cmd.Parameters.AddWithValue("@I59", transmission.I59);
                            cmd.Parameters.AddWithValue("@I510", transmission.I510);
                            cmd.Parameters.AddWithValue("@I6", transmission.I6);
                            cmd.Parameters.AddWithValue("@I7", transmission.I7);
                            cmd.Parameters.AddWithValue("@I8", transmission.I8);
                            cmd.Parameters.AddWithValue("@I9", transmission.I9);

                            cmd.ExecuteNonQuery();
                        }

                        string dutygovQuery = @"
                            INSERT INTO IVdutyGov (IdBPerson, IV11, IV12, IV13, IV2, IV3, IV4, IV51, IV52)
                            VALUES (@IdBPerson, @IV11, @IV12, @IV13, @IV2, @IV3, @IV4, @IV51, @IV52);
                        ";

                        using (var cmd = new SQLiteCommand(dutygovQuery, con, transaction)) {
                            cmd.Parameters.AddWithValue("@IdBPerson", idPerson);
                            cmd.Parameters.AddWithValue("@IV11", dutyGov.IV11);
                            cmd.Parameters.AddWithValue("@IV12", dutyGov.IV12);
                            cmd.Parameters.AddWithValue("@IV13", dutyGov.IV13);
                            cmd.Parameters.AddWithValue("@IV2", dutyGov.IV2);
                            cmd.Parameters.AddWithValue("@IV3", dutyGov.IV3);
                            cmd.Parameters.AddWithValue("@IV4", dutyGov.IV4);
                            cmd.Parameters.AddWithValue("@IV51", dutyGov.IV51);
                            cmd.Parameters.AddWithValue("@IV52", dutyGov.IV52);

                            cmd.ExecuteNonQuery();
                        }

                        string devsupportQuery = @"
                            INSERT INTO VdevSupport (IdBPerson, V1, V2, V3, V41, V42, V51, V52, V53)
                            VALUES (@IdBPerson, @V1, @V2, @V3, @V41, @V42, @V51, @V52, @V53);
                        ";

                        using (var cmd = new SQLiteCommand(devsupportQuery, con, transaction))
                        {
                            cmd.Parameters.AddWithValue("@IdBPerson", idPerson);
                            cmd.Parameters.AddWithValue("@V1", devSupport.V1);
                            cmd.Parameters.AddWithValue("@V2", devSupport.V2);
                            cmd.Parameters.AddWithValue("@V3", devSupport.V3);
                            cmd.Parameters.AddWithValue("@V41", devSupport.V41);
                            cmd.Parameters.AddWithValue("@V42", devSupport.V42);
                            cmd.Parameters.AddWithValue("@V51", devSupport.V51);
                            cmd.Parameters.AddWithValue("@V52", devSupport.V52);
                            cmd.Parameters.AddWithValue("@V53", devSupport.V53);

                            cmd.ExecuteNonQuery();
                        }

                        string partnercollabQuery = @"
                            INSERT INTO VIpartnerCollab (IdBPerson, VI1, VI2, VI3, VI4)
                            VALUES (@IdBPerson, @VI1, @VI2, @VI3, @VI4);
                        ";

                        using (var cmd = new SQLiteCommand(partnercollabQuery, con, transaction))
                        {
                            cmd.Parameters.AddWithValue("@IdBPerson", idPerson);
                            cmd.Parameters.AddWithValue("@VI1", partnerCollab.VI1);
                            cmd.Parameters.AddWithValue("@VI2", partnerCollab.VI2);
                            cmd.Parameters.AddWithValue("@VI3", partnerCollab.VI3);
                            cmd.Parameters.AddWithValue("@VI4", partnerCollab.VI4);

                            cmd.ExecuteNonQuery();
                        }

                        transaction.Commit();

                    } catch (Exception e)
                    {
                        transaction.Rollback();
                        throw new Exception("Erreur lors de l'ajout de la personne et de sa localisation : " + e.Message);

                    }
                }

                con.Close();
            }
        }

        //GET All Person
        public List<BPerson> GetAllPersons()
        {
            var persons = new List<BPerson>();
            using (var con = GetConnection())
            {
                con.Open();
                using (var cmd = new SQLiteCommand("SELECT Id, B1, B2, B3, Adress, Phone, Email FROM BPerson", con))
                using (var reader = cmd.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        persons.Add(new BPerson
                        {
                            Id = reader.GetInt32(0),
                            B1 = reader.GetString(1),
                            B2 = reader.GetInt32(2),
                            B3 = reader.GetString(3),
                            Adress = reader.GetString(4),
                            Phone = reader.GetString(5),
                            Email = reader.GetString(6)
                        });
                    }
                }
            }

            return persons;
        }


        public IIIright GetIIIrightById(int idBPerson)
        {
            using (var con = GetConnection())
            {
                con.Open();
                try
                {
                    var query = @"
                        SELECT Id, IdBPerson, III1, III2, III3, III21, III22, III23, III24, III25,
                               III31, III32, III33, III41, III42, III43, III51, III52, III53
                        FROM IIIright
                        WHERE IdBPerson = @IdBPerson";

                    using (var cmd = new SQLiteCommand(query, con))
                    {
                        cmd.Parameters.AddWithValue("@IdBPerson", idBPerson);
                        using (var reader = cmd.ExecuteReader())
                        {
                            if (reader.Read())
                            {
                                return new IIIright
                                {
                                    Id = reader.GetInt32(0),
                                    IdBPerson = reader.GetInt32(1),
                                    III1 = reader.GetString(2),
                                    III2 = reader.GetString(3),
                                    III3 = reader.GetString(4),
                                    III21 = reader.GetString(5),
                                    III22 = reader.GetString(6),
                                    III23 = reader.GetString(7),
                                    III24 = reader.GetString(8),
                                    III25 = reader.GetString(9),
                                    III31 = reader.GetString(10),
                                    III32 = reader.GetString(11),
                                    III33 = reader.GetString(12),
                                    III41 = reader.GetString(13),
                                    III42 = reader.GetString(14),
                                    III43 = reader.GetString(15),
                                    III51 = reader.GetString(16),
                                    III52 = reader.GetString(17),
                                    III53 = reader.GetString(18)
                                };
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    throw new Exception("Erreur lors de la récupération des droits : " + ex.Message);
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
                        SELECT Id, IdBPerson, A1, A2, A3, A4
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
                                    A4 = reader.GetString(5)
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
        public BPerson GetBPersonById(int id)
        {
            using (var con = GetConnection())
            {
                con.Open();
                try
                {
                    var query = @"
                        SELECT Id, B1, B2, B3, Adress, Phone, Email, B4, B42, B51, B52, B6, B61, B7, B71, B72, B73, B74, B8
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
                                    B2 = reader.GetInt32(2),
                                    B3 = reader.GetString(3),
                                    Adress = reader.GetString(4),
                                    Phone = reader.GetString(5),
                                    Email = reader.GetString(6),
                                    B4 = reader.GetString(7),
                                    B42 = reader.GetString(8),
                                    B51 = reader.GetString(9),
                                    B52 = reader.GetString(10),
                                    B6 = reader.GetString(11),
                                    B61 = reader.GetString(12),
                                    B7 = reader.GetString(13),
                                    B71 = reader.GetString(14),
                                    B72 = reader.GetString(15),
                                    B73 = reader.GetString(16),
                                    B74 = reader.GetString(17),
                                    B8 = reader.GetString(18)
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
        
            public IIapplicationCDPH GetIIapplicationCDPHById(int idBPerson)
            {
                using (var con = GetConnection())
                {
                    con.Open();
                    try
                    {
                        var query = @"
                        SELECT Id, IdBPerson, II1, II2, II3, II4, II5, II6
                        FROM IIapplicationCDPH
                        WHERE IdBPerson = @IdBPerson";

                        using (var cmd = new SQLiteCommand(query, con))
                        {
                            cmd.Parameters.AddWithValue("@IdBPerson", idBPerson);
                            using (var reader = cmd.ExecuteReader())
                            {
                                if (reader.Read())
                                {
                                    return new IIapplicationCDPH
                                    {
                                        Id = reader.GetInt32(0),
                                        IdBPerson = reader.GetInt32(1),
                                        II1 = reader.GetString(2),
                                        II2 = reader.GetString(3),
                                        II3 = reader.GetString(4),
                                        II4 = reader.GetString(5),
                                        II5 = reader.GetString(6),
                                        II6 = reader.GetString(7)
                                    };
                                }
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        throw new Exception("Erreur lors de la récupération de l'application CDPH : " + ex.Message);
                    }
                    finally
                    {
                        con.Close();
                    }
                }
                return null;
            }

        public Itransmission GetItransmissionById(int idBPerson)
        {
            using (var con = GetConnection())
            {
                con.Open();
                try
                {
                    var query = @"
                        SELECT Id, IdBPerson, I11, I12, I2, I3, I4, I51, I52, I53, I54, I55, I56, I57, I58, I59, I510, I6, I7, I8, I9
                        FROM Itransmission
                        WHERE IdBPerson = @IdBPerson";

                    using (var cmd = new SQLiteCommand(query, con))
                    {
                        cmd.Parameters.AddWithValue("@IdBPerson", idBPerson);
                        using (var reader = cmd.ExecuteReader())
                        {
                            if (reader.Read())
                            {
                                return new Itransmission
                                {
                                    Id = reader.GetInt32(0),
                                    IdBPerson = reader.GetInt32(1),
                                    I11 = reader.GetString(2),
                                    I12 = reader.GetString(3),
                                    I2 = reader.GetString(4),
                                    I3 = reader.GetString(5),
                                    I4 = reader.GetString(6),
                                    I51 = reader.GetString(7),
                                    I52 = reader.GetString(8),
                                    I53 = reader.GetString(9),
                                    I54 = reader.GetString(10),
                                    I55 = reader.GetString(11),
                                    I56 = reader.GetString(12),
                                    I57 = reader.GetString(13),
                                    I58 = reader.GetString(14),
                                    I59 = reader.GetString(15),
                                    I510 = reader.GetString(16),
                                    I6 = reader.GetString(17),
                                    I7 = reader.GetString(18),
                                    I8 = reader.GetString(19),
                                    I9 = reader.GetString(20)
                                };
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    throw new Exception("Erreur lors de la récupération de la transmission : " + ex.Message);
                }
                finally
                {
                    con.Close();
                }
            }
            return null;
        }

        public IVdutyGov GetIVdutyGovById(int idBPerson)
        {
            using (var con = GetConnection())
            {
                con.Open();
                try
                {
                    var query = @"
                        SELECT Id, IdBPerson, IV11, IV12, IV13, IV2, IV3, IV4, IV51, IV52
                        FROM IVdutyGov
                        WHERE IdBPerson = @IdBPerson";

                    using (var cmd = new SQLiteCommand(query, con))
                    {
                        cmd.Parameters.AddWithValue("@IdBPerson", idBPerson);
                        using (var reader = cmd.ExecuteReader())
                        {
                            if (reader.Read())
                            {
                                return new IVdutyGov
                                {
                                    Id = reader.GetInt32(0),
                                    IdBPerson = reader.GetInt32(1),
                                    IV11 = reader.GetString(2),
                                    IV12 = reader.GetString(3),
                                    IV13 = reader.GetString(4),
                                    IV2 = reader.GetString(5),
                                    IV3 = reader.GetString(6),
                                    IV4 = reader.GetString(7),
                                    IV51 = reader.GetString(8),
                                    IV52 = reader.GetString(9)
                                };
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    throw new Exception("Erreur lors de la récupération du devoir gouvernemental : " + ex.Message);
                }
                finally
                {
                    con.Close();
                }
            }
            return null;
        }

        public VdevSupport GetVdevSupportById(int idBPerson)
        {
            using (var con = GetConnection())
            {
                con.Open();
                try
                {
                    var query = @"
                        SELECT Id, IdBPerson, V1, V2, V3, V41, V42, V51, V52, V53
                        FROM VdevSupport
                        WHERE IdBPerson = @IdBPerson";

                    using (var cmd = new SQLiteCommand(query, con))
                    {
                        cmd.Parameters.AddWithValue("@IdBPerson", idBPerson);
                        using (var reader = cmd.ExecuteReader())
                        {
                            if (reader.Read())
                            {
                                return new VdevSupport
                                {
                                    Id = reader.GetInt32(0),
                                    IdBPerson = reader.GetInt32(1),
                                    V1 = reader.GetString(2),
                                    V2 = reader.GetString(3),
                                    V3 = reader.GetString(4),
                                    V41 = reader.GetString(5),
                                    V42 = reader.GetString(6),
                                    V51 = reader.GetString(7),
                                    V52 = reader.GetString(8),
                                    V53 = reader.GetString(9)
                                };
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    throw new Exception("Erreur lors de la récupération du support de développement : " + ex.Message);
                }
                finally
                {
                    con.Close();
                }
            }
            return null;
        }
        public VIpartnerCollab GetVIpartnerCollabById(int idBPerson)
        {
            using (var con = GetConnection())
            {
                con.Open();
                try
                {
                    var query = @"
                        SELECT Id, IdBPerson, VI1, VI2, VI3, VI4
                        FROM VIpartnerCollab
                        WHERE IdBPerson = @IdBPerson";

                    using (var cmd = new SQLiteCommand(query, con))
                    {
                        cmd.Parameters.AddWithValue("@IdBPerson", idBPerson);
                        using (var reader = cmd.ExecuteReader())
                        {
                            if (reader.Read())
                            {
                                return new VIpartnerCollab
                                {
                                    Id = reader.GetInt32(0),
                                    IdBPerson = reader.GetInt32(1),
                                    VI1 = reader.GetString(2),
                                    VI2 = reader.GetString(3),
                                    VI3 = reader.GetString(4),
                                    VI4 = reader.GetString(5)
                                };
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    throw new Exception("Erreur lors de la récupération de la collaboration partenaires : " + ex.Message);
                }
                finally
                {
                    con.Close();
                }
            }
            return null;
        }


        // UPDATE All 
        public void UpdateFullPerson(BPerson person, Alocalisation localisation,
            IIapplicationCDPH applicationCDPH, IIIright right, Itransmission transmission,
            IVdutyGov dutyGov, VdevSupport devSupport, VIpartnerCollab partnerCollab)
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
                            SET B1 = @B1, B2 = @B2, B3 = @B3, Adress = @Adress, Phone = @Phone, 
                                Email = @Email, B4 = @B4, B42 = @B42, B51 = @B51, B52 = @B52, B6 = @B6, B61 = @B61, 
                                B7 = @B7, B71 = @B71, B72 = @B72, B73 = @B73, B74 = @B74, B8 = @B8
                            WHERE Id = @Id";

                        using (var cmd = new SQLiteCommand(personQuery, con, transaction))
                        {
                            cmd.Parameters.AddWithValue("@Id", person.Id);
                            cmd.Parameters.AddWithValue("@B1", person.B1);
                            cmd.Parameters.AddWithValue("@B2", person.B2);
                            cmd.Parameters.AddWithValue("@B3", person.B3);
                            cmd.Parameters.AddWithValue("@Adress", person.Adress);
                            cmd.Parameters.AddWithValue("@Phone", person.Phone);
                            cmd.Parameters.AddWithValue("@Email", person.Email);
                            cmd.Parameters.AddWithValue("@B4", person.B4);
                            cmd.Parameters.AddWithValue("@B42", person.B42);
                            cmd.Parameters.AddWithValue("@B51", person.B51);
                            cmd.Parameters.AddWithValue("@B52", person.B52);
                            cmd.Parameters.AddWithValue("@B6", person.B6);
                            cmd.Parameters.AddWithValue("@B61", person.B61);
                            cmd.Parameters.AddWithValue("@B7", person.B7);
                            cmd.Parameters.AddWithValue("@B71", person.B71);
                            cmd.Parameters.AddWithValue("@B72", person.B72);
                            cmd.Parameters.AddWithValue("@B73", person.B73);
                            cmd.Parameters.AddWithValue("@B74", person.B74);
                            cmd.Parameters.AddWithValue("@B8", person.B8);

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
                                  SET A1 = @A1, A2 = @A2, A3 = @A3, A4 = @A4 
                                  WHERE IdBPerson = @IdBPerson" :
                                @"INSERT INTO Alocalisation (IdBPerson, A1, A2, A3, A4) 
                                  VALUES (@IdBPerson, @A1, @A2, @A3, @A4)";

                            using (var cmdLoc = new SQLiteCommand(localisationQuery, con, transaction))
                            {
                                cmdLoc.Parameters.AddWithValue("@IdBPerson", person.Id);
                                cmdLoc.Parameters.AddWithValue("@A1", localisation.A1);
                                cmdLoc.Parameters.AddWithValue("@A2", localisation.A2);
                                cmdLoc.Parameters.AddWithValue("@A3", localisation.A3);
                                cmdLoc.Parameters.AddWithValue("@A4", localisation.A4);

                                cmdLoc.ExecuteNonQuery();
                            }
                        }

                        // Mise à jour ou insertion de IIapplicationCDPH
                        var applicationExistsQuery = "SELECT COUNT(*) FROM IIapplicationCDPH WHERE IdBPerson = @IdBPerson";
                        using (var cmd = new SQLiteCommand(applicationExistsQuery, con, transaction))
                        {
                            cmd.Parameters.AddWithValue("@IdBPerson", person.Id);
                            var exists = (long)cmd.ExecuteScalar() > 0;

                            var applicationQuery = exists ?
                                @"UPDATE IIapplicationCDPH 
                                  SET II1 = @II1, II2 = @II2, II3 = @II3, II4 = @II4, II5 = @II5, II6 = @II6 
                                  WHERE IdBPerson = @IdBPerson" :
                                @"INSERT INTO IIapplicationCDPH (IdBPerson, II1, II2, II3, II4, II5, II6) 
                                  VALUES (@IdBPerson, @II1, @II2, @II3, @II4, @II5, @II6)";

                            using (var cmdApp = new SQLiteCommand(applicationQuery, con, transaction))
                            {
                                cmdApp.Parameters.AddWithValue("@IdBPerson", person.Id);
                                cmdApp.Parameters.AddWithValue("@II1", applicationCDPH.II1);
                                cmdApp.Parameters.AddWithValue("@II2", applicationCDPH.II2);
                                cmdApp.Parameters.AddWithValue("@II3", applicationCDPH.II3);
                                cmdApp.Parameters.AddWithValue("@II4", applicationCDPH.II4);
                                cmdApp.Parameters.AddWithValue("@II5", applicationCDPH.II5);
                                cmdApp.Parameters.AddWithValue("@II6", applicationCDPH.II6);

                                cmdApp.ExecuteNonQuery();
                            }
                        }

                        // Mise à jour ou insertion de IIIright
                        var rightExistsQuery = "SELECT COUNT(*) FROM IIIright WHERE IdBPerson = @IdBPerson";
                        using (var cmd = new SQLiteCommand(rightExistsQuery, con, transaction))
                        {
                            cmd.Parameters.AddWithValue("@IdBPerson", person.Id);
                            var exists = (long)cmd.ExecuteScalar() > 0;

                            var rightQuery = exists ?
                                @"UPDATE IIIright 
                                  SET III1 = @III1, III2 = @III2, III3 = @III3, III21 = @III21, III22 = @III22, 
                                      III23 = @III23, III24 = @III24, III25 = @III25, III31 = @III31, III32 = @III32, 
                                      III33 = @III33, III41 = @III41, III42 = @III42, III43 = @III43, 
                                      III51 = @III51, III52 = @III52, III53 = @III53 
                                  WHERE IdBPerson = @IdBPerson" :
                                @"INSERT INTO IIIright (IdBPerson, III1, III2, III3, III21, III22, III23, III24, III25, 
                                                       III31, III32, III33, III41, III42, III43, III51, III52, III53) 
                                  VALUES (@IdBPerson, @III1, @III2, @III3, @III21, @III22, @III23, @III24, @III25, 
                                          @III31, @III32, @III33, @III41, @III42, @III43, @III51, @III52, @III53)";

                            using (var cmdRight = new SQLiteCommand(rightQuery, con, transaction))
                            {
                                cmdRight.Parameters.AddWithValue("@IdBPerson", person.Id);
                                cmdRight.Parameters.AddWithValue("@III1", right.III1);
                                cmdRight.Parameters.AddWithValue("@III2", right.III2);
                                cmdRight.Parameters.AddWithValue("@III3", right.III3);
                                cmdRight.Parameters.AddWithValue("@III21", right.III21);
                                cmdRight.Parameters.AddWithValue("@III22", right.III22);
                                cmdRight.Parameters.AddWithValue("@III23", right.III23);
                                cmdRight.Parameters.AddWithValue("@III24", right.III24);
                                cmdRight.Parameters.AddWithValue("@III25", right.III25);
                                cmdRight.Parameters.AddWithValue("@III31", right.III31);
                                cmdRight.Parameters.AddWithValue("@III32", right.III32);
                                cmdRight.Parameters.AddWithValue("@III33", right.III33);
                                cmdRight.Parameters.AddWithValue("@III41", right.III41);
                                cmdRight.Parameters.AddWithValue("@III42", right.III42);
                                cmdRight.Parameters.AddWithValue("@III43", right.III43);
                                cmdRight.Parameters.AddWithValue("@III51", right.III51);
                                cmdRight.Parameters.AddWithValue("@III52", right.III52);
                                cmdRight.Parameters.AddWithValue("@III53", right.III53);

                                cmdRight.ExecuteNonQuery();
                            }
                        }

                        // Mise à jour ou insertion de Itransmission
                        var transmissionExistsQuery = "SELECT COUNT(*) FROM Itransmission WHERE IdBPerson = @IdBPerson";
                        using (var cmd = new SQLiteCommand(transmissionExistsQuery, con, transaction))
                        {
                            cmd.Parameters.AddWithValue("@IdBPerson", person.Id);
                            var exists = (long)cmd.ExecuteScalar() > 0;

                            var transmissionQuery = exists ?
                                @"UPDATE Itransmission 
                                  SET I11 = @I11, I12 = @I12, I2 = @I2, I3 = @I3, I4 = @I4, I51 = @I51, 
                                      I52 = @I52, I53 = @I53, I54 = @I54, I55 = @I55, I56 = @I56, I57 = @I57, 
                                      I58 = @I58, I59 = @I59, I510 = @I510, I6 = @I6, I7 = @I7, I8 = @I8, I9 = @I9 
                                  WHERE IdBPerson = @IdBPerson" :
                                @"INSERT INTO Itransmission (IdBPerson, I11, I12, I2, I3, I4, I51, I52, I53, I54, 
                                                            I55, I56, I57, I58, I59, I510, I6, I7, I8, I9) 
                                  VALUES (@IdBPerson, @I11, @I12, @I2, @I3, @I4, @I51, @I52, @I53, @I54, 
                                          @I55, @I56, @I57, @I58, @I59, @I510, @I6, @I7, @I8, @I9)";

                            using (var cmdTrans = new SQLiteCommand(transmissionQuery, con, transaction))
                            {
                                cmdTrans.Parameters.AddWithValue("@IdBPerson", person.Id);
                                cmdTrans.Parameters.AddWithValue("@I11", transmission.I11);
                                cmdTrans.Parameters.AddWithValue("@I12", transmission.I12);
                                cmdTrans.Parameters.AddWithValue("@I2", transmission.I2);
                                cmdTrans.Parameters.AddWithValue("@I3", transmission.I3);
                                cmdTrans.Parameters.AddWithValue("@I4", transmission.I4);
                                cmdTrans.Parameters.AddWithValue("@I51", transmission.I51);
                                cmdTrans.Parameters.AddWithValue("@I52", transmission.I52);
                                cmdTrans.Parameters.AddWithValue("@I53", transmission.I53);
                                cmdTrans.Parameters.AddWithValue("@I54", transmission.I54);
                                cmdTrans.Parameters.AddWithValue("@I55", transmission.I55);
                                cmdTrans.Parameters.AddWithValue("@I56", transmission.I56);
                                cmdTrans.Parameters.AddWithValue("@I57", transmission.I57);
                                cmdTrans.Parameters.AddWithValue("@I58", transmission.I58);
                                cmdTrans.Parameters.AddWithValue("@I59", transmission.I59);
                                cmdTrans.Parameters.AddWithValue("@I510", transmission.I510);
                                cmdTrans.Parameters.AddWithValue("@I6", transmission.I6);
                                cmdTrans.Parameters.AddWithValue("@I7", transmission.I7);
                                cmdTrans.Parameters.AddWithValue("@I8", transmission.I8);
                                cmdTrans.Parameters.AddWithValue("@I9", transmission.I9);

                                cmdTrans.ExecuteNonQuery();
                            }
                        }

                        // Mise à jour ou insertion de IVdutyGov
                        var dutyGovExistsQuery = "SELECT COUNT(*) FROM IVdutyGov WHERE IdBPerson = @IdBPerson";
                        using (var cmd = new SQLiteCommand(dutyGovExistsQuery, con, transaction))
                        {
                            cmd.Parameters.AddWithValue("@IdBPerson", person.Id);
                            var exists = (long)cmd.ExecuteScalar() > 0;

                            var dutyGovQuery = exists ?
                                @"UPDATE IVdutyGov 
                                  SET IV11 = @IV11, IV12 = @IV12, IV13 = @IV13, IV2 = @IV2, IV3 = @IV3, 
                                      IV4 = @IV4, IV51 = @IV51, IV52 = @IV52 
                                  WHERE IdBPerson = @IdBPerson" :
                                @"INSERT INTO IVdutyGov (IdBPerson, IV11, IV12, IV13, IV2, IV3, IV4, IV51, IV52) 
                                  VALUES (@IdBPerson, @IV11, @IV12, @IV13, @IV2, @IV3, @IV4, @IV51, @IV52)";

                            using (var cmdDuty = new SQLiteCommand(dutyGovQuery, con, transaction))
                            {
                                cmdDuty.Parameters.AddWithValue("@IdBPerson", person.Id);
                                cmdDuty.Parameters.AddWithValue("@IV11", dutyGov.IV11);
                                cmdDuty.Parameters.AddWithValue("@IV12", dutyGov.IV12);
                                cmdDuty.Parameters.AddWithValue("@IV13", dutyGov.IV13);
                                cmdDuty.Parameters.AddWithValue("@IV2", dutyGov.IV2);
                                cmdDuty.Parameters.AddWithValue("@IV3", dutyGov.IV3);
                                cmdDuty.Parameters.AddWithValue("@IV4", dutyGov.IV4);
                                cmdDuty.Parameters.AddWithValue("@IV51", dutyGov.IV51);
                                cmdDuty.Parameters.AddWithValue("@IV52", dutyGov.IV52);

                                cmdDuty.ExecuteNonQuery();
                            }
                        }

                        // Mise à jour ou insertion de VdevSupport
                        var devSupportExistsQuery = "SELECT COUNT(*) FROM VdevSupport WHERE IdBPerson = @IdBPerson";
                        using (var cmd = new SQLiteCommand(devSupportExistsQuery, con, transaction))
                        {
                            cmd.Parameters.AddWithValue("@IdBPerson", person.Id);
                            var exists = (long)cmd.ExecuteScalar() > 0;

                            var devSupportQuery = exists ?
                                @"UPDATE VdevSupport 
                                  SET V1 = @V1, V2 = @V2, V3 = @V3, V41 = @V41, V42 = @V42, 
                                      V51 = @V51, V52 = @V52, V53 = @V53 
                                  WHERE IdBPerson = @IdBPerson" :
                                @"INSERT INTO VdevSupport (IdBPerson, V1, V2, V3, V41, V42, V51, V52, V53) 
                                  VALUES (@IdBPerson, @V1, @V2, @V3, @V41, @V42, @V51, @V52, @V53)";

                            using (var cmdDev = new SQLiteCommand(devSupportQuery, con, transaction))
                            {
                                cmdDev.Parameters.AddWithValue("@IdBPerson", person.Id);
                                cmdDev.Parameters.AddWithValue("@V1", devSupport.V1);
                                cmdDev.Parameters.AddWithValue("@V2", devSupport.V2);
                                cmdDev.Parameters.AddWithValue("@V3", devSupport.V3);
                                cmdDev.Parameters.AddWithValue("@V41", devSupport.V41);
                                cmdDev.Parameters.AddWithValue("@V42", devSupport.V42);
                                cmdDev.Parameters.AddWithValue("@V51", devSupport.V51);
                                cmdDev.Parameters.AddWithValue("@V52", devSupport.V52);
                                cmdDev.Parameters.AddWithValue("@V53", devSupport.V53);

                                cmdDev.ExecuteNonQuery();
                            }
                        }

                        // Mise à jour ou insertion de VIpartnerCollab
                        var partnerCollabExistsQuery = "SELECT COUNT(*) FROM VIpartnerCollab WHERE IdBPerson = @IdBPerson";
                        using (var cmd = new SQLiteCommand(partnerCollabExistsQuery, con, transaction))
                        {
                            cmd.Parameters.AddWithValue("@IdBPerson", person.Id);
                            var exists = (long)cmd.ExecuteScalar() > 0;

                            var partnerCollabQuery = exists ?
                                @"UPDATE VIpartnerCollab 
                                  SET VI1 = @VI1, VI2 = @VI2, VI3 = @VI3, VI4 = @VI4  
                                  WHERE IdBPerson = @IdBPerson" :
                                @"INSERT INTO VIpartnerCollab (IdBPerson, VI1, VI2, VI3, VI4) 
                                  VALUES (@IdBPerson, @VI1, @VI2, @VI3, @VI4)";

                            using (var cmdPartner = new SQLiteCommand(partnerCollabQuery, con, transaction))
                            {
                                cmdPartner.Parameters.AddWithValue("@IdBPerson", person.Id);
                                cmdPartner.Parameters.AddWithValue("@VI1", partnerCollab.VI1);
                                cmdPartner.Parameters.AddWithValue("@VI2", partnerCollab.VI2);
                                cmdPartner.Parameters.AddWithValue("@VI3", partnerCollab.VI3);
                                cmdPartner.Parameters.AddWithValue("@VI4", partnerCollab.VI4);

                                cmdPartner.ExecuteNonQuery();
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
                        // Supprimer dans BPerson
                        using (var command = new SQLiteCommand("DELETE FROM BPerson WHERE Id = @Id", connection, transaction))
                        {
                            command.Parameters.AddWithValue("@Id", personId);
                            command.ExecuteNonQuery();
                        }

                        // Supprimer dans Alocalisation
                        using (var command = new SQLiteCommand("DELETE FROM Alocalisation WHERE IdBPerson = @Id", connection, transaction))
                        {
                            command.Parameters.AddWithValue("@Id", personId);
                            command.ExecuteNonQuery();
                        }

                        // Supprimer dans IIapplicationCDPH
                        using (var command = new SQLiteCommand("DELETE FROM IIapplicationCDPH WHERE IdBPerson = @Id", connection, transaction))
                        {
                            command.Parameters.AddWithValue("@Id", personId);
                            command.ExecuteNonQuery();
                        }

                        // Supprimer dans IIIright
                        using (var command = new SQLiteCommand("DELETE FROM IIIright WHERE IdBPerson = @Id", connection, transaction))
                        {
                            command.Parameters.AddWithValue("@Id", personId);
                            command.ExecuteNonQuery();
                        }

                        // Supprimer dans Itransmission
                        using (var command = new SQLiteCommand("DELETE FROM Itransmission WHERE IdBPerson = @Id", connection, transaction))
                        {
                            command.Parameters.AddWithValue("@Id", personId);
                            command.ExecuteNonQuery();
                        }

                        // Supprimer dans IVdutyGov
                        using (var command = new SQLiteCommand("DELETE FROM IVdutyGov WHERE IdBPerson = @Id", connection, transaction))
                        {
                            command.Parameters.AddWithValue("@Id", personId);
                            command.ExecuteNonQuery();
                        }

                        // Supprimer dans VdevSupport
                        using (var command = new SQLiteCommand("DELETE FROM VdevSupport WHERE IdBPerson = @Id", connection, transaction))
                        {
                            command.Parameters.AddWithValue("@Id", personId);
                            command.ExecuteNonQuery();
                        }

                        // Supprimer dans VIpartnerCollab
                        using (var command = new SQLiteCommand("DELETE FROM VIpartnerCollab WHERE IdBPerson = @Id", connection, transaction))
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
                using var command = new SQLiteCommand("SELECT COUNT(*) FROM BPerson WHERE B3 = @Gender", connection);
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
                using var command = new SQLiteCommand("SELECT COUNT(*) FROM BPerson WHERE B3 = @Gender", connection);
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
            var bPersonHeaders = new[] { "Id", "B1", "B2", "B3", "Adress", "Phone", "Email", "B4", "B42", "B51", "B52", "B6", "B61", "B7", "B71", "B72", "B73", "B74", "B8" };
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
                    ids.Add(reader.GetInt32(0)); // Stocker Id pour les tables liées
                    for (int i = 0; i < reader.FieldCount; i++)
                    {
                        worksheet.Cell(row, i + 1).Value = reader[i]?.ToString();
                    }
                    row++;
                }
            }

            // Exporter Alocalisation
            var alocHeaders = new[] { "Id", "IdBPerson", "A1", "A2", "A3", "A4" };
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

            // Exporter IIapplicationCDPH
            var appHeaders = new[] { "Id", "IdBPerson", "II1", "II2", "II3", "II4", "II5", "II6" };
            var worksheetApp = workbook.Worksheets.Add($"{exportType}_IIapplicationCDPH");
            for (int i = 0; i < appHeaders.Length; i++)
            {
                worksheetApp.Cell(1, i + 1).Value = appHeaders[i];
            }
            string appQuery = ids.Count > 0 ? $"SELECT * FROM IIapplicationCDPH WHERE IdBPerson IN ({string.Join(",", ids)})" : "SELECT * FROM IIapplicationCDPH WHERE 1=0";
            using (var commandApp = new SQLiteCommand(appQuery, connection))
            {
                using var readerApp = commandApp.ExecuteReader();
                int row = 2;
                while (readerApp.Read())
                {
                    for (int i = 0; i < readerApp.FieldCount; i++)
                    {
                        worksheetApp.Cell(row, i + 1).Value = readerApp[i]?.ToString();
                    }
                    row++;
                }
            }

            // Exporter IIIright
            var rightHeaders = new[] { "Id", "IdBPerson", "III1", "III2", "III3", "III21", "III22", "III23", "III24", "III25", "III31", "III32", "III33", "III41", "III42", "III43", "III51", "III52", "III53" };
            var worksheetRight = workbook.Worksheets.Add($"{exportType}_IIIright");
            for (int i = 0; i < rightHeaders.Length; i++)
            {
                worksheetRight.Cell(1, i + 1).Value = rightHeaders[i];
            }
            string rightQuery = ids.Count > 0 ? $"SELECT * FROM IIIright WHERE IdBPerson IN ({string.Join(",", ids)})" : "SELECT * FROM IIIright WHERE 1=0";
            using (var commandRight = new SQLiteCommand(rightQuery, connection))
            {
                using var readerRight = commandRight.ExecuteReader();
                int row = 2;
                while (readerRight.Read())
                {
                    for (int i = 0; i < readerRight.FieldCount; i++)
                    {
                        worksheetRight.Cell(row, i + 1).Value = readerRight[i]?.ToString();
                    }
                    row++;
                }
            }

            // Exporter Itransmission
            var transHeaders = new[] { "Id", "IdBPerson", "I11", "I12", "I2", "I3", "I4", "I51", "I52", "I53", "I54", "I55", "I56", "I57", "I58", "I59", "I510", "I6", "I7", "I8", "I9" };
            var worksheetTrans = workbook.Worksheets.Add($"{exportType}_Itransmission");
            for (int i = 0; i < transHeaders.Length; i++)
            {
                worksheetTrans.Cell(1, i + 1).Value = transHeaders[i];
            }
            string transQuery = ids.Count > 0 ? $"SELECT * FROM Itransmission WHERE IdBPerson IN ({string.Join(",", ids)})" : "SELECT * FROM Itransmission WHERE 1=0";
            using (var commandTrans = new SQLiteCommand(transQuery, connection))
            {
                using var readerTrans = commandTrans.ExecuteReader();
                int row = 2;
                while (readerTrans.Read())
                {
                    for (int i = 0; i < readerTrans.FieldCount; i++)
                    {
                        worksheetTrans.Cell(row, i + 1).Value = readerTrans[i]?.ToString();
                    }
                    row++;
                }
            }

            // Exporter IVdutyGov
            var dutyHeaders = new[] { "Id", "IdBPerson", "IV11", "IV12", "IV13", "IV2", "IV3", "IV4", "IV51", "IV52" };
            var worksheetDuty = workbook.Worksheets.Add($"{exportType}_IVdutyGov");
            for (int i = 0; i < dutyHeaders.Length; i++)
            {
                worksheetDuty.Cell(1, i + 1).Value = dutyHeaders[i];
            }
            string dutyQuery = ids.Count > 0 ? $"SELECT * FROM IVdutyGov WHERE IdBPerson IN ({string.Join(",", ids)})" : "SELECT * FROM IVdutyGov WHERE 1=0";
            using (var commandDuty = new SQLiteCommand(dutyQuery, connection))
            {
                using var readerDuty = commandDuty.ExecuteReader();
                int row = 2;
                while (readerDuty.Read())
                {
                    for (int i = 0; i < readerDuty.FieldCount; i++)
                    {
                        worksheetDuty.Cell(row, i + 1).Value = readerDuty[i]?.ToString();
                    }
                    row++;
                }
            }

            // Exporter VdevSupport
            var devHeaders = new[] { "Id", "IdBPerson", "V1", "V2", "V3", "V41", "V42", "V51", "V52", "V53" };
            var worksheetDev = workbook.Worksheets.Add($"{exportType}_VdevSupport");
            for (int i = 0; i < devHeaders.Length; i++)
            {
                worksheetDev.Cell(1, i + 1).Value = devHeaders[i];
            }
            string devQuery = ids.Count > 0 ? $"SELECT * FROM VdevSupport WHERE IdBPerson IN ({string.Join(",", ids)})" : "SELECT * FROM VdevSupport WHERE 1=0";
            using (var commandDev = new SQLiteCommand(devQuery, connection))
            {
                using var readerDev = commandDev.ExecuteReader();
                int row = 2;
                while (readerDev.Read())
                {
                    for (int i = 0; i < readerDev.FieldCount; i++)
                    {
                        worksheetDev.Cell(row, i + 1).Value = readerDev[i]?.ToString();
                    }
                    row++;
                }
            }

            // Exporter VIpartnerCollab
            var partnerHeaders = new[] { "Id", "IdBPerson", "VI1", "VI2", "VI3", "VI4" };
            var worksheetPartner = workbook.Worksheets.Add($"{exportType}_VIpartnerCollab");
            for (int i = 0; i < partnerHeaders.Length; i++)
            {
                worksheetPartner.Cell(1, i + 1).Value = partnerHeaders[i];
            }
            string partnerQuery = ids.Count > 0 ? $"SELECT * FROM VIpartnerCollab WHERE IdBPerson IN ({string.Join(",", ids)})" : "SELECT * FROM VIpartnerCollab WHERE 1=0";
            using (var commandPartner = new SQLiteCommand(partnerQuery, connection))
            {
                using var readerPartner = commandPartner.ExecuteReader();
                int row = 2;
                while (readerPartner.Read())
                {
                    for (int i = 0; i < readerPartner.FieldCount; i++)
                    {
                        worksheetPartner.Cell(row, i + 1).Value = readerPartner[i]?.ToString();
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

            // Fermer la connexion
            connection.Close();
        }
    }
}
   


