using System;
using System.Data.SQLite;
using System.IO;
using System.Diagnostics;
using gestion_concrets.Models;

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
                
                //if (!Directory.Exists(dataFolder))
                //{
                //    Debug.WriteLine($"[INFO] Création du dossier racine Data : {dataFolder}");
                //    Directory.CreateDirectory(dataFolder);
                //} else
                //{
                //    Debug.WriteLine($"[INFO] Dossier racine Data existe déjà : {dataFolder}");
                //}


                //if (!File.Exists(dbPath))
                //{
                //    Debug.WriteLine($"[INFO] Création du fichier database.db : {dbPath}");
                //    SQLiteConnection.CreateFile(dbPath);
                //} else
                //{
                //    Debug.WriteLine($"[INFO] Fichier database.db existe déjà : {dbPath}");
                //}
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
                        B51 TEXT NOT NULL,
                        B52 TEXT NOT NULL,
                        B6 TEXT NOT NULL,
                        B7 TEXT NOT NULL,
                        B71 TEXT NOT NULL,
                        B72 TEXT NOT NULL,
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

        // Method to Add BPerson
    //    public void AddPerson(BPerson person)
    //    {
    //        using var con = GetConnection();
    //        con.Open();

    //        string sql = @"
    //    INSERT INTO BPerson (B1, B2, B3, Adress, Phone, Email, B4, B51, B52, B6, B7, B71, B72, B8)
    //    VALUES (@B1, @B2, @B3, @Adress, @Phone, @Email, @B4, @B51, @B52, @B6, @B7, @B71, @B72, @B8);
    //";

    //        using var cmd = new SQLiteCommand(sql, con);
    //        cmd.Parameters.AddWithValue("@B1", person.B1);
    //        cmd.Parameters.AddWithValue("@B2", person.B2);
    //        cmd.Parameters.AddWithValue("@B3", person.B3);
    //        cmd.Parameters.AddWithValue("@Adress", person.Adress);
    //        cmd.Parameters.AddWithValue("@Phone", person.Phone);
    //        cmd.Parameters.AddWithValue("@Email", person.Email);
    //        cmd.Parameters.AddWithValue("@B4", person.B4);
    //        cmd.Parameters.AddWithValue("@B51", person.B51);
    //        cmd.Parameters.AddWithValue("@B52", person.B52);
    //        cmd.Parameters.AddWithValue("@B6", person.B6);
    //        cmd.Parameters.AddWithValue("@B7", person.B7);
    //        cmd.Parameters.AddWithValue("@B71", person.B71);
    //        cmd.Parameters.AddWithValue("@B72", person.B72);
    //        cmd.Parameters.AddWithValue("@B8", person.B8);

    //        cmd.ExecuteNonQuery();
    //    }

        //public int GetLastInsertedPersonId()
        //{
        //    using var con = GetConnection();
        //    con.Open();

        //    using var cmd = new SQLiteCommand("SELECT last_insert_rowid();", con);
        //    return Convert.ToInt32(cmd.ExecuteScalar());
        //}

        // Method to Add Alocalisation
    //    public void AddAlocalisation(Alocalisation model)
    //    {
    //        using var con = GetConnection();
    //        con.Open();

    //        string sql = @"
    //    INSERT INTO Alocalisation (IdBPerson, A1, A2, A3, A4)
    //    VALUES (@IdBPerson, @A1, @A2, @A3, @A4);
    //";

    //        using var cmd = new SQLiteCommand(sql, con);
    //        cmd.Parameters.AddWithValue("@IdBPerson", model.IdBPerson);
    //        cmd.Parameters.AddWithValue("@A1", model.A1);
    //        cmd.Parameters.AddWithValue("@A2", model.A2);
    //        cmd.Parameters.AddWithValue("@A3", model.A3);
    //        cmd.Parameters.AddWithValue("@A4", model.A4);

    //        cmd.ExecuteNonQuery();
    //    }

        // Method to Add IIapplicationCDPH
    //    public void AddIIapplicationCDPH(IIapplicationCDPH model)
    //    {
    //        using var con = GetConnection();
    //        con.Open();

    //        string sql = @"
    //    INSERT INTO IIapplicationCDPH (IdBPerson, II1, II2, II3, II4, II5, II6)
    //    VALUES (@IdBPerson, @II1, @II2, @II3, @II4, @II5, @II6);
    //";

    //        using var cmd = new SQLiteCommand(sql, con);
    //        cmd.Parameters.AddWithValue("@IdBPerson", model.IdBPerson);
    //        cmd.Parameters.AddWithValue("@II1", model.II1);
    //        cmd.Parameters.AddWithValue("@II2", model.II2);
    //        cmd.Parameters.AddWithValue("@II3", model.II3);
    //        cmd.Parameters.AddWithValue("@II4", model.II4);
    //        cmd.Parameters.AddWithValue("@II5", model.II5);
    //        cmd.Parameters.AddWithValue("@II6", model.II6);

    //        cmd.ExecuteNonQuery();
    //    }

        // Method to Add IIIright
    //    public void AddIIIright(IIIright model)
    //    {
    //        using var con = GetConnection();
    //        con.Open();

    //        string sql = @"
    //    INSERT INTO IIIright (IdBPerson, III1, III2, III3, III21, III22, III23, III24, III25,
    //                          III31, III32, III33, III41, III42, III43, III51, III52, III53)
    //    VALUES (@IdBPerson, @III1, @III2, @III3, @III21, @III22, @III23, @III24, @III25,
    //            @III31, @III32, @III33, @III41, @III42, @III43, @III51, @III52, @III53);
    //";

    //        using var cmd = new SQLiteCommand(sql, con);
    //        cmd.Parameters.AddWithValue("@IdBPerson", model.IdBPerson);
    //        cmd.Parameters.AddWithValue("@III1", model.III1);
    //        cmd.Parameters.AddWithValue("@III2", model.III2);
    //        cmd.Parameters.AddWithValue("@III3", model.III3);
    //        cmd.Parameters.AddWithValue("@III21", model.III21);
    //        cmd.Parameters.AddWithValue("@III22", model.III22);
    //        cmd.Parameters.AddWithValue("@III23", model.III23);
    //        cmd.Parameters.AddWithValue("@III24", model.III24);
    //        cmd.Parameters.AddWithValue("@III25", model.III25);
    //        cmd.Parameters.AddWithValue("@III31", model.III31);
    //        cmd.Parameters.AddWithValue("@III32", model.III32);
    //        cmd.Parameters.AddWithValue("@III33", model.III33);
    //        cmd.Parameters.AddWithValue("@III41", model.III41);
    //        cmd.Parameters.AddWithValue("@III42", model.III42);
    //        cmd.Parameters.AddWithValue("@III43", model.III43);
    //        cmd.Parameters.AddWithValue("@III51", model.III51);
    //        cmd.Parameters.AddWithValue("@III52", model.III52);
    //        cmd.Parameters.AddWithValue("@III53", model.III53);

    //        cmd.ExecuteNonQuery();
    //    }


        // Method to Add Itransmission
    //    public void AddItransmission(Itransmission model)
    //    {
    //        using var con = GetConnection();
    //        con.Open();

    //        string sql = @"
    //    INSERT INTO Itransmission (IdBPerson, I11, I12, I2, I3, I4, I51, I52, I53, I54, I55, I56, I57, I58, I59, I510, I6, I7, I8, I9)
    //    VALUES (@IdBPerson, @I11, @I12, @I2, @I3, @I4, @I51, @I52, @I53, @I54, @I55, @I56, @I57, @I58, @I59, @I510, @I6, @I7, @I8, @I9);
    //";

    //        using var cmd = new SQLiteCommand(sql, con);
    //        cmd.Parameters.AddWithValue("@IdBPerson", model.IdBPerson);
    //        cmd.Parameters.AddWithValue("@I11", model.I11);
    //        cmd.Parameters.AddWithValue("@I12", model.I12);
    //        cmd.Parameters.AddWithValue("@I2", model.I2);
    //        cmd.Parameters.AddWithValue("@I3", model.I3);
    //        cmd.Parameters.AddWithValue("@I4", model.I4);
    //        cmd.Parameters.AddWithValue("@I51", model.I51);
    //        cmd.Parameters.AddWithValue("@I52", model.I52);
    //        cmd.Parameters.AddWithValue("@I53", model.I53);
    //        cmd.Parameters.AddWithValue("@I54", model.I54);
    //        cmd.Parameters.AddWithValue("@I55", model.I55);
    //        cmd.Parameters.AddWithValue("@I56", model.I56);
    //        cmd.Parameters.AddWithValue("@I57", model.I57);
    //        cmd.Parameters.AddWithValue("@I58", model.I58);
    //        cmd.Parameters.AddWithValue("@I59", model.I59);
    //        cmd.Parameters.AddWithValue("@I510", model.I510);
    //        cmd.Parameters.AddWithValue("@I6", model.I6);
    //        cmd.Parameters.AddWithValue("@I7", model.I7);
    //        cmd.Parameters.AddWithValue("@I8", model.I8);
    //        cmd.Parameters.AddWithValue("@I9", model.I9);

    //        cmd.ExecuteNonQuery();
    //    }

        // Method to Add IVdutyGov
    //    public void AddIVdutyGov(IVdutyGov model)
    //    {
    //        using var con = GetConnection();
    //        con.Open();

    //        string sql = @"
    //    INSERT INTO IVdutyGov (IdBPerson, IV11, IV12, IV13, IV2, IV3, IV4, IV51, IV52)
    //    VALUES (@IdBPerson, @IV11, @IV12, @IV13, @IV2, @IV3, @IV4, @IV51, @IV52);
    //";

    //        using var cmd = new SQLiteCommand(sql, con);
    //        cmd.Parameters.AddWithValue("@IdBPerson", model.IdBPerson);
    //        cmd.Parameters.AddWithValue("@IV11", model.IV11);
    //        cmd.Parameters.AddWithValue("@IV12", model.IV12);
    //        cmd.Parameters.AddWithValue("@IV13", model.IV13);
    //        cmd.Parameters.AddWithValue("@IV2", model.IV2);
    //        cmd.Parameters.AddWithValue("@IV3", model.IV3);
    //        cmd.Parameters.AddWithValue("@IV4", model.IV4);
    //        cmd.Parameters.AddWithValue("@IV51", model.IV51);
    //        cmd.Parameters.AddWithValue("@IV52", model.IV52);

    //        cmd.ExecuteNonQuery();
    //    }


        // Method to Add VdevSupport
    //    public void AddVdevSupport(VdevSupport model)
    //    {
    //        using var con = GetConnection();
    //        con.Open();

    //        string sql = @"
    //    INSERT INTO VdevSupport (IdBPerson, V1, V2, V3, V41, V42, V51, V52, V53)
    //    VALUES (@IdBPerson, @V1, @V2, @V3, @V41, @V42, @V51, @V52, @V53);
    //";

    //        using var cmd = new SQLiteCommand(sql, con);
    //        cmd.Parameters.AddWithValue("@IdBPerson", model.IdBPerson);
    //        cmd.Parameters.AddWithValue("@V1", model.V1);
    //        cmd.Parameters.AddWithValue("@V2", model.V2);
    //        cmd.Parameters.AddWithValue("@V3", model.V3);
    //        cmd.Parameters.AddWithValue("@V41", model.V41);
    //        cmd.Parameters.AddWithValue("@V42", model.V42);
    //        cmd.Parameters.AddWithValue("@V51", model.V51);
    //        cmd.Parameters.AddWithValue("@V52", model.V52);
    //        cmd.Parameters.AddWithValue("@V53", model.V53);

    //        cmd.ExecuteNonQuery();
    //    }

        // Method to Add VIpartnerCollab    
    //    public void AddVIpartnerCollab(VIpartnerCollab model)
    //    {
    //        using var con = GetConnection();
    //        con.Open();

    //        string sql = @"
    //    INSERT INTO VIpartnerCollab (IdBPerson, VI1, VI2, VI3)
    //    VALUES (@IdBPerson, @VI1, @VI2, @VI3);
    //";

    //        using var cmd = new SQLiteCommand(sql, con);
    //        cmd.Parameters.AddWithValue("@IdBPerson", model.IdBPerson);
    //        cmd.Parameters.AddWithValue("@VI1", model.VI1);
    //        cmd.Parameters.AddWithValue("@VI2", model.VI2);
    //        cmd.Parameters.AddWithValue("@VI3", model.VI3);

    //        cmd.ExecuteNonQuery();
    //    }

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
                            INSERT INTO Alocalisation (IdBPerson, A1, A2, A3, A4)
                            VALUES (@IdBPerson, @A1, @A2, @A3, @A4); SELECT last_insert_rowid()";

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
                            cmd.Parameters.AddWithValue("@B51", person.B51);
                            cmd.Parameters.AddWithValue("@B52", person.B52);
                            cmd.Parameters.AddWithValue("@B6", person.B6);
                            cmd.Parameters.AddWithValue("@B7", person.B7);
                            cmd.Parameters.AddWithValue("@B71", person.B71);
                            cmd.Parameters.AddWithValue("@B72", person.B72);
                            cmd.Parameters.AddWithValue("@B8", person.B8);

                            idPerson = (long)cmd.ExecuteScalar();
                        }

                        var localisationQuery = @"
                                INSERT INTO Alocalisation (IdBPerson, A1, A2, A3, A4)
                                VALUES (@IdBPerson, @A1, @A2, @A3, @A4);
                            ";
                        using (var cmd = new SQLiteCommand(localisationQuery, con, transaction))
                        {
                            cmd.Parameters.AddWithValue("@IdBPerson", localisation.IdBPerson);
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
                            cmd.Parameters.AddWithValue("@IdBPerson", applicationCDPH.IdBPerson);
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
                             cmd.Parameters.AddWithValue("@IdBPerson", right.IdBPerson);
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
                            cmd.Parameters.AddWithValue("@IdBPerson", transmission.IdBPerson);
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
                            cmd.Parameters.AddWithValue("@IdBPerson", dutyGov.IdBPerson);
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
                            cmd.Parameters.AddWithValue("@IdBPerson", devSupport.IdBPerson);
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
                            INSERT INTO VIpartnerCollab (IdBPerson, VI1, VI2, VI3)
                            VALUES (@IdBPerson, @VI1, @VI2, @VI3);
                        ";

                        using (var cmd = new SQLiteCommand(partnercollabQuery, con, transaction))
                        {
                            cmd.Parameters.AddWithValue("@IdBPerson", partnerCollab.IdBPerson);
                            cmd.Parameters.AddWithValue("@VI1", partnerCollab.VI1);
                            cmd.Parameters.AddWithValue("@VI2", partnerCollab.VI2);
                            cmd.Parameters.AddWithValue("@VI3", partnerCollab.VI3);

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
    }
}

