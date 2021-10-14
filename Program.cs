using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Diagnostics;
using System.IO;
using System.Linq;
using ClosedXML;
using ClosedXML.Excel;

namespace sqltoexcel
{
    public class Program
    {

        public static void Main(string[] args)
        {
            var connectionString = "Data Source=localhost;Initial Catalog=Parametres;User Id=DevCaiman;pwd=Dev!Caiman"; // connexion à la BDD
            var requete = "select * from Clients"; // selection de la bonne requête
            var donnees = RenvoyerDataTableDepuisRequeteSQL(requete, connectionString); // 
            var fichierExcel = GenererFichierExcelDepuisDonnees(donnees);

            var temp = System.IO.Path.GetTempFileName(); // le fichier excel sera un fichier temporaire
            temp =  Path.ChangeExtension(temp, ".xlsx");
            System.IO.File.WriteAllBytes(temp, fichierExcel);
            Process.Start(temp);


            string[] filePaths = Directory.GetFiles(@"E:\Traitement\", "*.sql"); // on va chercher les fichiers 'Requêtes.sql", chemin puis extension
            foreach (string filePath in filePaths)
            {
                if (File.Exists(filePath)) // existence du fichier
                {
                    AfficherListe(filePath);
                }
                else if (Directory.Exists(filePath)) // existence du chemin
                {
                    CheminDesFichiers(filePath);
                }
                else
                {
                    Console.WriteLine(("{0} is not a valid file or directory.", filePath));
                }
            }
        }


        public static void CheminDesFichiers(string filePaths)
        {
            string[] entreeDuSousDossier = Directory.GetDirectories(filePaths);
            foreach (string sousDossier in entreeDuSousDossier)
                CheminDesFichiers(sousDossier);
        }

        public List<Fichier> RenvoyerListeFichier(string CheminDossier)
        {
            return System.IO.Directory.GetFiles(CheminDossier, "*.sql").Select(f => new Fichier(f)).ToList(); // Input/Output, Chemin des fichiers, get fichiers, + extension
            // Directory : expose des méthodes statiques pour créer, déplacer, énumérer à travers des dossiers et des sous-dossiers
        }

        public static void AfficherListe(string filePath)
        {
            Console.WriteLine("Processes file '{0}'.", filePath);

        }


        private static DataTable RenvoyerDataTableDepuisRequeteSQL(string requete, string connectionString) // Table de données = réulstat des requêtes SQL
        {
            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                SqlCommand command = new SqlCommand(requete, connection);
                command.Connection.Open(); // Ouvre la connexion avec la database

                SqlDataAdapter adapter = new SqlDataAdapter(command); // SqlDataAdapter représente un de commandes data et une connexion à la db qui sont utilités pour remplir la DataTable et update la BDD SQL

                DataTable datatable = new DataTable();
                adapter.Fill(datatable); // ici on remplit la datatable
                command.Connection.Close(); // ferme la connexion avec la database
                return datatable;
            }
        }

        private static Byte[] GenererFichierExcelDepuisDonnees(DataTable tableDonnees) // on génère un fichier, il faut un tableau de Bytes.
        {
            using (var fluxEcriture = new MemoryStream()) // on alloue un flux de mémoire
            {
                ClosedXML.Excel.XLWorkbook wbook = new ClosedXML.Excel.XLWorkbook(); // module ClosedXML pour créer le classeur Excel
                wbook.Worksheets.Add(tableDonnees, "Données"); // suite du module pour ajouter un sheet "Donneés" dans le classeur
                wbook.SaveAs(fluxEcriture);
                return fluxEcriture.ToArray(); // on retourne le résultat dans le tableau
            }
        }
    }
}


public class Fichier
{
    public Fichier()
    {
        NomFichier = string.Empty;
        ContenuFichier = string.Empty;
    }
    public string NomFichier { get; set; }
    public string ContenuFichier { get; set; }

    public Fichier(string CheminFichier) : this()
    {
        NomFichier = System.IO.Path.GetFileName(CheminFichier);
        ContenuFichier = System.IO.File.ReadAllText(CheminFichier);
    }
}

/*            //Récupération de la liste des fichiers
                    var listeFichier = RenvoyerListeFichier("");
                    //Récupération du fichier

                    var fichierSelection = listeFichier.First();
                    //Récupération des données

                    var tableDonnees = RenvoyerDonneesDepuisBaseDeDonnees(fichierSelection);
                    //Generation du fichier excel

                    var fichierExcel = genererFichierExceldpuisDonnees(tableDonnees);
                    //Ouverture du fichier excel
                    AfficherExcel(fichierExcel);*/

