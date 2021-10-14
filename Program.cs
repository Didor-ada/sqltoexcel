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
            var connectionString = "Data Source=localhost;Initial Catalog=Parametres;User Id=DevCaiman;pwd=Dev!Caiman"; // connexion � la BDD
            var requete = "select * from Clients"; // selection de la bonne requ�te
            var donnees = RenvoyerDataTableDepuisRequeteSQL(requete, connectionString); // 
            var fichierExcel = GenererFichierExcelDepuisDonnees(donnees);

            var temp = System.IO.Path.GetTempFileName(); // le fichier excel sera un fichier temporaire
            temp =  Path.ChangeExtension(temp, ".xlsx");
            System.IO.File.WriteAllBytes(temp, fichierExcel);
            Process.Start(temp);


            string[] filePaths = Directory.GetFiles(@"E:\Traitement\", "*.sql"); // on va chercher les fichiers 'Requ�tes.sql", chemin puis extension
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
            // Directory : expose des m�thodes statiques pour cr�er, d�placer, �num�rer � travers des dossiers et des sous-dossiers
        }

        public static void AfficherListe(string filePath)
        {
            Console.WriteLine("Processes file '{0}'.", filePath);

        }


        private static DataTable RenvoyerDataTableDepuisRequeteSQL(string requete, string connectionString) // Table de donn�es = r�ulstat des requ�tes SQL
        {
            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                SqlCommand command = new SqlCommand(requete, connection);
                command.Connection.Open(); // Ouvre la connexion avec la database

                SqlDataAdapter adapter = new SqlDataAdapter(command); // SqlDataAdapter repr�sente un de commandes data et une connexion � la db qui sont utilit�s pour remplir la DataTable et update la BDD SQL

                DataTable datatable = new DataTable();
                adapter.Fill(datatable); // ici on remplit la datatable
                command.Connection.Close(); // ferme la connexion avec la database
                return datatable;
            }
        }

        private static Byte[] GenererFichierExcelDepuisDonnees(DataTable tableDonnees) // on g�n�re un fichier, il faut un tableau de Bytes.
        {
            using (var fluxEcriture = new MemoryStream()) // on alloue un flux de m�moire
            {
                ClosedXML.Excel.XLWorkbook wbook = new ClosedXML.Excel.XLWorkbook(); // module ClosedXML pour cr�er le classeur Excel
                wbook.Worksheets.Add(tableDonnees, "Donn�es"); // suite du module pour ajouter un sheet "Donne�s" dans le classeur
                wbook.SaveAs(fluxEcriture);
                return fluxEcriture.ToArray(); // on retourne le r�sultat dans le tableau
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

/*            //R�cup�ration de la liste des fichiers
                    var listeFichier = RenvoyerListeFichier("");
                    //R�cup�ration du fichier

                    var fichierSelection = listeFichier.First();
                    //R�cup�ration des donn�es

                    var tableDonnees = RenvoyerDonneesDepuisBaseDeDonnees(fichierSelection);
                    //Generation du fichier excel

                    var fichierExcel = genererFichierExceldpuisDonnees(tableDonnees);
                    //Ouverture du fichier excel
                    AfficherExcel(fichierExcel);*/

