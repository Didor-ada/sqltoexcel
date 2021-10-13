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


            var connectionString = "Data Source=localhost;Initial Catalog=Parametres;User Id=DevCaiman;pwd=Dev!Caiman";
            var requete = "select * from Clients";
            var donnees = RenvoyerDataTableDepuisRequeteSQL(requete, connectionString);
            var fichierExcel = GenererFichierExcelDepuisDonnees(donnees);
            var temp = System.IO.Path.GetTempFileName();
           temp =  Path.ChangeExtension(temp, ".xlsx");
            System.IO.File.WriteAllBytes(temp, fichierExcel);
            Process.Start(temp);
            //Ouverture du fichier excel
            //AfficherExcel(fichierExcel);




            string[] filePaths = Directory.GetFiles(@"E:\Traitement\", "*.sql");
            foreach (string filePath in filePaths)
            {
                if (File.Exists(filePath))
                {
                    AfficherListe(filePath);
                }
                else if (Directory.Exists(filePath))
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
            string[] subdirectoryEntries = Directory.GetDirectories(filePaths);
            foreach (string subdirectory in subdirectoryEntries)
                CheminDesFichiers(subdirectory);
        }


        public static void AfficherListe(string filePath)
        {
            Console.WriteLine("Processes file '{0}'.", filePath);

        }


        public List<Fichier> RenvoyerListeFichier(string CheminDossier)
        {
            return System.IO.Directory.GetFiles(CheminDossier, "*.sql").Select(f => new Fichier(f)).ToList();
        }


        private static DataTable RenvoyerDataTableDepuisRequeteSQL(string requete, string connectionString)
        {
            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                SqlCommand command = new SqlCommand(requete, connection);
                command.Connection.Open();

                SqlDataAdapter adapter = new SqlDataAdapter(command);

                DataTable datatable = new DataTable();
                adapter.Fill(datatable);
                command.Connection.Close();
                return datatable;
            }
        }

        private static Byte[] GenererFichierExcelDepuisDonnees(DataTable tableDonnees)
        {
            //ExportDataToExcel_SSG.ExportDataTableToExcel(tableDonnees, @"E:\TraitementSQLAExcel.xls");
            //  ClosedXML.Excel.XLWorkbook wbook = new ClosedXML.Excel.XLWorkbook();
            using (var fluxEcriture = new MemoryStream())
            {
                ClosedXML.Excel.XLWorkbook wbook = new ClosedXML.Excel.XLWorkbook();
                wbook.Worksheets.Add(tableDonnees, "Données");
                wbook.SaveAs(fluxEcriture);
                return fluxEcriture.ToArray();
            }


            //Byte[]

            //return ficherExcel;
            return null;


        }
    }
}


/*
            DataTable dt = new DataTable();
            ExportDataToExcel_SSG.ExportDataTableToExcel(dt, @"d:\Filename.xls");
            DataTable[] dd = new DataTable[10];

            ExportDataToExcel_SSG.MultiDataTableAs_MultiExcelFile(dd, @"d:\File.xls");
            DataSet ds = new DataSet();
            ExportDataToExcel_SSG.MergingMultiDataSetAs_singleExcelSheet(ds, @"d:\File.xls", false);*/


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


/*        public IEnumerable<Ficher> RenvoyerListeFichier(string cheminDossier)
        {

        }
*/



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

