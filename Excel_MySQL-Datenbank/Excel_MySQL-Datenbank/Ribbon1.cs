using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Tools.Ribbon;
using MySql.Data.MySqlClient;
using System.Runtime.InteropServices;
using System.Reflection;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;
using Microsoft.Office.Tools.Excel;
using System.Data.Odbc;
using System.Configuration;

namespace Excel_MySQL_Datenbank
{
    public partial class Ribbon1
    {
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            int zeile = 1;
            string[] separator = { ",", "\r\n", };
            var ersteZeile = File.ReadLines("C:\\Users\\Ramona\\Desktop\\Praktikum\\MySQL\\auto.csv").First();
            //var --> definition einer nicht typisierten Variable (konkreter Typ ergibt sich aus dem zugewiesenen Wert)
            int zähler = ersteZeile.Split(separator, StringSplitOptions.None).Length;
            string spalte = Char.ConvertFromUtf32(zähler + 64);
            var reader = new StreamReader(File.OpenRead("C:\\Users\\Ramona\\Desktop\\Praktikum\\MySQL\\auto.csv"));
            List<string[]> values = new List<string[]>(); //new erstellt Objekt, ruft Konstruktor auf

            Excel.Window window = e.Control.Context; //aktives Fenster wird dargestellt
            Excel.Worksheet activeWorksheet = ((Excel.Worksheet)window.Application.ActiveSheet);
            //gibt ein Objekt zurück, das das aktuelle Blatt in der aktiven Arbeitsmappe darstellt 
            Excel.Range newWorksheet = activeWorksheet.get_Range("A" + zeile, spalte + zeile);
            //stellt eine Zelle/ein Zellenbereich dar   

            while (!reader.EndOfStream)
            {
                string line = reader.ReadLine();
                values.Add(line.Split(separator, StringSplitOptions.None)); //erste Zeile wird gesplittet hinzugefügt
                newWorksheet.Value2 = line.Split(separator, StringSplitOptions.None);
                zeile++;
                while (!reader.EndOfStream)
                {
                    Excel.Range newWorksheet2 = activeWorksheet.get_Range("A" + zeile, spalte + zeile);
                    string newline = reader.ReadLine();
                    values.Add(newline.Split(separator, StringSplitOptions.None));
                    newWorksheet2.Value2 = newline.Split(separator, StringSplitOptions.None);
                    zeile++;
                }
            }
        }

        private void button2_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                MySqlConnection mcon = new MySqlConnection("datasource=localhost;port=3306;username=root;password=");
                //Stellt eine Verbindung zur Datenbank her
                MySqlDataAdapter mda = new MySqlDataAdapter(
                "select * from auto.auto into outfile 'C://Users//Ramona//Desktop//Praktikum//Mysql2.csv'"
                + "Fields terminated By',' Lines Terminated By'\n' ", mcon);
                //bestimmte Tabelle wird in eine CSV-Datei konvertiert
                mcon.Open();
                DataSet ds = new DataSet();
                mda.Fill(ds, "auto"); //CSV-Datei wird mit den Werten aus "auto" gefüllt
                mcon.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        private void button3_Click(object sender, RibbonControlEventArgs e)
        {
            int zeile = 1;
            string[] separator = { ",", "\r\n", };
            var ersteZeile = File.ReadLines("C:\\Users\\Ramona\\Desktop\\Praktikum\\mysql2.csv").First();
            //var --> definition einer nicht typisierten Variable (konkreter Typ ergibt sich aus dem zugewiesenen Wert)
            int zähler = ersteZeile.Split(separator, StringSplitOptions.None).Length;
            string spalte = Char.ConvertFromUtf32(zähler + 64);
            var reader = new StreamReader(File.OpenRead("C:\\Users\\Ramona\\Desktop\\Praktikum\\mysql2.csv"));
            List<string[]> values = new List<string[]>();

            Excel.Window window = e.Control.Context;  //aktives Fenster wird dargestellt
            Excel.Worksheet activeWorksheet = ((Excel.Worksheet)window.Application.ActiveSheet);
            //gibt ein Objekt zurück, welches das aktuelle Blatt in der aktiven Arbeitsmappe darstellt 
            Excel.Range newWorksheet = activeWorksheet.get_Range("A" + zeile, spalte + zeile);
            //stellt eine Zelle/ein Zellenbereich dar 
            while (!reader.EndOfStream)
            {
                string line = reader.ReadLine();
                values.Add(line.Split(separator, StringSplitOptions.None)); //erste Zeile wird gesplittet hinzugefügt
                newWorksheet.Value2 = line.Split(separator, StringSplitOptions.None);
                zeile++;
                while (!reader.EndOfStream)
                {
                    Excel.Range newWorksheet2 = activeWorksheet.get_Range("A" + zeile, spalte + zeile);
                    string newline = reader.ReadLine();
                    values.Add(newline.Split(separator, StringSplitOptions.None));
                    newWorksheet2.Value2 = newline.Split(separator, StringSplitOptions.None);
                    zeile++;
                }
            }
        }

        private void button4_Click(object sender, RibbonControlEventArgs e)
        {
            string connectionString = "datasource=localhost;port=3306;username=root;password=";
            MySqlConnection connection = new MySqlConnection(connectionString);
            string MySql = "SELECT * FROM auto.auto";
            MySqlCommand cmd = new MySqlCommand(MySql, connection);
            MySqlDataReader reader;

            try
            {
                connection.Open();
                reader = cmd.ExecuteReader(); //es wird eine MySqlDataReader-Klasse erstellt
                while (reader.Read())
                {
                    MessageBox.Show(reader.GetValue(0) + " , " + reader.GetValue(1) + " , "
                    + reader.GetValue(2) + " , " + reader.GetValue(3));
                }
                reader.Close();
                cmd.Dispose();
                connection.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Can not open connection ! ");
            }
        }

        private void button5_Click(object sender, RibbonControlEventArgs e)
        {
            int zeile = 1;
            string[] separator = { ",", "\r\n" };
            string connectionString = "datasource=localhost;port=3306;username=root;password=";
            MySqlConnection connection = new MySqlConnection(connectionString);
            MySqlCommand command = connection.CreateCommand();
            command.CommandText = "SELECT * FROM auto.auto;";
            //command.CommandText = "SELECT * FROM test_excel.test_excel";
            MySqlDataReader reader;
            Excel.Window window = e.Control.Context; //aktives Fenster wird dargestellt
            Excel.Worksheet activeWorksheet = ((Excel.Worksheet)window.Application.ActiveSheet);
            //gibt ein Objekt zurück, das das aktuelle Blatt in der aktiven Arbeitsmappe darstellt 
            string commandLine =
            "SELECT COUNT(*) FROM information_schema.COLUMNS WHERE table_schema = 'auto' and table_name = 'auto'";
            //string commandLine = 
            //"SELECT COUNT(*) FROM information_schema.COLUMNS WHERE table_schema ='test_excel' and table_name='test_excel'";
            using (MySqlConnection connect = new MySqlConnection(connectionString))
            using (MySqlCommand cmd = new MySqlCommand(commandLine, connect))
            {
                connect.Open(); //Verbindung wird geöffnet
                try
                {
                    int i;
                    int zähler = Convert.ToInt32(cmd.ExecuteScalar()); //führt abfragen aus & gibt erste zeile&spalte zurück
                    string spalte = Char.ConvertFromUtf32(zähler + 64);
                    Excel.Range newWorksheet = activeWorksheet.get_Range("A" + zeile, spalte + zeile);
                    //stellt eine Zelle/ein Zellenbereich dar
                    connection.Open();
                    reader = command.ExecuteReader(); //es wird eine MySqlDataReader-Klasse erstellt
                    while (reader.Read())
                    {
                        string inhalt = "";
                        for (i = 0; i < reader.FieldCount; i++)//ruft die Anzahl der Spalten in der aktuelle Zeile ab
                        {
                            inhalt += reader.GetValue(i).ToString() + " , "; //weist links den Wert von links+rechts zu 
                        }
                        newWorksheet.Value2 = inhalt.Split(separator, StringSplitOptions.None);
                        zeile++;
                        while (reader.Read())
                        {
                            string inhalt2 = "";
                            Excel.Range newWorksheet2 = activeWorksheet.get_Range("A" + zeile, spalte + zeile);
                            for (i = 0; i < reader.FieldCount; i++)
                            {
                                inhalt2 += reader.GetValue(i).ToString() + " , ";
                            }
                            newWorksheet2.Value2 = inhalt2.Split(separator, StringSplitOptions.None);
                            zeile++;
                        }
                    }
                    reader.Close();
                    command.Dispose();
                    connection.Close();
                    connect.Close();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "Can not open connection ! ");
                }
            }
        }

    }
}
