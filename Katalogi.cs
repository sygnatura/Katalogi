using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using System.Text.RegularExpressions;
using System.Data.OleDb;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Threading;
using System.Globalization;
using System.Drawing.Imaging;

namespace Katalogi
{
    public partial class Form1 : Form
    {
        public DateTime dataTeraz;
        public string plik_log;
        public static string convert = @"C:\Program Files\ImageMagick\convert.exe";
        public static string baza_iza = null;

        public static Regex EXIF_INFO = new Regex("(\\w+): (.+)", RegexOptions.Compiled);
        public static Regex STRONY = new Regex("(\\d+)", RegexOptions.Compiled);
        public static Regex LICZBA = new Regex("^(\\d+)\\.?\\d?", RegexOptions.Compiled);
        public static Regex OKLADKA = new Regex("^0+$", RegexOptions.Compiled);
        public static Regex NUMER_STRONY = new Regex("^(\\d+)(_(\\d+))?$", RegexOptions.Compiled);
        public static Regex DANE = new Regex("(\\d+)x(\\d+) (\\d+)x(\\d+)[-+](\\d+)[+-](\\d+)", RegexOptions.Compiled);
        public static Regex ZERO = new Regex(Regex.Escape("0"), RegexOptions.Compiled);
        public static Regex kolorTest = new Regex("Alpha:.+?mean:.+?\\(([^\\)]+)", RegexOptions.Singleline | RegexOptions.Compiled);
        public static Regex DODANE = new Regex("^\\d+_[a-z]{1}$", RegexOptions.Compiled);
        public static Regex SERIA_REGEX = new Regex("(\\w+)\\.\\w+", RegexOptions.Compiled);

        public string nr_zespolu_old = "";
        public string cd_zespolu_old = "";
        public string seria_old = null;
        public string seria_old2 = "";
        public string sygnatura_old = "";
        Skan skan_old = null;
        public int strona_prev = -1;
        public int strona_old = 0;
        public bool czyBledy = false;
        public string czyKasowac = null;
        public string czyWszystkieMarginesy = null;
        public int ilosc_okladek = 1;
        public int tryb = -1;
        bool wzornik = false;
        int osoba_id = 0;
        int finansowanie_id = 0;
        int licz_skany = 0;
        List<string> dodaneStrony;
        List<string> pusteStrony;

        public Form1()
        {
            InitializeComponent();

            profilICC.Text = AppDomain.CurrentDomain.BaseDirectory + "data\\profil.icc";
            //if (File.Exists(convert) == false)
            convert = AppDomain.CurrentDomain.BaseDirectory + "bin\\im\\convert.exe";
            baza_iza = AppDomain.CurrentDomain.BaseDirectory + "data\\IZADANE.mdb";
        }

        private void button1_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog dialog = new FolderBrowserDialog();
            dialog.ShowDialog();
            if (dialog.SelectedPath != "")
            {
                katalogBox.Text = dialog.SelectedPath;
            }

        }

        private void DirSearch(string sDir)
        {
            // 0 naprawa, 1 import, 2 weryfikacja, 3 numerowanie, 4 metadane, 5 crc, 6 metryczki
            string rozszerzenie = "*.*";

            try
            {
                // GLOWNY KATALOG
                string[] pliki = Directory.GetFiles(sDir, rozszerzenie);
                int ilosc_plikow = pliki.Length;
                bool SMA = false;

                if (ilosc_plikow > 0)
                {
                    switch (tryb)
                    {
                        case 0:
                            this.Text = "Naprawianie skanów z " + sDir;
                            break;
                        case 1:
                            this.Text = "Importowanie skanów z " + sDir;
                            break;
                        case 2:
                            this.Text = "Weryfikacja katalogu: " + sDir;
                            weryfikacjaInfo.AppendText("\nWeryfikacja katalogu: " + sDir + "\n______________________________________________\n");
                            sygnatura_old = "";
                            strona_old = 0;
                            wzornik = false;
                            break;
                        case 3:
                            this.Text = "Sortowanie katalogu: " + sDir;
                            weryfikacjaInfo.AppendText("\nSortowanie katalogu: " + sDir + "\n______________________________________________\n");
                            sygnatura_old = "";
                            strona_old = 0;
                            break;
                        case 4:
                            this.Text = "Dodaję metadane dla skanów z: " + sDir;
                            string[] tiffy = Directory.GetFiles(sDir, "*.tif");
                            int ilosc_tiffy = tiffy.Length;
                            if (ilosc_tiffy > 0)
                            {
                                if (ilosc_tiffy > 2)
                                {
                                    int skan1 = 0;
                                    int skan2 = 0;
                                    int skan3 = 0;
                                    Random random = new Random();
                                    skan1 = random.Next(0, ilosc_tiffy);
                                    do
                                    {
                                        skan2 = random.Next(0, ilosc_tiffy);
                                    } while (skan1 == skan2);

                                    do
                                    {
                                        skan3 = random.Next(0, ilosc_tiffy);
                                    } while (skan3 == skan1 || skan3 == skan2);

                                    SMA = sprawdzSkaner(pliki[skan1], pliki[skan2], pliki[skan3]);
                                }
                                else if (ilosc_tiffy == 2) SMA = sprawdzSkaner(pliki[0], pliki[1], null);
                                else SMA = sprawdzSkaner(pliki[0], null, null);
                            }
                            break;
                        case 5:
                            this.Text = "Generowanie CRC dla skanów z: " + sDir;
                            break;
                        case 6:
                            this.Text = "Generowanie metryczek dla skanów z: " + sDir;
                            licz_skany = 0;
                            dodaneStrony = new List<string>();
                            pusteStrony = new List<string>();
                            break;
                        case 7:
                            this.Text = "Zastępowanie skanów z " + sDir;
                            break;
                    }
                }

                for (int i = 0; i < ilosc_plikow; i++)
                {
                    //MessageBox.Show("procent: " + procent);
                    int procent = (int)(0.5f + ((100f * (i + 1)) / ilosc_plikow));
                    // Report progress.
                    if (progressBar.Value != procent)
                    {
                        progressBar.Value = procent;
                        Application.DoEvents();
                        System.Threading.Thread.Sleep(1);

                        //Thread.Sleep(1);
                    }

                    switch (tryb)
                    {
                        case 0:
                            if (napraw(pliki[i]) == false) i = ilosc_plikow;
                            break;
                        case 1:
                            importuj(pliki[i]);
                            break;
                        case 2:
                            // jezeli ostatni plik z katalogu
                            if (i == ilosc_plikow - 1) weryfikuj(pliki[i], true);
                            else weryfikuj(pliki[i], false);
                            break;
                        case 3:
                            sortuj(pliki[i]);
                            if (i == (ilosc_plikow - 1)) weryfikacjaInfo.AppendText("Gotowe\n");
                            break;
                        case 4:
                            metadane(pliki[i], SMA);
                            break;
                        case 5:
                            liczcrc(pliki[i]);
                            break;
                        case 6:
                            // jezeli ostatni plik z katalogu
                            if (i == ilosc_plikow - 1)
                            {
                                if (metryczki(pliki[i], true, false))
                                {
                                    // szukaj innego ostatniego skanu
                                    int skan_numer = i - 1;
                                    while (metryczki(pliki[skan_numer], true, true) && skan_numer > 0)
                                    {
                                        skan_numer--;
                                    }
                                }
                            }
                            else metryczki(pliki[i], false, false);
                            break;
                        case 7:
                            zastapSkany(pliki[i], i, skanyCel.Text);
                            break;
                    }

                }

                // zamiana ewentualnych plikow z rozszerzeniem .new na .tif
                if (tryb == 3 && ilosc_plikow > 0)
                {
                    string[] pliki_new = Directory.GetFiles(sDir, "*.new");
                    for (int i = 0; i < pliki_new.Length; i++)
                    {
                        System.IO.File.Move(pliki_new[i], pliki_new[i].Replace(".new", ""));
                    }
                }


                // PODKATALOGI
                foreach (string d in Directory.GetDirectories(sDir)) DirSearch(d);
            }
            catch (System.Exception excpt)
            {
                MessageBox.Show(excpt.Message);
            }
        }

        public bool napraw(string plik)
        {
            Skan skan = new Skan(plik);
            if (skan.PobierzPlik() != null)
            {
                Skan skanDocelowy = null;

                if (checkBaza.Checked)
                {
                    Skan skanIza = new Skan(skan);

                    if (skan_old == null) skanIza.seria = pobierzSerie(skan.nr_zespolu, skan.sygnatura);
                    // jezeli juz pobrano poprzednie dane
                    else if (skan.nr_zespolu.Equals(skan_old.nr_zespolu) && skan.sygnatura.Equals(skan_old.sygnatura)) skanIza.seria = skan_old.seria;

                    if (skanIza.seria != null)
                    {
                        // powiadom ze jest wiecej mozliwych serwii dla tego zespolu
                        if (skanIza.seria.Equals("-1")) bledyBox.AppendText(plik + " --> kilka możliwych serii\n");
                        else skanDocelowy = new Katalogi.Skan(System.IO.Path.Combine(katalogDocelowy.Text + skanIza.GenerujStruktureKatalogow(), skanIza.GenerujNazwePliku()));

                    }
                    skan_old = skanIza;
                }

                if(skanDocelowy == null) skanDocelowy = new Skan(System.IO.Path.Combine(katalogDocelowy.Text + skan.GenerujStruktureKatalogow(), skan.GenerujNazwePliku()));

                //skanDocelowy 
                string katalog_docelowy = katalogDocelowy.Text + skanDocelowy.GenerujStruktureKatalogow();
                if (skanyWeryfikacja.Text.Equals(katalog_docelowy) == false)
                {
                    skanyWeryfikacja.Text = katalog_docelowy;
                    Clipboard.SetText(katalog_docelowy);
                }

                bool exists = System.IO.Directory.Exists(katalog_docelowy);

                if (!exists) System.IO.Directory.CreateDirectory(katalog_docelowy);


                // jezeli w katalogu docelowym istnieje plik o takiej samej nazwie to policz crc dla obu i porownaj
                if (System.IO.File.Exists(skanDocelowy.PobierzPlik()))
                {
                    string crc1 = skan.GenerujCRC();
                    string crc2 = skanDocelowy.GenerujCRC();
                    
                    // jezeli to te same pliki to skasuj plik zrodlowy
                    if (crc1 != null && crc2 != null && crc1.Equals(crc2))
                    {
                        bledyBox.AppendText(plik + " --> usuwanie duplikatu\n");
                        System.IO.File.Delete(plik);
                        return true;
                    }
                    else bledyBox.AppendText(plik + " --> " + katalog_docelowy + "\n");
                    return false;
                }
                else
                {
                    System.IO.File.Move(skan.PobierzPlik(), skanDocelowy.PobierzPlik());
                    try
                    {
                        if (logiCheck.Checked)
                        {
                            System.IO.StreamWriter file = new System.IO.StreamWriter(plik_log, true);
                            file.WriteLine(skan.PobierzPlik() + "\t" + skanDocelowy.PobierzPlik());
                            file.Close();
                        }
                    }
                    catch
                    {
                    }
                }
            }
            else bledyBox.AppendText(plik + " nie pasuje do wzorca\n");

            return true;

        }

        private void button2_Click(object sender, EventArgs e)
        {
            //MessageBox.Show(pobierzSerie("81", "1"));
            if (katalogDocelowy.Text.Length != 0 && katalogBox.Text.Length != 0)
            {
                // dodawanie \ na koncu katalogu docelowego
                if (!katalogDocelowy.Text.EndsWith("\\")) katalogDocelowy.Text = katalogDocelowy.Text + "\\";
                bledyBox.Clear();
                tabControl1.Enabled = false;
                this.Text = "Katalogi - naprawianie";
                tryb = 0;
                progressBar.Visible = true;

                if (logiCheck.Checked)
                {
                    dataTeraz = DateTime.Now;
                    plik_log = katalogDocelowy.Text + dataTeraz.ToString().Replace(":", "") + ".txt";
                }

                if (katalogBox.Text.EndsWith(".txt"))
                {
                    try
                    {
                        StreamReader reader = new StreamReader(katalogBox.Text);
                        string linia;

                        while ((linia = reader.ReadLine()) != null)
                        {
                            if (linia.EndsWith("\\"))
                            {
                                int koniec = linia.LastIndexOf("\\");
                                string katalog_glowny = linia.Substring(0, koniec);

                                DirSearch(katalog_glowny);
                            }
                            else DirSearch(linia);

                            // kasuj puste katalogi
                            czyscPuste(linia);
                        }
                    }
                    catch
                    {
                    }
                }
                else if (katalogBox.Text.EndsWith("\\"))
                {
                    int koniec = katalogBox.Text.LastIndexOf("\\");
                    string katalog_glowny = katalogBox.Text.Substring(0, koniec);
                    DirSearch(katalog_glowny);
                }
                else DirSearch(katalogBox.Text);

                // kasuj puste katalogi
                czyscPuste(katalogBox.Text);
                tabControl1.Enabled = true;
                this.Text = "Katalogi";
                progressBar.Visible = false;
                //MessageBox.Show("Gotowe");
            }
            else MessageBox.Show("Proszę wskazać katalog do naprawy oraz katalog docelowy");
        }

        public bool IsDirectoryEmpty(string path)
        {
            return !Directory.EnumerateFileSystemEntries(path).Any();
        }

        public void czyscPuste(string path)
        {
            if (Directory.Exists(path))
            {
                string[] katalogi = Directory.GetDirectories(path);
                string[] pliki = Directory.GetFiles(path);
                // jezeli wskazany katalog ma podkatalogi to rekurencja
                if (katalogi.Length > 0)
                {
                    // PODKATALOGI
                    foreach (string d in katalogi)
                    {
                        czyscPuste(d);
                    }
                    // sprawdz ponownie czy sa jakies katalogi
                    katalogi = Directory.GetDirectories(path);
                }
                // jezeli nie ma plikow ani katalogow to skasuj
                if (pliki.Length == 0 && katalogi.Length == 0)
                {
                    try
                    {
                        Directory.Delete(path);
                    }
                    catch
                    {
                    }

                }
            }

        }

        private void button3_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog dialog = new FolderBrowserDialog();
            dialog.ShowDialog();
            if (dialog.SelectedPath != "")
            {
                katalogDocelowy.Text = dialog.SelectedPath;
            }
        }

        private void ustawieniaToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

        private void wyjdźToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        string pobierzSerie(string nr_zespolu, string sygnatura)
        {
            string dane = null;
            string connectionString =
            "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=\""
            + baza_iza + "\";User Id=;Password=;";

            //MessageBox.Show(connectionString);

            // Provide the query string with a parameter placeholder.
            //SELECT SERIA FROM INWENTARZ WHERE NRZESPOLU=126 AND SYGNATURA="25M";
            string queryString =
                "SELECT seria from INWENTARZ "
                    + "WHERE NRZESPOLU = @zespolPole AND SYGNATURA = @sygnaturaPole;";

            // Create and open the connection in a using block. This
            // ensures that all resources will be closed and disposed
            // when the code exits.
            using (OleDbConnection connection =
                new OleDbConnection(connectionString))
            {
                // Create the Command and Parameter objects.
                OleDbCommand command = new OleDbCommand(queryString, connection);
                command.Parameters.AddWithValue("@zespolPole", nr_zespolu);
                command.Parameters.AddWithValue("@sygnaturaPole", sygnatura);

                // Open the connection in a try/catch block. 
                // Create and execute the DataReader, writing the result
                // set to the console window.
                try
                {
                    connection.Open();
                    OleDbDataReader reader = command.ExecuteReader();
                    if (reader.Read())
                    {
                        dane = "" + reader[0];
                        // jezeli jest wiecej rekordow zwroc -1
                        if (reader.Read()) dane = "-1";
                        else
                        {
                            if (dane.Equals("-")) dane = "0";
                            if (dane.Equals("")) dane = null;
                        }
                    }
                    reader.Close();

                    // oczyszczenie danych
                    if (dane != null)
                    {
                        if (dane.Contains(" ")) dane = dane.Replace(" ", "");
                        // do poprawy
                        if (dane.EndsWith(".")) dane = dane.Replace(".", "");
                    }

                    string podseria = pobierzPodserie(nr_zespolu, sygnatura);
                    if (podseria != null && podseria.Length > 0) dane = dane + "." + podseria;
                    return dane;
                }
                catch (Exception ex)
                {
                    //MessageBox.Show(""+ex.Message);
                }
                //Console.ReadLine();
            }
            return dane;
        }

        string pobierzPodserie(string nr_zespolu, string sygnatura)
        {
            string dane = null;
            string connectionString =
            "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=\""
            + baza_iza + "\";User Id=;Password=;";

            //MessageBox.Show(connectionString);

            // Provide the query string with a parameter placeholder.
            //SELECT SERIA FROM INWENTARZ WHERE NRZESPOLU=126 AND SYGNATURA="25M";
            string queryString =
                "SELECT podseria from INWENTARZ "
                    + "WHERE NRZESPOLU = @zespolPole AND SYGNATURA = @sygnaturaPole;";

            // Create and open the connection in a using block. This
            // ensures that all resources will be closed and disposed
            // when the code exits.
            using (OleDbConnection connection =
                new OleDbConnection(connectionString))
            {
                // Create the Command and Parameter objects.
                OleDbCommand command = new OleDbCommand(queryString, connection);
                command.Parameters.AddWithValue("@zespolPole", nr_zespolu);
                command.Parameters.AddWithValue("@sygnaturaPole", sygnatura);

                // Open the connection in a try/catch block. 
                // Create and execute the DataReader, writing the result
                // set to the console window.
                try
                {
                    connection.Open();
                    OleDbDataReader reader = command.ExecuteReader();
                    if (reader.Read())
                    {
                        dane = "" + reader[0];
                        // jezeli jest wiecej rekordow zwroc -1
                        if (reader.Read()) dane = "-1";
                        else
                        {
                            if (dane.Equals("-")) dane = "0";
                            if (dane.Equals("")) dane = null;
                        }
                    }
                    reader.Close();

                    // kasowanie poczatkowych zer itd
                    if (dane != null)
                    {
                        while (dane.StartsWith("0"))
                        {
                            dane = ZERO.Replace(dane, "", 1);
                        }
                        if (dane.Contains(" ")) dane = dane.Replace(" ", "");
                        // do poprawy
                        if (dane.EndsWith(".")) dane = dane.Replace(".", "");
                    }

                    return dane;
                }
                catch (Exception ex)
                {
                    //MessageBox.Show(""+ex.Message);
                }
                //Console.ReadLine();
            }
            return dane;
        }

        int pobierzStrony(string nr_zespolu, string seria, string sygnatura)
        {

            string dane = null;
            int stron = 0;
            string connectionString =
            "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=\"" + baza_iza + "\";User Id=;Password=;";

            // jezeli seria posiada . to pobierz wartosc po lewej stronie
            if (SERIA_REGEX.Match(seria).Success)
            {
                Match match = SERIA_REGEX.Match(seria);
                seria = match.Groups[1].Value;
            }

            //MessageBox.Show(connectionString);

            // Provide the query string with a parameter placeholder.
            //SELECT SERIA FROM INWENTARZ WHERE NRZESPOLU=126 AND SYGNATURA="25M";
            string queryString =
                "SELECT strony from INWENTARZ "
                    + "WHERE NRZESPOLU = @zespolPole AND SERIA = @seriaPole AND SYGNATURA = @sygnaturaPole;";

            // Create and open the connection in a using block. This
            // ensures that all resources will be closed and disposed
            // when the code exits.
            using (OleDbConnection connection =
                new OleDbConnection(connectionString))
            {
                // Create the Command and Parameter objects.
                OleDbCommand command = new OleDbCommand(queryString, connection);
                command.Parameters.AddWithValue("@zespolPole", nr_zespolu);
                command.Parameters.AddWithValue("@seriaPole", seria);
                command.Parameters.AddWithValue("@sygnaturaPole", sygnatura);

                // Open the connection in a try/catch block. 
                // Create and execute the DataReader, writing the result
                // set to the console window.
                try
                {
                    connection.Open();
                    OleDbDataReader reader = command.ExecuteReader();
                    if (reader.Read())
                    {
                        dane = "" + reader[0];
                        Match match = STRONY.Match(dane);

                        if (match != null && match.Success) stron = int.Parse(match.Groups[1].Value);
                    }
                }
                catch (Exception ex)
                {
                    //MessageBox.Show(""+ex.Message);
                }
                //Console.ReadLine();
            }
            return stron;
        }


        void publikujSkany(string sciezka_skany, string sciezka_docelowa, string baza_danych)
        {
            string katalog_zrodlowy = null;
            string connectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=\"" + baza_danych + "\";User Id=;Password=;";

            // Provide the query string with a parameter placeholder.
            string queryString = "SELECT NRAP, NRZESPOLU, CDNUMERU, SERIA, SYGNATURA FROM OPIS WHERE STATUS_UDOSTEPNIENIA=2 AND TYTUL Is Not Null;";


            // Create and open the connection in a using block. This
            // ensures that all resources will be closed and disposed
            // when the code exits.
            using (OleDbConnection connection =
                new OleDbConnection(connectionString))
            {
                // Create the Command and Parameter objects.
                OleDbCommand command = new OleDbCommand(queryString, connection);

                // Open the connection in a try/catch block. 
                // Create and execute the DataReader, writing the result
                // set to the console window.
                try
                {
                    connection.Open();
                    OleDbDataReader reader = command.ExecuteReader();
                    while (reader.Read())
                    {
                        katalog_zrodlowy = sciezka_skany + reader[0] + "\\" + reader[1] + "\\" + reader[2] + "\\" + reader[3] + "\\" + reader[4] + "\\";
                        DirectoryCopy(katalog_zrodlowy, sciezka_docelowa);
                    }
                    reader.Close();
                }
                catch (Exception ex)
                {
                    bledyKopiowanie.AppendText(ex.Message);
                }
                //Console.ReadLine();
            }
        }

        private void importujButton_Click(object sender, EventArgs e)
        {
            if (skanySciezka.Text.Length != 0 && bazaSciezka.Text.Length != 0 && listaOsoby.Items.Count > 0 && listaFinansowanie.Items.Count > 0)
            {
                tabControl1.Enabled = false;
                this.Text = "Katalogi - importowanie";
                tryb = 1;
                progressBar.Visible = true;
                osoba_id = listaOsoby.SelectedIndex + 1;
                finansowanie_id = listaFinansowanie.SelectedIndex + 1;

                if (skanySciezka.Text.EndsWith("\\"))
                {
                    int koniec = skanySciezka.Text.LastIndexOf("\\");
                    string katalog_glowny = skanySciezka.Text.Substring(0, koniec);
                    DirSearch(katalog_glowny);
                }
                else DirSearch(skanySciezka.Text);

                tabControl1.Enabled = true;
                this.Text = "Katalogi";
                progressBar.Visible = false;
                //MessageBox.Show("Gotowe");
            }
            else MessageBox.Show("Proszę wskazać katalog ze skanami oraz poprawną bazę danych");

        }

        private void katalogSkanyButton_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog dialog = new FolderBrowserDialog();
            dialog.ShowDialog();
            if (dialog.SelectedPath != "")
            {
                skanySciezka.Text = dialog.SelectedPath;
            }
        }

        private void bazaButton_Click(object sender, EventArgs e)
        {
            OpenFileDialog dialog = new OpenFileDialog();
            dialog.Filter = "Baza danych Microsoft Access|*.mdb";
            dialog.Title = "Wskaż plik bazy danych";
            DialogResult result = dialog.ShowDialog();
            if (result == DialogResult.OK) // Test result.
            {
                bazaSciezka.Text = dialog.FileName;
                uzupelnijFormatki();
            }
        }

        public void importuj(string plik)
        {
            Skan skan = new Skan(plik);

            if (skan.PobierzPlik() != null)
            {
                if (skan.rozszerzenie.Equals("tif") || skan.rozszerzenie.Equals("tiff"))
                {
                    var dane = getImageInfo(plik);
                    FileInfo f = new FileInfo(plik);
                    string format = "yyyy-MM-dd HH:mm:ss";
                    string dpi = null;
                    string szerokosc = null;
                    string wysokosc = null;
                    DateTime data_utworzenia = f.CreationTime;
                    DateTime data_modyfikacji = f.LastWriteTime;
                    string rozmiar = f.Length.ToString();

                    if (dane.ContainsKey("XResolution")) dpi = (string)dane["XResolution"];
                    if (dane.ContainsKey("ImageWidth")) szerokosc = (string)dane["ImageWidth"];
                    if (dane.ContainsKey("ImageHeight")) wysokosc = (string)dane["ImageHeight"];

                    try
                    {
                        if (dane.ContainsKey("ModifyDate")) data_utworzenia = DateTime.ParseExact((string)dane["ModifyDate"], format, CultureInfo.InvariantCulture);
                        else if (DateTime.Compare(data_modyfikacji, data_utworzenia) < 0) data_utworzenia = data_modyfikacji;
                    }
                    catch { }



                    // jezeli inna seria to dodaj wpis
                    if (nr_zespolu_old != skan.nr_zespolu || cd_zespolu_old != skan.cd_zespolu || seria_old2 != skan.seria || sygnatura_old != skan.sygnatura)
                    {
                        dodajRekordOpis(skan, Convert.ToString(osoba_id), Convert.ToString(finansowanie_id));
                    }

                    //MessageBox.Show("dpi: " + dpi + " data utworzenia:" + data_utworzenia + " szerokosc:" + szerokosc + " wysokosc:" + wysokosc +" rozmiar: "+rozmiar);

                    dodajRekordSkany(skan, dpi, rozmiar, szerokosc, wysokosc, data_utworzenia.ToString(format));

                    nr_zespolu_old = skan.nr_zespolu;
                    cd_zespolu_old = skan.cd_zespolu;
                    seria_old2 = skan.seria;
                    sygnatura_old = skan.sygnatura;
                }

            }
        }

        public void dodajRekordOpis(Skan skan, string id_osoby, string finansowanie)
        {
            // otwieranie bazy danych
            string connectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=\"" + bazaSciezka.Text + "\";User Id=;Password=;";

            // Create and open the connection in a using block. This
            // ensures that all resources will be closed and disposed
            // when the code exits.
            using (OleDbConnection connection = new OleDbConnection(connectionString))
            {
                // Open the connection in a try/catch block. 
                // Create and execute the DataReader, writing the result
                // set to the console window.
                try
                {
                    OleDbCommand cmd = new OleDbCommand();
                    cmd.CommandType = CommandType.Text;
                    cmd.CommandText = "insert into OPIS ([NRZESPOLU], [CDNUMERU], [SERIA], [SYGNATURA], [ID_OSOBY], [FINANSOWANIE]) values (?,?,?,?,?,?)";
                    cmd.Parameters.AddWithValue("@nr_zespolu", Convert.ToDecimal(skan.nr_zespolu));
                    cmd.Parameters.AddWithValue("@cd_zespolu", Convert.ToDecimal(skan.cd_zespolu));
                    cmd.Parameters.AddWithValue("@seria", skan.seria);
                    cmd.Parameters.AddWithValue("@sygnatura", skan.sygnatura);
                    cmd.Parameters.AddWithValue("@id_osoby", Convert.ToDecimal(id_osoby));
                    cmd.Parameters.AddWithValue("@finansowanie", Convert.ToDecimal(finansowanie));
                    cmd.Connection = connection;
                    connection.Open();
                    cmd.ExecuteNonQuery();
                }
                catch (Exception ex)
                {
                    //MessageBox.Show(nr_zespolu +" "+cd_zespolu+" "+seria+" "+sygnatura+" "+id_osoby+" "+finansowanie+" "+ex.Message);
                }
                //Console.ReadLine();
            }
        }

        public void dodajRekordSkany(Skan skan, string dpi, string wiekosc, string szerokosc, string wysokosc, string data_utworzenia)
        {

            // otwieranie bazy danych
            string connectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=\"" + bazaSciezka.Text + "\";User Id=;Password=;";
            if (dpi.Contains("."))
            {
                int index = dpi.IndexOf(".");
                dpi = dpi.Substring(0, index);
            }

            // Create and open the connection in a using block. This
            // ensures that all resources will be closed and disposed
            // when the code exits.
            using (OleDbConnection connection = new OleDbConnection(connectionString))
            {
                // Open the connection in a try/catch block. 
                // Create and execute the DataReader, writing the result
                // set to the console window.
                try
                {
                    OleDbCommand cmd = new OleDbCommand();
                    cmd.CommandType = CommandType.Text;
                    cmd.CommandText = "insert into SKANY ([NRAP],[NRZESPOLU],[CDNUMERU],[SERIA],[SYGNATURA],[STRONA],[DPI],[FORMAT],[WIELKOSC],[SZEROKOSC],[WYSOKOSC],[DATA_UTWORZENIA]) values (?,?,?,?,?,?,?,?,?,?,?,?)";
                    cmd.Parameters.AddWithValue("@nr_archiwum", Convert.ToDecimal(skan.nr_archiwum));
                    cmd.Parameters.AddWithValue("@nr_zespolu", Convert.ToDecimal(skan.nr_zespolu));
                    cmd.Parameters.AddWithValue("@cd_zespolu", Convert.ToDecimal(skan.cd_zespolu));
                    cmd.Parameters.AddWithValue("@seria", skan.seria);
                    cmd.Parameters.AddWithValue("@sygnatura", skan.sygnatura);
                    cmd.Parameters.AddWithValue("@strona", skan.strona);
                    cmd.Parameters.AddWithValue("@dpi", Convert.ToInt64(dpi));
                    cmd.Parameters.AddWithValue("@format", skan.rozszerzenie);
                    cmd.Parameters.AddWithValue("@wiekosc", Convert.ToInt64(wiekosc));
                    cmd.Parameters.AddWithValue("@szerokosc", Convert.ToDecimal(szerokosc));
                    cmd.Parameters.AddWithValue("@wysokosc", Convert.ToDecimal(wysokosc));
                    cmd.Parameters.AddWithValue("@data_utworzenia", data_utworzenia);

                    cmd.Connection = connection;
                    connection.Open();
                    cmd.ExecuteNonQuery();
                }
                catch (Exception ex)
                {
                    //MessageBox.Show("nr_archiwum " + nr_archiwum + " nr_zespolu " + nr_zespolu + " cd_zespolu " + cd_zespolu + " seria " + seria + " sygnatura " + sygnatura + " strona " + strona + " dpi " + dpi + " format " + format + " wielkosc " + wiekosc + " szerokosc " + szerokosc + " wysokosc " + wysokosc + " data " + data_utworzenia);
                    //MessageBox.Show("" + ex.Message);
                }
                //Console.ReadLine();
            }
        }

        private void startButton_Click(object sender, EventArgs e)
        {
            bledyKopiowanie.Clear();
            if (publikacjaSkany.Text.Length != 0 && publikacjaDocelowy.Text.Length != 0 && publikacjaBaza.Text.Length != 0)
            {
                tabControl1.Enabled = false;
                this.Text = "Katalogi - przygotowanie skanów do publikacji";
                progressBar.Visible = true;

                if (publikacjaSkany.Text.EndsWith("\\") == false) publikacjaSkany.Text = publikacjaSkany.Text + "\\";
                if (publikacjaDocelowy.Text.EndsWith("\\") == false) publikacjaDocelowy.Text = publikacjaDocelowy.Text + "\\";
                publikujSkany(publikacjaSkany.Text, publikacjaDocelowy.Text, publikacjaBaza.Text);

                this.Text = "Katalogi";
                tabControl1.Enabled = true;
                progressBar.Visible = false;
                //MessageBox.Show("Gotowe");
            }
            else MessageBox.Show("Proszę wskazać katalog ze skanami oraz bazę danych SKANY");
        }

        private void DirectoryCopy(string sourceDirName, string destDirName)
        {
            // Get the subdirectories for the specified directory.
            DirectoryInfo dir = new DirectoryInfo(sourceDirName);

            if (!dir.Exists)
            {
                bledyKopiowanie.AppendText("Katalog źródłowy nie istnieje bądź nie może zostać znaleziony: " + sourceDirName + "\n");
            }

            DirectoryInfo[] dirs = dir.GetDirectories();
            // If the destination directory doesn't exist, create it.
            if (!Directory.Exists(destDirName))
            {
                Directory.CreateDirectory(destDirName);
            }

            // Get the files in the directory and copy them to the new location.
            FileInfo[] files = dir.GetFiles();
            int licznik = 1;
            foreach (FileInfo file in files)
            {
                int procent = (int)(0.5f + ((100f * licznik++) / files.Length));
                // Report progress.
                progressBar.Value = procent;
                Application.DoEvents();
                System.Threading.Thread.Sleep(1);

                string temppath = Path.Combine(destDirName, file.Name);
                file.CopyTo(temppath, false);
            }
        }

        private void p_KatalogButton_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog dialog = new FolderBrowserDialog();
            dialog.ShowDialog();
            if (dialog.SelectedPath != "")
            {
                publikacjaSkany.Text = dialog.SelectedPath;
            }
        }

        private void p_DocelowyButton_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog dialog = new FolderBrowserDialog();
            dialog.ShowDialog();
            if (dialog.SelectedPath != "")
            {
                publikacjaDocelowy.Text = dialog.SelectedPath;
            }
        }

        private void p_BazaButton_Click(object sender, EventArgs e)
        {
            OpenFileDialog dialog = new OpenFileDialog();
            dialog.Filter = "Baza danych Microsoft Access|*.mdb";
            dialog.Title = "Wskaż plik bazy danych";
            DialogResult result = dialog.ShowDialog();
            if (result == DialogResult.OK) // Test result.
            {
                publikacjaBaza.Text = dialog.FileName;
            }
        }

        private void katalogWeryfikacjaWskaz_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog dialog = new FolderBrowserDialog();
            dialog.ShowDialog();
            if (dialog.SelectedPath != "")
            {
                skanyWeryfikacja.Text = dialog.SelectedPath;
            }
        }

        private void weryfikacjaStart_Click(object sender, EventArgs e)
        {
            if (skanyWeryfikacja.Text.Length != 0)
            {
                weryfikacjaInfo.Clear();
                czyKasowac = null;
                czyWszystkieMarginesy = null;
                tabControl1.Enabled = false;
                this.Text = "Katalogi - weryfikacja skanów";
                tryb = 2;

                progressBar.Visible = true;

                try
                {
                    ilosc_okladek = int.Parse(iloscOkladek.Text);
                }
                catch (Exception ex) { }

                if (skanyWeryfikacja.Text.EndsWith("\\"))
                {
                    int koniec = skanyWeryfikacja.Text.LastIndexOf("\\");
                    string katalog_glowny = skanyWeryfikacja.Text.Substring(0, koniec);
                    DirSearch(katalog_glowny);
                }
                else DirSearch(skanyWeryfikacja.Text);

                this.Text = "Katalogi";
                tabControl1.Enabled = true;
                progressBar.Visible = false;
            }
            else MessageBox.Show("Proszę wskazać katalog ze skanami do weryfikacji");
        }

        public void weryfikuj(string plik, bool last)
        {
            //MessageBox.Show(plik);
            string strona_temp;
            Skan skan = new Skan(plik);

            if (skan.PobierzPlik() != null)
            {
                string sygnatura_new = skan.nr_archiwum + "_" + skan.nr_zespolu + "_" + skan.cd_zespolu + "_" + skan.seria + "_" + skan.sygnatura;

                // sprawdzenie czy plik ma dpi wieksze lub rowne 300
                if (checkDPI.Checked && (skan.rozszerzenie.Equals("tif") || skan.rozszerzenie.Equals("tiff") || skan.rozszerzenie.Equals("jpg") || skan.rozszerzenie.Equals("jpeg")))
                {
                    try
                    {
                        var dane = getImageInfo(plik);
                        int dpi = 0;
                        if (dane.ContainsKey("XResolution")) dpi = Convert.ToInt32((string)dane["XResolution"]);

                        if (dpi > 0 && dpi < 300) weryfikacjaInfo.AppendText("- skan " + plik + " posiada za małą rozdzielczość: " + Convert.ToString(dpi) + " DPI\n");
                        else if (dpi > 0 && dpi % 100 > 0) weryfikacjaInfo.AppendText("- skan " + plik + " posiada nieprawidłową rozdzielczość: " + Convert.ToString(dpi) + " DPI\n");

                        string kompresja = null;
                        if (dane.ContainsKey("Compression")) kompresja = (string)dane["Compression"];
                        if (kompresja != null && kompresja.Equals("") == false && kompresja.Contains("Uncompressed") == false)
                        {
                            weryfikacjaInfo.AppendText("- skan " + plik + " jest skompresowany przy wykorzystaniu: " + kompresja + "\n");
                        }
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message);
                    }
                }


                // jeżeli w katalou są skany z różnych sygnatur to poinformuj
                if (sygnatura_old.Length > 0 && sygnatura_old.Equals(sygnatura_new) == false)
                {
                    weryfikacjaInfo.AppendText("- skany z różnych sygnatur: " + sygnatura_old + " oraz " + sygnatura_new + "\n");
                }
                else
                {
                    // jezeli skany pochodza z tej samej sygnatury to zweryfikuj paginacje stron (puste strony)
                    //if (checkBlank.Checked)
                    //{
                    if (skan.strona_start == 0)
                    {
                        weryfikacjaInfo.AppendText("- nieprawidłowy licznik dla skanu: " + plik + "\n");
                    }
                    else if (skan.strona_start != -1)
                    {
                        // sprawdzenie czy są zera wiodące
                        if (skan.strona.Length < 4 && skan.strona.StartsWith("0") == false) weryfikacjaInfo.AppendText("- licznik bez zer wiodących: " + plik + "\n");

                        int roznica = skan.strona_start - strona_old;

                        if (skan.strona_start < 9999 && roznica > 1)
                        {
                            for (int i = 1; i < roznica; i++)
                            {
                                strona_old = strona_old + 1;

                                if (checkBlank.Checked)
                                {
                                    string pusta_strona = Convert.ToString(strona_old);
                                    // uzupełnianie 0 przed cyfrą
                                    pusta_strona = dopelnij(pusta_strona, skan.strona);

                                    string nowy_plik = skan.nr_archiwum + "_" + skan.nr_zespolu + "_" + skan.cd_zespolu + "_" + skan.seria + "_" + skan.sygnatura + "_" + pusta_strona + ".txt";
                                    File.Create(System.IO.Path.Combine(System.IO.Path.GetDirectoryName(plik), nowy_plik)).Dispose();
                                    weryfikacjaInfo.AppendText("- dodano pustą stronę: " + strona_old + "\n");
                                }
                                else weryfikacjaInfo.AppendText("- brakuje strony: " + strona_old + "\n");
                            }
                        }
                        if (skan.strona_end != -1)
                        {
                            strona_old = skan.strona_end;
                            int test = skan.strona_end - skan.strona_start;
                            if (test > 1) weryfikacjaInfo.AppendText("- według licznika plik " + plik + " zawiera " + (test + 1) + " skanów\n");
                        }
                        else strona_old = skan.strona_start;
                    }
                    //}
                }
                sygnatura_old = sygnatura_new;

                // sprawdzanie czy plik nie jest pusty
                if (skan.rozszerzenie.Equals("txt") == false)
                {
                    FileInfo f = new FileInfo(plik);
                    if (f.Length == 0) weryfikacjaInfo.AppendText("- plik " + plik + " jest pusty\n");
                    else
                    {
                        if (szukajPuste.Checked && skan.rozszerzenie.Equals("tif"))
                        {
                            int marginesySkan = czyZlySkan(plik, true);
                            switch (marginesySkan)
                            {
                                case -1:
                                    weryfikacjaInfo.AppendText("- " + plik + ": nieudało się zweryfikować marginesów\n");
                                    break;
                                case 1:
                                    weryfikacjaInfo.AppendText("- " + plik + ": zbyt mały margines\n");
                                    break;
                                case 2:
                                    weryfikacjaInfo.AppendText("- " + plik + ": zbyt duży margines lub źle zeskanowany dokument\n");
                                    // jezeli pierwsze wyswietlanie komunikatu
                                    /*if (czyKasowac == null)
                                    {
                                        DialogResult dialogResult = MessageBox.Show("Czy skasować źle wykonane skany ?", "Pytanie", MessageBoxButtons.YesNo);
                                        if (dialogResult == DialogResult.Yes)
                                        {
                                            czyKasowac = "tak";
                                        }
                                        else if (dialogResult == DialogResult.No)
                                        {
                                            czyKasowac = "nie";
                                        }
                                    }
                                    // kasowanie zle wykonanych skanów
                                    if (czyKasowac.Equals("tak")) System.IO.File.Delete(plik); */

                                    // czy naprawic marginesy dla wszystkich zlych
                                    if (fixMargins.Checked)
                                    {
                                        /*if (czyWszystkieMarginesy == null)
                                        {
                                            DialogResult dialogResult = MessageBox.Show("Czy naprawiać marginesy dla wszystkich skanów ?", "Pytanie", MessageBoxButtons.YesNo);
                                            if (dialogResult == DialogResult.Yes)
                                            {
                                                czyWszystkieMarginesy = "tak";
                                            }
                                            else if (dialogResult == DialogResult.No)
                                            {
                                                czyWszystkieMarginesy = "nie";
                                            }
                                        }*/

                                        string czulosc = "10";
                                        int marginesy = 30;
                                        if (NUMER_STRONY.IsMatch(czuloscBox.Text)) czulosc = czuloscBox.Text;
                                        if (NUMER_STRONY.IsMatch(marginesyBox.Text)) marginesy = int.Parse(marginesyBox.Text);

                                        naprawMarginesy(plik, czulosc, marginesy);
                                        //if (czyWszystkieMarginesy.Equals("tak")) naprawMarginesy(plik, czulosc, marginesy);
                                        //else if (strona_start != -1 && (strona_start <= ilosc_okladek || strona.EndsWith("okladka1") || strona.EndsWith("okladka"))) naprawMarginesy(plik, czulosc, marginesy);

                                    }

                                    break;
                            }
                        }
                    }
                }


                if (checkWzorniki.Checked && System.IO.File.Exists(plik))
                {
                    if (wzornik == false)
                    {
                        if (last) weryfikacjaInfo.AppendText("- jednostka: " + skan.sygnatura + " prawdopodobnie nie posiada wzornika\n");
                        else if (skan.strona.Contains("okladka") || skan.strona.EndsWith("001") || skan.strona.EndsWith("002") || skan.strona.Contains("001_0"))
                        {
                            wzornik = sprawdzWzorniki(plik);
                        }
                    }
                }


                // weryfikacja nazewnictwa okladek i pozostalych plikow
                if (System.IO.File.Exists(plik) && skan.strona_start != -1)
                {
                    string nowy_plik = null;

                    // zachowanie dla plikow ktore maja byc okladka
                    if (skan.strona_start < ilosc_okladek)
                    {
                        strona_temp = "0000";

                        if (skan.strona_start == 0) nowy_plik = skan.nr_archiwum + "_" + skan.nr_zespolu + "_" + skan.cd_zespolu + "_" + skan.seria + "_" + skan.sygnatura + "_" + strona_temp + "_okladka." + skan.rozszerzenie;
                        else nowy_plik = skan.nr_archiwum + "_" + skan.nr_zespolu + "_" + skan.cd_zespolu + "_" + skan.seria + "_" + skan.sygnatura + "_" + strona_temp + "_okladka" + skan.strona_start + "." + skan.rozszerzenie;
                        nowy_plik = System.IO.Path.Combine(System.IO.Path.GetDirectoryName(plik), nowy_plik);

                        // jezeli dokonano jakiejs zmiany na pliku i plik docelowy nie istnieje
                        if (skan.rozszerzenie.Equals("txt") == false && nowy_plik.Equals(plik) == false && System.IO.File.Exists(nowy_plik) == false)
                        {
                            weryfikacjaInfo.AppendText("- poprawiono nazwę pliku dla okładki: " + nowy_plik + "\n");
                            System.IO.File.Move(plik, nowy_plik);
                        }
                    }

                    // pozostale pliki
                    /*else
                    {
                        strona_temp = Convert.ToString(strona_start - (ilosc_okladek - 1));
                        string strona_temp_end = null;
                        strona_temp = dopelnij(strona_temp, match.Groups[1].Value);

                        if (strona_end != -1)
                        {
                            strona_temp_end = Convert.ToString(strona_end - (ilosc_okladek - 1));
                            strona_temp_end = dopelnij(strona_temp_end, match.Groups[1].Value);
                            nowy_plik = nr_archiwum + "_" + nr_zespolu + "_" + cd_zespolu + "_" + seria + "_" + sygnatura + "_" + strona_temp + "_" + strona_temp_end + "." + rozszerzenie;

                            strona_temp = strona_temp_end;
                        }
                        else nowy_plik = nr_archiwum + "_" + nr_zespolu + "_" + cd_zespolu + "_" + seria + "_" + sygnatura + "_" + strona_temp + "." + rozszerzenie;
                    }*/

                    // sprawdz czy ostatnia strona skanu odpowiada ilosci skanow w inwentarzu
                    if (last && checkLast.Checked && baza_iza.EndsWith(".mdb"))
                    {
                        int stron_iza = pobierzStrony(skan.nr_zespolu, skan.seria, skan.sygnatura);
                        int strona_last = 0;

                        // jezeli ostatni plik istnieje to pobierz z niego strone inaczej pobierz z przedostatniego
                        if (System.IO.File.Exists(plik)) strona_last = skan.strona_start;
                        //else strona_last = int.Parse(strona_temp);

                        if (stron_iza > 0 && strona_last > 0)
                        {
                            int roznica = stron_iza - strona_last;
                            if (roznica > 0)
                            {
                                weryfikacjaInfo.AppendText("- według paginacji brakuje stron: " + roznica + " z " + stron_iza + "\n");

                                if (checkBlank.Checked)
                                {
                                    for (int i = 1; i <= roznica; i++)
                                    {
                                        string pusta_strona = Convert.ToString(strona_last + i);
                                        // uzupełnianie 0 przed cyfrą
                                        pusta_strona = dopelnij(pusta_strona, skan.strona);

                                        nowy_plik = skan.nr_archiwum + "_" + skan.nr_zespolu + "_" + skan.cd_zespolu + "_" + skan.seria + "_" + skan.sygnatura + "_" + pusta_strona + ".txt";
                                        File.Create(System.IO.Path.Combine(System.IO.Path.GetDirectoryName(plik), nowy_plik)).Dispose();
                                        weryfikacjaInfo.AppendText("- dodano pustą stronę: " + (strona_last + i) + "\n");
                                    }
                                }
                                //else weryfikacjaInfo.AppendText("- brakuje strony: " + (strona_last + i) + "\n");
                            }
                            else if (roznica < 0) weryfikacjaInfo.AppendText("- według paginacji jednostka ma za dużo skanów o: " + roznica * -1 + "\n");
                        }
                    }
                }
            }
            else weryfikacjaInfo.AppendText("- plik " + plik + " nie pasuje do wzorca\n");
        }

        public int czyZlySkan(string plik, bool brightness)
        {
            string mean = null;

            ProcessStartInfo startInfo = new ProcessStartInfo();
            startInfo.FileName = convert;
            if (brightness) startInfo.Arguments = "\"" + plik + "\"" + " -resize \"1000x1000>\" -brightness-contrast 40x100 -blur 0x1 -bordercolor black -border 1x1 -fuzz 50% -trim -quiet info:";
            else startInfo.Arguments = "\"" + plik + "\"" + " -resize \"1000x1000>\" -brightness-contrast 0x100 -blur 0x1 -bordercolor black -border 1x1 -fuzz 50% -trim -quiet info:";
            startInfo.CreateNoWindow = true;
            startInfo.WindowStyle = ProcessWindowStyle.Hidden;
            startInfo.ErrorDialog = false;
            startInfo.UseShellExecute = false;
            startInfo.RedirectStandardOutput = true;
            startInfo.RedirectStandardError = false;

            try
            {
                using (Process exeProcess = Process.Start(startInfo))
                {
                    //exeProcess.PriorityClass = ProcessPriorityClass.AboveNormal;
                    mean = exeProcess.StandardOutput.ReadToEnd();
                    exeProcess.WaitForExit();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Błąd podczas weryfikacji zawartości skanu " + plik + ", " + ex.Message, "Błąd");
            }

            if (mean != null)
            {
                Match match = DANE.Match(mean);
                if (match.Success)
                {
                    int x1 = Convert.ToInt32(match.Groups[1].Value);
                    int y1 = Convert.ToInt32(match.Groups[2].Value);
                    int x2 = Convert.ToInt32(match.Groups[3].Value);
                    int y2 = Convert.ToInt32(match.Groups[4].Value);
                    int roznica = ((x2 - x1) + (y2 - y1));

                    //MessageBox.Show(plik + ": " + roznica);

                    // brak marginesu
                    if (roznica < 40)
                    {
                        // jezeli to test bez rozjasniania to wynik jest pewny
                        if (brightness == false) return 1;
                        // sprawdz czy bez rozjasniania tez brak marginesow
                        else
                        {
                            int test = czyZlySkan(plik, false);
                            return test;
                        }
                    }
                    // za duzy margines
                    else if (roznica >= 300) return 2;
                    // skan ok
                    else return 0;
                }
            }

            return -1;
        }

        public void naprawMarginesy(string plik, string procent, int margines)
        {
            double offset = Math.Round((margines * 300) / 25.4, 2);
            string filename = convert;
            string output = null;
            if (File.Exists(convert)) filename = convert;

            ProcessStartInfo startInfo = new ProcessStartInfo();
            startInfo.FileName = filename;
            startInfo.Arguments = "\"" + plik + "\"" + " -bordercolor black -border 1x1 -fuzz " + procent + "% -trim info:";
            startInfo.CreateNoWindow = true;
            startInfo.WindowStyle = ProcessWindowStyle.Hidden;
            startInfo.ErrorDialog = false;
            startInfo.UseShellExecute = false;
            startInfo.RedirectStandardOutput = true;
            startInfo.RedirectStandardError = true;

            try
            {
                using (Process exeProcess = Process.Start(startInfo))
                {
                    exeProcess.PriorityClass = ProcessPriorityClass.AboveNormal;
                    output = exeProcess.StandardOutput.ReadToEnd();
                    exeProcess.WaitForExit();
                }
            }
            catch
            {
                MessageBox.Show("Błąd podczas naprawy marginesów dla " + plik, "Błąd");
            }

            if (output != null)
            {
                Match match = DANE.Match(output);
                if (match != null && match.Success)
                {
                    int szerokosc = int.Parse(match.Groups[1].Value) + Convert.ToInt32(offset);
                    int wysokosc = int.Parse(match.Groups[2].Value) + Convert.ToInt32(offset);
                    int offset_x = int.Parse(match.Groups[5].Value) - Convert.ToInt32(Math.Round(offset / 2));
                    int offset_y = int.Parse(match.Groups[6].Value) - Convert.ToInt32(Math.Round(offset / 2));

                    string offset_X = (offset_x >= 0) ? "+" + Convert.ToString(offset_x) : Convert.ToString(offset_x);
                    string offset_Y = (offset_y >= 0) ? "+" + Convert.ToString(offset_y) : Convert.ToString(offset_y);

                    startInfo = new ProcessStartInfo();
                    startInfo.FileName = convert;
                    startInfo.Arguments = "\"" + plik + "\"" + " -crop " + Convert.ToString(szerokosc) + "x" + Convert.ToString(wysokosc) + offset_X + offset_Y + " " + "\"" + plik + "\"";
                    startInfo.CreateNoWindow = true;
                    startInfo.WindowStyle = ProcessWindowStyle.Hidden;
                    startInfo.ErrorDialog = false;
                    startInfo.UseShellExecute = false;
                    startInfo.RedirectStandardOutput = true;
                    startInfo.RedirectStandardError = true;

                    try
                    {
                        using (Process exeProcess = Process.Start(startInfo))
                        {
                            exeProcess.PriorityClass = ProcessPriorityClass.AboveNormal;
                            //output = exeProcess.StandardOutput.ReadToEnd();
                            exeProcess.WaitForExit();
                        }
                    }
                    catch
                    {
                        MessageBox.Show("Błąd podczas naprawy marginesów dla " + plik, "Błąd");
                    }
                }
            }

        }

        private void fixMargins_CheckedChanged(object sender, EventArgs e)
        {
            if (fixMargins.Checked) groupParametry.Enabled = true;
            else groupParametry.Enabled = false;
        }

        private void buttonSort_Click(object sender, EventArgs e)
        {
            if (skanyWeryfikacja.Text.Length != 0)
            {
                tabControl1.Enabled = false;
                this.Text = "Katalogi - sortowanie";
                tryb = 3;
                progressBar.Visible = true;

                if (skanyWeryfikacja.Text.EndsWith("\\"))
                {
                    int koniec = skanyWeryfikacja.Text.LastIndexOf("\\");
                    string katalog_glowny = skanyWeryfikacja.Text.Substring(0, koniec);
                    DirSearch(katalog_glowny);
                }
                else DirSearch(skanyWeryfikacja.Text);

                tabControl1.Enabled = true;
                this.Text = "Katalogi";
                progressBar.Visible = false;
                //MessageBox.Show("Skany zostały ponumerowane");
            }
            else MessageBox.Show("Proszę wskazać katalog ze skanami do weryfikacji");
        }

        private void sortuj(string plik)
        {
            Skan skan = new Skan(plik);

            if (skan.PobierzPlik() != null)
            {
                string sygnatura_new = skan.nr_archiwum + "_" + skan.nr_zespolu + "_" + skan.cd_zespolu + "_" + skan.seria + "_" + skan.sygnatura;

                // jeżeli w katalou są skany z różnych sygnatur to poinformuj
                if (sygnatura_old.Length > 0 && sygnatura_old.Equals(sygnatura_new) == false)
                {
                    weryfikacjaInfo.AppendText("- skany z różnych sygnatur: " + sygnatura_old + " oraz " + sygnatura_new + "\n");
                }
                else
                {
                    // jezeli skany pochodza z tej samej sygnatury to zweryfikuj paginacje stron (sortowanie)
                    if (skan.strona_start != -1)
                    {
                        int roznica = skan.strona_start - strona_old;
                        if (roznica > 1 || roznica == 0)
                        {
                            strona_old++;
                            string strona_temp = Convert.ToString(strona_old);
                            string strona_temp_end = null;
                            // uzupełnianie 0 przed cyfrą
                            strona_temp = dopelnij(strona_temp, skan.strona);
                            if (skan.strona_end != -1)
                            {
                                roznica = skan.strona_end - skan.strona_start;
                                strona_old = strona_old + roznica;
                                strona_temp_end = Convert.ToString(strona_old);
                            }

                            string nowy_plik = skan.nr_archiwum + "_" + skan.nr_zespolu + "_" + skan.cd_zespolu + "_" + skan.seria + "_" + skan.sygnatura + "_" + strona_temp + "." + skan.rozszerzenie;
                            if (strona_temp_end != null) nowy_plik = skan.nr_archiwum + "_" + skan.nr_zespolu + "_" + skan.cd_zespolu + "_" + skan.seria + "_" + skan.sygnatura + "_" + strona_temp + "_" + strona_temp_end + "." + skan.rozszerzenie;

                            nowy_plik = System.IO.Path.Combine(System.IO.Path.GetDirectoryName(plik), nowy_plik);

                            //jezeli plik o nazwie docelowej juz istnieje to zmien mu rozszerzenie na new
                            if (System.IO.File.Exists(nowy_plik)) nowy_plik = nowy_plik + ".new";
                            File.Move(plik, nowy_plik);
                        }
                        else
                        {
                            if (skan.strona_end != -1) strona_old = skan.strona_end;
                            else strona_old = skan.strona_start;
                        }
                    }
                }
                sygnatura_old = sygnatura_new;
            }

        }

        public Hashtable getImageInfo(string plik)
        {
            Hashtable dane = new Hashtable();
            //var dane = new List<KeyValuePair<string, string>>();
            ProcessStartInfo startInfo = new ProcessStartInfo();
            startInfo.FileName = AppDomain.CurrentDomain.BaseDirectory + "bin\\exiftool.exe";
            startInfo.Arguments = "-fast -q -s2 -d \"%Y-%m-%d %H:%M:%S\" -EXIF:XResolution -EXIF:ImageWidth -EXIF:ImageHeight -EXIF:Compression -EXIF:ModifyDate -EXIF:Make -EXIF:Model -EXIF:Orientation -EXIF:Software -EXIF:DocumentName " + "\"" + plik + "\"";
            startInfo.CreateNoWindow = true;
            startInfo.WindowStyle = ProcessWindowStyle.Hidden;
            startInfo.ErrorDialog = false;
            startInfo.UseShellExecute = false;
            startInfo.RedirectStandardOutput = true;
            startInfo.RedirectStandardError = true;
            startInfo.StandardOutputEncoding = Encoding.UTF8;

            try
            {
                // Start the process with the info we specified.
                // Call WaitForExit and then the using-statement will close.
                using (Process exeProcess = Process.Start(startInfo))
                {
                    string output = exeProcess.StandardOutput.ReadToEnd();
                    foreach (Match m in EXIF_INFO.Matches(output)) dane.Add(m.Groups[1].Value.Trim(), m.Groups[2].Value.Trim());
                    exeProcess.WaitForExit();
                }
            }
            catch
            {
                MessageBox.Show("Błąd odczytu informacji o skanie: " + plik, "Błąd");
            }
            return dane;
        }

        public string dopelnij(string dane_wejsciowe, string wzorzec)
        {
            int index = wzorzec.IndexOf("_");
            if (index > 0) wzorzec.Substring(0, index);
            int ilosc_cyfr = wzorzec.Length - dane_wejsciowe.Length;

            for (int j = 0; j < ilosc_cyfr; j++)
            {
                dane_wejsciowe = "0" + dane_wejsciowe;
            }

            return dane_wejsciowe;
        }

        private void backgroundWorker1_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            // Change the value of the ProgressBar to the BackgroundWorker progress.
            progressBar.Value = e.ProgressPercentage;
            // Set the text.
            //this.Text = e.ProgressPercentage.ToString();
        }

        private void katalogMetadaneWskaz_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog dialog = new FolderBrowserDialog();
            dialog.ShowDialog();
            if (dialog.SelectedPath != "")
            {
                skanyMetadane.Text = dialog.SelectedPath;
            }
        }

        private void profilWskaz_Click(object sender, EventArgs e)
        {
            OpenFileDialog dialog = new OpenFileDialog();
            dialog.Filter = "Profil ICC|*.icc";
            dialog.Title = "Wskaż profil barwny";
            DialogResult result = dialog.ShowDialog();
            if (result == DialogResult.OK) // Test result.
            {
                profilICC.Text = dialog.FileName;
            }
        }

        private void metadaneButton_Click(object sender, EventArgs e)
        {
            if (skanyMetadane.Text.Length != 0 || profilICC.Text.Length != 0)
            {
                tabControl1.Enabled = false;
                this.Text = "Katalogi - dodawanie metadanych";
                tryb = 4;
                progressBar.Visible = true;

                if (skanyMetadane.Text.EndsWith("\\"))
                {
                    int koniec = skanyMetadane.Text.LastIndexOf("\\");
                    string katalog_glowny = skanyMetadane.Text.Substring(0, koniec);
                    DirSearch(katalog_glowny);
                }
                else DirSearch(skanyMetadane.Text);

                this.Text = "Katalogi";
                tabControl1.Enabled = true;
                progressBar.Visible = false;
            }
            else MessageBox.Show("Proszę wskazać katalog ze skanami do weryfikacji oraz wskazać profil ICC");
        }

        private void checkArtist_CheckedChanged(object sender, EventArgs e)
        {
            if (checkArtist.Checked) metaArtist.Enabled = true;
            else metaArtist.Enabled = false;
        }

        private void checkCopyright_CheckedChanged(object sender, EventArgs e)
        {
            if (checkCopyright.Checked) metaCopyright.Enabled = true;
            else metaCopyright.Enabled = false;
        }

        private void checkUserComment_CheckedChanged(object sender, EventArgs e)
        {
            if (checkUserComment.Checked) metaUserComment.Enabled = true;
            else metaUserComment.Enabled = false;
        }

        public void metadane(string plik, bool SMA)
        {
            Skan skan = new Skan(plik);

            if (skan.PobierzPlik() != null)
            {
                if (skan.rozszerzenie.Equals("tif") || skan.rozszerzenie.Equals("tiff"))
                {
                    // -ModifyDate -Make -Model -Orientation -Software 
                    var dane = getImageInfo(plik);
                    string DocumentName = null;
                    if (dane.ContainsKey("DocumentName")) DocumentName = (string)dane["DocumentName"];

                    // jezeli plik nie posiada specjalnych metadanych to je zapisz
                    if (overwriteBox.Checked == true || DocumentName == null)
                    {
                        FileInfo f = new FileInfo(plik);
                        string format = "yyyy-MM-dd HH:mm:ss";
                        string ModifyDate = null;
                        string Make = null;
                        string Model = null;
                        string Orientation = null;
                        string Software = null;

                        string argumenty = "-charset Latin2 -q -P -overwrite_original \"-icc_profile<=" + profilICC.Text + "\" -DocumentName=\"" + f.Name + "\"";

                        if (dane.ContainsKey("ModifyDate")) ModifyDate = (string)dane["ModifyDate"];
                        if (dane.ContainsKey("Make")) Make = (string)dane["Make"];
                        if (dane.ContainsKey("Model")) Model = (string)dane["Model"];
                        if (dane.ContainsKey("Orientation")) Orientation = (string)dane["Orientation"];
                        if (dane.ContainsKey("Software")) Software = (string)dane["Software"];

                        // jezeli dokument powstal na scan master 2
                        if (SMA)
                        {
                            if (smaModel.Text.Length > 0) Model = smaModel.Text.Replace("\"", "\\\"");
                            if (smaMake.Text.Length > 0) Make = smaMake.Text.Replace("\"", "\\\"");
                            if (smaSoftware.Text.Length > 0) Software = smaSoftware.Text.Replace("\"", "\\\"");
                        }
                        else
                        {
                            if (metisModel.Text.Length > 0) Model = metisModel.Text.Replace("\"", "\\\"");
                            if (metisMake.Text.Length > 0) Make = metisMake.Text.Replace("\"", "\\\"");
                            if (metisSoftware.Text.Length > 0) Software = metisSoftware.Text.Replace("\"", "\\\"");
                        }

                        if (checkArtist.Checked && metaArtist.Text.Length > 0) argumenty += " -EXIF:Artist=\"" + metaArtist.Text.Replace("\"", "\\\"") + "\"";
                        if (checkCopyright.Checked && metaCopyright.Text.Length > 0) argumenty += " -EXIF:Copyright=\"" + metaCopyright.Text.Replace("\"", "\\\"") + "\"";
                        if (checkUserComment.Checked && metaUserComment.Text.Length > 0) argumenty += " -EXIF:UserComment=\"" + metaUserComment.Text.Replace("\"", "\\\"") + "\"";
                        if (ModifyDate == null)
                        {
                            DateTime data_utworzenia = f.CreationTime;
                            DateTime data_modyfikacji = f.LastWriteTime;
                            if (DateTime.Compare(data_modyfikacji, data_utworzenia) < 0) argumenty += " -EXIF:ModifyDate=\"" + data_modyfikacji.ToString(format) + "\"";
                            else argumenty += " -EXIF:ModifyDate=\"" + data_utworzenia.ToString(format) + "\"";
                        }
                        if (Orientation == null) argumenty += " -EXIF:Orientation=\"Horizontal (normal)\"";
                        argumenty += " -EXIF:Make=\"" + Make + "\"";
                        argumenty += " -EXIF:Model=\"" + Model + "\"";
                        argumenty += " -EXIF:Software=\"" + Software + "\"";

                        //MessageBox.Show((string)dane["UserComment"]);
                        // ponizej potrzebna konwersja do prawidłowego wyświetlania polskich znaków
                        byte[] unicodebytes = Encoding.UTF8.GetBytes(argumenty + " \"" + plik + "\"");

                        ProcessStartInfo startInfo = new ProcessStartInfo();
                        startInfo.FileName = AppDomain.CurrentDomain.BaseDirectory + "bin\\exiftool.exe";
                        startInfo.Arguments = Encoding.Default.GetString(Encoding.Convert(Encoding.UTF8, Encoding.GetEncoding(1250), unicodebytes));
                        //startInfo.Arguments = argumenty + " \"" + plik + "\"";
                        startInfo.CreateNoWindow = true;
                        startInfo.WindowStyle = ProcessWindowStyle.Hidden;
                        startInfo.ErrorDialog = false;
                        startInfo.UseShellExecute = false;
                        startInfo.RedirectStandardOutput = true;
                        startInfo.RedirectStandardError = true;
                        startInfo.StandardOutputEncoding = Encoding.GetEncoding(1250);

                        try
                        {
                            // Start the process with the info we specified.
                            // Call WaitForExit and then the using-statement will close.
                            using (Process exeProcess = Process.Start(startInfo))
                            {
                                exeProcess.WaitForExit();
                            }
                        }
                        catch
                        {
                            MessageBox.Show("Błąd zapisywania metadanych w " + plik, "Błąd");
                        }
                    }

                }

            }
        }

        public bool sprawdzSkaner(string plik1, string plik2, string plik3)
        {
            string output = null;
            string args = "-fast -q -s3 -XMP:CreatorTool";
            if (plik1 != null) args += " \"" + plik1 + "\"";
            if (plik2 != null) args += " \"" + plik2 + "\"";
            if (plik3 != null) args += " \"" + plik3 + "\"";
            ProcessStartInfo startInfo = new ProcessStartInfo();
            startInfo.FileName = AppDomain.CurrentDomain.BaseDirectory + "bin\\exiftool.exe";
            startInfo.Arguments = args;
            startInfo.CreateNoWindow = true;
            startInfo.WindowStyle = ProcessWindowStyle.Hidden;
            startInfo.ErrorDialog = false;
            startInfo.UseShellExecute = false;
            startInfo.RedirectStandardOutput = true;
            startInfo.RedirectStandardError = true;

            try
            {
                // Start the process with the info we specified.
                // Call WaitForExit and then the using-statement will close.
                using (Process exeProcess = Process.Start(startInfo))
                {
                    output = exeProcess.StandardOutput.ReadToEnd();
                    exeProcess.WaitForExit();
                }
            }
            catch
            {
                MessageBox.Show("Nie udało się sprawdzić rodzaju skanera przy wykorzystaniu polecenia: " + args, "Błąd");
            }

            if (output != null && output.Contains("SMA-A2")) return true;
            else return false;
        }

        public bool sprawdzWzorniki(string plik)
        {
            string red = plik + " -alpha set -fuzz 30% -transparent yellow -verbose info:";
            string yellow = plik + " -alpha set -fuzz 30% -transparent red -verbose info:";
            string output = "";

            ProcessStartInfo startInfo = new ProcessStartInfo();
            startInfo.FileName = convert;
            startInfo.Arguments = yellow;
            startInfo.CreateNoWindow = true;
            startInfo.WindowStyle = ProcessWindowStyle.Hidden;
            startInfo.ErrorDialog = false;
            startInfo.UseShellExecute = false;
            startInfo.RedirectStandardOutput = true;
            startInfo.RedirectStandardError = false;

            try
            {
                using (Process exeProcess = Process.Start(startInfo))
                {
                    output = exeProcess.StandardOutput.ReadToEnd();
                    exeProcess.WaitForExit();
                }
            }
            catch
            {
                MessageBox.Show("Błąd podczas sprawdzania wzornika dla " + plik, "Błąd");
            }

            Match m = kolorTest.Match(output);

            if (m != null && m.Success)
            {
                string procent = m.Groups[1].Value;
                //MessageBox.Show(plik + ": " + procent);
                //MessageBox.Show(procent);
                if (procent.StartsWith("0.9"))
                {
                    output = "";
                    startInfo = new ProcessStartInfo();
                    startInfo.FileName = convert;
                    startInfo.Arguments = red;
                    startInfo.CreateNoWindow = true;
                    startInfo.WindowStyle = ProcessWindowStyle.Hidden;
                    startInfo.ErrorDialog = false;
                    startInfo.UseShellExecute = false;
                    startInfo.RedirectStandardOutput = true;
                    startInfo.RedirectStandardError = false;

                    try
                    {
                        using (Process exeProcess = Process.Start(startInfo))
                        {
                            output = exeProcess.StandardOutput.ReadToEnd();
                            exeProcess.WaitForExit();
                        }
                    }
                    catch
                    {
                        MessageBox.Show("Błąd podczas sprawdzania wzornika dla " + plik, "Błąd");
                    }

                    m = kolorTest.Match(output);
                    if (m != null && m.Success)
                    {
                        procent = m.Groups[1].Value;
                        //MessageBox.Show(plik + ": " + procent);
                        if (procent.StartsWith("0.9")) return true;
                    }
                }
            }
            return false;
        }

        private void katalogSkanyCRC_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog dialog = new FolderBrowserDialog();
            dialog.ShowDialog();
            if (dialog.SelectedPath != "")
            {
                skanyCRC.Text = dialog.SelectedPath;
            }
        }

        private void generujSumyCRC_Click(object sender, EventArgs e)
        {
            if (skanyCRC.Text.Length != 0 || plikCRC.Text.Length != 0)
            {
                tabControl1.Enabled = false;
                this.Text = "Katalogi - generowanie sum kontrolnych";
                tryb = 5;
                progressBar.Visible = true;

                if (skanyCRC.Text.EndsWith("\\"))
                {
                    int koniec = skanyCRC.Text.LastIndexOf("\\");
                    string katalog_glowny = skanyCRC.Text.Substring(0, koniec);
                    DirSearch(katalog_glowny);
                }
                else DirSearch(skanyCRC.Text);

                this.Text = "Katalogi";
                tabControl1.Enabled = true;
                progressBar.Visible = false;
            }
            else MessageBox.Show("Proszę wskazać katalog ze skanami oraz miejsce zapisu sum kontrolnych");
        }

        private void przyciskPlikCRC_Click(object sender, EventArgs e)
        {
            SaveFileDialog dialog = new SaveFileDialog();
            dialog.Filter = "SHA-256|*.sha256";
            dialog.Title = "Wskaż miejsce zapisu sum kontrolnych";
            DialogResult result = dialog.ShowDialog();
            if (result == DialogResult.OK) // Test result.
            {
                plikCRC.Text = dialog.FileName;
            }
        }

        private void liczcrc(string plik)
        {
            Skan skan = new Skan(plik);

            if (skan.PobierzPlik() != null)
            {
                if (skan.rozszerzenie.Equals("tif") || skan.rozszerzenie.Equals("tiff") || skan.rozszerzenie.Equals("jpg") || skan.rozszerzenie.Equals("jpeg"))
                {
                    StreamWriter plik_sumy = new StreamWriter(plikCRC.Text, true);
                    plik_sumy.WriteLine(skan.GenerujCRC() + " *" + skan.GenerujNazwePliku());
                    plik_sumy.Close();
                }
            }
        }


        private void skanyMetryczkiButton_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog dialog = new FolderBrowserDialog();
            dialog.ShowDialog();
            if (dialog.SelectedPath != "")
            {
                skanyMetryczki.Text = dialog.SelectedPath;
            }
        }

        private void metryczkiStart_Click(object sender, EventArgs e)
        {
            if (skanyMetryczki.Text.Length == 0) MessageBox.Show("Proszę wskazać katalog ze skanami", "Błąd");
            else if (digitalizatorBox.Text.Length == 0) MessageBox.Show("Proszę uzupełnić pole digitalizator", "Błąd");
            else if (kontrolerBox.Text.Length == 0) MessageBox.Show("Proszę uzupełnić pole kontroler", "Błąd");
            else if (pracowniaBox.Text.Length == 0) MessageBox.Show("Proszę uzupełnić pole pracownia", "Błąd");
            else
            {
                // dodawanie \ na koncu katalogu docelowego
                bledyBox.Clear();
                tabControl1.Enabled = false;
                this.Text = "Metryczki - generowanie";
                tryb = 6;
                progressBar.Visible = true;

                if (skanyMetryczki.Text.EndsWith("\\"))
                {
                    int koniec = skanyMetryczki.Text.LastIndexOf("\\");
                    string katalog_glowny = skanyMetryczki.Text.Substring(0, koniec);
                    DirSearch(katalog_glowny);
                }
                else DirSearch(skanyMetryczki.Text);

                tabControl1.Enabled = true;
                this.Text = "Katalogi";
                progressBar.Visible = false;
                //MessageBox.Show("Gotowe");
            }
        }

        public bool metryczki(string plik, bool last, bool revers)
        {
            Skan skan = new Skan(plik);

            if (skan.PobierzPlik() != null)
            {
                string sygnatura_new = skan.nr_archiwum + "_" + skan.nr_zespolu + "_" + skan.cd_zespolu + "_" + skan.seria + "_" + skan.sygnatura;

                // jezeli skany nie leca od konca w poszukiwaniu ostatniego
                if (revers == false)
                {
                    if ((skan.rozszerzenie.Equals("tif") || skan.rozszerzenie.Equals("tiff")) && (skan.strona.Contains("metryczka") == false && skan.strona.Contains("tablica") == false))
                    {
                        //MessageBox.Show(plik);
                        licz_skany++;
                        // jezeli dodana strona
                        if (DODANE.IsMatch(skan.strona)) dodaneStrony.Add(skan.strona.TrimStart('0').Replace("_", ""));
                    }
                    else if (skan.rozszerzenie.Equals("txt")) pusteStrony.Add(skan.strona.TrimStart('0'));
                }

                // jezeli ostatnia strona to stworz metryczke
                if (last)
                {
                    if (skan.strona_start == -1) return true;
                    else
                    {
                        int stron = 0;
                        // jezeli nowe zarzadzenie to pobierz ilosc stron z bazy IZA
                        if (radioNowe.Checked) stron = pobierzStrony(skan.nr_zespolu, skan.seria, skan.sygnatura);
                        // jezeli w izie nie ma podanej strony to policz po liczniku
                        if (radioStare.Checked)
                        {
                            stron = skan.strona_start;
                            if (skan.strona_end != -1) stron = skan.strona_end;
                        }

                        FileInfo f = new FileInfo(plik);
                        string format = "dd.MM.yyyy";
                        DateTime data_utworzenia = f.CreationTime;
                        DateTime data_modyfikacji = f.LastWriteTime;

                        if ((skan.rozszerzenie.Equals("tif") || skan.rozszerzenie.Equals("tiff")))
                        {
                            try
                            {
                                var dane = getImageInfo(plik);
                                if (dane.ContainsKey("ModifyDate")) data_utworzenia = DateTime.ParseExact((string)dane["ModifyDate"], format, CultureInfo.InvariantCulture);
                            }
                            catch { }
                        }
                        // jezeli data modyfikacji jest wczesniejsza to przyjmij ta wczesniejsza za date utworzenia
                        if (DateTime.Compare(data_modyfikacji, data_utworzenia) < 0) data_utworzenia = data_modyfikacji;

                        if (licz_skany > 0 && stron > 0) tworzMetryczke(f.DirectoryName, skan.nr_archiwum, skan.nr_zespolu, skan.cd_zespolu, skan.seria, skan.sygnatura, Convert.ToString(licz_skany), Convert.ToString(stron), data_utworzenia.ToString(format));
                        else MessageBox.Show("Brak informacji o ilości stron dla skanów w katalogu: " + f.DirectoryName, "Błąd podczas generowania metryczki");
                    }
                }
            }
            return false;
        }

        public void tworzMetryczke(string katalog, string nr_archiwum, string nr_zespolu, string cd_zespolu, string seria, string sygnatura, string licz_skany, string stron, string data)
        {
            //string tytul_jednostki = pobierzNazweJednostki(nr_archiwum, nr_zespolu, cd_zespolu, seria, sygnatura);
            string nazwa_zespolu = pobierzNazweZespolu(nr_archiwum, nr_zespolu, cd_zespolu);
            string dodane = "";
            string puste = "";

            if (nazwa_zespolu != null && nazwa_zespolu.Length > 0)
            {
                if (dodaneStrony.Count > 0)
                {
                    dodane = "Dodano str.: ";
                    for (int i = 0; i < dodaneStrony.Count; i++)
                    {
                        if (i > 0) dodane += ", " + dodaneStrony[i];
                        else dodane += dodaneStrony[i];
                    }
                }

                if (pusteStrony.Count > 0)
                {
                    puste = "Brak str.: ";
                    for (int i = 0; i < pusteStrony.Count; i++)
                    {
                        if (i > 0) puste += "," + pusteStrony[i];
                        else puste += pusteStrony[i];
                    }
                }

                if (radioNowe.Checked)
                {
                    string metryczka = AppDomain.CurrentDomain.BaseDirectory + "data\\metryczka.html";
                    string metryczka_htm = System.IO.Path.Combine(katalog, nr_archiwum + "_" + nr_zespolu + "_" + cd_zespolu + "_" + seria + "_" + sygnatura + "_0000_metryczka.htm");
                    string metryczka_jpg = System.IO.Path.Combine(katalog, nr_archiwum + "_" + nr_zespolu + "_" + cd_zespolu + "_" + seria + "_" + sygnatura + "_0000_metryczka.jpg");

                    File.Copy(metryczka, metryczka_htm, true);

                    string text = File.ReadAllText(metryczka_htm);
                    text = text.Replace("&lt;&lt;ARCHIWUM&gt;&gt;", archiwumBox.Text);
                    text = text.Replace("&lt;&lt;ZESPOL&gt;&gt;", nazwa_zespolu);
                    text = text.Replace("&lt;&lt;NRZESPOLU&gt;&gt;", nr_zespolu);
                    text = text.Replace("&lt;&lt;CDNUMERU&gt;&gt;", cd_zespolu);
                    text = text.Replace("&lt;&lt;SERIA&gt;&gt;", seria);
                    text = text.Replace("&lt;&lt;SYGNATURA&gt;&gt;", sygnatura);
                    text = text.Replace("&lt;&lt;STRON&gt;&gt;", stron);
                    text = text.Replace("&lt;&lt;SKANOW&gt;&gt;", licz_skany);
                    text = text.Replace("&lt;&lt;PRACOWNIA&gt;&gt;", pracowniaBox.Text);
                    text = text.Replace("&lt;&lt;DATA&gt;&gt;", data);
                    text = text.Replace("&lt;&lt;SKANOWAL&gt;&gt;", digitalizatorBox.Text);
                    text = text.Replace("&lt;&lt;KONTROLER&gt;&gt;", kontrolerBox.Text);
                    text = text.Replace("&lt;&lt;PUSTE&gt;&gt;", puste);
                    text = text.Replace("&lt;&lt;DODANE&gt;&gt;", dodane);

                    File.WriteAllText(metryczka_htm, text);

                    ProcessStartInfo startInfo = new ProcessStartInfo();
                    startInfo.FileName = AppDomain.CurrentDomain.BaseDirectory + "bin\\IECapt.exe";
                    startInfo.Arguments = " --url=file:\"" + metryczka_htm + "\"" + " --out=\"" + metryczka_jpg + "\" --min-width=1024";
                    startInfo.CreateNoWindow = true;
                    startInfo.WindowStyle = ProcessWindowStyle.Hidden;
                    startInfo.ErrorDialog = false;
                    startInfo.UseShellExecute = false;
                    startInfo.RedirectStandardOutput = true;
                    startInfo.RedirectStandardError = true;

                    try
                    {
                        using (Process exeProcess = Process.Start(startInfo))
                        {
                            exeProcess.WaitForExit();
                            File.Delete(metryczka_htm);
                        }
                    }
                    catch
                    {
                        MessageBox.Show("Błąd tworzenia metryczki " + metryczka_htm, "Błąd");
                    }
                }
                // stare zarzadzenie
                else
                {
                    int licz_skany_old = Convert.ToInt32(licz_skany) + 3;
                    string metryczka = AppDomain.CurrentDomain.BaseDirectory + "data\\metryczka_old.html";
                    string metryczka_htm = System.IO.Path.Combine(katalog, nr_archiwum + "_" + nr_zespolu + "_" + cd_zespolu + "_" + seria + "_" + sygnatura + "_0000_metryczka.htm");
                    string metryczka_tif = System.IO.Path.Combine(katalog, nr_archiwum + "_" + nr_zespolu + "_" + cd_zespolu + "_" + seria + "_" + sygnatura + "_0000_metryczka.tif");
                    string tablica_poczatkowa = AppDomain.CurrentDomain.BaseDirectory + "data\\tablica_poczatkowa.html";
                    string tablica_poczatkowa_htm = System.IO.Path.Combine(katalog, nr_archiwum + "_" + nr_zespolu + "_" + cd_zespolu + "_" + seria + "_" + sygnatura + "_0000_tablica_poczatkowa.htm");
                    string tablica_poczatkowa_tif = System.IO.Path.Combine(katalog, nr_archiwum + "_" + nr_zespolu + "_" + cd_zespolu + "_" + seria + "_" + sygnatura + "_0000_tablica_poczatkowa.tif");
                    string tablica_koncowa = AppDomain.CurrentDomain.BaseDirectory + "data\\tablica_koncowa.html";
                    string tablica_koncowa_htm = System.IO.Path.Combine(katalog, nr_archiwum + "_" + nr_zespolu + "_" + cd_zespolu + "_" + seria + "_" + sygnatura + "_9999_tablica_koncowa.htm");
                    string tablica_koncowa_tif = System.IO.Path.Combine(katalog, nr_archiwum + "_" + nr_zespolu + "_" + cd_zespolu + "_" + seria + "_" + sygnatura + "_9999_tablica_koncowa.tif");


                    // metryczka
                    File.Copy(metryczka, metryczka_htm, true);

                    string text = File.ReadAllText(metryczka_htm);
                    text = text.Replace("&lt;&lt;ARCHIWUM&gt;&gt;", archiwumBox.Text);
                    text = text.Replace("&lt;&lt;ZESPOL&gt;&gt;", nazwa_zespolu);
                    text = text.Replace("&lt;&lt;NRZESPOLU&gt;&gt;", nr_zespolu);
                    text = text.Replace("&lt;&lt;SERIA&gt;&gt;", seria);
                    text = text.Replace("&lt;&lt;SYGNATURA&gt;&gt;", sygnatura);
                    text = text.Replace("&lt;&lt;STRON&gt;&gt;", stron);
                    text = text.Replace("&lt;&lt;SKANOW&gt;&gt;", Convert.ToString(licz_skany_old));
                    text = text.Replace("&lt;&lt;PRACOWNIA&gt;&gt;", pracowniaBox.Text);
                    text = text.Replace("&lt;&lt;DATA&gt;&gt;", data);
                    text = text.Replace("&lt;&lt;PUSTE&gt;&gt;", puste);
                    text = text.Replace("&lt;&lt;DODANE&gt;&gt;", dodane);

                    File.WriteAllText(metryczka_htm, text);

                    ProcessStartInfo startInfo = new ProcessStartInfo();
                    startInfo.FileName = AppDomain.CurrentDomain.BaseDirectory + "bin\\IECapt.exe";
                    startInfo.Arguments = " --url=file:\"" + metryczka_htm + "\"" + " --out=\"" + metryczka_tif + "\" --min-width=1024";
                    startInfo.CreateNoWindow = true;
                    startInfo.WindowStyle = ProcessWindowStyle.Hidden;
                    startInfo.ErrorDialog = false;
                    startInfo.UseShellExecute = false;
                    startInfo.RedirectStandardOutput = true;
                    startInfo.RedirectStandardError = true;

                    try
                    {
                        using (Process exeProcess = Process.Start(startInfo))
                        {
                            exeProcess.WaitForExit();
                            File.Delete(metryczka_htm);
                        }
                    }
                    catch
                    {
                        MessageBox.Show("Błąd tworzenia metryczki " + metryczka_htm, "Błąd");
                    }

                    // tablica poczatkowa
                    File.Copy(tablica_poczatkowa, tablica_poczatkowa_htm, true);

                    text = File.ReadAllText(tablica_poczatkowa_htm);
                    text = text.Replace("&lt;&lt;ZESPOL&gt;&gt;", nazwa_zespolu);
                    text = text.Replace("&lt;&lt;NRZESPOLU&gt;&gt;", nr_zespolu);
                    text = text.Replace("&lt;&lt;SERIA&gt;&gt;", seria);
                    text = text.Replace("&lt;&lt;SYGNATURA&gt;&gt;", sygnatura);
                    text = text.Replace("&lt;&lt;DATA&gt;&gt;", data);
                    text = text.Replace("&lt;&lt;SKANOWAL&gt;&gt;", digitalizatorBox.Text);
                    text = text.Replace("&lt;&lt;PRACOWNIA&gt;&gt;", pracowniaBox.Text);

                    File.WriteAllText(tablica_poczatkowa_htm, text);

                    startInfo = new ProcessStartInfo();
                    startInfo.FileName = AppDomain.CurrentDomain.BaseDirectory + "bin\\IECapt.exe";
                    startInfo.Arguments = " --url=file:\"" + tablica_poczatkowa_htm + "\"" + " --out=\"" + tablica_poczatkowa_tif + "\" --min-width=1024";
                    startInfo.CreateNoWindow = true;
                    startInfo.WindowStyle = ProcessWindowStyle.Hidden;
                    startInfo.ErrorDialog = false;
                    startInfo.UseShellExecute = false;
                    startInfo.RedirectStandardOutput = true;
                    startInfo.RedirectStandardError = true;

                    try
                    {
                        using (Process exeProcess = Process.Start(startInfo))
                        {
                            exeProcess.WaitForExit();
                            File.Delete(tablica_poczatkowa_htm);
                        }
                    }
                    catch
                    {
                        MessageBox.Show("Błąd tworzenia tablicy początkowej " + tablica_poczatkowa_htm, "Błąd");
                    }

                    // tablica koncowa
                    File.Copy(tablica_koncowa, tablica_koncowa_htm, true);

                    text = File.ReadAllText(tablica_koncowa_htm);
                    text = text.Replace("&lt;&lt;NRZESPOLU&gt;&gt;", nr_zespolu);
                    text = text.Replace("&lt;&lt;SERIA&gt;&gt;", seria);
                    text = text.Replace("&lt;&lt;SYGNATURA&gt;&gt;", sygnatura);
                    text = text.Replace("&lt;&lt;DATA&gt;&gt;", data);
                    text = text.Replace("&lt;&lt;KONTROLER&gt;&gt;", kontrolerBox.Text);

                    File.WriteAllText(tablica_koncowa_htm, text);

                    startInfo = new ProcessStartInfo();
                    startInfo.FileName = AppDomain.CurrentDomain.BaseDirectory + "bin\\IECapt.exe";
                    startInfo.Arguments = " --url=file:\"" + tablica_koncowa_htm + "\"" + " --out=\"" + tablica_koncowa_tif + "\" --min-width=1024";
                    startInfo.CreateNoWindow = true;
                    startInfo.WindowStyle = ProcessWindowStyle.Hidden;
                    startInfo.ErrorDialog = false;
                    startInfo.UseShellExecute = false;
                    startInfo.RedirectStandardOutput = true;
                    startInfo.RedirectStandardError = true;

                    try
                    {
                        using (Process exeProcess = Process.Start(startInfo))
                        {
                            exeProcess.WaitForExit();
                            File.Delete(tablica_koncowa_htm);
                        }
                    }
                    catch
                    {
                        MessageBox.Show("Błąd tworzenia tablicy końcowej " + tablica_poczatkowa_htm, "Błąd");
                    }
                }
            }
            else MessageBox.Show("Nieznaleziono opisu zespołu nr: " + nr_zespolu + " w bazie danych IZA", "Błąd");


        }

        string pobierzNazweJednostki(string nr_archiwum, string nr_zespolu, string cd_zespolu, string seria, string sygnatura)
        {
            // jezeli seria posiada . to pobierz wartosc po lewej stronie
            if (SERIA_REGEX.Match(seria).Success)
            {
                Match match = SERIA_REGEX.Match(seria);
                seria = match.Groups[1].Value;
            }

            string dane = null;
            string connectionString =
            "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=\""
            + baza_iza + "\";User Id=;Password=;";

            //MessageBox.Show(connectionString);

            // Provide the query string with a parameter placeholder.
            //SELECT SERIA FROM INWENTARZ WHERE NRZESPOLU=126 AND SYGNATURA="25M";
            string queryString =
                "SELECT tytul from INWENTARZ "
                    + "WHERE NRAP = @nrapPole AND NRZESPOLU = @zespolPole AND CDNUMERU = @cdnumeruPole AND SERIA = @seriaPole AND SYGNATURA = @sygnaturaPole;";

            // Create and open the connection in a using block. This
            // ensures that all resources will be closed and disposed
            // when the code exits.
            using (OleDbConnection connection =
                new OleDbConnection(connectionString))
            {
                // Create the Command and Parameter objects.
                OleDbCommand command = new OleDbCommand(queryString, connection);
                command.Parameters.AddWithValue("@nrapPole", nr_archiwum);
                command.Parameters.AddWithValue("@zespolPole", nr_zespolu);
                command.Parameters.AddWithValue("@cdnumeruPole", cd_zespolu);
                command.Parameters.AddWithValue("@seriaPole", seria);
                command.Parameters.AddWithValue("@sygnaturaPole", sygnatura);

                // Open the connection in a try/catch block. 
                // Create and execute the DataReader, writing the result
                // set to the console window.
                try
                {
                    connection.Open();
                    OleDbDataReader reader = command.ExecuteReader();
                    if (reader.Read()) dane = "" + reader[0];
                    reader.Close();
                    return dane;
                }
                catch (Exception ex)
                {
                    //MessageBox.Show(""+ex.Message);
                }
                //Console.ReadLine();
            }
            return dane;
        }

        string pobierzNazweZespolu(string nr_archiwum, string nr_zespolu, string cd_zespolu)
        {
            string dane = null;
            string connectionString =
            "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=\""
            + baza_iza + "\";User Id=;Password=;";


            // Provide the query string with a parameter placeholder.
            //SELECT SERIA FROM INWENTARZ WHERE NRZESPOLU=126 AND SYGNATURA="25M";
            string queryString =
                "SELECT tytul from ZESPOLY "
                    + "WHERE NRAP = @nrapPole AND NRZESPOLU = @zespolPole AND CDNUMERU = @cdnumeruPole;";

            // Create and open the connection in a using block. This
            // ensures that all resources will be closed and disposed
            // when the code exits.
            using (OleDbConnection connection =
                new OleDbConnection(connectionString))
            {
                // Create the Command and Parameter objects.
                OleDbCommand command = new OleDbCommand(queryString, connection);
                command.Parameters.AddWithValue("@nrapPole", nr_archiwum);
                command.Parameters.AddWithValue("@zespolPole", nr_zespolu);
                command.Parameters.AddWithValue("@cdnumeruPole", cd_zespolu);

                // Open the connection in a try/catch block. 
                // Create and execute the DataReader, writing the result
                // set to the console window.
                try
                {
                    connection.Open();
                    OleDbDataReader reader = command.ExecuteReader();
                    if (reader.Read()) dane = "" + reader[0];
                    reader.Close();
                    return dane;
                }
                catch (Exception ex)
                {
                    MessageBox.Show("" + ex.Message);
                }
                //Console.ReadLine();
            }
            return dane;
        }

        string pobierzAutora(string nr_zespolu, string sygnatura)
        {
            string dane = null;
            string connectionString =
            "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=\"D:\\skany zlecenie\\SKANY ZLECENIE.mdb\";User Id=;Password=;";


            // Provide the query string with a parameter placeholder.
            //SELECT SERIA FROM INWENTARZ WHERE NRZESPOLU=126 AND SYGNATURA="25M";
            string queryString =
                "SELECT osoba from PRACE "
                    + "WHERE NRZESPOLU = @zespolPole AND SYGNATURA = @sygnaturaPole;";

            // Create and open the connection in a using block. This
            // ensures that all resources will be closed and disposed
            // when the code exits.
            using (OleDbConnection connection =
                new OleDbConnection(connectionString))
            {
                // Create the Command and Parameter objects.
                OleDbCommand command = new OleDbCommand(queryString, connection);
                command.Parameters.AddWithValue("@zespolPole", nr_zespolu);
                command.Parameters.AddWithValue("@sygnaturaPole", sygnatura);

                // Open the connection in a try/catch block. 
                // Create and execute the DataReader, writing the result
                // set to the console window.
                try
                {
                    connection.Open();
                    OleDbDataReader reader = command.ExecuteReader();
                    if (reader.Read()) dane = "" + reader[0];
                    reader.Close();
                    return dane;
                }
                catch (Exception ex)
                {
                    //MessageBox.Show(""+ex.Message);
                }
                //Console.ReadLine();
            }
            return dane;
        }

        public void uzupelnijFormatki()
        {

            string dane = null;
            string connectionString =
            "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=\""
            + bazaSciezka.Text + "\";User Id=;Password=;";

            // Provide the query string with a parameter placeholder.
            //SELECT SERIA FROM INWENTARZ WHERE NRZESPOLU=126 AND SYGNATURA="25M";
            string queryString = "SELECT TYP_FINANSOWANIA from FINANSOWANIE;";

            // Create and open the connection in a using block. This
            // ensures that all resources will be closed and disposed
            // when the code exits.
            using (OleDbConnection connection =
                new OleDbConnection(connectionString))
            {
                // Create the Command and Parameter objects.
                OleDbCommand command = new OleDbCommand(queryString, connection);

                // Open the connection in a try/catch block. 
                // Create and execute the DataReader, writing the result
                // set to the console window.
                try
                {
                    connection.Open();
                    OleDbDataReader reader = command.ExecuteReader();
                    while (reader.Read())
                    {
                        dane = Convert.ToString(reader[0]);
                        listaFinansowanie.Items.Add(dane);
                    }
                }
                catch (Exception ex)
                {
                    //MessageBox.Show(""+ex.Message);
                }
                //Console.ReadLine();
            }
            if (listaFinansowanie.Items.Count > 0) listaFinansowanie.SelectedIndex = 0;

            queryString = "SELECT IMIE_I_NAZWISKO from OSOBY;";

            using (OleDbConnection connection =
                new OleDbConnection(connectionString))
            {
                // Create the Command and Parameter objects.
                OleDbCommand command = new OleDbCommand(queryString, connection);

                // Open the connection in a try/catch block. 
                // Create and execute the DataReader, writing the result
                // set to the console window.
                try
                {
                    connection.Open();
                    OleDbDataReader reader = command.ExecuteReader();
                    while (reader.Read())
                    {
                        dane = Convert.ToString(reader[0]);
                        listaOsoby.Items.Add(dane);
                    }
                }
                catch (Exception ex)
                {
                    //MessageBox.Show(""+ex.Message);
                }
                //Console.ReadLine();
            }
            if (listaOsoby.Items.Count > 0) listaOsoby.SelectedIndex = 0;

        }

        private void buttonStop_Click(object sender, EventArgs e)
        {
            tryb = -1;
        }

        private void katalogZrodloWskaz_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog dialog = new FolderBrowserDialog();
            dialog.ShowDialog();
            if (dialog.SelectedPath != "")
            {
                skanyZrodlo.Text = dialog.SelectedPath;
            }
        }

        private void skanyCelWskaz_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog dialog = new FolderBrowserDialog();
            dialog.ShowDialog();
            if (dialog.SelectedPath != "")
            {
                skanyCel.Text = dialog.SelectedPath;
            }
        }

        private void zastapSkanyStart_Click(object sender, EventArgs e)
        {
            if (skanyZrodlo.Text.Length != 0 || skanyCel.Text.Length != 0)
            {
                tabControl1.Enabled = false;
                this.Text = "Katalogi - zastępowanie skanów";
                tryb = 7;
                progressBar.Visible = true;

                if (skanyZrodlo.Text.EndsWith("\\"))
                {
                    int koniec = skanyZrodlo.Text.LastIndexOf("\\");
                    string katalog_glowny = skanyZrodlo.Text.Substring(0, koniec);
                    DirSearch(katalog_glowny);
                }
                else DirSearch(skanyZrodlo.Text);

                this.Text = "Katalogi";
                tabControl1.Enabled = true;
                progressBar.Visible = false;
            }
            else MessageBox.Show("Proszę wskazać katalog ze skanami oraz katalog docelowy");
        }

        private void zastapSkany(string plik, int licznik, string katalogCel)
        {
            if (katalogCel.EndsWith("\\"))
            {
                int koniec = katalogCel.LastIndexOf("\\");
                katalogCel = katalogCel.Substring(0, koniec);
            }

            string strona_zrodlo_text = null;
            string strona_org = null;
            int strona_zrodlo = -1;

            Skan skan = new Skan(plik);

            if (skan.PobierzPlik() != null)
            {
                strona_zrodlo_text = skan.strona;
                strona_org = strona_zrodlo_text;
                try
                {
                    strona_zrodlo = int.Parse(strona_zrodlo_text);
                    // kolejne numery
                    if (strona_prev != -1 && strona_zrodlo - strona_prev == 1) strona_zrodlo += licznik - 1;
                    else strona_zrodlo += licznik;
                    strona_zrodlo_text = dopelnij(strona_zrodlo.ToString(), strona_org);

                    strona_prev = int.Parse(strona_org);
                }
                catch (Exception ex)
                {
                    strona_zrodlo = -1;
                }
            }

            if(strona_zrodlo != -1)
            {
                string destFile = System.IO.Path.Combine(katalogCel, System.IO.Path.GetFileName(plik).Replace(strona_org + ".tif", strona_zrodlo_text + ".tif"));
                if (System.IO.File.Exists(destFile))
                {
                    string[] pliki = Directory.GetFiles(katalogCel, "*.tif");

                    for (int i = 0; i < pliki.Length; i++)
                    {
                        string strona_cel_text = null;
                        int strona_cel = -1;
                        Skan skan_test = new Skan(pliki[i]);

                        if (skan_test.PobierzPlik() != null)
                        {
                            strona_cel_text = skan_test.strona;
                            try
                            {
                                strona_cel = int.Parse(strona_cel_text);
                            }
                            catch (Exception ex)
                            {
                                strona_cel = -1;
                            }

                        }

                        // dodawanie +1 do kazdej strony
                        if (strona_cel != -1 && (strona_zrodlo <= strona_cel))
                        {
                            int nowa_strona = strona_cel + 1;
                            string nowa_strona_text = dopelnij(nowa_strona.ToString(), strona_zrodlo_text);
                            string nowy_plik = pliki[i].Replace(strona_cel_text + ".tif", nowa_strona_text + ".tif.new");
                            System.IO.File.Move(pliki[i], nowy_plik);
                        }

                    }
                    // ponowna zamiana rozszerzen
                    string[] pliki_new = Directory.GetFiles(katalogCel, "*.new");
                    for (int i = 0; i < pliki_new.Length; i++)
                    {
                        System.IO.File.Move(pliki_new[i], pliki_new[i].Replace(".new", ""));
                    }
                }
                System.IO.File.Move(plik, destFile);

            }

        }

    }
}
