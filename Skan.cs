using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace Katalogi
{
    public class Skan
    {
        string plik = null;
        public string nr_archiwum { get; set; }
        public string nr_zespolu { get; set; }
        public string cd_zespolu { get; set; }
        public string seria { get; set; }
        public string sygnatura { get; set; }
        public string strona { get; private set; }
        public string rozszerzenie { get; private set; }
        const string format = "yyyy-MM-dd HH:mm:ss";
        public int strona_start { get; private set; }
        public int strona_end { get; private set; }
        // metadane
        Hashtable metadane;
        // REGEX
        Regex EXIF_INFO = new Regex("(\\w+): (.+)", RegexOptions.Compiled);
        Regex STARY_ZAPIS = new Regex("(\\d+)_(\\d+)_(\\d+)_(\\d+)_(\\d+)_(\\w+)\\.(\\w+)", RegexOptions.Compiled);
        Regex NOWY_ZAPIS = new Regex("(\\d+)_(\\d+)_(\\d+)_([A-Za-z0-9.]+)_([A-Za-z0-9]+)_(\\w+)\\.(\\w+)", RegexOptions.Compiled);
        Regex NUMER_STRONY = new Regex("^(\\d+)(_(\\d+))?$", RegexOptions.Compiled);
        Regex ZERO = new Regex(Regex.Escape("0"), RegexOptions.Compiled);

        public Skan(string plik)
        {
            Match match = null;
            if (STARY_ZAPIS.Match(plik).Success) match = STARY_ZAPIS.Match(plik);
            else if (NOWY_ZAPIS.Match(plik).Success) match = NOWY_ZAPIS.Match(plik);

            if (match != null && match.Success)
            {
                this.plik = plik;

                nr_archiwum = match.Groups[1].Value;
                nr_zespolu = match.Groups[2].Value;
                cd_zespolu = match.Groups[3].Value;
                seria = match.Groups[4].Value;
                sygnatura = match.Groups[5].Value;
                strona = match.Groups[6].Value;
                rozszerzenie = match.Groups[7].Value;

                strona_start = -1;
                strona_end = -1;

                if (nr_archiwum.StartsWith("0") && nr_archiwum.Length > 1) nr_archiwum = ZERO.Replace(nr_archiwum, "", 1);
                if (nr_zespolu.StartsWith("0") && nr_zespolu.Length > 1) nr_zespolu = ZERO.Replace(nr_zespolu, "", 1);
                if (cd_zespolu.StartsWith("0") && cd_zespolu.Length > 1) cd_zespolu = ZERO.Replace(cd_zespolu, "", 1);
                if (seria.StartsWith("0") && seria.Length > 1) seria = ZERO.Replace(seria, "", 1);
                if (sygnatura.StartsWith("0") && sygnatura.Length > 1) sygnatura = ZERO.Replace(sygnatura, "", 1);

                match = NUMER_STRONY.Match(strona);
                if (match != null && match.Success)
                {
                    try
                    {
                        strona_start = int.Parse(match.Groups[1].Value);
                        string end = match.Groups[3].Value;
                        if (end != null && end.Length > 0) strona_end = int.Parse(end);
                        else strona_end = -1;
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("" + match.Groups[3].Value);
                    }
                }

            }
        }

        public Skan(Skan skan) : this(skan.PobierzPlik()) { }

        private void pobierzMetadane()
        {
            metadane = new Hashtable();
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
                    foreach (Match m in EXIF_INFO.Matches(output)) metadane.Add(m.Groups[1].Value.Trim(), m.Groups[2].Value.Trim());
                    exeProcess.WaitForExit();
                }
            }
            catch
            {
                MessageBox.Show("Błąd odczytu informacji o skanie: " + plik, "Błąd");
            }
        }
        
        public string PobierzPlik()
        {
            return plik;
        }
        public string GenerujNazwePliku()
        {
            return nr_archiwum + "_" + nr_zespolu + "_" + cd_zespolu + "_" + seria + "_" + sygnatura + "_" + strona + "." + rozszerzenie;
        }
        public string GenerujStruktureKatalogow()
        {
            return nr_archiwum + "\\" + nr_zespolu + "\\" + cd_zespolu + "\\" + seria + "\\" + sygnatura + "\\";
        }

        public string GenerujCRC()
        {
            using (FileStream stream = File.OpenRead(plik))
            {
                var sha = new SHA256Managed();
                byte[] checksum = sha.ComputeHash(stream);
                return BitConverter.ToString(checksum).Replace("-", String.Empty);
            }
        }
    }
}
