using System;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Xml;

namespace EplanGetProjectVersion
{
    internal class Program
    {
        private static void Main(string[] args)
        {
            try
            {
                string ProjectFile = args[0];
                const string QUOTE = "\"";
                const string NEWLINE = "\n";
                string IniFile = Path.GetDirectoryName(Assembly.GetEntryAssembly().Location) + @"\Settings.ini";
                string SevenZip = Path.GetDirectoryName(Assembly.GetEntryAssembly().Location) + @"\7za.exe";
                string TempDir = Path.GetTempPath();
                string TempFile = TempDir + "ProjectInfo.xml";
                string clArg = "e " + QUOTE + ProjectFile + QUOTE + " -o" + QUOTE + TempDir + QUOTE +
                               " ProjectInfo.xml -r";

                Console.WriteLine(
                    "==> INI file: " + IniFile + NEWLINE +
                    "==> 7zip location: " + SevenZip + NEWLINE +
                    "==> Project: " + ProjectFile + NEWLINE +
                    "==> Temp directory: " + TempDir + NEWLINE +
                    "==> Temp file: " + TempFile +
                    "==> Argument: " + clArg
                    );

                Console.WriteLine("<<< Search for 'ProjectInfo.xml'... >>>");

                // Delete tempfile
                if (File.Exists(TempFile))
                {
                    File.Delete(TempFile);
                    Console.WriteLine("==> TempFile deleted.");
                }

                // Unzip
                Console.WriteLine("<<< Arguments >>>");
                Process t7zip = new Process();
                t7zip.StartInfo.FileName = SevenZip;
                t7zip.StartInfo.Arguments = clArg;
                t7zip.StartInfo.WindowStyle = ProcessWindowStyle.Hidden;
                t7zip.StartInfo.RedirectStandardOutput = true;
                t7zip.StartInfo.UseShellExecute = false;
                t7zip.EnableRaisingEvents = true;
                t7zip.Start();
                Console.WriteLine(t7zip.StandardOutput.ReadToEnd());

                // Check ProjectInfo.xml
                if (!File.Exists(TempFile))
                {
                    Console.WriteLine("==> No 'ProjectInfo.xml' found :(");
                    Console.ReadLine();
                    return;
                }

                // Search for XML property
                Console.WriteLine("<<< Search for last used version... >>>");
                string LastVersion = ReadXml(TempFile, 10043);
                Console.WriteLine("==> " + LastVersion);

                // EPLAN start
                string[] Ini = File.ReadAllLines(IniFile);
                foreach (string s in Ini)
                {
                    string[] setting = s.Split('\t');
                    if (setting[0].Equals(LastVersion))
                    {
                        Process.Start(setting[1]);
                        Console.WriteLine("==> EPLAN " + LastVersion + " started...");
                        return;
                    }
                }

                // Wait
                Console.WriteLine("==> No EPLAN " + LastVersion + " found in settings file :(");
                Console.ReadLine();
                return;
            }
            catch (Exception ex)
            {
                Console.WriteLine("ERROR: " + ex);
                Console.ReadLine();
            }
        }

        private static string ReadXml(string filename, int ID)
        {
            string strLastVersion;

            try
            {
                XmlTextReader reader = new XmlTextReader(filename);
                while (reader.Read())
                {
                    if (reader.HasAttributes)
                    {
                        while (reader.MoveToNextAttribute())
                        {
                            if (reader.Name == "id")
                            {
                                if (reader.Value == ID.ToString())
                                {
                                    strLastVersion = reader.ReadString();
                                    reader.Close();
                                    return strLastVersion;
                                }
                            }
                        }
                    }

                }
            }
            catch (Exception ex)
            {
                return ex.Message;
            }
            return "Property 10043 not found :(";
        }

    } // Class
} // Namespace