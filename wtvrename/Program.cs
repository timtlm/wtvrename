using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml;

using Shell32;

namespace wtvrename
{
    class Program
    {
        private static string ApiKey = "C555F3021FFC7B90";

        private static string seasonEpisodeFormat = @"S\d\dE\d\d";
        private static string skipPrefixFormat = "^__SKIP__";

        private static List<string> cleanupFileList = new List<string>();
        private static int countFilesProcessed = 0;
        private static string logFileName = "wtvrename_log.txt";

        static void Main(string[] args)
        {
            DumpLog();
            Log("---------START PROCESSING----------");
            RenameThem();
            Log("---------END PROCESSING------------");
        }

        private static void RenameThem()
        {
            DirectoryInfo d = new DirectoryInfo(Directory.GetCurrentDirectory());
            FileInfo[] infos = d.GetFiles("*.wtv");
            foreach (FileInfo f in infos)
            {
                try
                {
                    if (!Regex.Match(f.Name, seasonEpisodeFormat).Success && !Regex.Match(f.Name, skipPrefixFormat).Success)
                    {
                        Log("Processing " + f.Name + "....");
                        ShellClass sh = new ShellClass();

                        List<DetailedFileInfo> fileDetails = GetDetailedFileInfo(f.FullName);

                        string seriesName = (from a in fileDetails
                                             where a.Name == "Title"
                                             select a.Value).FirstOrDefault();

                        string episodeName = (from a in fileDetails
                                           where a.Name == "Episode name"
                                           select a.Value).FirstOrDefault();

                        string airDate = (from a in fileDetails
                                          where a.Name == "Broadcast date"
                                          select a.Value).FirstOrDefault();

                        string rerun = (from a in fileDetails
                                        where a.Name == "Rerun"
                                        select a.Value).FirstOrDefault();

                        // If the series name is null it is likely the recording is still in progress
                        // So, we want to skip these files
                        if (seriesName == null)
                        {
                            Log("Skipping. Recording likely in progress.");
                            break;
                        }

                        //seriesName = "The Office";
                        //episodeName = "The List";
                        //airDate = "1/1/2013";
                        //rerun = "Yes";

                        // Replace hidden characters that cause problems with date parsing
                        airDate = Regex.Replace(airDate, @"[^a-zA-Z0-9%/\: ._]", string.Empty);

                        string episode = GetEpisode(seriesName, episodeName, DateTime.Parse(airDate));

                        // Do the renaming here
                        if (episode != null)
                        {
                            string fileName = episode + f.Extension;
                            File.Move(f.FullName, Path.Combine(f.DirectoryName, fileName));
                            Log(episode);
                            countFilesProcessed++;
                        }
                        else
                        {
                            //File.Move(f.FullName, Path.Combine(f.DirectoryName, "__UNKNOWN__" + f.Name));
                            Log("Episode not found");
                        }
                    }
                }
                catch (Exception e)
                {
                    Log(e.Message);
                }
            }
            CleanupXmlFiles();
            Log("TOTAL FILES PROCESSED: " + countFilesProcessed);
        }

        // Gets a list of all of the extended properties for a file
        // (right-click > properties > details)
        private static List<DetailedFileInfo> GetDetailedFileInfo(string sFile)
        {
            List<DetailedFileInfo> aReturn = new List<DetailedFileInfo>();
            if (sFile.Length > 0)
            {
                try
                {
                    ShellClass sh = new ShellClass();
                    Folder dir = sh.NameSpace(Path.GetDirectoryName(sFile));
                    FolderItem item = dir.ParseName(Path.GetFileName(sFile));
                    for (int i = 0; i < 300; i++)
                    {
                        string det = dir.GetDetailsOf(item, i);
                        if (det != "")
                        {
                            DetailedFileInfo oFileInfo = new DetailedFileInfo();
                            oFileInfo.ID = i;
                            oFileInfo.Value = det;
                            oFileInfo.Name = dir.GetDetailsOf(null, i);

                            aReturn.Add(oFileInfo);
                        }
                    }

                }
                catch (Exception)
                {

                }
            }
            return aReturn;
        }

        private static string GetEpisode(string seriesName, string episodeName, DateTime airDate)
        {
            string seriesUrl = "http://thetvdb.com/api/GetSeries.php?seriesname=" + seriesName;
            XmlDocument seriesListXml = new XmlDocument();
            seriesListXml.Load(seriesUrl);
            XmlNodeList seriesList = seriesListXml.SelectNodes("//Data/Series");

            string seriesId = "0";
            XmlNode episode = null;

            // Find episode by air date to get the series id
            // If we have an episode name use that to find the episode in the series
            // Otherwise we use the episode found by air date
            // Sometimes there are multiple episodes for the same air date
            foreach (XmlNode series in seriesList)
            {
                seriesId = series.SelectSingleNode("seriesid").InnerText;
                episode = GetEpisodeByAirDate(seriesId, airDate);

                if (episode != null && episodeName != null)
                {
                    episode = GetEpisodeByName(seriesId, episodeName);
                    break;
                }
            }

            if (episode != null)
            {
                int seasonNumber = Convert.ToInt32(episode.SelectSingleNode("SeasonNumber").InnerText);
                int episodeNumber = Convert.ToInt32(episode.SelectSingleNode("EpisodeNumber").InnerText);
                return seriesName + " - " + "S" + seasonNumber.ToString("00") + "E" + episodeNumber.ToString("00");
            }
            else
            {
                return null;
            }
        }

        private static XmlNode GetEpisodeByAirDate(string seriesId, DateTime airDate)
        {
            string Url = "http://thetvdb.com/api/GetEpisodeByAirDate.php?apikey=" + ApiKey + "&seriesid=" + seriesId + "&airdate=" + airDate.ToString("yyyy-MM-dd");
            XmlDocument episodeXml = new XmlDocument();
            episodeXml.Load(Url);

            XmlNode episode = episodeXml.SelectSingleNode("//Episode");
            if (episode.ChildNodes.Count > 0)
            {
                return episode;
            }
            else
            {
                return null;
            }
        }

        private static XmlNode GetEpisodeByName(string seriesId, string episodeName)
        {
            XmlDocument seriesXml = new XmlDocument();
            try
            {
                seriesXml.Load(seriesId + ".xml");
            }
            catch
            {
                seriesXml.Load("http://www.thetvdb.com/api/" + ApiKey + "/series/" + seriesId + "/all/");
                seriesXml.Save(seriesId + ".xml");
                cleanupFileList.Add(seriesId + ".xml");
            }

            XmlNodeList episodeList = seriesXml.SelectNodes("//Data/Episode");
            foreach (XmlNode e in episodeList)
            {
                string eName = e.SelectSingleNode("EpisodeName").InnerText;
                if (eName == episodeName)
                {
                    return e;
                }
            }

            return null;
        }

        private static void CleanupXmlFiles()
        {
            foreach (string s in cleanupFileList)
            {
                File.Delete(s);
            }
        }

        private static void Log(string logText)
        {
            StreamWriter log;

            if (!File.Exists(logFileName))
            {
              log = new StreamWriter(logFileName);
            }
            else
            {
              log = File.AppendText(logFileName);
            }

            log.WriteLine(DateTime.Now + " - " + logText);
            log.Close();
        }

        private static void DumpLog()
        {
            File.WriteAllText(logFileName, String.Empty);
        }
    }
}
