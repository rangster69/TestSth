using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using System.Windows.Forms;
using System.Xml;
using System.IO;
using Newtonsoft.Json;
using System.Net;
using Newtonsoft.Json.Linq;


namespace Cricket
{
    public class WebScraping
    {
        // The suffix will change every year
        public const string sBaseURL = "https://www.espncricinfo.com/series/vitality-blast-2022-1297657";
        public const string sCOMPETITION_NAME = "T20Blast-South";
        List<Player> NewPlayers = new List<Player>();

        private string GetDate(int iArrNo, JObject obj)
        {
            JToken jtok;
            jtok = obj["props"]["appPageProps"]["data"]["data"]["content"]["matches"][iArrNo]["startTime"];
            string sStartTime = jtok.Value<string>();
            jtok = obj["props"]["appPageProps"]["data"]["data"]["content"]["matches"][iArrNo]["ground"]["town"]["timezone"];
            string sTimeZone = jtok.Value<string>();

            DateTime dtStartTime = DateTime.ParseExact(sStartTime, "MM/dd/yyyy HH:mm:ss", System.Globalization.CultureInfo.InvariantCulture);
            TimeZoneInfo tzi1 = OlsonTimeZoneToTimeZoneInfo(sTimeZone);
            TimeZoneInfo tzi2 = OlsonTimeZoneToTimeZoneInfo("Australia/Brisbane");
            TimeSpan tsTimeDiifBrisMinusLocal = tzi2.BaseUtcOffset - tzi1.BaseUtcOffset;
            DateTime BrisbaneDateTime = dtStartTime.Add(tsTimeDiifBrisMinusLocal);

            string sStart = BrisbaneDateTime.ToString("dd=MM=yyyy");
            DateTime dtStart = BrisbaneDateTime;

            return sStart;
        }


        private int GetArrNo(HighestGameAndLastGame hglg, string sMatchResultString)
        {
            SlugObjctID soi = new SlugObjctID();
            ExceWebScraping ews = new ExceWebScraping();
            JObject objRes = CreateJObject(sMatchResultString);
            JToken jtok;

            string sStart = GetDate(0, objRes);
            DateTime dtCurrGame = DateTime.ParseExact(hglg.sStart, "dd=MM=yyyy", null);
            DateTime dtLastGameCompleted = DateTime.ParseExact(sStart, "dd=MM=yyyy", null);
            int iCmp = dtCurrGame.CompareTo(dtLastGameCompleted);
            if (iCmp > 0)
            {
                return -1;
            }
            int iCountRes = objRes["props"]["appPageProps"]["data"]["data"]["content"]["matches"].Count();
            int iMaxArrNo = iCountRes - hglg.iHighestGame;
            string sSlug = "";
            string sTeamHome = "";
            string sTeamAway = "";
            string sTeamHomeAbbrev = "";
            string sTeamAwayAbbrev = "";
            for (int i = 0; i <= iCountRes; i++)
            {
                jtok = objRes["props"]["appPageProps"]["data"]["data"]["content"]["matches"][i]["slug"];
                sSlug = jtok.Value<string>();
                jtok = objRes["props"]["appPageProps"]["data"]["data"]["content"]["matches"][i]["teams"][0]["team"]["slug"];
                sTeamHome = jtok.Value<string>();
                sTeamHomeAbbrev = ews.GetTeamNameAbbrev(sTeamHome);
                jtok = objRes["props"]["appPageProps"]["data"]["data"]["content"]["matches"][i]["teams"][1]["team"]["slug"];
                sTeamAway = jtok.Value<string>();
                sTeamAwayAbbrev = ews.GetTeamNameAbbrev(sTeamAway);

                if ((sTeamHomeAbbrev == hglg.sTeamHome && sTeamAwayAbbrev == hglg.sTeamAway) || (sTeamHomeAbbrev == hglg.sTeamAway && sTeamAwayAbbrev == hglg.sTeamHome))
                {
                    string sCurrStart = GetDate(i, objRes);
                    if (sCurrStart == hglg.sStart)
                    {
                        return i;
                    }
                }
            }
            MessageBox.Show("Error in GetMatchNo: No Match Found");
            return -2;
        }

        public bool Main(CricketForm fm)
        {
            XmlDocument xmlDocMatch = new XmlDocument();
            WebClient WbClnt = new WebClient();
            ExceWebScraping ews = new ExceWebScraping();
            HighestGameAndLastGame hglg = new HighestGameAndLastGame();

            NewPlayers = new List<Player>();
            hglg = ews.GetLastMatchNo(fm);
            string sMtchRslts = sBaseURL + "/match-results";
            string sMtchReslt = WbClnt.DownloadString(sMtchRslts);
            string sMatchResultString = GetMatchResultString(sMtchReslt);
            int iArrNo = GetArrNo(hglg, sMatchResultString);
            if (iArrNo == -1)
            {
                fm.AddData("\r\n\r\n" + "Game number " + hglg.iHighestGame + "  has NOT been played yet or is greater than the total number of games for the season\r\n");
                fm.AddData("\r\n\r\n" + "All operations have been completed successfully!\r\n\r\n");

                return true;
            }
            MatchStrings ms = GetURLs(hglg.iHighestGame, iArrNo, fm);
            if (ms == null)
            {
                fm.AddData("\r\n\r\n" + "All operations have been completed successfully!\r\n\r\n");
                System.Media.SystemSounds.Exclamation.Play();
                return true;
            }

            string sMatchSource = WbClnt.DownloadString(ms.sBallByBall);
            string sMatchString = GetMatchResultString(sMatchSource);

            string sScrCrdSource = WbClnt.DownloadString(ms.sFullScoreCard);
            string sScrCrdString = GetMatchResultString(sScrCrdSource);

            //string sMatchResult = WbClnt.DownloadString(ms.sMatchResults);
            //string sMatchResultString = GetMatchResultString(sMatchResult);

            string sBallByBall = WbClnt.DownloadString(ms.sBallByBall);
            string sBallByBallString = GetMatchResultString(sBallByBall);

            string sCompName = GetCompName(sMatchResultString, fm);

            LoadBatsmenData(0, ms.iMatchNumber, iArrNo, sMatchString, sScrCrdString, sMatchResultString, ews, fm);
            LoadBatsmenData(1, ms.iMatchNumber, iArrNo, sMatchString, sScrCrdString, sMatchResultString, ews, fm);

            ResultsBowler rsbowlerFrstInn = new ResultsBowler();
            ResultsBowler rsbowlerSndInn = new ResultsBowler();

            rsbowlerFrstInn.IntegerFour = 0;
            rsbowlerFrstInn.IntegerSix = 0;
            rsbowlerFrstInn = LoadBowlingData(0, ms.iMatchNumber, iArrNo, sMatchString, sScrCrdString, sMatchResultString, rsbowlerFrstInn, ews, fm);

            rsbowlerSndInn.IntegerFour = 0;
            rsbowlerSndInn.IntegerSix = 0;
            rsbowlerSndInn = LoadBowlingData(1, ms.iMatchNumber, iArrNo, sMatchString, sScrCrdString, sMatchResultString, rsbowlerSndInn, ews, fm);

            string sMatchCompString = GetMatchResultString(sMatchSource);

            if (hglg.bIsFirstGame)
            {
                ews.CalcNext();
            }
            if (!hglg.bIsLastGame)
            {
                ews.CalcNext();
            }

            LoadMatchResultData(0, ms.iMatchNumber, iArrNo, sMatchCompString, sScrCrdString, sMatchResultString, rsbowlerFrstInn, ews, fm);
            LoadMatchResultData(1, ms.iMatchNumber, iArrNo, sMatchCompString, sScrCrdString, sMatchResultString, rsbowlerSndInn, ews, fm);

            ews.FormatTeam();
            ews.UpdateAllPlayersDOB();



            List<Player> ExistingPlayers = new List<Player>();
            ExistingPlayers = ews.LdPlyrLst();
            ews.UpdatePlyrLst(NewPlayers, ExistingPlayers);

            ews.CreateBackups(ms.iMatchNumber);

            fm.AddData("\r\n\r\n" + "All operations have been completed successfully!\r\n\r\n");
            System.Media.SystemSounds.Exclamation.Play();

            return false;
        }

        public bool LoadFixtures(CricketForm fm)
        {
            HighestGameAndLastGame hglg = new HighestGameAndLastGame();
            ExceWebScrapingDraw ewsd = new ExceWebScrapingDraw();
            WebClient WbClnt = new WebClient();
            ExceWebScraping ews = new ExceWebScraping();

            int iLastInDrawCurrent = ewsd.GetLastInDrawCurrent(fm);

            string sMatchResults = sBaseURL + "/match-schedule-fixtures";
            string sMatchResultsSource = WbClnt.DownloadString(sMatchResults);
            string sMatchResultsString = GetMatchResultString(sMatchResultsSource);

            LoadAllFixtures(iLastInDrawCurrent + 1, sMatchResultsString, ewsd, fm);

            fm.AddData("FIXTURES in Sheet:DrawCurrent have been loaded successfully!\r\n\r\n");
            System.Media.SystemSounds.Exclamation.Play();

            return true;
        }

        public bool LoadResults(CricketForm fm)
        {
            HighestGameAndLastGame hglg = new HighestGameAndLastGame();
            ExceWebScrapingDraw ewsd = new ExceWebScrapingDraw();
            WebClient WbClnt = new WebClient();
            ExceWebScraping ews = new ExceWebScraping();

            int iLastInDrawCurrent = ewsd.GetLastInDrawCurrent(fm);

            string sMatchResults = sBaseURL + "/match-results";
            string sMatchResultsSource = WbClnt.DownloadString(sMatchResults);
            string sMatchResultsString = GetMatchResultString(sMatchResultsSource);

            LoadAllResults(iLastInDrawCurrent + 1, sMatchResultsString, ewsd, fm);

            fm.AddData("RESULTS in Sheet:DrawCurrent have been loaded successfully!\r\n\r\n");
            System.Media.SystemSounds.Exclamation.Play();

            return true;
        }

        private JObject CreateJObject(string sString)
        {
            byte[] byteArrayComp = Encoding.UTF8.GetBytes(sString);
            MemoryStream sStream = new MemoryStream(byteArrayComp);
            XmlDocument xmlDocMatchComp = new XmlDocument();
            xmlDocMatchComp.Load(sStream);
            XmlNode root = xmlDocMatchComp.FirstChild;
            string sXML = root.InnerXml;
            object json = Newtonsoft.Json.JsonConvert.DeserializeObject(sXML);
            string jsonString = JsonConvert.SerializeObject(json, Newtonsoft.Json.Formatting.None);
            JObject obj = JObject.Parse(jsonString);

            return obj;
        }

        private string GetCompName(string sMatchResultString, CricketForm fm)
        {
            JObject obj = CreateJObject(sMatchResultString);

            // CompetitionName
            JToken jtok = obj["props"]["appPageProps"]["data"]["data"]["series"]["longName"];
            string sCompName = jtok.Value<string>();

            return sCompName;
        }
        public bool LoadAllInDraw(CricketForm fm)
        {
            LoadResults(fm);
            LoadFixtures(fm);
            return true;
        }

        private void LoadMatchScheduleData(int iNewMatch, string ssMatchScheduleFixturesString, ExceWebScrapingDraw ewsd, CricketForm fm)
        {
            DOBchanged dc = new DOBchanged();
            XmlDocument xmlDocMatch = new XmlDocument();

            JObject obj = CreateJObject(ssMatchScheduleFixturesString);

            int iCount = obj["props"]["appPageProps"]["data"]["data"]["content"]["matches"].Count();
            JToken jtok;
            string sStartTime = "";
            string sTimeZone = "";
            string sTeamHome = "";
            string sTeamAway = "";
            string sGroundName = "";
            for (int i = 0; i < iCount; i++)
            {
                jtok = obj["props"]["appPageProps"]["data"]["data"]["content"]["matches"][i]["teams"][0]["team"]["slug"];
                sTeamHome = jtok.Value<string>();
                jtok = obj["props"]["appPageProps"]["data"]["data"]["content"]["matches"][i]["teams"][1]["team"]["slug"];
                sTeamAway = jtok.Value<string>();
                jtok = obj["props"]["appPageProps"]["data"]["data"]["content"]["matches"][0]["startTime"];
                sStartTime = jtok.Value<string>();
                jtok = obj["props"]["appPageProps"]["data"]["data"]["content"]["matches"][0]["ground"]["name"];
                sGroundName = jtok.Value<string>();
                jtok = obj["props"]["appPageProps"]["data"]["data"]["content"]["matches"][0]["ground"]["town"]["timezone"];
                sTimeZone = jtok.Value<string>();

                if (!(sTeamHome == "tba" || sTeamAway == "tba"))
                {
                    DateTime dtStartTime = DateTime.ParseExact(sStartTime, "MM/dd/yyyy HH:mm:ss", System.Globalization.CultureInfo.InvariantCulture);
                    TimeZoneInfo tzi1 = OlsonTimeZoneToTimeZoneInfo(sTimeZone);
                    TimeZoneInfo tzi2 = OlsonTimeZoneToTimeZoneInfo("Australia/Brisbane");
                    TimeSpan tsTimeDiifBrisMinusLocal = tzi2.BaseUtcOffset - tzi1.BaseUtcOffset;
                    //TimeSpan a2 = tzi1.BaseUtcOffset + tsTimeDiifBrisMinusLocal;
                    DateTime BrisbaneDateTime = dtStartTime.Add(tsTimeDiifBrisMinusLocal);

                    GameInDraw gid = new GameInDraw();
                    gid.iGameNo = iNewMatch;
                    gid.sTeamHome = ewsd.GetTeamNameAbbrev(sTeamHome);
                    gid.sTeamAway = ewsd.GetTeamNameAbbrev(sTeamAway);
                    gid.sGround = ewsd.GetGroundNameAbbrev(sGroundName);

                    gid.sStart = BrisbaneDateTime.ToString("dd=MM=yyyy");

                    gid.dtStart = BrisbaneDateTime;
                    ewsd.AddGameInDraw(gid);
                }
                else
                {
                    return;
                }
            }
            fm.AddData("\r\n\r\n" + "All operations have been completed successfully!\r\n\r\n");
            System.Media.SystemSounds.Exclamation.Play();
        }

        public bool LoadMatchSchedule(CricketForm fm)
        {
            HighestGameAndLastGame hglg = new HighestGameAndLastGame();
            ExceWebScrapingDraw ewsd = new ExceWebScrapingDraw();
            WebClient WbClnt = new WebClient();
            ExceWebScraping ews = new ExceWebScraping();

            hglg = ews.GetLastMatchNo(fm);
            int iHghstMtchSoFar = hglg.iHighestGame;
            int iLastGame = hglg.iLastGame;
            int iNewMatch = iHghstMtchSoFar + 1;

            int iLastInDrawCurrent = ewsd.GetLastInDrawCurrent(fm);

            if ((iLastInDrawCurrent != -1) && (iNewMatch > iLastInDrawCurrent))
            {
                string sMatchScheduleFixtures = sBaseURL + "/match-schedule-fixtures";
                string ssMatchScheduleFixturesSource = WbClnt.DownloadString(sMatchScheduleFixtures);
                string ssMatchScheduleFixturesString = GetMatchResultString(ssMatchScheduleFixturesSource);

                fm.AddData("\r\n Loading game number " + iNewMatch.ToString() + " in Sheet:DrawCurrent\r\n");


                LoadMatchScheduleData(iNewMatch, ssMatchScheduleFixturesString, ewsd, fm);

                fm.AddData("Game number " + iNewMatch.ToString() + " has been added in Sheet:DrawCurrent\r\n");
                fm.AddData("\r\n\r\n" + "All operations have been completed successfully!\r\n\r\n");
                System.Media.SystemSounds.Exclamation.Play();

                return true;
            }
            else
            {
                fm.AddData("\r\nCan't Load Match " + iNewMatch.ToString() + " in Sheet:DrawCurrent because this match is already recorded\r\n");
                return false;
            }
        }


        private int GetMatchInteger(string sLastMatch)
        {
            string a = sLastMatch;
            string b = string.Empty;
            int val = -1;
            for (int i = 0; i < a.Length; i++)
            {
                if (Char.IsDigit(a[i]))
                    b += a[i];
                else if (b.Length != 0)
                    break;
            }
            if (b.Length > 0)
            {
                val = int.Parse(b);
            }
            return val;
        }

        private string GetMatchResultString(string sMatchURL)
        {
            if (sMatchURL == null)
            {
                MessageBox.Show("Error in GetBaseURLxmlString: No valid BaseURL source code");
                return null;
            }
            if (sMatchURL == "")
            {
                MessageBox.Show("Error in GetBaseURLxmlString: No valid BaseURL source code");
                return "";
            }

            // Get a substring between two strings     
            int iFirstPosition = sMatchURL.IndexOf("__NEXT_DATA__") - 12;
            int iSecondPosition = sMatchURL.IndexOf("</html>");
            string sNEXTDATAtoEnd = sMatchURL.Substring(iFirstPosition,
            iSecondPosition - iFirstPosition + 7);

            iFirstPosition = 0;
            iSecondPosition = sNEXTDATAtoEnd.IndexOf("</script>") + 9;
            string sNEXTDATAtoSCRIPT = sNEXTDATAtoEnd.Substring(iFirstPosition,
            iSecondPosition - iFirstPosition);
            return sNEXTDATAtoSCRIPT;
        }




        private string GameNoIntToString(int iGameNo)
        {
            string sGameNo = iGameNo.ToString();
            char lastCharacter = sGameNo[sGameNo.Length - 1];
            if (iGameNo != 11 && iGameNo != 12 && iGameNo != 13)
            {
                switch (lastCharacter)
                {
                    case '1':
                        return sGameNo + "st";
                    case '2':
                        return sGameNo + "nd";
                    case '3':
                        return sGameNo + "rd";
                    default:
                        return sGameNo + "th";
                }
            }
            return "Error01";
        }

        private int GetMatchNo(string sSlug)
        {
            string[] aAll;

            aAll = sSlug.Split('-');
            string sMatchNum = "";

            int iNoElements = aAll.Count();

            if (iNoElements == 10)
            {
                sMatchNum = aAll[8];
            }
            if (iNoElements == 9)
            {
                sMatchNum = aAll[7];
            }
            else if (iNoElements == 8)
            {
                sMatchNum = aAll[6];
            }
            else if (iNoElements == 7)
            {
                sMatchNum = aAll[5];
            }
            else
            {
                sMatchNum = aAll[4];
            }
            return Convert.ToInt32(sMatchNum.Remove(sMatchNum.Length - 2));
        }

        private bool PlayerAlreadyInNewPlayerList(Player plyr, List<Player> NewPlayers, ExceWebScraping ews)
        {
            for (int i = 0; i < NewPlayers.Count; i++)
            {
                if (ews.IsSamePlayer(plyr, NewPlayers[i]))
                {
                    return true;
                }
            }
            return false;
        }

        /*
        private void PlayMP3(string sFilePath)
        {
            Sounds sounds = new Sounds();
            sounds.PlayFile(sFilePath);
        }
        */

        private string TestStringForIllegals(string sText, CricketForm fm)
        {
            string sNewText = "";

            for (int i = 0; i < sText.Length; i++)
            {
                char c = sText[i];
                if (c == 'Â')
                {
                    fm.AddData("\r\nIllegal Character " + c.ToString() + " is in Text\r\n");
                    sNewText = sNewText + ' ';
                    //return null;
                }
                else
                {
                    sNewText = sNewText + c.ToString();
                }
            }
            return sNewText;
        }

        private Player FormatFullname(string sFullName, CricketForm fm)
        {
            Player plyr = new Player();
            string[] aAll;
            bool bNextIsLower = false;
            char cFirstChar;
            char cNextChar;
            bool bIsLower;

            if (sFullName == "Naveen-ul-Haq")
            {
                plyr.LastName = "Naveen-ul-Haq";
                plyr.MiddleName = "---";
                plyr.FirstName = "---";
                return plyr;
            }

            string sFllNm = TestStringForIllegals(sFullName, fm);
            aAll = sFllNm.Split(' ');
            int iAllCount = aAll.Count();
            if (iAllCount > 5)
            {
                MessageBox.Show("Error: There is MORE than 5 names for Player " + sFullName);
                plyr.FirstName = "---";
                plyr.MiddleName = "---";
                plyr.LastName = "---";
            }
            else if (iAllCount == 5)
            {
                plyr.FirstName = aAll[0].Trim();
                plyr.MiddleName = aAll[1].Trim();

                cFirstChar = aAll[2][0];
                bIsLower = Char.IsLower(cFirstChar);
                if (bIsLower)
                {
                    plyr.LastName = aAll[2].Trim() + " " + aAll[3].Trim() + " " + aAll[4].Trim();
                }
                else
                {
                    plyr.MiddleName = plyr.MiddleName + " " + aAll[2].Trim();
                    cNextChar = aAll[3][0];
                    bNextIsLower = Char.IsLower(cNextChar);
                    if (bNextIsLower)
                    {
                        plyr.LastName = aAll[3].Trim() + " " + aAll[4].Trim();
                    }
                    else
                    {
                        plyr.MiddleName = plyr.MiddleName + " " + aAll[3].Trim();
                        plyr.LastName = aAll[4].Trim();
                    }
                }
            }
            else if (iAllCount == 4)
            {
                plyr.FirstName = aAll[0].Trim();

                cFirstChar = aAll[1][0];
                bIsLower = Char.IsLower(cFirstChar);
                if (bIsLower)
                {
                    plyr.MiddleName = "---";
                    plyr.LastName = aAll[1].Trim() + " " + aAll[2].Trim() + " " + aAll[3].Trim();
                }
                else
                {
                    plyr.MiddleName = aAll[1].Trim();
                    cNextChar = aAll[2][0];
                    bNextIsLower = Char.IsLower(cNextChar);
                    if (bNextIsLower)
                    {
                        plyr.LastName = aAll[2].Trim() + " " + aAll[3].Trim();
                    }
                    else
                    {
                        plyr.MiddleName = plyr.MiddleName + " " + aAll[2].Trim();
                        plyr.LastName = aAll[3].Trim();
                    }
                }
            }
            else if (iAllCount == 3)
            {
                plyr.FirstName = aAll[0].Trim();

                cFirstChar = aAll[1][0];
                bIsLower = Char.IsLower(cFirstChar);
                if (bIsLower)
                {
                    plyr.MiddleName = "---";
                    plyr.LastName = aAll[1].Trim() + " " + aAll[2].Trim();
                }
                else
                {
                    plyr.MiddleName = aAll[1].Trim();
                    plyr.LastName = aAll[2].Trim();
                }
            }
            else if (iAllCount == 2)
            {
                plyr.FirstName = aAll[0].Trim();
                plyr.MiddleName = "---";
                plyr.LastName = aAll[1].Trim();
            }
            return plyr;
        }

        private Player GetPlayerData(string sFullName, string sTeamCode, string sCountryID, string sCountryName, DOBchanged dc, ExceWebScraping ews, CricketForm fm)
        {
            Player plyr = new Player();
            plyr = FormatFullname(sFullName, fm);
            plyr.DOB = ews.StringToDate(dc.NewDOB);
            plyr.TeamCode = sTeamCode.Trim();
            plyr.CountryID = sCountryID;
            plyr.Country = sCountryName;
            return plyr;
        }

        private string GetCountryName(string sCountryID)
        {
            switch (sCountryID)
            {
                case "1":
                    return "England";
                case "2":
                    return "Australia";
                case "3":
                    return "South Africa";
                case "4":
                    return "West Indies";
                case "5":
                    return "New Zealand";
                case "6":
                    return "India";
                case "8":
                    return "Sri Lanka";
                case "23":
                    return "Singapore";
                case "40":
                    return "Afghanistan";
                default:
                    return sCountryID;
            }
        }

        private MatchStrings GetURLs(int iNewMatch, int iArrNo, CricketForm fm)
        {
            MatchStrings ms = new MatchStrings();
            WebClient WbClnt = new WebClient();

            string sMatchResults = sBaseURL + "/match-results";

            string sMatchResultsSource = WbClnt.DownloadString(sMatchResults);
            string sMatchResultsString = GetMatchResultString(sMatchResultsSource);

            JObject obj = CreateJObject(sMatchResultsString);

            JToken jtok;

            int iMatchCount = obj["props"]["appPageProps"]["data"]["data"]["content"]["matches"].Count();
            //int iCurrMatchArrayNo = iArrNo;

            if (iArrNo < 0)
            {
                fm.AddData("\r\n\r\n" + "Game number " + iNewMatch.ToString() + "  has NOT been played yet or is greater than the total number of games for the season\r\n");
                return null;
            }
            jtok = obj["props"]["appPageProps"]["data"]["data"]["content"]["matches"][iArrNo]["objectId"];
            string sObjectID = jtok.Value<string>();
            jtok = obj["props"]["appPageProps"]["data"]["data"]["content"]["matches"][iArrNo]["slug"];
            string sSlug = jtok.Value<string>();

            ms.sBallByBall = sBaseURL + "/" + sSlug + "-" + sObjectID + "/ball-by-ball-commentary";

            ms.sMatchOversComparison = sBaseURL + "/" + sSlug + "-" + sObjectID + "/match-overs-comparison";
            ms.sFullScoreCard = sBaseURL + "/" + sSlug + "-" + sObjectID + "/full-scorecard";
            ms.sMatchResults = sBaseURL + "/match-results";
            ms.iMatchNumber = iNewMatch;
            return ms;

        }

        /// <summary>
        /// Converts an Olson time zone ID to a Windows time zone ID.
        /// </summary>
        /// <param name="olsonTimeZoneId">An Olson time zone ID. See http://unicode.org/repos/cldr-tmp/trunk/diff/supplemental/zone_tzid.html. </param>
        /// <returns>
        /// The TimeZoneInfo corresponding to the Olson time zone ID, 
        /// or null if you passed in an invalid Olson time zone ID.
        /// </returns>
        /// <remarks>
        /// See http://unicode.org/repos/cldr-tmp/trunk/diff/supplemental/zone_tzid.html
        /// </remarks>
        public static TimeZoneInfo OlsonTimeZoneToTimeZoneInfo(string olsonTimeZoneId)
        {
            var olsonWindowsTimes = new Dictionary<string, string>()
            {
                { "Africa/Bangui", "W. Central Africa Standard Time" },
                { "Africa/Cairo", "Egypt Standard Time" },
                { "Africa/Casablanca", "Morocco Standard Time" },
                { "Africa/Harare", "South Africa Standard Time" },
                { "Africa/Johannesburg", "South Africa Standard Time" },
                { "Africa/Lagos", "W. Central Africa Standard Time" },
                { "Africa/Monrovia", "Greenwich Standard Time" },
                { "Africa/Nairobi", "E. Africa Standard Time" },
                { "Africa/Windhoek", "Namibia Standard Time" },
                { "America/Anchorage", "Alaskan Standard Time" },
                { "America/Argentina/San_Juan", "Argentina Standard Time" },
                { "America/Asuncion", "Paraguay Standard Time" },
                { "America/Bahia", "Bahia Standard Time" },
                { "America/Bogota", "SA Pacific Standard Time" },
                { "America/Buenos_Aires", "Argentina Standard Time" },
                { "America/Caracas", "Venezuela Standard Time" },
                { "America/Cayenne", "SA Eastern Standard Time" },
                { "America/Chicago", "Central Standard Time" },
                { "America/Chihuahua", "Mountain Standard Time (Mexico)" },
                { "America/Cuiaba", "Central Brazilian Standard Time" },
                { "America/Denver", "Mountain Standard Time" },
                { "America/Fortaleza", "SA Eastern Standard Time" },
                { "America/Godthab", "Greenland Standard Time" },
                { "America/Guatemala", "Central America Standard Time" },
                { "America/Halifax", "Atlantic Standard Time" },
                { "America/Indianapolis", "US Eastern Standard Time" },
                { "America/Indiana/Indianapolis", "US Eastern Standard Time" },
                { "America/La_Paz", "SA Western Standard Time" },
                { "America/Los_Angeles", "Pacific Standard Time" },
                { "America/Mexico_City", "Mexico Standard Time" },
                { "America/Montevideo", "Montevideo Standard Time" },
                { "America/New_York", "Eastern Standard Time" },
                { "America/Noronha", "UTC-02" },
                { "America/Phoenix", "US Mountain Standard Time" },
                { "America/Regina", "Canada Central Standard Time" },
                { "America/Santa_Isabel", "Pacific Standard Time (Mexico)" },
                { "America/Santiago", "Pacific SA Standard Time" },
                { "America/Sao_Paulo", "E. South America Standard Time" },
                { "America/St_Johns", "Newfoundland Standard Time" },
                { "America/Tijuana", "Pacific Standard Time" },
                { "Antarctica/McMurdo", "New Zealand Standard Time" },
                { "Atlantic/South_Georgia", "UTC-02" },
                { "Asia/Almaty", "Central Asia Standard Time" },
                { "Asia/Amman", "Jordan Standard Time" },
                { "Asia/Baghdad", "Arabic Standard Time" },
                { "Asia/Baku", "Azerbaijan Standard Time" },
                { "Asia/Bangkok", "SE Asia Standard Time" },
                { "Asia/Beirut", "Middle East Standard Time" },
                { "Asia/Calcutta", "India Standard Time" },
                { "Asia/Colombo", "Sri Lanka Standard Time" },
                { "Asia/Damascus", "Syria Standard Time" },
                { "Asia/Dhaka", "Bangladesh Standard Time" },
                { "Asia/Dubai", "Arabian Standard Time" },
                { "Asia/Irkutsk", "North Asia East Standard Time" },
                { "Asia/Jerusalem", "Israel Standard Time" },
                { "Asia/Kabul", "Afghanistan Standard Time" },
                { "Asia/Kamchatka", "Kamchatka Standard Time" },
                { "Asia/Karachi", "Pakistan Standard Time" },
                { "Asia/Katmandu", "Nepal Standard Time" },
                { "Asia/Kolkata", "India Standard Time" },
                { "Asia/Krasnoyarsk", "North Asia Standard Time" },
                { "Asia/Kuala_Lumpur", "Singapore Standard Time" },
                { "Asia/Kuwait", "Arab Standard Time" },
                { "Asia/Magadan", "Magadan Standard Time" },
                { "Asia/Muscat", "Arabian Standard Time" },
                { "Asia/Novosibirsk", "N. Central Asia Standard Time" },
                { "Asia/Oral", "West Asia Standard Time" },
                { "Asia/Rangoon", "Myanmar Standard Time" },
                { "Asia/Riyadh", "Arab Standard Time" },
                { "Asia/Seoul", "Korea Standard Time" },
                { "Asia/Shanghai", "China Standard Time" },
                { "Asia/Singapore", "Singapore Standard Time" },
                { "Asia/Taipei", "Taipei Standard Time" },
                { "Asia/Tashkent", "West Asia Standard Time" },
                { "Asia/Tbilisi", "Georgian Standard Time" },
                { "Asia/Tehran", "Iran Standard Time" },
                { "Asia/Tokyo", "Tokyo Standard Time" },
                { "Asia/Ulaanbaatar", "Ulaanbaatar Standard Time" },
                { "Asia/Vladivostok", "Vladivostok Standard Time" },
                { "Asia/Yakutsk", "Yakutsk Standard Time" },
                { "Asia/Yekaterinburg", "Ekaterinburg Standard Time" },
                { "Asia/Yerevan", "Armenian Standard Time" },
                { "Atlantic/Azores", "Azores Standard Time" },
                { "Atlantic/Cape_Verde", "Cape Verde Standard Time" },
                { "Atlantic/Reykjavik", "Greenwich Standard Time" },
                { "Australia/Adelaide", "Cen. Australia Standard Time" },
                { "Australia/Brisbane", "E. Australia Standard Time" },
                { "Australia/Darwin", "AUS Central Standard Time" },
                { "Australia/Hobart", "Tasmania Standard Time" },
                { "Australia/Perth", "W. Australia Standard Time" },
                { "Australia/Sydney", "AUS Eastern Standard Time" },
                { "Etc/GMT", "UTC" },
                { "Etc/GMT+11", "UTC-11" },
                { "Etc/GMT+12", "Dateline Standard Time" },
                { "Etc/GMT+2", "UTC-02" },
                { "Etc/GMT-12", "UTC+12" },
                { "Europe/Amsterdam", "W. Europe Standard Time" },
                { "Europe/Athens", "GTB Standard Time" },
                { "Europe/Belgrade", "Central Europe Standard Time" },
                { "Europe/Berlin", "W. Europe Standard Time" },
                { "Europe/Brussels", "Romance Standard Time" },
                { "Europe/Budapest", "Central Europe Standard Time" },
                { "Europe/Dublin", "GMT Standard Time" },
                { "Europe/Helsinki", "FLE Standard Time" },
                { "Europe/Istanbul", "GTB Standard Time" },
                { "Europe/Kiev", "FLE Standard Time" },
                { "Europe/London", "GMT Standard Time" },
                { "Europe/Minsk", "E. Europe Standard Time" },
                { "Europe/Moscow", "Russian Standard Time" },
                { "Europe/Paris", "Romance Standard Time" },
                { "Europe/Sarajevo", "Central European Standard Time" },
                { "Europe/Warsaw", "Central European Standard Time" },
                { "Indian/Mauritius", "Mauritius Standard Time" },
                { "Pacific/Apia", "Samoa Standard Time" },
                { "Pacific/Auckland", "New Zealand Standard Time" },
                { "Pacific/Fiji", "Fiji Standard Time" },
                { "Pacific/Guadalcanal", "Central Pacific Standard Time" },
                { "Pacific/Guam", "West Pacific Standard Time" },
                { "Pacific/Honolulu", "Hawaiian Standard Time" },
                { "Pacific/Pago_Pago", "UTC-11" },
                { "Pacific/Port_Moresby", "West Pacific Standard Time" },
                { "Pacific/Tongatapu", "Tonga Standard Time" }
            };

            var windowsTimeZoneId = default(string);
            var windowsTimeZone = default(TimeZoneInfo);
            if (olsonWindowsTimes.TryGetValue(olsonTimeZoneId, out windowsTimeZoneId))
            {
                try { windowsTimeZone = TimeZoneInfo.FindSystemTimeZoneById(windowsTimeZoneId); }
                catch (TimeZoneNotFoundException) { }
                catch (InvalidTimeZoneException) { }
            }
            return windowsTimeZone;
        }

        private ResultsBowler LoadBowlingData(int iInningsNumber, int iHighestMatch, int iArrNo, string sMatchString, string sScrCrdString, string sMatchResultString, ResultsBowler resbowler, ExceWebScraping ews, CricketForm fm)
        {
            DOBchanged dc = new DOBchanged();
            XmlDocument xmlDocMatch = new XmlDocument();
            //WebClient WbClnt = new WebClient();
            int iMatchFours = 0;
            int iMatchSixes = 0;
            ResultsBowler rbwlr = new ResultsBowler();
            ExceWebScrapingDraw ewsd = new ExceWebScrapingDraw();

            JObject obj = CreateJObject(sMatchString);
            JObject objRes = CreateJObject(sMatchResultString);

            int iCountRes = objRes["props"]["appPageProps"]["data"]["data"]["content"]["matches"].Count();


            //int iMatchNo = iCountRes - iHighestMatch;
            //int iMatchNo = iArrNo;


            JToken jtokMatchStatus = objRes["props"]["appPageProps"]["data"]["data"]["content"]["matches"][iArrNo]["statusEng"];
            string sMatchStatus = jtokMatchStatus.Value<string>();

            if (sMatchStatus != "ABANDONED" && sMatchStatus != "NO RESULT")
            {
                int iCount = obj["props"]["appPageProps"]["data"]["content"]["inningsScore"][iInningsNumber]["inningBowlers"].Count();

                fm.AddData("\r\n" + "WebScraping is creating data for " + iCount.ToString() + " Bowlers in innings number " + (iInningsNumber + 1).ToString() + "\r\n\r\n");


                JObject objScrCrd = CreateJObject(sScrCrdString);

                for (int i = 0; i < iCount; i++)
                {
                    Player plyr = new Player();
                    JToken jtokPlayerName = obj["props"]["appPageProps"]["data"]["content"]["inningsScore"][iInningsNumber]["inningBowlers"][i]["player"]["longName"];
                    string sFullName = jtokPlayerName.Value<string>();
                    jtokPlayerName = obj["props"]["appPageProps"]["data"]["content"]["inningsScore"][iInningsNumber]["inningBowlers"][i]["player"]["dateOfBirth"]["year"];
                    string sDOByear = jtokPlayerName.Value<string>();
                    jtokPlayerName = obj["props"]["appPageProps"]["data"]["content"]["inningsScore"][iInningsNumber]["inningBowlers"][i]["player"]["dateOfBirth"]["month"];
                    string sDOBmonth = jtokPlayerName.Value<string>();
                    jtokPlayerName = obj["props"]["appPageProps"]["data"]["content"]["inningsScore"][iInningsNumber]["inningBowlers"][i]["player"]["dateOfBirth"]["date"];
                    string sDOBday = jtokPlayerName.Value<string>();
                    string sFullDOB = sDOBday.Trim() + "=" + sDOBmonth.Trim() + "=" + sDOByear.Trim();


                    JToken jtok = objScrCrd["props"]["appPageProps"]["data"]["content"]["scorecard"]["innings"][iInningsNumber]["inningBowlers"][i]["player"]["countryTeamId"];
                    string sCountryID = jtok.Value<string>();

                    string sCountryName = GetCountryName(sCountryID);

                    jtok = objScrCrd["props"]["appPageProps"]["data"]["content"]["scorecard"]["innings"][iInningsNumber]["inningBowlers"][i]["player"]["longName"];
                    string sRes = jtok.Value<string>();

                    int iBowlingTeam = -1;
                    if (iInningsNumber == 0)
                    {
                        iBowlingTeam = 1;
                    }
                    else
                    {
                        iBowlingTeam = 0;
                    }
                    jtokPlayerName = obj["props"]["appPageProps"]["data"]["content"]["inningsScore"][iBowlingTeam]["team"]["slug"];

                    string sTeamName = jtokPlayerName.Value<string>();
                    string sTeamCode = ewsd.GetTeamNameAbbrev(sTeamName);

                    dc = ews.UpdatePlayerDOB(sFullDOB);
                    plyr = GetPlayerData(sFullName, sTeamCode, sCountryID, sCountryName, dc, ews, fm);

                    JToken jtokBowler = obj["props"]["appPageProps"]["data"]["content"]["inningsScore"][iInningsNumber]["inningBowlers"][i]["wickets"];
                    string sWickets = jtokBowler.Value<string>();
                    jtokBowler = obj["props"]["appPageProps"]["data"]["content"]["inningsScore"][iInningsNumber]["inningBowlers"][i]["conceded"];
                    string sConceded = jtokBowler.Value<string>();
                    jtokBowler = obj["props"]["appPageProps"]["data"]["content"]["inningsScore"][iInningsNumber]["inningBowlers"][i]["overs"];
                    string sOvers = jtokBowler.Value<string>();

                    jtokBowler = obj["props"]["appPageProps"]["data"]["content"]["inningsScore"][iInningsNumber]["inningBowlers"][i]["fours"];
                    string sFoursThisBowler = jtokBowler.Value<string>();
                    jtokBowler = obj["props"]["appPageProps"]["data"]["content"]["inningsScore"][iInningsNumber]["inningBowlers"][i]["sixes"];
                    string sSixesThisBowler = jtokBowler.Value<string>();



                    iMatchFours = iMatchFours + Convert.ToInt32(sFoursThisBowler);
                    iMatchSixes = iMatchSixes + Convert.ToInt32(sSixesThisBowler);

                    rbwlr.Wickets = sWickets;
                    rbwlr.RunsConceded = sConceded;
                    rbwlr.OversBowled = sOvers;

                    string sInningsNo = "";
                    if (iInningsNumber == 0)
                    {
                        sInningsNo = "1st";
                    }
                    else
                    {
                        sInningsNo = "2nd";
                    }

                    fm.AddData("WebScraping is adding Bowler data for team " + plyr.TeamCode + " - Match Number is " + iHighestMatch.ToString() + " and is the " + sInningsNo + " Innings\r\n");
                    ews.UpdateBowler(iHighestMatch, plyr, rbwlr, fm);

                    if (PlayerAlreadyInNewPlayerList(plyr, NewPlayers, ews) == false)
                    {
                        NewPlayers.Add(plyr);
                    }
                }

                rbwlr.IntegerFour = iMatchFours;
                rbwlr.IntegerSix = iMatchSixes;


                return rbwlr;
            }
            else
            {
                jtokMatchStatus = objRes["props"]["appPageProps"]["data"]["data"]["content"]["matches"][iArrNo]["slug"];
                string sSlug = jtokMatchStatus.Value<string>();
                fm.AddData("Game number " + iHighestMatch.ToString() + " between " + sSlug + " was ABANDONED or NO RESULT - NO Bowling Data for innings number " + (iInningsNumber + 1).ToString() + " is being loaded\r\n");
                return null;
            }
        }

        private void LoadBatsmenData(int iInningsNumber, int iHighestMatch, int iArrNo, string sMatchString, string sScrCrdString, string sMatchResultString, ExceWebScraping ews, CricketForm fm)
        {
            DOBchanged dc = new DOBchanged();
            XmlDocument xmlDocMatch = new XmlDocument();
            ExceWebScrapingDraw ewsd = new ExceWebScrapingDraw();

            JObject obj = CreateJObject(sMatchString);
            JObject objRes = CreateJObject(sMatchResultString);

            int iCountRes = objRes["props"]["appPageProps"]["data"]["data"]["content"]["matches"].Count();
            //int iMatchNo = iCountRes - iHighestMatch;
            //int iMatchNo = iArrNo;

            JToken jtokMatchStatus = objRes["props"]["appPageProps"]["data"]["data"]["content"]["matches"][iArrNo]["statusEng"];
            string sMatchStatus = jtokMatchStatus.Value<string>();
            if (sMatchStatus != "ABANDONED" && sMatchStatus != "NO RESULT")
            {
                int iCount = obj["props"]["appPageProps"]["data"]["content"]["inningsScore"][iInningsNumber]["inningBatsmen"].Count();
                fm.AddData("\r\n\r\n" + "Game number is " + iHighestMatch.ToString() + " - WebScraping is creating data for several Batsmen - Batsmen Count is " + iCount.ToString() + " BUT only loading first 5 in innings number " + (iInningsNumber + 1).ToString() + "\r\n");
                JObject objScrCrd = CreateJObject(sScrCrdString);
                for (int i = 0; i < 5; i++)
                {
                    Player plyr = new Player();
                    JToken jtokPlayerName = obj["props"]["appPageProps"]["data"]["content"]["inningsScore"][iInningsNumber]["inningBatsmen"][i]["player"]["longName"];
                    string sFullName = jtokPlayerName.Value<string>();
                    jtokPlayerName = obj["props"]["appPageProps"]["data"]["content"]["inningsScore"][iInningsNumber]["inningBatsmen"][i]["player"]["dateOfBirth"]["year"];
                    string sDOByear = jtokPlayerName.Value<string>();
                    jtokPlayerName = obj["props"]["appPageProps"]["data"]["content"]["inningsScore"][iInningsNumber]["inningBatsmen"][i]["player"]["dateOfBirth"]["month"];
                    string sDOBmonth = jtokPlayerName.Value<string>();
                    jtokPlayerName = obj["props"]["appPageProps"]["data"]["content"]["inningsScore"][iInningsNumber]["inningBatsmen"][i]["player"]["dateOfBirth"]["date"];
                    string sDOBday = jtokPlayerName.Value<string>();
                    string sFullDOB = sDOBday.Trim() + "=" + sDOBmonth.Trim() + "=" + sDOByear.Trim();
                    jtokPlayerName = obj["props"]["appPageProps"]["data"]["content"]["inningsScore"][iInningsNumber]["team"]["slug"];
                    string sTeamName = jtokPlayerName.Value<string>();
                    string sTeamCode = ewsd.GetTeamNameAbbrev(sTeamName);

                    dc = ews.UpdatePlayerDOB(sFullDOB);
                    JToken jtok = objScrCrd["props"]["appPageProps"]["data"]["content"]["scorecard"]["innings"][iInningsNumber]["inningBatsmen"][i]["player"]["countryTeamId"];
                    string sCountryID = jtok.Value<string>();
                    string sCountryName = GetCountryName(sCountryID);
                    plyr = GetPlayerData(sFullName, sTeamCode, sCountryID, sCountryName, dc, ews, fm);

                    JToken jtokBatsman = obj["props"]["appPageProps"]["data"]["content"]["inningsScore"][iInningsNumber]["inningBatsmen"][i]["runs"];
                    string sRuns = jtokBatsman.Value<string>();
                    jtokBatsman = obj["props"]["appPageProps"]["data"]["content"]["inningsScore"][iInningsNumber]["inningBatsmen"][i]["balls"];
                    string sBallsFaced = jtokBatsman.Value<string>();
                    jtokBatsman = obj["props"]["appPageProps"]["data"]["content"]["inningsScore"][iInningsNumber]["inningBatsmen"][i]["fours"];
                    string sFours = jtokBatsman.Value<string>();
                    jtokBatsman = obj["props"]["appPageProps"]["data"]["content"]["inningsScore"][iInningsNumber]["inningBatsmen"][i]["sixes"];
                    string sSixes = jtokBatsman.Value<string>();
                    jtokBatsman = obj["props"]["appPageProps"]["data"]["content"]["inningsScore"][iInningsNumber]["inningBatsmen"][i]["isOut"];
                    string sIsOut = jtokBatsman.Value<string>();

                    ResultsBatsman rb = new ResultsBatsman();
                    if (sIsOut == "False")
                    {
                        rb.NotOut = true;
                    }
                    else
                    {
                        rb.NotOut = false;
                    }
                    if ((sIsOut == "False") && (Convert.ToInt32(sRuns) <= 19))
                    {
                        rb.IsSignificant = false;
                    }
                    else
                    {
                        rb.IsSignificant = true;
                    }
                    rb.Runs = sRuns;
                    rb.BallsFaced = sBallsFaced;
                    rb.Four = sFours;
                    rb.Six = sSixes;

                    string sInningsNo = "";
                    if (iInningsNumber == 0)
                    {
                        sInningsNo = "1st";
                    }
                    else
                    {
                        sInningsNo = "2nd";
                    }

                    fm.AddData("\r\n" + "WebScraping is adding Batting data for team " + plyr.TeamCode + " - Match Number is " + iHighestMatch.ToString() + " and is the " + sInningsNo + " Innings\r\n");
                    ews.UpdateBatsman(iHighestMatch, plyr, rb, fm);
                    if (PlayerAlreadyInNewPlayerList(plyr, NewPlayers, ews) == false)
                    {
                        NewPlayers.Add(plyr);
                    }
                }
            }
            else
            {
                jtokMatchStatus = objRes["props"]["appPageProps"]["data"]["data"]["content"]["matches"][iArrNo]["slug"];
                string sSlug = jtokMatchStatus.Value<string>();
                fm.AddData("Game number " + iHighestMatch.ToString() + " between " + sSlug + " was ABANDONED or NO RESULT - NO Batting Data for innings number " + (iInningsNumber + 1).ToString() + " is being loaded\r\n");

            }
        }

        private void LoadMatchResultData(int iInningsNumber, int iHighestMatch, int iArrNo, string sMatchCompString, string sScrCrdString, string sMatchResultString, ResultsBowler resbwlr, ExceWebScraping ews, CricketForm fm)
        {
            JObject objComp = CreateJObject(sMatchCompString);
            ExceWebScrapingDraw ewsd = new ExceWebScrapingDraw();
            JToken jtokComp;
            JObject objRes = CreateJObject(sMatchResultString);

            int iCountRes = objRes["props"]["appPageProps"]["data"]["data"]["content"]["matches"].Count();

            //int iMatchNo = iCountRes - iHighestMatch;
            //int iMatchNo = iArrNo;

            JToken jtokMatchStatus = objRes["props"]["appPageProps"]["data"]["data"]["content"]["matches"][iArrNo]["statusEng"];
            string sMatchStatus = jtokMatchStatus.Value<string>();

            if (sMatchStatus != "ABANDONED" && sMatchStatus != "NO RESULT")
            {

                // Match Result - need to do 4's and 6's by adding up bowlers
                jtokComp = objComp["props"]["appPageProps"]["data"]["content"]["inningsScore"][iInningsNumber]["wickets"];
                string sWktsDown = jtokComp.Value<string>();
                jtokComp = objComp["props"]["appPageProps"]["data"]["content"]["inningsScore"][iInningsNumber]["runs"];
                string sMatchScore = jtokComp.Value<string>();
                jtokComp = objComp["props"]["appPageProps"]["data"]["content"]["inningsScore"][iInningsNumber]["overs"];
                string sMatchOvers = jtokComp.Value<string>();
                jtokComp = objComp["props"]["appPageProps"]["data"]["content"]["inningsScore"][iInningsNumber]["team"]["slug"];

                string sTeamName = jtokComp.Value<string>();
                string sTeamCode = ewsd.GetTeamNameAbbrev(sTeamName);

                JObject objScrCrd = CreateJObject(sScrCrdString);
                JToken jtok;
                string sOpeningPartnership = "";
                try
                {
                    jtok = objScrCrd["props"]["appPageProps"]["data"]["content"]["scorecard"]["innings"][iInningsNumber]["inningFallOfWickets"][0]["fowRuns"];
                    sOpeningPartnership = jtok.Value<string>();
                }
                catch (Exception ex)
                {
                    sOpeningPartnership = sMatchScore;
                    fm.AddData("\r\n No wickets have fallen in this innings therefore the opening partnership is the innings score\r\n");
                }
                finally
                {
                }

                jtok = objScrCrd["props"]["appPageProps"]["data"]["content"]["scorecard"]["innings"][iInningsNumber]["inningOvers"][gbl.iXOvers - 1]["totalRuns"];
                string sXoverScore = jtok.Value<string>();
                string sCountryID = jtok.Value<string>();

                ResultsMatch rm = new ResultsMatch();
                rm.WicketsDown = sWktsDown;
                rm.TotalScore = sMatchScore;
                rm.OversFaced = sMatchOvers;
                rm.Four = resbwlr.IntegerFour.ToString().Trim();
                rm.Six = resbwlr.IntegerSix.ToString().Trim();
                rm.OpeningPartnership = sOpeningPartnership;
                rm.XOverScore = sXoverScore;

                string sInningsNo = "";
                if (iInningsNumber == 0)
                {
                    sInningsNo = "1st";
                    rm.BattedFirstOrSecond = "f";
                }
                else
                {
                    sInningsNo = "2nd";
                    rm.BattedFirstOrSecond = "s";
                }
                fm.AddData("WebScraping is adding match data for team " + sTeamCode + " - Match Number is " + iHighestMatch.ToString() + " and is the " + sInningsNo + " Innings\r\n");
                ews.UpdateMatch(iHighestMatch, resbwlr.IntegerFour, resbwlr.IntegerSix, rm, sTeamCode, fm);
            }
            else
            {
                string[] aAll;

                jtokMatchStatus = objRes["props"]["appPageProps"]["data"]["data"]["content"]["matches"][iArrNo]["slug"];
                string sTeams = jtokMatchStatus.Value<string>();
                aAll = sTeams.Split('-');
                string sTeamHome = aAll[0];
                string sTeamAway = aAll[2];
                string sTeamCode;
                if (iInningsNumber == 0)
                {
                    sTeamCode = ewsd.GetTeamNameAbbrev(sTeamHome);
                }
                else
                {
                    sTeamCode = ewsd.GetTeamNameAbbrev(sTeamAway);
                }
                ews.UpdateMatchAbandoned(iInningsNumber, iHighestMatch, sTeamCode, fm);
                fm.AddData("Game number " + iHighestMatch.ToString() + " between " + sTeams + " was ABANDONED or NO RESULT - NO Data for innings number " + (iInningsNumber + 1).ToString() + " is being loaded\r\n");
            }
        }

        private void LoadAllFixtures(int iNewMatch, string sMatchResultsString, ExceWebScrapingDraw ewsd, CricketForm fm)
        {
            DOBchanged dc = new DOBchanged();
            XmlDocument xmlDocMatch = new XmlDocument();

            JObject obj = CreateJObject(sMatchResultsString);

            int iCount = obj["props"]["appPageProps"]["data"]["data"]["content"]["matches"].Count();
            JToken jtok;
            string sStartTime = "";
            string sTimeZone = "";
            string sTeamHome = "";
            string sTeamAway = "";
            string sGroundName = "";

            int iArrNo = iCount - 1;

            for (int i = 0; i < iArrNo; i++)
            {
                jtok = obj["props"]["appPageProps"]["data"]["data"]["content"]["matches"][i]["slug"];
                string sSlug = jtok.Value<string>();
                string[] aAll;
                aAll = sSlug.Split('-');
                int iAllCount = aAll.Count();
                string sNthSth = aAll[iAllCount - 2];
                bool bIsNorth = (sNthSth == "north");
                if (sSlug == "hampshire-vs-kent-north-group")
                {
                    bIsNorth = false;
                }
                if (!bIsNorth)
                {
                    jtok = obj["props"]["appPageProps"]["data"]["data"]["content"]["matches"][i]["teams"][0]["team"]["slug"];
                    sTeamHome = jtok.Value<string>();
                    jtok = obj["props"]["appPageProps"]["data"]["data"]["content"]["matches"][i]["teams"][1]["team"]["slug"];
                    sTeamAway = jtok.Value<string>();
                    if (!(sTeamHome == "tba" || sTeamAway == "tba"))
                    {
                        jtok = obj["props"]["appPageProps"]["data"]["data"]["content"]["matches"][i]["startTime"];
                        sStartTime = jtok.Value<string>();
                        jtok = obj["props"]["appPageProps"]["data"]["data"]["content"]["matches"][i]["ground"]["name"];
                        sGroundName = jtok.Value<string>();
                        jtok = obj["props"]["appPageProps"]["data"]["data"]["content"]["matches"][i]["ground"]["town"]["timezone"];
                        sTimeZone = jtok.Value<string>();

                        DateTime dtStartTime = DateTime.ParseExact(sStartTime, "MM/dd/yyyy HH:mm:ss", System.Globalization.CultureInfo.InvariantCulture);
                        TimeZoneInfo tzi1 = OlsonTimeZoneToTimeZoneInfo(sTimeZone);
                        TimeZoneInfo tzi2 = OlsonTimeZoneToTimeZoneInfo("Australia/Brisbane");
                        TimeSpan tsTimeDiifBrisMinusLocal = tzi2.BaseUtcOffset - tzi1.BaseUtcOffset;
                        DateTime BrisbaneDateTime = dtStartTime.Add(tsTimeDiifBrisMinusLocal);

                        GameInDraw gid = new GameInDraw();
                        gid.iGameNo = iNewMatch + i;
                        gid.sTeamHome = ewsd.GetTeamNameAbbrev(sTeamHome);
                        gid.sTeamAway = ewsd.GetTeamNameAbbrev(sTeamAway);
                        gid.sGround = ewsd.GetGroundNameAbbrev(sGroundName);

                        gid.sStart = BrisbaneDateTime.ToString("dd=MM=yyyy");

                        gid.dtStart = BrisbaneDateTime;
                        ewsd.AddGameInDraw(gid);
                        fm.AddData("Game number " + gid.iGameNo.ToString() + " (" + sSlug + ") in FIXTURES has been added in Sheet:DrawCurrent\r\n");
                    }
                    else
                    {
                        return;
                    }
                }
            }
        }

        private void LoadAllResults(int iNewMatch, string sMatchResultsString, ExceWebScrapingDraw ewsd, CricketForm fm)
        {
            DOBchanged dc = new DOBchanged();
            XmlDocument xmlDocMatch = new XmlDocument();

            JObject obj = CreateJObject(sMatchResultsString);

            int iCount = obj["props"]["appPageProps"]["data"]["data"]["content"]["matches"].Count();
            JToken jtok;
            string sStartTime = "";
            string sTimeZone = "";
            string sTeamHome = "";
            string sTeamAway = "";
            string sGroundName = "";

            int iArrNo = iCount - iNewMatch;
            int iCtr = iNewMatch - 1;

            for (int i = iArrNo; i >= 0; i--)
            {
                iCtr++;
                jtok = obj["props"]["appPageProps"]["data"]["data"]["content"]["matches"][i]["slug"];
                string sSlug = jtok.Value<string>();
                string[] aAll;
                aAll = sSlug.Split('-');
                int iAllCount = aAll.Count();
                string sNthSth = aAll[iAllCount - 2];

                bool bIsNorth = (sNthSth == "north");
                if (sSlug == "hampshire-vs-kent-north-group")
                {
                    bIsNorth = false;
                }
                if (!bIsNorth)
                {
                    jtok = obj["props"]["appPageProps"]["data"]["data"]["content"]["matches"][i]["teams"][0]["team"]["slug"];
                    sTeamHome = jtok.Value<string>();
                    jtok = obj["props"]["appPageProps"]["data"]["data"]["content"]["matches"][i]["teams"][1]["team"]["slug"];
                    sTeamAway = jtok.Value<string>();
                    jtok = obj["props"]["appPageProps"]["data"]["data"]["content"]["matches"][i]["startTime"];
                    sStartTime = jtok.Value<string>();
                    jtok = obj["props"]["appPageProps"]["data"]["data"]["content"]["matches"][i]["ground"]["name"];
                    sGroundName = jtok.Value<string>();
                    jtok = obj["props"]["appPageProps"]["data"]["data"]["content"]["matches"][i]["ground"]["town"]["timezone"];
                    sTimeZone = jtok.Value<string>();
                    if (!(sTeamHome == "tba" || sTeamAway == "tba"))
                    {
                        DateTime dtStartTime = DateTime.ParseExact(sStartTime, "MM/dd/yyyy HH:mm:ss", System.Globalization.CultureInfo.InvariantCulture);
                        TimeZoneInfo tzi1 = OlsonTimeZoneToTimeZoneInfo(sTimeZone);
                        TimeZoneInfo tzi2 = OlsonTimeZoneToTimeZoneInfo("Australia/Brisbane");
                        TimeSpan tsTimeDiifBrisMinusLocal = tzi2.BaseUtcOffset - tzi1.BaseUtcOffset;
                        DateTime BrisbaneDateTime = dtStartTime.Add(tsTimeDiifBrisMinusLocal);
                        GameInDraw gid = new GameInDraw();
                        gid.iGameNo = iCtr;
                        gid.sTeamHome = ewsd.GetTeamNameAbbrev(sTeamHome);
                        gid.sTeamAway = ewsd.GetTeamNameAbbrev(sTeamAway);
                        gid.sGround = ewsd.GetGroundNameAbbrev(sGroundName);
                        gid.sStart = BrisbaneDateTime.ToString("dd=MM=yyyy");
                        gid.dtStart = BrisbaneDateTime;
                        ewsd.AddGameInDraw(gid);
                        fm.AddData("Game number " + gid.iGameNo.ToString() + " (" + sSlug + ") in RESULTS has been added in Sheet:DrawCurrent\r\n");
                    }
                    else
                    {
                        return;
                    }
                }
            }
        }
    }
}
