using System;
using System.Globalization;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using System.Collections.Generic;
using Microsoft.Office.Interop.Excel;
using System.Diagnostics;
using System.Linq;
using System.IO;

namespace Cricket
{
    // Blue Text means "Not Out"
    public class ExcelWorkbook
    {
        public Excel.Application xlApp;
        public Excel.Workbook xlWorkbook;
        public Excel._Worksheet xlWorksheet;
        public Excel._Worksheet xlWrkshtPlyrStts;
        public Excel._Worksheet xlWrkshtMtchs;
        public Excel.Range xlRange;

        public void InitExcel()
        {
            //Create COM Objects. Create a COM object for everything that is referenced
            xlApp = new Excel.Application();
            xlWorkbook = xlApp.Workbooks.Open(gbl.sEXCELPATH);
            xlWrkshtPlyrStts = (Excel.Worksheet)xlWorkbook.Worksheets[gbl.SheetType.sPLAYERSTATS];
            xlWrkshtMtchs = (Excel.Worksheet)xlWorkbook.Worksheets[gbl.SheetType.sMATCHES];
        }

        public ExcelWorkbook()
        {
            xlApp = null;
            xlWorkbook = null;
            xlWorksheet = null;
            xlWrkshtPlyrStts = null;
            xlWrkshtMtchs = null;
        }

        public void CleanUpExcel()
        {

            GC.Collect();
            GC.WaitForPendingFinalizers();
            //rule of thumb for releasing com objects:
            //  never use two dots, all COM objects must be referenced and released individually
            //  ex: [somthing].[something].[something] is bad
            //release com objects to fully kill excel process from running in the background
            if (xlRange != null)
            {
                Marshal.ReleaseComObject(xlRange);
            }
            if (xlWorksheet != null)
            {
                Marshal.ReleaseComObject(xlWorksheet);
            }
            if (xlWrkshtPlyrStts != null)
            {
                Marshal.ReleaseComObject(xlWrkshtPlyrStts);
            }
            if (xlWrkshtMtchs != null)
            {
                Marshal.ReleaseComObject(xlWrkshtMtchs);
            }

            if (xlWorkbook != null)
            {
                xlWorkbook.Close();
                Marshal.ReleaseComObject(xlWorkbook);
            }
            xlApp.Quit();
            Marshal.ReleaseComObject(xlApp);
            Process[] pro = Process.GetProcessesByName("excel");
            pro[0].Kill();
            pro[0].WaitForExit();
            xlRange = null;
            xlWorksheet = null;
            xlWorkbook = null;
            xlWrkshtPlyrStts = null;
            xlWrkshtMtchs = null;
            xlApp = null;
        }

        ~ExcelWorkbook()
        {
            //cleanup
            GC.Collect();
            GC.WaitForPendingFinalizers();
            //rule of thumb for releasing com objects:
            //  never use two dots, all COM objects must be referenced and released individually
            //  ex: [somthing].[something].[something] is bad
            //release com objects to fully kill excel process from running in the background
            if (xlRange != null)
            {
                Marshal.ReleaseComObject(xlRange);
            }
            if (xlWorksheet != null)
            {
                Marshal.ReleaseComObject(xlWorksheet);
            }
            if (xlWrkshtPlyrStts != null)
            {
                Marshal.ReleaseComObject(xlWrkshtPlyrStts);
            }
            if (xlWrkshtMtchs != null)
            {
                Marshal.ReleaseComObject(xlWrkshtMtchs);
            }
            //close and release
            if (xlWorkbook != null)
            {
                xlWorkbook.Close();
                Marshal.ReleaseComObject(xlWorkbook);
            }
            //quit and release
            if (xlApp != null)
            {
                xlApp.Quit();
                Marshal.ReleaseComObject(xlApp);
            }
        }

        public string GetTeamNameAbbrev(string sTeamNameFull)
        {
            string sAbbrev = "";

            InitExcel();
            xlWorksheet = (Excel.Worksheet)xlWorkbook.Worksheets[gbl.SheetType.sDRAWCURRENT];
            int iLastCol = GetLastCol(1, gbl.SheetType.sDRAWCURRENT);
            for (int i = 1; i <= iLastCol; i = i + 3)
            {
                string sCurrTeamFull = xlWorksheet.Cells[1, i].Value.ToString().Trim();
                if (sCurrTeamFull == sTeamNameFull)
                {
                    sAbbrev = xlWorksheet.Cells[1, i + 1].Value.ToString().Trim();
                    CleanUpExcel();
                    return sAbbrev;
                }
            }
            CleanUpExcel();
            return "Error: " + sTeamNameFull + " is NOT in Sheet:DrawCurrent";
        }

        protected SumAve AverageTwoValues(string sFirst, string sSecond)
        {
            double dFirst = -1;
            double dSecond = -1;
            SumAve SmAv = new SumAve();

            if (sFirst != "")
            {
                dFirst = Convert.ToDouble(sFirst);
            }
            if (sSecond != "")
            {
                dSecond = Convert.ToDouble(sSecond);
            }

            if (dFirst == -1 && dSecond == -1)
            {
                SmAv.Sum = -1;
                SmAv.Ave = -1;
            }
            else if (dFirst == -1 && dSecond != -1)
            {
                SmAv.Sum = dSecond;
                SmAv.Ave = SmAv.Sum / 1;
            }
            else if (dFirst != -1 && dSecond == -1)
            {
                SmAv.Sum = dFirst;
                SmAv.Ave = SmAv.Sum / 1;
            }
            else
            {
                SmAv.Sum = dFirst + dSecond;
                SmAv.Ave = SmAv.Sum / 2;
            }
            return SmAv;
        }

        public bool IsSamePlayer(Player FrstPlyr, Player SndPlyr)
        {
            string sFirstDOB = DateToString(FrstPlyr.DOB);
            string sSecondDOB = DateToString(SndPlyr.DOB);
            if ((FrstPlyr.FirstName == SndPlyr.FirstName) && (FrstPlyr.LastName == SndPlyr.LastName) && (sFirstDOB == sSecondDOB))
            {
                return true;
            }
            else
            {
                return false;
            }
        }

        public Competition GetCompetition(string sSheetType, CricketForm fm)
        {
            string[] aSplitString;
            Competition Cmpttn = new Competition();
            string sComp = "";

            xlWorksheet = xlWorkbook.Sheets[sSheetType];

            if (sSheetType == gbl.SheetType.sTEAM)
            {
                sComp = xlWorksheet.Cells[1, 19].Value.ToString().Trim();
            }
            else
            {
                sComp = xlWorksheet.Cells[1, 4].Value.ToString().Trim();
            }
            aSplitString = sComp.Split('=');
            Cmpttn.CompetitionCode = aSplitString[0];
            Cmpttn.Season = aSplitString[1];

            fm.AddData("\r\n" + "Excel is Loading Competiton: " + Cmpttn.CompetitionCode + "=" + Cmpttn.Season + "\r\n");
            return Cmpttn;
        }

        public string GetColorType(double dColorCode)
        {
            if (dColorCode == 5287936)
            {
                return gbl.ColorType.sGREEN;
            }
            else if (dColorCode == 12611584)
            {
                return gbl.ColorType.sBLUE;
            }
            else if (dColorCode == 0)
            {
                return gbl.ColorType.sBLACK;
            }
            else if (dColorCode == 255)
            {
                return gbl.ColorType.sRED;
            }
            else
            {
                return "Error";
            }
        }

        public DateTime StringToDate(string sDate)
        {
            string[] aSplitString;
            aSplitString = sDate.Split('=');
            if (aSplitString[0].Length == 1)
            {
                aSplitString[0] = "0" + aSplitString[0];
            }
            if (aSplitString[1].Length == 1)
            {
                aSplitString[1] = "0" + aSplitString[1];
            }
            if (aSplitString[2].Length == 2)
            {
                int iYear = Convert.ToInt32(aSplitString[2]);
                if (iYear < 65)
                {
                    aSplitString[2] = "20" + aSplitString[2];
                }
                else
                {
                    aSplitString[2] = "19" + aSplitString[2];
                }
            }
            sDate = (aSplitString[0].Trim() + "/" + aSplitString[1].Trim() + "/" + aSplitString[2].Trim().ToString());
            aSplitString = sDate.Split(' ');
            sDate = aSplitString[0] + " 12:00:00";
            DateTime date = DateTime.ParseExact(sDate, "dd/MM/yyyy HH:mm:ss", CultureInfo.InvariantCulture);
            return date;
        }

        protected String DateToString(DateTime dDate)
        {
            string[] aAll;
            string sDate = "";

            sDate = dDate.ToString();
            aAll = sDate.Split(' ');
            sDate = aAll[0];
            aAll = sDate.Split('/');
            return aAll[0].Trim() + "=" + aAll[1].Trim() + "=" + aAll[2].Trim();
        }


        public int GetLastRowFast(int iCol, Excel._Worksheet xlWrksht)
        {
            string sCurrentCell = "";
            bool bNotFound = true;
            int i = gbl.iMAXROWS;

            xlRange = xlWrksht.UsedRange;
            while ((bNotFound) && (i != 0))
            {
                sCurrentCell = xlRange.Cells[i, iCol]?.Value?.ToString();
                if (sCurrentCell != null)
                {
                    bNotFound = false;
                    return i;
                }
                i--;
            }
            return 0;
        }

        public int GetLastRow(int iCol, string sMemberType)
        {
            string sCurrentCell = "";
            bool bNotFound = true;
            int i = gbl.iMAXROWS;

            xlWorksheet = xlWorkbook.Sheets[sMemberType];
            xlRange = xlWorksheet.UsedRange;

            while ((bNotFound) && (i != 0))
            {
                sCurrentCell = xlRange.Cells[i, iCol]?.Value?.ToString();

                if (sCurrentCell != null)
                {
                    bNotFound = false;
                    return i;
                }
                i--;
            }
            return 0;
        }

        public int GetLastRowPlyrSts(int iCol)
        {
            string sCurrentCell = "";
            bool bNotFound = true;
            int i = gbl.iMAXROWS;

            xlWrkshtPlyrStts = xlWorkbook.Sheets[gbl.SheetType.sPLAYERSTATS];
            while ((bNotFound) && (i != 0))
            {
                sCurrentCell = xlWrkshtPlyrStts.Cells[i, iCol]?.Value?.ToString();

                if (sCurrentCell != null)
                {
                    bNotFound = false;
                    return i;
                }
                i--;
            }
            return 0;
        }

        protected string TeamShortToLong(string sTeamCode)
        {
            int iColCount = 0;
            string sCurentTeam = "";
            int i = 0;

            xlWorksheet = xlWorkbook.Sheets[gbl.SheetType.sDRAWCURRENT];
            xlRange = xlWorksheet.UsedRange;
            iColCount = GetLastCol(1, gbl.SheetType.sDRAWCURRENT);
            i = 2;
            while (i <= (iColCount - 1))
            {
                sCurentTeam = xlRange.Cells[1, i]?.Value?.ToString().Trim();
                if (sCurentTeam == sTeamCode)
                {
                    return xlRange.Cells[1, i - 1]?.Value?.ToString().Trim();
                }
                i = i + 3;
            }
            MessageBox.Show("Error in GetThisTeam: No Team ");
            return "Error";
        }

        protected string TeamLongToShort(string sTeamLong)
        {
            int iColCount = 0;
            string sCurentTeamLong = "";
            int i = 0;

            xlWorksheet = xlWorkbook.Sheets[gbl.SheetType.sDRAWCURRENT];
            xlRange = xlWorksheet.UsedRange;
            iColCount = GetLastCol(1, gbl.SheetType.sDRAWCURRENT);
            i = 1;
            while (i <= (iColCount - 2))
            {
                sCurentTeamLong = xlRange.Cells[1, i]?.Value?.ToString().Trim();
                if (sCurentTeamLong == sTeamLong)
                {
                    return xlRange.Cells[1, i + 1]?.Value?.ToString().Trim();
                }
                i = i + 3;
            }
            MessageBox.Show("Error in GetThisTeam: No Team ");
            return "Error";
        }

        public int GetLastCol(int iRow, string sMemberType)
        {
            string sCurrentCell = "";
            bool bNotFound = true;
            int i = gbl.iMAXCOLUMNS;

            xlWorksheet = xlWorkbook.Sheets[sMemberType];
            xlRange = xlWorksheet.UsedRange;

            while ((bNotFound) && (i != 0))
            {
                sCurrentCell = xlRange.Cells[iRow, i]?.Value?.ToString();

                if (sCurrentCell != null)
                {
                    bNotFound = false;
                    return i;
                }
                i--;
            }
            return 0;
        }

        public int GetLastColFast(int iRow, Excel._Worksheet xlWrksht)
        {
            string sCurrentCell = "";
            bool bNotFound = true;
            int i = gbl.iMAXCOLUMNS;

            while ((bNotFound) && (i != 0))
            {
                sCurrentCell = xlWrksht.Cells[iRow, i].Value.ToString();

                if (sCurrentCell != null)
                {
                    bNotFound = false;
                    return i;
                }
                i--;
            }
            return 0;
        }

        public int GetLastMatch(CricketForm fm)
        {
            int iRowCount = -1;
            int iLastSoFar = -1;
            string sCurrentCell = "";
            string sGameNo = "";
            int iGameNo = -1;
            string[] aSplitString;
            xlWorksheet = xlWorkbook.Sheets[gbl.SheetType.sTEAM];
            xlRange = xlWorksheet.UsedRange;

            iRowCount = GetLastRow(1, gbl.SheetType.sTEAM);
            int iNoTeams = (iRowCount - 2) / 13;
            //int k = 0;
            for (int i = 0; i < iNoTeams; i++)
            {
                int iRowNo = (i * 13) + 13;
                int iLastCol = GetLastCol(iRowNo, gbl.SheetType.sTEAM);

                sCurrentCell = xlRange.Cells[iRowNo, iLastCol]?.Value?.ToString();
                if (sCurrentCell != "Opponent")
                {
                    aSplitString = sCurrentCell.Split(')');
                    sGameNo = aSplitString[0];
                    iGameNo = Convert.ToInt32(sGameNo);

                    if (iGameNo > iLastSoFar)
                    {
                        iLastSoFar = iGameNo;
                    }
                }
            }
            return iLastSoFar;
        }

        public int GetNoGamesForSeason(CricketForm fm)
        {
            int iRowCount = -1;
            xlWorksheet = xlWorkbook.Sheets[gbl.SheetType.sTEAM];
            xlRange = xlWorksheet.UsedRange;
            iRowCount = GetLastRow(1, gbl.SheetType.sTEAM);
            int iNoTeams = (iRowCount - 1) / 13;
            int k = 0;
            for (int i = 0; i < iNoTeams; i++)
            {
                int iRowNo = (i * 13) + 12;
                int iLastCol = GetLastCol(iRowNo, gbl.SheetType.sTEAM);
                for (int j = 3; j <= iLastCol; j++)
                {
                    k++;
                }

            }
            return (k / 2);
        }
    }

    //public class ExcelSavePlrTm : ExcelWorkbook
    //{
    //}

    public class ExcelSeason : ExcelWorkbook
    {
        public Player GtNxtBtsmn(int iRowNo, CricketForm fm)
        {
            string sNames;
            string[] aSplitString;
            string sFirstName;
            string sMiddleName;
            string sDate = "";

            xlWorksheet = xlWorkbook.Sheets[gbl.SheetType.sBATSMEN];
            xlRange = xlWorksheet.UsedRange;
            Player plyrCurrent = new Player();
            plyrCurrent.LastName = xlRange.Cells[iRowNo - 4, 2].Value.ToString().Trim();
            sNames = xlRange.Cells[iRowNo - 3, 2].Value.ToString().Trim();
            aSplitString = sNames.Split(' ');
            sFirstName = aSplitString[0];
            sMiddleName = aSplitString[1];
            plyrCurrent.FirstName = sFirstName.Trim();
            plyrCurrent.MiddleName = sMiddleName.Trim();

            sDate = xlRange.Cells[iRowNo - 1, 2].Value.ToString().Trim();
            plyrCurrent.DOB = StringToDate(sDate);
            plyrCurrent.Country = xlRange.Cells[iRowNo, 2].Value.ToString().Trim();
            return plyrCurrent;
        }

        private bool IsSignificant(string sData)
        {
            string sLastChar = sData.Substring(sData.Length - 1);
            if (sLastChar == "*")
            {
                return false;
            }
            else
            {
                return true;
            }
        }

        public Season GetNextSeason(string sMatchID, string sGameNo, string sPlayerType, int iRow, int iCol, CricketForm fm)
        {
            string sTeamCode = "";
            string sData = "";
            string[] aSplitString;
            string[] aSplitString1 = new string[2]; ;
            string sTeam = "";
            Season ssn = new Season();
            string sColor = "";
            string sOppTeamCode = "";

            if (sPlayerType == gbl.SheetType.sBATSMEN)
            {
                xlWorksheet = xlWorkbook.Sheets[gbl.SheetType.sBATSMEN];
                xlRange = xlWorksheet.UsedRange;
                if (xlRange.Cells[iRow - 4, iCol].Value == null)
                {
                    //prf.BatsmanRes = null;
                }
                else
                {

                    sData = xlRange.Cells[iRow - 4, iCol].Value.ToString().Trim();
                    double dColor = xlRange.Cells[iRow - 4, iCol].Font.Color;
                    sColor = GetColorType(dColor);


                    ResultsBatsman resbat = new ResultsBatsman();
                    if (sColor == gbl.ColorType.sBLACK)
                    {
                        resbat.NotOut = false;
                    }
                    else if (sColor == gbl.ColorType.sBLUE)
                    {
                        resbat.NotOut = true;
                    }
                    else
                    {
                        MessageBox.Show("Error in GetNextPerformance: Cell Text should be Black or Blue");
                        resbat.NotOut = true;
                    }
                    resbat.IsSignificant = IsSignificant(sData);
                    aSplitString = sData.Split('=');
                    resbat.Runs = aSplitString[0].Trim();
                    resbat.BallsFaced = aSplitString[1].Trim();
                    resbat.Four = aSplitString[2].Trim();
                    resbat.Six = aSplitString[3].Trim();
                }
                aSplitString1 = GetGameNo(iRow, iCol, gbl.SheetType.sBATSMEN);
                if (aSplitString1 != null)
                {
                    sOppTeamCode = aSplitString1[1].ToString().Trim();
                    xlWorksheet = xlWorkbook.Sheets[gbl.SheetType.sBATSMEN];
                    xlRange = xlWorksheet.UsedRange;
                    sTeamCode = xlRange.Cells[iRow, 1].Value.ToString().Trim();
                    aSplitString = sTeamCode.Split('.');
                    sTeamCode = aSplitString[0];
                    sTeam = TeamShortToLong(sTeamCode);
                    ssn.TeamMine = sTeamCode;
                }
            }
            return ssn;
        }

        public string[] GetGameNo(int iRow, int iCol, string sPlayerType)
        {
            // Returns GameNo) OppTeamCode
            // Based on Game Row
            string[] aSplitString = new string[2];
            string sGameNoAndOpp = "";

            if (sPlayerType == gbl.SheetType.sBATSMEN)
            {
                xlWorksheet = xlWorkbook.Sheets[gbl.SheetType.sBATSMEN];
                xlRange = xlWorksheet.UsedRange;
                sGameNoAndOpp = xlRange.Cells[iRow, iCol].Value.ToString().Trim();
                aSplitString = sGameNoAndOpp.Split(')');
                return aSplitString;
            }
            else //(sPlayerType == gbl.SheetType.sBOWLER)
            {
                xlWorksheet = xlWorkbook.Sheets[gbl.SheetType.sBOWLER];
                xlRange = xlWorksheet.UsedRange;
                sGameNoAndOpp = xlRange.Cells[iRow, iCol].Value;
                if (sGameNoAndOpp == null)
                {
                    MessageBox.Show("Error in GetGameNo: Game Number and Opponent is null");
                    return null;
                }
                else
                {
                    aSplitString = sGameNoAndOpp.Split(')');
                    return aSplitString;
                }
            }
        }

        public ResultsBowler GetBowlingData(Player plyr, Match mtch)
        {
            string[] aSplitString;
            string[] aSplitString1 = new string[2];
            string[] aSplitString2 = new string[3];
            Player CurrentPlayer = new Player();
            ResultsBowler resbow = new ResultsBowler();


            xlWorksheet = xlWorkbook.Sheets[gbl.SheetType.sBOWLER];
            xlRange = xlWorksheet.UsedRange;

            int iRowCount = GetLastRow(1, gbl.SheetType.sBOWLER);
            int iNoBowlers = (iRowCount - 1) / 5;

            for (int i = 0; i < iNoBowlers; i++)
            {

                int iRow = (i * 5) + 5;
                xlWorksheet = xlWorkbook.Sheets[gbl.SheetType.sBOWLER];
                xlRange = xlWorksheet.UsedRange;
                string sNames = xlRange.Cells[iRow - 2, 2].Value;

                aSplitString = sNames.Split('=');
                CurrentPlayer.FirstName = aSplitString[0].Trim();
                CurrentPlayer.MiddleName = aSplitString[1].Trim();
                CurrentPlayer.LastName = xlRange.Cells[iRow - 3, 2].Value.ToString().Trim();
                string sDate = xlRange.Cells[iRow - 1, 2].Value.ToString().Trim();
                //sDate = FrmtDate(sDate);
                //DateTime date = DateTime.ParseExact(sDate, "dd/MM/yyyy HH:mm:ss", CultureInfo.InvariantCulture);
                CurrentPlayer.DOB = StringToDate(sDate);
                CurrentPlayer.Country = xlRange.Cells[iRow, 2].Value.ToString().Trim();
                if (CurrentPlayer.FirstName == plyr.FirstName && CurrentPlayer.MiddleName == plyr.MiddleName && CurrentPlayer.LastName == plyr.LastName && CurrentPlayer.DOB == plyr.DOB && CurrentPlayer.Country == plyr.Country)
                {
                    int iCol = GetLastCol(iRow, gbl.SheetType.sBOWLER);
                    for (int j = 3; j <= iCol; j++)
                    {
                        aSplitString1 = this.GetGameNo(iRow, j, gbl.SheetType.sBOWLER);
                        //aSplitString1 = this.GetGameNo(iRow, j, gbl.SheetType.sBOWLER);
                        if (aSplitString1 != null)
                        {
                            string sGameNo = aSplitString1[0];
                            string sOpponent = aSplitString1[1].Trim();
                            string sOppFullName = TeamShortToLong(sOpponent);
                            if (sGameNo == mtch.GameNumber.ToString().Trim())
                            {
                                xlWorksheet = xlWorkbook.Sheets[gbl.SheetType.sBOWLER];
                                xlRange = xlWorksheet.UsedRange;
                                string sData = xlRange.Cells[iRow - 3, j].Value;
                                if (sData == null)
                                {
                                    return null;
                                }
                                else
                                {
                                    resbow.IsSignificant = IsSignificant(sData.ToString());
                                    aSplitString = sData.ToString().Split('=');
                                    resbow.Wickets = aSplitString[0].Trim();
                                    resbow.RunsConceded = aSplitString[1].Trim();
                                    resbow.OversBowled = aSplitString[2].Trim();
                                    return resbow;

                                }
                            }
                        }
                    }
                }
            }
            return null;
        }

        public Player GetNextPlayerBatsman(int iPlyrNo, CricketForm fm)
        {
            string sNames;
            string[] aSplitString;
            string sFirstName;
            string sMiddleName;
            string sDate = "";

            fm.AddData("\r\n" + "Excel is Loading next Player...");

            xlWorksheet = xlWorkbook.Sheets[gbl.SheetType.sBATSMEN];
            xlRange = xlWorksheet.UsedRange;

            Player plyrCurrent = new Player();
            int iFirstRow;
            iFirstRow = (iPlyrNo * 6) + 4;
            plyrCurrent.LastName = xlRange.Cells[iFirstRow, 2].Value.ToString().Trim();

            sNames = xlRange.Cells[iFirstRow + 1, 2].Value.ToString().Trim();
            aSplitString = sNames.Split('=');
            sFirstName = aSplitString[0];
            sMiddleName = aSplitString[1];

            plyrCurrent.FirstName = sFirstName.Trim();
            plyrCurrent.MiddleName = sMiddleName.Trim();


            sDate = xlRange.Cells[iFirstRow + 3, 2].Value.ToString().Trim();
            //sDate = FrmtDate(sDate);
            //DateTime date = DateTime.ParseExact(sDate, "dd/MM/yyyy HH:mm:ss", CultureInfo.InvariantCulture);
            plyrCurrent.DOB = StringToDate(sDate);

            plyrCurrent.Country = xlRange.Cells[iFirstRow + 4, 2].Value.ToString().Trim();
            return plyrCurrent;
        }

        public bool PlayerPlayedThisGame(int iRow, int iCol, CricketForm fm)
        {
            string sGame = "";

            fm.AddData("\r\n" + "Checking if player played this game...");

            xlWorksheet = xlWorkbook.Sheets[gbl.SheetType.sBATSMEN];
            xlRange = xlWorksheet.UsedRange;
            sGame = xlRange.Cells[iRow, iCol].Value;

            if (sGame != null)
            {
                fm.AddData("\r\n" + "Player DID play this game...\r\n");
                return true;
            }
            else
            {
                fm.AddData("\r\n" + "Player DID NOT play this game...\r\n");
                return false;
            }
        }



        private Player GetPlayer(int iRowNo)
        {
            Player ply = new Player();
            string[] aAll;
            string sFirstNames = "";
            string sDB = "";

            ply.LastName = xlRange.Cells[iRowNo, 2].Value.ToString().Trim();
            if (xlRange.Cells[iRowNo, 2].Value == null)
            {
                MessageBox.Show("Error in GetPlayer: No LastName Entered for Player at Cell [" + iRowNo.ToString() + ", 2]");
                return null;
            }

            iRowNo++;
            sFirstNames = xlRange.Cells[iRowNo, 2].Value;
            if (sFirstNames != null)
            {
                bool bHasEquals = sFirstNames.Contains('=');
                if (bHasEquals)
                {
                    sFirstNames = sFirstNames.ToString().Trim();
                    aAll = sFirstNames.Split('=');
                    ply.FirstName = aAll[0];
                    ply.MiddleName = aAll[1];
                }
                else
                {
                    MessageBox.Show("Error in GetPlayer: No = sign Entered for First names at Cell [" + iRowNo.ToString() + ", 2]");
                    return null;
                }
            }
            else
            {
                MessageBox.Show("Error in GetPlayer: No First Names Entered for Player at Cell [" + iRowNo.ToString() + ", 2]");
                return null;
            }


            iRowNo = iRowNo + 2; ;
            if (xlRange.Cells[iRowNo, 2].Value == null)
            {
                MessageBox.Show("Error in GetPlayer: No DOB Entered for Player at Cell [" + iRowNo.ToString() + ", 2]");
                return null;
            }
            else
            {
                // System.DateTime
                sDB = xlRange.Cells[iRowNo, 2].Value.ToString().Trim();
                ply.DOB = StringToDate(sDB);
            }

            iRowNo++;
            ply.Country = xlRange.Cells[iRowNo, 2].Value;
            if (ply.Country == null)
            {
                MessageBox.Show("Error in GetPlayer: No Country Entered for Player at Cell [" + iRowNo.ToString() + ", 2]");
                return null;
            }
            return ply;
        }

        protected bool IsNotOut(int iRow, int iCol)
        {
            string sColor = "";
            double dColor = xlRange.Cells[iRow, iCol].Font.Color;

            sColor = GetColorType(dColor);
            if (sColor == gbl.ColorType.sBLACK)
            {
                return false;
            }
            else if (sColor == gbl.ColorType.sBLUE)
            {
                return true;
            }
            else
            {
                MessageBox.Show("Error in IsNotOut: Cell Text should be Black or Blue");
                return false;
            }
        }

        private bool PlayerCompSeasonAlreadyInDB(Season CurrPlayer, List<Season> ssns)
        {
            for (int i = 0; i < ssns.Count; i++)
            {
                Season PlyrInDB = new Season();
                PlyrInDB = ssns[i];
                if (IsSamePlayer(CurrPlayer.Playr, PlyrInDB.Playr))
                {
                    if (CurrPlayer.Comp.CompetitionCode == PlyrInDB.Comp.CompetitionCode)
                    {
                        if (CurrPlayer.Comp.Season == PlyrInDB.Comp.Season)
                        {
                            return true;
                        }
                    }
                }
            }
            return false;
        }
        public List<Season> LoadExcelBatBowl(string sSheetName, List<Season> ssns, CricketForm fm)
        {
            int iLastRow = -1;
            int iLastCol = -1;
            string sInngs = "";
            string[] aAll;
            List<Season> AllPlayers = new List<Season>();
            string sGameNoAndOpp = "";

            InitExcel();
            xlWorksheet = xlWorkbook.Sheets[sSheetName];
            xlRange = xlWorksheet.UsedRange;
            iLastRow = GetLastRow(1, sSheetName);
            Competition cmp = new Competition();
            cmp = GetCompetition(sSheetName, fm);
            if (iLastRow < 9)
            {
                iLastRow = 0;
            }
            else
            {
                double iCount = iLastRow - 3;
                double iMod = iCount % 6;
                if (iMod != 0)
                {
                    MessageBox.Show("Error in LoadExcelBatsmen: (1stColumn - 3) is NOT a Multiple of 6");
                }
            }
            for (int i = 4; i <= iLastRow - 5; i = i + 6)
            {
                List<AllInnFrSsn> AllInnList = new List<AllInnFrSsn>();
                Player plyr = new Player();
                Season OneBatsman = new Season();
                plyr = GetPlayer(i);
                if (plyr == null)
                {
                    return null;
                }
                OneBatsman.Playr = plyr;
                OneBatsman.Comp = cmp;
                OneBatsman.TeamMine = xlRange.Cells[i + 2, 2].Value.ToString().Trim();
                if (!PlayerCompSeasonAlreadyInDB(OneBatsman, ssns))
                {
                    iLastCol = GetLastCol(i + 4, sSheetName);
                    for (int j = 3; j <= iLastCol; j++)
                    {
                        AllInnFrSsn OneInngs = new AllInnFrSsn();
                        sGameNoAndOpp = xlRange.Cells[i + 4, j].Value.ToString().Trim();
                        OneInngs.GameNoAndOpp = sGameNoAndOpp;
                        aAll = sGameNoAndOpp.Split(')');
                        OneInngs.GameNumber = aAll[0].ToString().Trim();
                        OneInngs.TeamOpposition = aAll[1].ToString().Trim();
                        sInngs = xlRange.Cells[i, j].Value;
                        if (sInngs != null)
                        {
                            sInngs = sInngs.ToString().Trim();
                            aAll = sInngs.Split('=');
                            if (sSheetName == gbl.SheetType.sBATSMEN)
                            {
                                ResultsBatsman RsBat = new ResultsBatsman();
                                RsBat.Runs = aAll[0];
                                RsBat.BallsFaced = aAll[1];
                                RsBat.Four = aAll[2];
                                RsBat.Six = aAll[3];
                                RsBat.NotOut = IsNotOut(i, j);
                                RsBat.IsSignificant = IsSignificant(sInngs);
                                OneInngs.ResBat = RsBat;
                            }
                            else if (sSheetName == gbl.SheetType.sBOWLER)
                            {
                                ResultsBowler RsBwl = new ResultsBowler();
                                RsBwl.Wickets = aAll[0];
                                RsBwl.RunsConceded = aAll[1];
                                RsBwl.OversBowled = aAll[2];
                                RsBwl.IsSignificant = IsSignificant(sInngs);
                                OneInngs.ResBowl = RsBwl;
                            }
                        }
                        else
                        {
                            OneInngs.ResBat = null;
                            OneInngs.ResBowl = null;
                        }
                        AllInnList.Add(OneInngs);
                    }
                    OneBatsman.AllInnings = AllInnList;
                    AllPlayers.Add(OneBatsman);

                }
            }
            CleanUpExcel();
            return AllPlayers;
        }

        private bool PlayerAlreadyInDBAnsSameCompSeason(Season OnePlayer, List<Season> ssns, CricketForm fm)
        {
            if (ssns == null)
            {
                return false;
            }
            else
            {
                for (int i = 0; i < ssns.Count; i++)
                {
                    if (IsSamePlayer(OnePlayer.Playr, ssns[i].Playr))
                    {
                        if ((OnePlayer.Comp.CompetitionCode == ssns[i].Comp.CompetitionCode)
                            && (OnePlayer.Comp.Season == ssns[i].Comp.Season))
                        {
                            fm.AddData("Player " + OnePlayer.Playr.FirstName + " " + OnePlayer.Playr.LastName + " IS already in the DB for Comp=Season: " +
                                 OnePlayer.Comp.CompetitionCode + "=" + OnePlayer.Comp.Season + "\r\n");


                            return true;
                        }
                    }
                }
            }
            return false;
        }

        public List<Season> MergeBatBowl(List<Season> aBatsmen, List<Season> aBowlers, List<Season> ssns, CricketForm fm)
        {
            List<Season> aMerged = new List<Season>();

            if ((aBatsmen == null) && (aBowlers == null))
            {
                return null;
            }
            for (int i = 0; i <= aBatsmen.Count - 1; i++)
            {
                Season OneBatsman = new Season();
                OneBatsman = aBatsmen[i];

                if (!PlayerAlreadyInDBAnsSameCompSeason(OneBatsman, ssns, fm))
                {
                    List<AllInnFrSsn> AllInng = new List<AllInnFrSsn>();
                    for (int j = 0; j <= OneBatsman.AllInnings.Count - 1; j++)
                    {
                        ResultsBowler ResBwl = new ResultsBowler();
                        ResBwl = null;
                        if (aBowlers.Count > 0)
                        {
                            for (int x = 0; x < aBowlers.Count; x++)
                            {
                                Season OneBowler = new Season();
                                OneBowler = aBowlers[x];

                                // games no's must match in size -- fix this
                                if ((OneBatsman.AllInnings.Count != OneBowler.AllInnings.Count) && (IsSamePlayer(OneBatsman.Playr, OneBowler.Playr)))
                                {
                                    MessageBox.Show("Error in MergeBatBowl: Number of Games must be the same for Player in Batsman Innings and Bowling Innings: Player " + OneBatsman.Playr.FirstName + " " + OneBatsman.Playr.LastName);
                                    return null;
                                }
                                else
                                {
                                    if (IsSamePlayer(OneBatsman.Playr, OneBowler.Playr))
                                    {
                                        // Load Full Season of Bowling Data
                                        for (int y = OneBowler.AllInnings.Count - 1; y >= 0; y--)
                                        {
                                            OneBatsman.AllInnings[y].ResBowl = OneBowler.AllInnings[y].ResBowl;
                                            if ((OneBatsman.AllInnings[y].ResBat == null) && (OneBatsman.AllInnings[y].ResBowl == null))
                                            {
                                                OneBatsman.AllInnings.RemoveAt(y);
                                            }
                                        }
                                        aBowlers.RemoveAt(x);
                                    }
                                }
                            }
                        }
                        else
                        {
                            for (int z = OneBatsman.AllInnings.Count - 1; z >= 0; z--)
                            {
                                if (OneBatsman.AllInnings[z].ResBat == null)
                                {
                                    OneBatsman.AllInnings.RemoveAt(z);
                                }
                            }
                        }
                    }
                    if (OneBatsman.AllInnings.Count != 0)
                    {
                        aMerged.Add(OneBatsman);
                    }
                }
            }


            for (int x = aBowlers.Count - 1; x >= 0; x--)
            {
                Season OneBowler = new Season();
                OneBowler = aBowlers[x];
                for (int y = OneBowler.AllInnings.Count - 1; y >= 0; y--)
                {
                    if (OneBowler.AllInnings[y].ResBowl == null)
                    {
                        OneBowler.AllInnings.RemoveAt(y);
                    }
                }
                if (OneBowler.AllInnings.Count != 0)
                {
                    aMerged.Add(OneBowler);
                }
            }

            return aMerged;
        }
    }

    public class ExcelLoadPlyrs : ExcelWorkbook
    {
        public object HorizontalAlignType { get; private set; }

        private string GetDOB(string sDOBdb)
        {
            string[] aSplitString;
            string sPre = "";

            aSplitString = sDOBdb.Split(' ');
            sPre = aSplitString[0];
            aSplitString = sPre.Split('/');

            return (aSplitString[0] + "=" + aSplitString[1] + "=" + aSplitString[2]).Trim();
        }

        private bool IsBowlerInBowlingSheet(Player plyr)
        {
            Player pCurrentPlayer = new Player();
            int iRow;
            string[] aSplitString;

            Excel.Worksheet xlWorksheet = (Excel.Worksheet)xlWorkbook.Worksheets[gbl.SheetType.sTESTBOWLER];
            // xlRange = xlWorksheet.UsedRange;
            // get last row
            iRow = GetLastRow(1, gbl.SheetType.sTESTBOWLER);
            if (iRow == 0)
            {
                return false;
            }
            else
            {

                int iNoBowlers = iRow / 5;
                for (int i = 0; i < iNoBowlers; i++)
                {
                    int iCurrentRow = (i * 5) + 2;
                    string Names = xlWorksheet.Cells[iCurrentRow + 1, 2].Value;
                    aSplitString = Names.Split(' ');
                    pCurrentPlayer.FirstName = aSplitString[0].Trim();
                    pCurrentPlayer.MiddleName = aSplitString[1].Trim();
                    pCurrentPlayer.LastName = xlWorksheet.Cells[iCurrentRow, 2].Value.Trim();
                    string sDate = xlWorksheet.Cells[iCurrentRow + 2, 2].Value.Trim();
                    pCurrentPlayer.DOB = StringToDate(sDate);
                    pCurrentPlayer.Country = xlWorksheet.Cells[iCurrentRow + 3, 2].Value.Trim();
                    if (!PlayerIsNew(pCurrentPlayer, plyr))
                    {
                        return true;
                    }

                }
                return false;
            }
        }

        private void AddPlayerToBowlingSheetIfNotThere(Season ssn)
        {
            Player plr = new Player();
            int iRow = -1;
            string sTeamCtr = "";



            Excel.Worksheet xlWorksheet = (Excel.Worksheet)xlWorkbook.Worksheets[gbl.SheetType.sTESTBOWLER];
            // get last row
            iRow = GetLastRow(1, gbl.SheetType.sTESTBOWLER);

            string sMyTeamCode = ssn.TeamMine;
            string sLastName = ssn.Playr.LastName;
            string sFirstName = ssn.Playr.FirstName;
            string sMiddleName = ssn.Playr.MiddleName;
            string sDOB = ssn.Playr.DOB.ToString();
            sDOB = GetDOB(sDOB);
            string sCountry = ssn.Playr.Country;
            //int iTeamCtr = 1;

            plr = ssn.Playr;

            bool bBowlerExists = IsBowlerInBowlingSheet(plr);

            if (!bBowlerExists)
            {
                int iTeamCtr = GetTeamCountBowler(sMyTeamCode);

                if (iTeamCtr <= 9)
                {
                    sTeamCtr = "0" + iTeamCtr.ToString().Trim();
                }
                else
                {
                    sTeamCtr = iTeamCtr.ToString().Trim();
                }

                iRow++;
                xlWorksheet.Cells[iRow + 1, 2] = sLastName;
                xlWorksheet.Cells[iRow + 1, 2].Font.Underline = true;
                xlWorksheet.Cells[iRow + 2, 2] = sFirstName + " " + sMiddleName;
                xlWorksheet.Cells[iRow + 3, 2] = sDOB;
                xlWorksheet.Cells[iRow + 4, 2] = sCountry;

                xlWorksheet.Cells[iRow + 1, 1] = sMyTeamCode + "." + sTeamCtr + "." + "1";
                xlWorksheet.Cells[iRow + 2, 1] = sMyTeamCode + "." + sTeamCtr + "." + "2";
                xlWorksheet.Cells[iRow + 3, 1] = sMyTeamCode + "." + sTeamCtr + "." + "3";
                xlWorksheet.Cells[iRow + 4, 1] = sMyTeamCode + "." + sTeamCtr + "." + "4";
                xlWorksheet.Cells[iRow + 5, 1] = sMyTeamCode + "." + sTeamCtr + "." + "5";

                // add bowler to after last row
            }
        }

        private int GetTeamCountBowler(string sTeamCode)
        {
            int Ctr = 0;
            int iLstRw = GetLastRow(1, gbl.SheetType.sTESTBOWLER);
            int iNoBwlrs = iLstRw / 5;
            string[] aSplitString;

            Excel.Worksheet xlWorksheet = (Excel.Worksheet)xlWorkbook.Worksheets[gbl.SheetType.sTESTBOWLER];

            for (int i = 0; i < iNoBwlrs; i = i + 5)
            {
                int iCurrentRow = i * 5;
                string sCurrentTeam = xlWorksheet.Cells[iCurrentRow + 2, 1].Value;
                aSplitString = sCurrentTeam.Split('.');
                sCurrentTeam = aSplitString[0];
                if (sCurrentTeam == sTeamCode)
                {
                    Ctr++;
                }
            }
            return Ctr + 1;
        }

        private string GetTwoSeasonsAgo(string sLstSeason)
        {
            string[] aAll;


            aAll = sLstSeason.Split('-');
            string sFirstYear = aAll[0];
            string sSecondYear = aAll[1];
            int iFirstYear = Int32.Parse(sFirstYear);
            int iSecondYear = Int32.Parse(sSecondYear);
            iFirstYear = iFirstYear - 1;
            iSecondYear = iSecondYear - 1;

            sFirstYear = iFirstYear.ToString();
            sSecondYear = iSecondYear.ToString();
            return (sFirstYear + "-" + sSecondYear);
        }

        private string GetNextSeason(string sLstSeason)
        {
            string[] aAll;


            aAll = sLstSeason.Split('-');
            string sFirstYear = aAll[0];
            string sSecondYear = aAll[1];
            int iFirstYear = Int32.Parse(sFirstYear);
            int iSecondYear = Int32.Parse(sSecondYear);
            iFirstYear = iFirstYear + 1;
            iSecondYear = iSecondYear + 1;

            sFirstYear = iFirstYear.ToString();
            sSecondYear = iSecondYear.ToString();
            return (sFirstYear + "-" + sSecondYear);
        }

        private void LoadHeading(int iRowNo)
        {
            //xlWorksheet = (Excel.Worksheet)xlWorkbook.Worksheets[xlWrkshtPlayerStats];
            xlWrkshtPlyrStts.Cells[iRowNo, 3] = "Batting";
            xlWrkshtPlyrStts.Cells[iRowNo, 4] = "Inngs";
            xlWrkshtPlyrStts.Cells[iRowNo, 5] = "Runs";
            xlWrkshtPlyrStts.Cells[iRowNo, 6] = "Ave";
            xlWrkshtPlyrStts.Cells[iRowNo, 7] = "SR";
            xlWrkshtPlyrStts.Cells[iRowNo, 8] = "HS";
            xlWrkshtPlyrStts.Cells[iRowNo, 9] = "NOs";
            xlWrkshtPlyrStts.Cells[iRowNo, 10] = "Ave 4's";
            xlWrkshtPlyrStts.Cells[iRowNo, 11] = "Ave 6's";
            xlWrkshtPlyrStts.Cells[iRowNo, 12] = "BllFcd";
            xlWrkshtPlyrStts.Cells[iRowNo, 13] = "Bowling";
            xlWrkshtPlyrStts.Cells[iRowNo, 14] = "Inngs";
            xlWrkshtPlyrStts.Cells[iRowNo, 15] = "Ave";
            xlWrkshtPlyrStts.Cells[iRowNo, 16] = "Wkts";
            xlWrkshtPlyrStts.Cells[iRowNo, 17] = "Rns Cncd";
            xlWrkshtPlyrStts.Cells[iRowNo, 18] = "Ovrs";
            xlWrkshtPlyrStts.Cells[iRowNo, 19] = "RPO";
            xlWrkshtPlyrStts.Cells[iRowNo, 20] = "SR";
            xlWrkshtPlyrStts.Cells[iRowNo, 21] = "BllBwld";
            xlWrkshtPlyrStts.Cells[iRowNo + 4, 3] = "Both";
        }

        public void LoadPlyrsStatsThisLast(int i, List<Season> seasns, string sThisLast, CricketForm fm)
        {
            // Load ALL players from this comp /season
            // Search Players Worksheet
            int iLastRow = GetLastRowPlyrSts(1);
            int iNoPlayers = -1;
            Player CurrPlyr = new Player();
            Season OnePlayer = new Season();
            OnePlayer = seasns[i];


            if (iLastRow < 6)
            {
                iNoPlayers = -1;
            }
            else
            {
                iNoPlayers = (iLastRow - 3) / 6;
            }
            for (int x = 0; x <= iNoPlayers; x++)
            {
                int iThisLastRow = (x * 6) + 6;
            }
        }

        private string GetCompFromCompSeason(string sThisAndLastSeason)
        {
            string[] aAll;

            bool bHasEqualsSign = sThisAndLastSeason.Contains('=');
            if (bHasEqualsSign)
            {
                aAll = sThisAndLastSeason.Split('=');
                return aAll[0];
            }
            else
            {
                MessageBox.Show("Error in GetCompFromCompSeason: No = sign in sThisAndLastSeason");
                return "Error";
            }
        }
        private string GetSeasonFromCompSeason(string sThisAndLastSeason)
        {
            string[] aAll;

            bool bHasEqualsSign = sThisAndLastSeason.Contains('=');
            if (bHasEqualsSign)
            {
                aAll = sThisAndLastSeason.Split('=');
                return aAll[1];
            }
            else
            {
                MessageBox.Show("Error in GetSeasonFromCompSeason: No = sign in sThisAndLastSeason");
                return "Error";
            }
        }

        private PlyrSts PopulatePlayStats(int iFirstLastRow)
        {

            PlyrSts PlaySts = new PlyrSts();
            StatsBat SttsBt = new StatsBat();
            StatsBowl SttsBwl = new StatsBowl();

            string sCompSsn = xlWrkshtPlyrStts.Cells[1, 4].Value;
            string sComp = GetCompFromCompSeason(sCompSsn);
            string sSeason = GetSeasonFromCompSeason(sCompSsn);
            PlaySts.Comp = sComp;
            PlaySts.ThisLastSsn = sSeason;
            PlaySts.BeforeThisLastSsn = GetTwoSeasonsAgo(sSeason);


            SttsBt.Inngs = xlWrkshtPlyrStts.Cells[iFirstLastRow, 4].Text as string;
            SttsBt.Runs = xlWrkshtPlyrStts.Cells[iFirstLastRow, 5].Text as string;
            SttsBt.Ave = xlWrkshtPlyrStts.Cells[iFirstLastRow, 6].Text as string;
            SttsBt.SR = xlWrkshtPlyrStts.Cells[iFirstLastRow, 7].Text as string;
            SttsBt.HS = xlWrkshtPlyrStts.Cells[iFirstLastRow, 8].Text as string;
            SttsBt.NOs = xlWrkshtPlyrStts.Cells[iFirstLastRow, 9].Text as string;
            SttsBt.Ave4s = xlWrkshtPlyrStts.Cells[iFirstLastRow, 10].Text as string;
            SttsBt.Ave6s = xlWrkshtPlyrStts.Cells[iFirstLastRow, 11].Text as string;
            PlaySts.SttsBt = SttsBt;


            SttsBwl.Inngs = xlWrkshtPlyrStts.Cells[iFirstLastRow, 14].Text as string;
            SttsBwl.Ave = xlWrkshtPlyrStts.Cells[iFirstLastRow, 15].Text as string;
            SttsBwl.Wkts = xlWrkshtPlyrStts.Cells[iFirstLastRow, 16].Text as string;
            SttsBwl.Runs = xlWrkshtPlyrStts.Cells[iFirstLastRow, 17].Text as string;
            SttsBwl.Ovrs = xlWrkshtPlyrStts.Cells[iFirstLastRow, 18].Text as string;
            SttsBwl.RPO = xlWrkshtPlyrStts.Cells[iFirstLastRow, 19].Text as string;
            SttsBwl.SR = xlWrkshtPlyrStts.Cells[iFirstLastRow, 20].Text as string;
            SttsBwl.BllBwl = xlWrkshtPlyrStts.Cells[iFirstLastRow, 21].Text as string;
            PlaySts.SttsBwl = SttsBwl;

            return PlaySts;
        }



        public List<Player> LoadPlayerList()
        {
            List<Player> plyrs = new List<Player>();
            int iNoPlayers = -1;

            InitExcel();
            xlWrkshtPlyrStts = (Excel.Worksheet)xlWorkbook.Worksheets[gbl.SheetType.sPLAYERLIST];
            int iLastRow = GetLastRow(2, gbl.SheetType.sPLAYERLIST);
            if (iLastRow < 4)
            {
                iNoPlayers = 0;
                return null;
            }
            else
            {
                iNoPlayers = (iLastRow + 1) / 6;
            }

            for (int i = 0; i < iNoPlayers; i++)
            {
                Player OnePlayer = new Player();
                string[] aAll;

                OnePlayer.LastName = xlWrkshtPlyrStts.Cells[(i * 6) + 4, 2].Value.ToString().Trim();
                string sFirstNames = xlWrkshtPlyrStts.Cells[(i * 6) + 5, 2].Value.ToString().Trim();
                aAll = sFirstNames.Split('=');
                OnePlayer.FirstName = aAll[0];
                OnePlayer.MiddleName = aAll[1];
                xlWrkshtPlyrStts.Cells[(i * 6) + 6, 2].Value = "";

                string sDOB = xlWrkshtPlyrStts.Cells[(i * 6) + 7, 2].Value.ToString().Trim();

                OnePlayer.DOB = StringToDate(sDOB);
                OnePlayer.Country = xlWrkshtPlyrStts.Cells[(i * 6) + 8, 2].Value.ToString().Trim();
                plyrs.Add(OnePlayer);
            }
            CleanUpExcel();
            return plyrs;
        }

        public List<Player> AddToPlayerListIfNotThere(Player CurrPlyr, List<Player> plyrs)
        {
            if (CurrPlyr == null)
            {
                return plyrs;
            }
            else
            {
                if (plyrs != null)
                {
                    for (int i = 0; i < plyrs.Count; i++)
                    {
                        if (IsSamePlayer(CurrPlyr, plyrs[i]))
                        {
                            return plyrs;
                        }
                    }
                }

                xlWrkshtPlyrStts = (Excel.Worksheet)xlWorkbook.Worksheets[gbl.SheetType.sPLAYERLIST];

                int iLastRow = GetLastRow(2, gbl.SheetType.sPLAYERLIST);
                int iNewRow = -1;
                if (iLastRow < 5)
                {
                    iNewRow = iLastRow + 4;
                }
                else
                {
                    iNewRow = iLastRow + 5;
                }

                xlWrkshtPlyrStts.Cells[iNewRow, 2].Value = CurrPlyr.LastName.ToString().Trim();
                string sFirstNames = CurrPlyr.FirstName + "=" + CurrPlyr.MiddleName;
                xlWrkshtPlyrStts.Cells[iNewRow + 1, 2].Value = sFirstNames;
                xlWrkshtPlyrStts.Cells[iNewRow + 2, 2].Value = "";
                xlWrkshtPlyrStts.Cells[iNewRow + 3, 2].Value = DateToString(CurrPlyr.DOB);
                xlWrkshtPlyrStts.Cells[iNewRow + 4, 2].Value = CurrPlyr.Country;

                xlWrkshtPlyrStts.Columns[2].Font.Bold = true;
                xlWrkshtPlyrStts.Columns[2].HorizontalAlignment = XlHAlign.xlHAlignCenter;


                //xlWrkshtPlyrStts = (Excel.Worksheet)xlWorkbook.Worksheets[xlWrkshtPlayerStats];

                if (plyrs == null)
                {
                    List<Player> pls = new List<Player>();
                    pls.Add(CurrPlyr);
                    return pls;
                }
                else
                {
                    plyrs.Add(CurrPlyr);
                    return plyrs;
                }
            }
        }

        private List<Player> GetPlayerStatsList()
        {
            List<Player> plysts = new List<Player>();

            int iNoPlayers = -1;
            string[] aAll;


            //xlWrkshtPlyrStts = (Excel.Worksheet)xlWorkbook.Worksheets[xlWrkshtPlayerStats];
            int iLastRow = GetLastRowPlyrSts(1);
            if (iLastRow < 9)
            {
                iNoPlayers = 0;
                return plysts;
            }
            else
            {
                iNoPlayers = (iLastRow - 3) / 6;
            }
            for (int i = 0; i < iNoPlayers; i++)
            {
                int iFirstRow = (i * 6) + 4;
                Player CurrPlyr = new Player();
                CurrPlyr.LastName = xlWrkshtPlyrStts.Cells[iFirstRow, 2].Value;
                string sFirstNames = xlWrkshtPlyrStts.Cells[iFirstRow + 1, 2].Value;
                aAll = sFirstNames.Split('=');
                CurrPlyr.FirstName = aAll[0];
                CurrPlyr.MiddleName = aAll[1];
                string sDOB = xlWrkshtPlyrStts.Cells[iFirstRow + 3, 2].Value;
                DateTime dDOB = StringToDate(sDOB);
                CurrPlyr.DOB = dDOB;
                CurrPlyr.Country = xlWrkshtPlyrStts.Cells[iFirstRow + 4, 2].Value;
                plysts.Add(CurrPlyr);
            }

            return plysts;
        }



        private int CurrPlayerIsInPlayerStats(Player CurrPlayer, List<Player> plyrsts)
        {  // Returns -1 if Player is NOT in plyrsts list otherwise returns the first row number of CurrPlayer in PlayerStats
            int iPlayerRowNo;
            for (int i = 0; i < plyrsts.Count; i++)
            {
                iPlayerRowNo = (i * 6) + 4;
                if (IsSamePlayer(CurrPlayer, plyrsts[i]))
                {
                    return iPlayerRowNo;
                }
            }
            return -1;
        }

        private void MoveOnePlayerUp(int j)
        {
            for (int k = 1; k <= 21; k++)
            {
                xlWrkshtPlyrStts.Cells[j, k] = xlWrkshtPlyrStts.Cells[j + 6, k].Text as string;
                xlWrkshtPlyrStts.Cells[j - 1, k] = xlWrkshtPlyrStts.Cells[j + 5, k].Text as string;
                xlWrkshtPlyrStts.Cells[j - 2, k] = xlWrkshtPlyrStts.Cells[j + 4, k].Text as string;
                xlWrkshtPlyrStts.Cells[j - 3, k] = xlWrkshtPlyrStts.Cells[j + 3, k].Text as string;
                xlWrkshtPlyrStts.Cells[j - 4, k] = xlWrkshtPlyrStts.Cells[j + 2, k].Text as string;
                xlWrkshtPlyrStts.Cells[j - 5, k] = xlWrkshtPlyrStts.Cells[j + 1, k].Text as string;
            }
        }

        private void MoveAllPlayerStatsDown(string sNewSeason)
        {
            int iLastRow = GetLastRowPlyrSts(1);
            int iNoPlayers = -1;
            string[] aAll;

            if (iLastRow < 9)
            {
                iNoPlayers = 0;
                return;
            }
            else
            {
                iNoPlayers = (iLastRow - 3) / 6;
            }
            aAll = sNewSeason.Split('=');
            string sNewSsn = aAll[1];
            aAll = sNewSsn.Split('-');
            sNewSsn = aAll[0];
            int iNewSsn = -1;
            int i = iLastRow;
            int iNewCurrPlyRow = -1;
            int iOldCurrPlyRow = iLastRow + 6;
            while (i >= 4)
            {


                iNewCurrPlyRow = iOldCurrPlyRow - 6;

                string sUpper = xlWrkshtPlyrStts.Cells[i - 4, 3].Text as string;
                string sLower = xlWrkshtPlyrStts.Cells[i - 3, 3].Text as string;

                MovePlayerStatsDown(i);


                int iUpper = -1;
                if (sUpper != "")
                {
                    string sPre = sUpper;
                    aAll = sPre.Split('-');
                    sPre = aAll[0];
                    iUpper = Convert.ToInt32(sPre);
                    iNewSsn = Convert.ToInt32(sNewSsn);
                }

                if (iNewSsn - iUpper >= 2)
                //if ((sLower == "") && (sUpper == ""))
                {    // Shuffle everything up
                    for (int j = iNewCurrPlyRow; j <= iLastRow + 6; j = j + 6)
                    {
                        MoveOnePlayerUp(j);
                    }
                    iLastRow = iLastRow - 6;
                }
                i = i - 6;
                iOldCurrPlyRow = iNewCurrPlyRow;
            }
        }

        private void MovePlayerStatsDown(int iRow)
        {
            // newrow := upper
            for (int i = 3; i <= 21; i++)
            {

                string sUpper = xlWrkshtPlyrStts.Cells[iRow - 4, i].Text as string;
                string sLower = xlWrkshtPlyrStts.Cells[iRow - 3, i].Text as string;

                xlWrkshtPlyrStts.Cells[iRow - 3, i].Value = sUpper;
                xlWrkshtPlyrStts.Cells[iRow - 4, i] = "";
                xlWrkshtPlyrStts.Cells[iRow - 1, i].Value = "";

            }
        }

        private void MovePlayerStatsUp(int iRow)
        {

            for (int i = 3; i <= 21; i++)
            {

                string sUpper = xlWrkshtPlyrStts.Cells[iRow - 4, i].Text as string;
                string sLower = xlWrkshtPlyrStts.Cells[iRow - 3, i].Text as string;

                xlWrkshtPlyrStts.Cells[iRow - 4, i].Value = sLower;
                xlWrkshtPlyrStts.Cells[iRow - 3, i] = "";
                xlWrkshtPlyrStts.Cells[iRow - 1, i].Value = sLower;
            }
            xlWrkshtPlyrStts.Cells[iRow - 1, 3].Value = "Both";
        }

        private void MoveSingleRowUpOne()
        {
            int iLastRow = GetLastRowPlyrSts(1);

            if (iLastRow < 9)
            {
                return;
            }
            for (int i = 9; i <= iLastRow; i = i + 6)
            {
                string sUpper = xlWrkshtPlyrStts.Cells[i - 4, 3].Text as string;
                string sLower = xlWrkshtPlyrStts.Cells[i - 3, 3].Text as string;
                if (sUpper == "" && sLower != "")
                {
                    MovePlayerStatsUp(i);
                }
            }
        }

        public void LoadPlyrStatsThisLast(List<Season> ssns, CricketForm fm)
        {
            string sMyTeamCode = "";
            string sFirstName = "";
            string sMiddleName = "";
            string sTeamCtr = "";
            List<Player> plyrs = new List<Player>();
            List<Player> plyrsts = new List<Player>();
            int iFirstRow = -1; ;

            if (ssns.Count == 0)
            {
                return;
            }
            InitExcel();
            plyrs = LoadPlayerList();
            xlWorksheet = (Excel.Worksheet)xlWorkbook.Worksheets[gbl.SheetType.sPLAYERSTATS];
            //xlWorksheet = (Excel.Worksheet)xlWorkbook.Worksheets[gbl.SheetType.sMATCHES];
            string sCurrentTeamCode = ssns[0].TeamMine;
            string sCOMPLETEDSEASON = xlWorksheet.Cells[1, 4].Value;
            MoveAllPlayerStatsDown(sCOMPLETEDSEASON);
            plyrsts = GetPlayerStatsList();
            for (int i = 0; i < ssns.Count; i++)
            {
                plyrs = AddToPlayerListIfNotThere(ssns[i].Playr, plyrs);
                string sCurrPlyrComp = ssns[i].Comp.CompetitionCode + "=" + ssns[i].Comp.Season;
                if (sCurrPlyrComp == sCOMPLETEDSEASON)
                {  // PlayerStats MUST be updated
                    Player CurrPlayer = new Player();
                    CurrPlayer = ssns[i].Playr;
                    // Either this player is already in PlayerSats or he isn't
                    int iFstRowForPlayerFound = CurrPlayerIsInPlayerStats(CurrPlayer, plyrsts);
                    if (iFstRowForPlayerFound != -1)
                    {  // Player IS already in PlayerStats 
                        xlWorksheet.Cells[iFstRowForPlayerFound + 1, 3] = ssns[i].Comp.Season;
                        PopulateDataBat(iFstRowForPlayerFound, ssns[i]);
                        PopulateDataBowl(iFstRowForPlayerFound, ssns[i]);
                    }
                    else
                    {
                        // Player IS NOT already in PlayerStats 
                        int iLastRow = GetLastRowPlyrSts(1);
                        if (iLastRow < 9)
                        {
                            iFirstRow = 4;
                        }
                        else
                        {
                            iFirstRow = iLastRow + 1;
                        }
                        xlWorksheet.Cells[iFirstRow, 2] = ssns[i].Playr.LastName;
                        xlWorksheet.Cells[iFirstRow, 2].Font.Underline = true;
                        sFirstName = ssns[i].Playr.FirstName;
                        sMiddleName = ssns[i].Playr.MiddleName;
                        xlWorksheet.Cells[iFirstRow + 1, 2] = sFirstName + "=" + sMiddleName;
                        xlWorksheet.Cells[iFirstRow + 2, 2] = ssns[i].TeamMine;
                        string sLastSeason = GetTwoSeasonsAgo(ssns[i].Comp.Season);
                        DateTime dDOB = ssns[i].Playr.DOB;
                        string sDBrth = dDOB.ToString();
                        sDBrth = GetDOB(sDBrth);
                        xlWorksheet.Cells[iFirstRow + 3, 2] = sDBrth;
                        xlWorksheet.Cells[iFirstRow + 4, 2] = ssns[i].Playr.Country;
                        sMyTeamCode = ssns[i].TeamMine;
                        sTeamCtr = GetTeamPlayerCount(sMyTeamCode);
                        int iTeamCtr = Convert.ToInt32(sTeamCtr);

                        if (iTeamCtr <= 9)
                        {
                            sTeamCtr = "0" + sTeamCtr;
                        }
                        xlWorksheet.Cells[iFirstRow, 1] = sMyTeamCode + "." + sTeamCtr + "." + "1";
                        xlWorksheet.Cells[iFirstRow + 1, 1] = sMyTeamCode + "." + sTeamCtr + "." + "2";
                        xlWorksheet.Cells[iFirstRow + 2, 1] = sMyTeamCode + "." + sTeamCtr + "." + "3";
                        xlWorksheet.Cells[iFirstRow + 3, 1] = sMyTeamCode + "." + sTeamCtr + "." + "4";
                        xlWorksheet.Cells[iFirstRow + 4, 1] = sMyTeamCode + "." + sTeamCtr + "." + "5";
                        xlWorksheet.Cells[iFirstRow + 5, 1] = sMyTeamCode + "." + sTeamCtr + "." + "6";
                        iTeamCtr++;
                        xlWorksheet.Cells[iFirstRow + 1, 3] = ssns[i].Comp.Season;
                        LoadHeading(iFirstRow);
                        PopulateDataBat(iFirstRow, ssns[i]);
                        PopulateDataBowl(iFirstRow, ssns[i]);
                    }
                }
            }
            MoveSingleRowUpOne();
            xlApp.DisplayAlerts = false;
            xlWorkbook.SaveAs(gbl.sEXCELPATH, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing,
                 Type.Missing, Type.Missing);
            CleanUpExcel();
        }



        private string GetTeamPlayerCount(string sMyTeamCode)
        {
            int iLastRow = GetLastRowPlyrSts(1);
            int iNoPlayers = -1;
            string[] aAll;
            int iTeamCtr = 0;

            if (iLastRow < 9)
            {
                iNoPlayers = 0;
                return "1";

            }
            else
            {
                iNoPlayers = (iLastRow - 3) / 6;
            }
            for (int i = 0; i < iNoPlayers; i++)
            {
                int iFirstRow = (i * 6) + 4;
                string sCurrFstRow = xlWrkshtPlyrStts.Cells[iFirstRow, 1].Text as string;
                aAll = sCurrFstRow.Split('.');
                string sCurrTeam = aAll[0];
                if (sCurrTeam == sMyTeamCode)
                {
                    iTeamCtr++;
                }
            }
            iTeamCtr++;
            return iTeamCtr.ToString();
        }

        private bool PlayerIsNew(Player plyr, Player currentplr)
        {
            if (plyr.FirstName == currentplr.FirstName && plyr.MiddleName == currentplr.MiddleName && plyr.LastName == currentplr.LastName && plyr.DOB == currentplr.DOB && plyr.Country == currentplr.Country)
            {
                return false;
            }
            else
            {
                return true;
            }
        }



        private void PopulateDataBat(int iRow, Season OnePlayer)
        {
            //PopulateData Data for 1 Batsman
            InngsSsnBat BatSsn = new InngsSsnBat();
            List<AllInnFrSsn> AllInn = new List<AllInnFrSsn>();
            ResultsBatsman rBat = new ResultsBatsman();
            List<AllInnFrSsn> AllInnForBowlers = new List<AllInnFrSsn>();

            BatSsn.Inngs = 0;
            BatSsn.Runs = 0;
            BatSsn.Total4s = 0;
            BatSsn.Total6s = 0;
            BatSsn.NOs = 0;
            BatSsn.HS = 0;
            BatSsn.BallsFaced = 0;

            AllInn = OnePlayer.AllInnings;
            for (int i = 0; i < AllInn.Count; i++)
            {
                rBat = AllInn[i].ResBat;
                if (rBat != null)
                {
                    if (rBat.IsSignificant == true)
                    {
                        if (rBat.NotOut == false)
                        {
                            BatSsn.Inngs = BatSsn.Inngs + 1;
                        }
                        else
                        {
                            BatSsn.NOs = BatSsn.NOs + 1;
                        }
                        int iRuns = Convert.ToInt32(rBat.Runs);
                        BatSsn.Runs = BatSsn.Runs + iRuns;

                        int iFours = Convert.ToInt32(rBat.Four);
                        BatSsn.Total4s = BatSsn.Total4s + iFours;

                        int iSixs = Convert.ToInt32(rBat.Six);
                        BatSsn.Total6s = BatSsn.Total6s + iSixs;

                        int iBallsFaced = Convert.ToInt32(rBat.BallsFaced);
                        BatSsn.BallsFaced = BatSsn.BallsFaced + iBallsFaced;

                        if (iRuns > BatSsn.HS)
                        {
                            BatSsn.HS = iRuns;
                        }
                    }
                }
            }
            string sAv = "";
            decimal dAv = 0;
            decimal dTotalRuns = Convert.ToDecimal(BatSsn.Runs);
            decimal dTotalInnings = Convert.ToDecimal(BatSsn.Inngs);
            decimal dTotal4s = Convert.ToDecimal(BatSsn.Total4s);
            decimal dTotal6s = Convert.ToDecimal(BatSsn.Total6s);
            decimal dBallsFaced = Convert.ToDecimal(BatSsn.BallsFaced);

            if (BatSsn.Inngs > 0)
            {
                dAv = dTotal4s / dTotalInnings;
                dAv = Math.Round(dAv, 2);
                sAv = dAv.ToString().Trim();

                xlWrkshtPlyrStts.Cells[iRow + 1, 10] = sAv;


                dAv = dTotal6s / dTotalInnings;
                dAv = Math.Round(dAv, 2);
                sAv = dAv.ToString().Trim();
                xlWrkshtPlyrStts.Cells[iRow + 1, 11] = sAv;

                dAv = dTotalRuns / dTotalInnings;
                dAv = Math.Round(dAv, 2);
                sAv = dAv.ToString().Trim();
                xlWrkshtPlyrStts.Cells[iRow + 1, 6] = sAv;
            }
            else
            {
                xlWrkshtPlyrStts.Cells[iRow + 1, 10] = "";
                xlWrkshtPlyrStts.Cells[iRow + 1, 11] = "";
                xlWrkshtPlyrStts.Cells[iRow + 1, 6] = "";
            }
            if ((BatSsn.BallsFaced > 0) && (BatSsn.Inngs > 0))
            {
                dAv = (dTotalRuns / BatSsn.BallsFaced) * 100;
                dAv = Math.Round(dAv, 2);
                sAv = dAv.ToString().Trim();
                xlWrkshtPlyrStts.Cells[iRow + 1, 7] = sAv;
            }
            else
            {
                xlWrkshtPlyrStts.Cells[iRow + 1, 7] = "";
            }

            if (BatSsn.Inngs > 0)
            {
                xlWrkshtPlyrStts.Cells[iRow + 1, 4] = BatSsn.Inngs.ToString().Trim();
                xlWrkshtPlyrStts.Cells[iRow + 1, 5] = BatSsn.Runs.ToString().Trim();
                xlWrkshtPlyrStts.Cells[iRow + 1, 8] = BatSsn.HS.ToString().Trim();
                xlWrkshtPlyrStts.Cells[iRow + 1, 9] = BatSsn.NOs.ToString().Trim(); ;
                xlWrkshtPlyrStts.Cells[iRow + 1, 12] = BatSsn.BallsFaced;
            }

            string BeforeRowCol = xlWrkshtPlyrStts.Cells[iRow + 2, 4].Text as string;
            StatsBatDbl BothSsns = new StatsBatDbl();

            SumAve SmAv = new SumAve();


            //NO's
            SmAv = AverageTwoValues(xlWrkshtPlyrStts.Cells[iRow + 1, 9].Text as string, xlWrkshtPlyrStts.Cells[iRow + 2, 9].Text as string);
            double dSUMnos = SmAv.Sum;
            BothSsns.NOs = SmAv.Ave;
            //double  = Convert.ToDouble(xlWrkshtPlyrStts.Cells[iRow + 1, 9].Text as string) + Convert.ToDouble(xlWrkshtPlyrStts.Cells[iRow + 2, 9].Text as string);


            //BallsFaced
            SmAv = new SumAve();
            SmAv = AverageTwoValues(xlWrkshtPlyrStts.Cells[iRow + 1, 12].Text as string, xlWrkshtPlyrStts.Cells[iRow + 2, 12].Text as string);
            double dSUMbf = SmAv.Sum;
            BothSsns.BllFcd = SmAv.Ave;


            // Inngs Total
            SmAv = new SumAve();
            SmAv = AverageTwoValues(xlWrkshtPlyrStts.Cells[iRow + 1, 4].Text as string, xlWrkshtPlyrStts.Cells[iRow + 2, 4].Text as string);
            double dSUMInngs = SmAv.Sum;
            BothSsns.Inngs = SmAv.Ave;


            // Net Innings
            if (dSUMInngs == -1)
            {
                BothSsns.Inngs = -1;
            }

            //Runs Total
            SmAv = new SumAve();
            SmAv = AverageTwoValues(xlWrkshtPlyrStts.Cells[iRow + 1, 5].Text as string, xlWrkshtPlyrStts.Cells[iRow + 2, 5].Text as string);
            double dSUMRuns = SmAv.Sum;

            //BothSsns.Runs = SmAv.Ave;
            BothSsns.Runs = dSUMRuns;
            // Average
            if (dSUMInngs > 0)
            {
                BothSsns.Ave = dSUMRuns / dSUMInngs;
            }
            else
            {
                BothSsns.Ave = -1;
            }

            // SR
            // (Runs / BF) * 100
            if ((dSUMbf > 0) && (BatSsn.Inngs > 0))
            {
                BothSsns.SR = (dSUMRuns / dSUMbf) * 100;
            }
            else
            {
                BothSsns.SR = -1;
            }

            string sSsn1HS = xlWrkshtPlyrStts.Cells[iRow + 1, 8].Text as string;
            string sSsn2HS = xlWrkshtPlyrStts.Cells[iRow + 2, 8].Text as string;
            if (sSsn1HS == "")
            {
                if (sSsn2HS != "")
                {
                    BothSsns.HS = Convert.ToDouble(sSsn2HS);
                }
                else
                {
                    BothSsns.HS = -1;
                }
            }
            if (sSsn1HS != "")
            {
                if (sSsn2HS == "")
                {
                    BothSsns.HS = Convert.ToDouble(sSsn1HS);
                }
                else
                {
                    double dFirst = Convert.ToDouble(sSsn1HS);
                    double dSecond = Convert.ToDouble(sSsn2HS);

                    if (dFirst > dSecond)
                    {
                        BothSsns.HS = dFirst;
                    }
                    else
                    {
                        BothSsns.HS = dSecond;
                    }
                }
            }


            string sSsn1Inngs = xlWrkshtPlyrStts.Cells[iRow + 1, 4].Text as string;
            string sSsn1Ave4s = xlWrkshtPlyrStts.Cells[iRow + 1, 10].Text as string;

            string sSsn2Inngs = xlWrkshtPlyrStts.Cells[iRow + 2, 4].Text as string;
            string sSsn2Ave4s = xlWrkshtPlyrStts.Cells[iRow + 2, 10].Text as string;



            BothSsns.Ave4s = GetAve4Ave6(sSsn1Inngs, sSsn1Ave4s, sSsn2Inngs, sSsn2Ave4s);
            //BatSsn.Inngs 
            //BothSsns.Inngs


            //Ave 6's
            string dSsn1Ave6s = xlWrkshtPlyrStts.Cells[iRow + 1, 11].Text as string;
            string dSsn2Ave6s = xlWrkshtPlyrStts.Cells[iRow + 2, 11].Text as string;
            BothSsns.Ave6s = GetAve4Ave6(sSsn1Inngs, dSsn1Ave6s, sSsn2Inngs, dSsn2Ave6s);

            if ((BothSsns.Inngs != -1) && (BothSsns.Inngs != 0))
            {
                xlWrkshtPlyrStts.Cells[iRow + 4, 4] = BothSsns.Inngs.ToString();
            }
            else
            {
                xlWrkshtPlyrStts.Cells[iRow + 4, 4] = "";
            }

            if (BothSsns.Runs != -1)
            {
                //BothSsns.Runs = dSUMRuns / dSUMInngs;
                xlWrkshtPlyrStts.Cells[iRow + 4, 5] = BothSsns.Runs.ToString();
            }
            else
            {
                xlWrkshtPlyrStts.Cells[iRow + 4, 5] = "";
            }


            if (BothSsns.Ave != -1)
            {
                xlWrkshtPlyrStts.Cells[iRow + 4, 6] = BothSsns.Ave.ToString();
            }
            else
            {
                xlWrkshtPlyrStts.Cells[iRow + 4, 6] = "";
            }


            if (BothSsns.SR != -1)
            {
                xlWrkshtPlyrStts.Cells[iRow + 4, 7] = BothSsns.SR.ToString();
            }
            else
            {
                xlWrkshtPlyrStts.Cells[iRow + 4, 7] = "";
            }


            if (BothSsns.HS != -1)
            {
                xlWrkshtPlyrStts.Cells[iRow + 4, 8] = BothSsns.HS.ToString();
            }
            else
            {
                xlWrkshtPlyrStts.Cells[iRow + 4, 8] = "";
            }
            if (BothSsns.NOs != -1)
            {
                BothSsns.NOs = dSUMnos / 2;
                xlWrkshtPlyrStts.Cells[iRow + 4, 9] = BothSsns.NOs.ToString();
            }
            else
            {
                xlWrkshtPlyrStts.Cells[iRow + 4, 9] = "";
            }

            if (BothSsns.BllFcd != -1)
            {
                xlWrkshtPlyrStts.Cells[iRow + 4, 12] = BothSsns.BllFcd.ToString();
            }
            else
            {
                xlWrkshtPlyrStts.Cells[iRow + 4, 12] = "";
            }

            if (BothSsns.Ave4s != -1)
            {
                xlWrkshtPlyrStts.Cells[iRow + 4, 10] = BothSsns.Ave4s.ToString();
            }
            else
            {
                xlWrkshtPlyrStts.Cells[iRow + 4, 10] = "";
            }

            if (BothSsns.Ave6s != -1)
            {
                xlWrkshtPlyrStts.Cells[iRow + 4, 11] = BothSsns.Ave6s.ToString();
            }
            else
            {
                xlWrkshtPlyrStts.Cells[iRow + 4, 11] = "";
            }
            xlWrkshtPlyrStts.Cells[iRow + 4, 3] = "Both";
        }

        private double GetAve4Ave6(string sSsn1Inngs, string sSsn1Ave, string sSsn2Inngs, string sSsn2Ave)
        {

            if ((sSsn1Inngs == "") && (sSsn2Inngs == ""))
            {
                return -1;
            }
            else if ((sSsn1Inngs != "") && (sSsn2Inngs == ""))
            {
                return Convert.ToDouble(sSsn1Ave);
            }
            else if ((sSsn1Inngs == "") && (sSsn2Inngs != ""))
            {
                return Convert.ToDouble(sSsn2Ave);
            }
            else
            {
                double dSsn1Inngs = Convert.ToDouble(sSsn1Inngs);
                double dSsn1Ave = Convert.ToDouble(sSsn1Ave);
                double dSsn2Inngs = Convert.ToDouble(sSsn2Inngs);
                double dSsn2Ave = Convert.ToDouble(sSsn2Ave);

                return ((dSsn1Ave * dSsn1Inngs) + (dSsn2Ave * dSsn2Inngs)) / (dSsn1Inngs + dSsn2Inngs);
            }
        }


        private int GetBallsBowled(int OversBowled)
        {
            string[] aAll;
            string sOversBwld = OversBowled.ToString().Trim();
            bool bHasFullstop = sOversBwld.Contains('.');
            if (bHasFullstop)
            {
                aAll = sOversBwld.Split('.');
                string sPre = aAll[0];
                string sPost = aAll[1];
                int iPre = Convert.ToInt32(sPre);
                int iPost = Convert.ToInt32(sPost);
                return (iPre * 6) + (iPost);
            }
            else
            {
                return OversBowled * 6;
            }
        }

        private void PopulateDataBowl(int iRow, Season OnePlayer)
        {
            //PopulateData Data for 1 Bowler
            InngsSsnBowl BowlSsn = new InngsSsnBowl();
            List<AllInnFrSsn> AllInn = new List<AllInnFrSsn>();
            ResultsBowler rBowl = new ResultsBowler();

            BowlSsn.Wkts = 0;
            BowlSsn.RunsConc = 0;
            BowlSsn.OversBwld = 0;
            BowlSsn.NoInn = 0;
            BowlSsn.BallsBwld = 0;
            AllInn = OnePlayer.AllInnings;

            decimal dWckts = 0;
            decimal dRunsConceded = 0;
            decimal dOversBowled = 0;
            for (int i = 0; i < AllInn.Count; i++)
            {

                rBowl = AllInn[i].ResBowl;
                if (rBowl != null)
                {
                    if (rBowl.IsSignificant == true)
                    {
                        BowlSsn.NoInn = BowlSsn.NoInn + 1;

                        dWckts = Convert.ToDecimal(rBowl.Wickets);
                        BowlSsn.Wkts = BowlSsn.Wkts + dWckts;

                        dRunsConceded = Convert.ToDecimal(rBowl.RunsConceded);
                        BowlSsn.RunsConc = BowlSsn.RunsConc + dRunsConceded;

                        dOversBowled = Convert.ToDecimal(rBowl.OversBowled);
                        int iBallsBowled = GetBallsBowled(Convert.ToInt32(rBowl.OversBowled));
                        dOversBowled = Convert.ToDecimal(iBallsBowled / 6);
                        BowlSsn.OversBwld = BowlSsn.OversBwld + dOversBowled;
                        BowlSsn.BallsBwld = BowlSsn.BallsBwld + Convert.ToDecimal(iBallsBowled);
                    }
                }
            }

            if (BowlSsn.NoInn != 0)
            {

                string sAv = "";
                decimal dAv = 0;
                decimal RPO = 0;

                if (BowlSsn.OversBwld > 0)
                {
                    RPO = BowlSsn.RunsConc / BowlSsn.OversBwld;
                    xlWrkshtPlyrStts.Cells[iRow + 1, 19] = RPO.ToString().Trim();
                }
                else
                {
                    xlWrkshtPlyrStts.Cells[iRow + 1, 19] = 0;
                }

                if (BowlSsn.NoInn > 0)
                {
                    dAv = BowlSsn.RunsConc / BowlSsn.Wkts;
                    dAv = Math.Round(dAv, 2);
                    sAv = dAv.ToString().Trim();
                    xlWrkshtPlyrStts.Cells[iRow + 1, 15] = sAv;
                }
                else
                {
                    xlWrkshtPlyrStts.Cells[iRow + 1, 15] = 0;
                }
                if (BowlSsn.Wkts > 0)
                {
                    // SR
                    dAv = BowlSsn.BallsBwld / BowlSsn.Wkts;
                    dAv = Math.Round(dAv, 2);
                    sAv = dAv.ToString().Trim();
                    xlWrkshtPlyrStts.Cells[iRow + 1, 20] = sAv;
                }
                else
                {
                    xlWrkshtPlyrStts.Cells[iRow + 1, 20] = 0;
                }
                xlWrkshtPlyrStts.Cells[iRow + 1, 14] = BowlSsn.NoInn.ToString().Trim();
                xlWrkshtPlyrStts.Cells[iRow + 1, 16] = BowlSsn.Wkts.ToString().Trim();
                xlWrkshtPlyrStts.Cells[iRow + 1, 17] = BowlSsn.RunsConc.ToString().Trim();
                xlWrkshtPlyrStts.Cells[iRow + 1, 21] = BowlSsn.BallsBwld.ToString().Trim();
                xlWrkshtPlyrStts.Cells[iRow + 1, 18] = BowlSsn.OversBwld.ToString().Trim();


                string BeforeRowCol = xlWrkshtPlyrStts.Cells[iRow + 2, 4].Text as string;
                StatsBowlDbl BothInngs = new StatsBowlDbl();
                if (!string.IsNullOrWhiteSpace(BeforeRowCol))
                {
                    // BallsBowled
                    SumAve SmAv = new SumAve();
                    SmAv = AverageTwoValues(xlWrkshtPlyrStts.Cells[iRow + 1, 21].Text as string, xlWrkshtPlyrStts.Cells[iRow + 2, 21].Text as string);
                    double dSUMBllBwld = SmAv.Sum;
                    BothInngs.BllBwl = SmAv.Ave;

                    // Inngs Total
                    SmAv = new SumAve();
                    SmAv = AverageTwoValues(xlWrkshtPlyrStts.Cells[iRow + 1, 14].Text as string, xlWrkshtPlyrStts.Cells[iRow + 2, 14].Text as string);
                    double dSUMInngs = SmAv.Sum;
                    BothInngs.Inngs = SmAv.Ave;

                    // RunsCncd
                    SmAv = new SumAve();
                    SmAv = AverageTwoValues(xlWrkshtPlyrStts.Cells[iRow + 1, 17].Text as string, xlWrkshtPlyrStts.Cells[iRow + 2, 17].Text as string);
                    double dSUMRnsCncd = SmAv.Sum;
                    BothInngs.Runs = SmAv.Ave;

                    // Wkts
                    SmAv = new SumAve();
                    SmAv = AverageTwoValues(xlWrkshtPlyrStts.Cells[iRow + 1, 16].Text as string, xlWrkshtPlyrStts.Cells[iRow + 2, 16].Text as string);
                    double dSUMwkts = SmAv.Sum;
                    BothInngs.Wkts = SmAv.Ave;

                    // Average
                    BothInngs.Ave = dSUMRnsCncd / dSUMwkts;

                    // Ovrs
                    SmAv = new SumAve();
                    SmAv = AverageTwoValues(xlWrkshtPlyrStts.Cells[iRow + 1, 18].Text as string, xlWrkshtPlyrStts.Cells[iRow + 2, 18].Text as string);
                    double dSUMOvrs = SmAv.Sum;
                    BothInngs.Ovrs = SmAv.Ave;


                    // SR
                    // Balls Bowled per Wickets taken
                    if (BothInngs.Wkts > 0)
                    {
                        BothInngs.SR = dSUMBllBwld / dSUMwkts;
                    }
                    else
                    {
                        BothInngs.SR = 0;
                    }


                    // RPO
                    BothInngs.RPO = BothInngs.Runs / BothInngs.Ovrs;

                    xlWrkshtPlyrStts.Cells[iRow + 4, 14] = BothInngs.Inngs.ToString();
                    xlWrkshtPlyrStts.Cells[iRow + 4, 15] = BothInngs.Ave.ToString();

                    xlWrkshtPlyrStts.Cells[iRow + 4, 16] = (BothInngs.Wkts / BothInngs.Inngs).ToString();

                    xlWrkshtPlyrStts.Cells[iRow + 4, 17] = (BothInngs.Runs / BothInngs.Inngs).ToString();

                    xlWrkshtPlyrStts.Cells[iRow + 4, 18] = (BothInngs.Ovrs / BothInngs.Inngs).ToString();


                    xlWrkshtPlyrStts.Cells[iRow + 4, 19] = BothInngs.RPO.ToString();
                    xlWrkshtPlyrStts.Cells[iRow + 4, 20] = BothInngs.SR.ToString();
                    xlWrkshtPlyrStts.Cells[iRow + 4, 21] = BothInngs.BllBwl.ToString();

                }
                else
                {
                    // Both is ThisAndLast
                    //xlWrkshtPlyrStts.Cells[iRow + 4, 3] = "Both";
                    for (int k = 14; k < 22; k++)
                    {
                        xlWrkshtPlyrStts.Cells[iRow + 4, k] = xlWrkshtPlyrStts.Cells[iRow + 1, k];
                    }
                }
            }
            //else
            //{

            //}
        }
    }


    public class ExcelLoadMatches : ExcelWorkbook
    {
        //private List<TeamTotal> teamsOneYrTotal = new List<TeamTotal> ();
        //private List<TeamTotal> teamsTwoYrTotal = new List<TeamTotal>();

        private void PopulateTeamCode(string sTeam, int iFrstRow, string sWkshtName)
        {
            string sRow = "";

            xlWorksheet = (Excel.Worksheet)xlWorkbook.Worksheets[sWkshtName];

            iFrstRow = iFrstRow - 1;

            for (int i = 1; i <= 5; i++)
            {
                sRow = "0" + i.ToString().Trim();
                xlWorksheet.Cells[iFrstRow + i, 1] = sTeam + "." + sRow;
            }
        }

        private void PopulateHeadings(int iFirstRow)
        {
            xlWrkshtMtchs.Cells[iFirstRow, 3] = "WcktsDwn";
            xlWrkshtMtchs.Cells[iFirstRow, 4] = "Scr";
            xlWrkshtMtchs.Cells[iFirstRow, 5] = "OvrsFcd";
            xlWrkshtMtchs.Cells[iFirstRow, 6] = "Frs";
            xlWrkshtMtchs.Cells[iFirstRow, 7] = "Sxs";
            xlWrkshtMtchs.Cells[iFirstRow, 8] = "OP";
            xlWrkshtMtchs.Cells[iFirstRow, 9] = "PwrPly";
            xlWrkshtMtchs.Cells[iFirstRow, 10] = "Wns";
            xlWrkshtMtchs.Cells[iFirstRow, 11] = "GmsPlyd";
            xlWrkshtMtchs.Cells[iFirstRow, 12] = "W%";
            xlWrkshtMtchs.Cells[iFirstRow, 13] = "L%";
            xlWrkshtMtchs.Cells[iFirstRow, 14] = "WinsBy15%";
            xlWrkshtMtchs.Cells[iFirstRow, 15] = "WinsBy25%";
        }
        private void InitMatches()
        {
            xlWorksheet = (Excel.Worksheet)xlWorkbook.Worksheets[gbl.SheetType.sDRAWCURRENT];
            int iLastCol = GetLastCol(1, gbl.SheetType.sDRAWCURRENT);

            int x = 4;

            for (int i = 2; i <= iLastCol; i = i + 3)
            {
                xlWorksheet = (Excel.Worksheet)xlWorkbook.Worksheets[gbl.SheetType.sDRAWCURRENT];
                string sCurrTeam = xlWorksheet.Cells[1, i].Text as string;
                xlWrkshtMtchs.Cells[x, 2] = sCurrTeam;
                xlWrkshtMtchs.Cells[x + 1, 2] = "Last Year";
                xlWrkshtMtchs.Cells[x + 3, 2] = "Last 2 Years";
                PopulateHeadings(x);
                PopulateTeamCode(sCurrTeam, x, gbl.SheetType.sMATCHES);
                x = x + 5;
            }
        }

        private string GetLastSeason(string sCOMPLETEDSEASON)
        {
            string[] aAll;
            string sLastSsn = "";
            string sLastComp = "";
            int iLastSsnPre = -1;
            int iLastSsnPost = -1;
            string sLastSsnPre = "";
            string sLastSsnPost = "";

            aAll = sCOMPLETEDSEASON.Split('=');
            sLastComp = aAll[0];
            sLastSsn = aAll[1];
            aAll = sLastSsn.Split('-');
            sLastSsnPre = aAll[0];
            sLastSsnPost = aAll[1];
            iLastSsnPre = Convert.ToInt32(sLastSsnPre);
            iLastSsnPost = Convert.ToInt32(sLastSsnPost);
            iLastSsnPre = iLastSsnPre - 1;
            iLastSsnPost = iLastSsnPost - 1;

            return (sLastComp + "=" + iLastSsnPre.ToString() + "-" + iLastSsnPost.ToString());
        }

        public void LoadMatches(List<Match> matchesInDB, CricketForm fm)
        {
            InitExcel();

            xlWorksheet = (Excel.Worksheet)xlWorkbook.Worksheets[gbl.SheetType.sMATCHES];
            string sAlreadyInitilized = xlWorksheet.Cells[4, 1].Text as string;
            //if (sAlreadyInitilized == "")
            //{
            InitMatches();
            //}

            LoadMatchesForSeason(matchesInDB, fm);

            xlApp.DisplayAlerts = false;
            xlWorkbook.SaveAs(gbl.sEXCELPATH, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing,
                 Type.Missing, Type.Missing);
            CleanUpExcel();
        }



        public void LoadMatchesForSeason(List<Match> mtchs, CricketForm fm)
        {
            List<TeamMatches> teamsOneYrTotal = new List<TeamMatches>();
            List<TeamMatches> teamsTwoYrsTotal = new List<TeamMatches>();

            if (mtchs.Count == 0)
            {
                return;
            }
            xlWrkshtMtchs = (Excel.Worksheet)xlWorkbook.Worksheets[gbl.SheetType.sMATCHES];
            string sCurrTeamCodeHome = mtchs[0].TeamHome;
            string sCurrTeamCodeAway = mtchs[0].TeamAway;

            xlWorksheet = (Excel.Worksheet)xlWorkbook.Worksheets[gbl.SheetType.sTEAM];
            string sCOMPLETEDSEASON = xlWorksheet.Cells[1, 9].Value;
            string sLASTSEASON = GetLastSeason(sCOMPLETEDSEASON);

            teamsOneYrTotal = GetTeamListTotal();
            teamsTwoYrsTotal = GetTeamListTotal();


            for (int i = 0; i < mtchs.Count; i++)
            {
                string sCurrPlyrComp = mtchs[i].Cmpttn.CompetitionCode + "=" + mtchs[i].Cmpttn.Season;
                string sCurrTeamHome = mtchs[i].TeamHome;
                string sCurrTeamAway = mtchs[i].TeamAway;
                if (sCurrPlyrComp == sCOMPLETEDSEASON)
                {
                    teamsOneYrTotal = PopulateOneMatch(sCurrTeamHome, sCurrTeamAway, mtchs[i], ref teamsOneYrTotal);
                    teamsTwoYrsTotal = PopulateOneMatch(sCurrTeamHome, sCurrTeamAway, mtchs[i], ref teamsTwoYrsTotal);
                }
                else if (sCurrPlyrComp == sLASTSEASON)
                {
                    teamsTwoYrsTotal = PopulateOneMatch(sCurrTeamHome, sCurrTeamAway, mtchs[i], ref teamsTwoYrsTotal);
                }
            }
            PopulateMatches(false, teamsOneYrTotal);
            PopulateMatches(true, teamsTwoYrsTotal);

        }

        private List<TeamMatches> GetTeamListTotal()
        {
            List<TeamMatches> tmtotal = new List<TeamMatches>();
            //TeamSeason OneTeam = new TeamSeason();

            xlWorksheet = (Excel.Worksheet)xlWorkbook.Worksheets[gbl.SheetType.sDRAWCURRENT];
            int iLastCol = GetLastCol(1, gbl.SheetType.sDRAWCURRENT);
            if (iLastCol == 0)
            {
                MessageBox.Show("Error in GetTeamListTotal");
                return null;
            }
            for (int i = 1; i <= iLastCol; i = i + 3)
            {
                TeamMatches OneTeam = new TeamMatches();
                OneTeam.TeamName = xlWorksheet.Cells[1, i + 1].Text as string;

                //TeamMatches TmMtch = new TeamMatches();
                OneTeam.WinCount = 0;
                OneTeam.WcktsDwn = 0;
                OneTeam.Scr = 0;
                OneTeam.OvrsFcd = 0;
                OneTeam.Frs = 0;
                OneTeam.Sxs = 0;
                OneTeam.OP = 0;
                OneTeam.PwrPly = 0;
                //OneTeam.Wns = 0;
                OneTeam.GmsPlyd = 0;

                //OneTeam.Totals = TmMtch;
                tmtotal.Add(OneTeam);

            }

            return tmtotal;
        }

        private List<TeamMatches> PopulateOneMatch(string sTeamHome, string sTeamAway, Match mtch, ref List<TeamMatches> teamsTotal)
        {
            //List<TeamMatches> TmTotal = new List<TeamMatches>();
            //int i = 0;

            //TmSsn = OneOrTwoYrs;
            bool bIsHomeTm = true;
            teamsTotal = UpdateTeamMatches(bIsHomeTm, sTeamHome, mtch, ref teamsTotal);
            bIsHomeTm = false;
            teamsTotal = UpdateTeamMatches(bIsHomeTm, sTeamAway, mtch, ref teamsTotal);


            return teamsTotal;
        }

        private List<TeamMatches> UpdateTeamMatches(bool bIsHomeTeam, string sTeam, Match mtch, ref List<TeamMatches> tmsTotal)
        {
            int i = 0;
            bool bNotFound = true;
            while (bNotFound)
            {
                if (tmsTotal[i].TeamName == sTeam)
                {
                    bNotFound = false;
                }
                else
                {
                    i++;
                }
            }
            if (tmsTotal[i] == null)
            {
                MessageBox.Show("Error in UpdateTeamMatches");
                return null;
            }

            TeamMatches tm = new TeamMatches();
            double dMyScore = -1;
            double dOppScore = -1;
            double dMyOversBowled = -1;
            double dOppOversBowled = -1;

            if (bIsHomeTeam)
            {
                tm.WcktsDwn = Convert.ToDouble(mtch.HomeTeamData.WicketsDown);
                tm.Scr = Convert.ToDouble(mtch.HomeTeamData.TotalScore);
                tm.OvrsFcd = Convert.ToDouble(mtch.HomeTeamData.OversFaced);
                tm.Frs = Convert.ToDouble(mtch.HomeTeamData.Four);
                tm.Sxs = Convert.ToDouble(mtch.HomeTeamData.Six);
                tm.OP = Convert.ToDouble(mtch.HomeTeamData.OpeningPartnership);
                tm.PwrPly = Convert.ToDouble(mtch.HomeTeamData.XOverScore);
                dMyScore = Convert.ToDouble(mtch.HomeTeamData.TotalScore);
                dOppScore = Convert.ToDouble(mtch.AwayTeamData.TotalScore);
                dMyOversBowled = Convert.ToDouble(mtch.HomeTeamData.OversFaced);
                dOppOversBowled = Convert.ToDouble(mtch.AwayTeamData.OversFaced);

            }
            else
            {
                tm.WcktsDwn = Convert.ToDouble(mtch.AwayTeamData.WicketsDown);
                tm.Scr = Convert.ToDouble(mtch.AwayTeamData.TotalScore);
                tm.OvrsFcd = Convert.ToDouble(mtch.AwayTeamData.OversFaced);
                tm.Frs = Convert.ToDouble(mtch.AwayTeamData.Four);
                tm.Sxs = Convert.ToDouble(mtch.AwayTeamData.Six);
                tm.OP = Convert.ToDouble(mtch.AwayTeamData.OpeningPartnership);
                tm.PwrPly = Convert.ToDouble(mtch.AwayTeamData.XOverScore);
                dMyScore = Convert.ToDouble(mtch.AwayTeamData.TotalScore);
                dOppScore = Convert.ToDouble(mtch.HomeTeamData.TotalScore);
                dMyOversBowled = Convert.ToDouble(mtch.AwayTeamData.OversFaced);
                dOppOversBowled = Convert.ToDouble(mtch.HomeTeamData.OversFaced);
            }

            tmsTotal[i].GmsPlyd = tmsTotal[i].GmsPlyd + 1;
            tmsTotal[i].WcktsDwn = tmsTotal[i].WcktsDwn + tm.WcktsDwn;
            tmsTotal[i].Scr = tmsTotal[i].Scr + tm.Scr;
            tmsTotal[i].OvrsFcd = tmsTotal[i].OvrsFcd + tm.OvrsFcd;
            tmsTotal[i].Frs = tmsTotal[i].Frs + tm.Frs;
            tmsTotal[i].Sxs = tmsTotal[i].Sxs + tm.Sxs;
            tmsTotal[i].OP = tmsTotal[i].OP + tm.OP;
            tmsTotal[i].PwrPly = tmsTotal[i].PwrPly + tm.PwrPly;
            if (mtch.TeamWinner == sTeam)
            {
                tmsTotal[i].WinCount = tmsTotal[i].WinCount + 1;
            }
            double dWinBy = -1;
            if (sTeam == mtch.BatFirst)
            {  // MyTeam Batted First
                dWinBy = dMyScore - dOppScore;
                if (dWinBy >= 30)
                {
                    tmsTotal[i].WinByFifteenPerc = tmsTotal[i].WinByFifteenPerc + 1;
                }
                else if (dWinBy >= 50)
                {
                    tmsTotal[i].WinByTwentyfivePerc = tmsTotal[i].WinByTwentyfivePerc + 1;
                }
            }
            else
            {  // MyTeam Bowled First
                dWinBy = dMyOversBowled - dOppOversBowled;
                if (dWinBy >= 3)
                {
                    tmsTotal[i].WinByFifteenPerc = tmsTotal[i].WinByFifteenPerc + 1;
                }
                else if (dWinBy >= 5)
                {
                    tmsTotal[i].WinByTwentyfivePerc = tmsTotal[i].WinByTwentyfivePerc + 1;
                }
            }

            return tmsTotal;
        }



        private void PopulateMatches(bool bIsTwoYears, List<TeamMatches> TeamOneYr)
        {
            int iRowNo = -1;
            if (bIsTwoYears)
            {
                iRowNo = 7;
            }
            else
            {
                iRowNo = 5;
            }
            if (TeamOneYr == null)
            {
                return;
            }
            if (TeamOneYr.Count == 0)
            {
                return;
            }
            xlWorksheet = (Excel.Worksheet)xlWorkbook.Worksheets[gbl.SheetType.sMATCHES];
            for (int i = 0; i < TeamOneYr.Count; i++)
            {
                if (TeamOneYr[i].OvrsFcd == 0)
                {
                    xlWorksheet.Cells[(i * 5) + iRowNo, 3] = "";
                    xlWorksheet.Cells[(i * 5) + iRowNo, 4] = "";
                    xlWorksheet.Cells[(i * 5) + iRowNo, 5] = "";
                    xlWorksheet.Cells[(i * 5) + iRowNo, 6] = "";
                    xlWorksheet.Cells[(i * 5) + iRowNo, 7] = "";
                    xlWorksheet.Cells[(i * 5) + iRowNo, 8] = "";
                    xlWorksheet.Cells[(i * 5) + iRowNo, 9] = "";
                    xlWorksheet.Cells[(i * 5) + iRowNo, 10] = "";
                    xlWorksheet.Cells[(i * 5) + iRowNo, 11] = "";
                    xlWorksheet.Cells[(i * 5) + iRowNo, 12] = "";
                    xlWorksheet.Cells[(i * 5) + iRowNo, 13] = "";
                    xlWorksheet.Cells[(i * 5) + iRowNo, 14] = "";
                    xlWorksheet.Cells[(i * 5) + iRowNo, 15] = "";
                }
                else
                {
                    xlWorksheet.Cells[(i * 5) + iRowNo, 3] = Math.Round((Double)TeamOneYr[i].WcktsDwn / TeamOneYr[i].GmsPlyd, 0).ToString().Trim();
                    xlWorksheet.Cells[(i * 5) + iRowNo, 4] = Math.Round((Double)TeamOneYr[i].Scr / TeamOneYr[i].GmsPlyd, 0).ToString().Trim();
                    xlWorksheet.Cells[(i * 5) + iRowNo, 5] = Math.Round((Double)TeamOneYr[i].OvrsFcd / TeamOneYr[i].GmsPlyd, 0).ToString().Trim();
                    xlWorksheet.Cells[(i * 5) + iRowNo, 6] = Math.Round((Double)TeamOneYr[i].Frs / TeamOneYr[i].GmsPlyd, 0).ToString().Trim();
                    xlWorksheet.Cells[(i * 5) + iRowNo, 7] = Math.Round((Double)TeamOneYr[i].Sxs / TeamOneYr[i].GmsPlyd, 0).ToString().Trim();
                    xlWorksheet.Cells[(i * 5) + iRowNo, 8] = Math.Round((Double)TeamOneYr[i].OP / TeamOneYr[i].GmsPlyd, 0).ToString().Trim();
                    xlWorksheet.Cells[(i * 5) + iRowNo, 9] = Math.Round((Double)TeamOneYr[i].PwrPly / TeamOneYr[i].GmsPlyd, 0).ToString().Trim();
                    xlWorksheet.Cells[(i * 5) + iRowNo, 10] = Math.Round((Double)TeamOneYr[i].WinCount / TeamOneYr[i].GmsPlyd, 0).ToString().Trim();
                    xlWorksheet.Cells[(i * 5) + iRowNo, 11] = TeamOneYr[i].GmsPlyd.ToString().Trim();

                    double dWinRatio = Math.Round((Double)(TeamOneYr[i].WinCount / TeamOneYr[i].GmsPlyd), 0);
                    double dLosses = Math.Round((Double)(TeamOneYr[i].GmsPlyd - TeamOneYr[i].WinCount), 0);
                    double dLossRatio = dLosses / TeamOneYr[i].GmsPlyd;
                    // W%
                    xlWorksheet.Cells[(i * 5) + iRowNo, 12] = Math.Round((Double)(dWinRatio * 100), 0).ToString().Trim();
                    // L%
                    xlWorksheet.Cells[(i * 5) + iRowNo, 13] = Math.Round((Double)(dLossRatio * 100), 0).ToString().Trim();
                    // WinBy15%
                    xlWorksheet.Cells[(i * 5) + iRowNo, 14] = Math.Round((Double)(TeamOneYr[i].WinByFifteenPerc), 0).ToString().Trim();
                    // WinBy25%
                    xlWorksheet.Cells[(i * 5) + iRowNo, 15] = Math.Round((Double)(TeamOneYr[i].WinByTwentyfivePerc), 0).ToString().Trim();
                }
            }
        }
    }

    public class ExcelMatch : ExcelWorkbook
    {
        public Competition Comp = new Competition();


        public string[] GetAllCompetitionData(string sCompetitionCode)
        {
            int i = 2;
            bool bNotFound = true;
            string sCurrentCompCode = "";
            string[] aData = new string[5];

            aData[0] = "Error";
            aData[1] = "Error";
            aData[2] = "Error";
            aData[3] = "Error";
            aData[4] = "Error";

            xlWorksheet = xlWorkbook.Sheets[gbl.SheetType.sCOMPETITION];
            xlRange = xlWorksheet.UsedRange;
            bNotFound = true;

            while (bNotFound)
            {
                sCurrentCompCode = xlRange.Cells[i, 1].Value.ToString().Trim();
                sCurrentCompCode = sCurrentCompCode + "=" + xlRange.Cells[i, 2].Value.ToString().Trim();
                if (sCurrentCompCode == sCompetitionCode)
                {
                    bNotFound = false;
                    aData[0] = xlRange.Cells[i, 1].Value.ToString().Trim();
                    aData[1] = xlRange.Cells[i, 2].Value.ToString().Trim();
                    aData[2] = xlRange.Cells[i, 3].Value.ToString().Trim();
                    aData[3] = xlRange.Cells[i, 4].Value.ToString().Trim();
                    aData[4] = xlRange.Cells[i, 5].Value.ToString().Trim();
                    xlWorksheet = xlWorkbook.Sheets["Teams"];
                    xlRange = xlWorksheet.UsedRange;
                    return aData;
                }
                i++;
            }
            MessageBox.Show("Error in GetCompetitionName: CompetitionCode NOT in Worksheet Competition");
            //xlWorksheet = xlWorkbook.Sheets["Teams"];
            //xlRange = xlWorksheet.UsedRange;
            return aData;
        }

        private string GetTeamName(int iRowNo)
        {
            string[] aSplitString;
            string sTeamCode = "";

            xlWorksheet = xlWorkbook.Sheets[gbl.SheetType.sTEAM];
            xlRange = xlWorksheet.UsedRange;

            sTeamCode = xlRange.Cells[iRowNo, 1]?.Value?.ToString();
            aSplitString = sTeamCode.Split('.');
            sTeamCode = aSplitString[1].ToString().Trim();

            return TeamShortToLong(sTeamCode);
        }

        private string GetTeamCode(int iRowNo)
        {
            string[] aSplitString;
            string sTeamCode = "";

            xlWorksheet = xlWorkbook.Sheets[gbl.SheetType.sTEAM];
            xlRange = xlWorksheet.UsedRange;

            sTeamCode = xlRange.Cells[iRowNo, 1]?.Value?.ToString();
            aSplitString = sTeamCode.Split('.');
            sTeamCode = aSplitString[1].ToString().Trim();

            return sTeamCode;

            //return TeamShortToLong(sTeamCode);
        }

        private int[] GetOtherTeam(int iGameNostring, string sOtherTeamCode)
        {
            int[] aData = new int[2];
            int iRowNo = 0;
            int iCol = 0;
            string[] aSplitString;
            string sTeamCode;
            string sCurrentCode = "";
            string sGameData = "";

            xlWorksheet = xlWorkbook.Sheets[gbl.SheetType.sTEAM];
            xlRange = xlWorksheet.UsedRange;
            iRowNo = 13;
            while (true)
            {
                sTeamCode = xlRange.Cells[iRowNo, 1]?.Value?.ToString().Trim();
                aSplitString = sTeamCode.Split('.');
                sCurrentCode = aSplitString[1];
                if (sCurrentCode == sOtherTeamCode)
                {
                    iCol = 3;
                    while (true)
                    {
                        sGameData = xlRange.Cells[iRowNo, iCol]?.Value?.ToString().Trim();
                        aSplitString = sGameData.Split(')');
                        sGameData = aSplitString[0].Trim();
                        ; if (sGameData == iGameNostring.ToString().Trim())
                        {
                            aData[0] = iRowNo;
                            aData[1] = iCol;
                            return aData;
                        }
                        iCol++;
                    }
                }
                iRowNo = iRowNo + 13;
            }
        }

        public Match GetNextMatch(int iGameNoThisSeason, CricketForm fm)
        {
            //string sCompCode = "";
            //string[] aCompData = new string[5];
            //int[] iFirstTeamFound = new int[2];
            //int[] iOtherTeamFound = new int[2];

            //xlWorksheet = xlWorkbook.Sheets[gbl.SheetType.sTEAM];
            //xlRange = xlWorksheet.UsedRange;
            Match plyrMatch = new Match();
            //Competition Cmp = new Competition();
            //sCompCode = xlRange.Cells[1, 9].Value;
            //if (sCompCode == null)
            //{
            //    MessageBox.Show("Error in GetNextMatch: No Competition/Season code in SHEET:Teams Cell[1,9]");
            //}
            //else
            //{
            //Cmp = Comp;
            plyrMatch.Cmpttn = Comp;
            plyrMatch = GetMatchData(iGameNoThisSeason, plyrMatch, fm);
            //}
            return plyrMatch;
        }

        private Match PopulateTeam(string sThsTm, string sHomeAway, int iMatchRow, int iMatchCol, Match mch)
        {
            string sGameData = "";
            string[] aSplitString1;
            string[] aSplitString2;
            string sFours = "";
            string sSixes = "";
            string sSixOverScore = "";
            string sOP = "";
            string sBatFirstOrSecond = "";
            string sRemainder = "";
            string sWktsDown = "";
            string sTotalScore = "";
            string sOversFaced = "";
            string sWinner = "";
            Match mh = new Match();

            xlWorksheet = xlWorkbook.Sheets[gbl.SheetType.sTEAM];
            xlRange = xlWorksheet.UsedRange;
            sGameData = xlRange.Cells[iMatchRow, iMatchCol]?.Value?.ToString().Trim();
            if (sGameData == null)
            {
                return null;
            }
            else
            {
                aSplitString1 = sGameData.Split('=');
                sFours = aSplitString1[0].Trim();
                sSixes = aSplitString1[1].Trim();
                sSixOverScore = aSplitString1[2].Trim();
                sOP = aSplitString1[3].Trim();
                sBatFirstOrSecond = aSplitString1[4].Trim();
                sRemainder = aSplitString1[5].Trim();
                aSplitString2 = sRemainder.Split('-');
                sWktsDown = aSplitString2[0].Trim();
                sTotalScore = aSplitString2[1].Trim();
                sOversFaced = aSplitString2[2].Trim();
                sWinner = aSplitString1[6].Trim();
                if (sBatFirstOrSecond == "f")
                {
                    mch.BatFirst = TeamLongToShort(sThsTm);
                }
                else if (sBatFirstOrSecond == "s")
                {
                    mch.BatSecond = TeamLongToShort(sThsTm);
                }
                if (sWinner == "W")
                {
                    mch.TeamWinner = TeamLongToShort(sThsTm);
                }
                else if (sWinner == "L")
                {
                    mch.TeamLoser = TeamLongToShort(sThsTm);
                }
                else if (sWinner == "N")
                {
                    mch.TeamWinner = gbl.sNORESULT;
                    mch.TeamLoser = gbl.sNORESULT;
                }
                else if (sWinner == "D")
                {
                    mch.TeamWinner = gbl.sDRAW;
                    mch.TeamLoser = gbl.sDRAW;
                }
                else if (sWinner == "T")
                {
                    mch.TeamWinner = gbl.sTIE;
                    mch.TeamLoser = gbl.sTIE;
                }
                ResultsMatch res = new ResultsMatch();
                res.Four = sFours;
                res.Six = sSixes;
                res.XOverScore = sSixOverScore;
                res.OpeningPartnership = sOP;
                res.WicketsDown = sWktsDown;
                res.TotalScore = sTotalScore;
                res.OversFaced = sOversFaced;
                if (sHomeAway == gbl.sHOMETEAM)
                {
                    //mch.TeamHome = sThsTm;
                    mch.HomeTeamData = res;
                }
                else if (sHomeAway == gbl.sAWAYTEAM)
                {
                    //mch.TeamAway = sThsTm;
                    mch.AwayTeamData = res;
                }
                else
                {
                    //MessageBox.Show("Error in PopulateTeam: Team should be BatFirst or BatSecond");
                }
                return mch;
            }
        }

        private Match GetMatchData(int iGameNoThisSeason, Match mtch, CricketForm fm)
        {
            // Get Data for 1 match
            int iRowCount = 0;
            int iTeamRow = 0;
            int iLastCol = 0;
            string sCurrentGame = "";
            int j = 0;
            string[] aSplitString1;
            string[] aSplitString2;
            string sGameDate = "";
            string sCurrentGameNo = "";
            //string sGameNo = "";
            string sHomeAwayGround = "";
            string sHomeAway = "";
            string sGround = "";
            string sThisTeam = "";
            int iMatchRow = 0;
            //string sRemainder1 = "";
            //string sRemainder2 = "";
            string sOtherTeam = "";
            int[] aRowCol = new int[2];
            string sFullOtherTeam = "";
            bool bNotFound = true;
            string sThisTeamCode = "";

            xlWorksheet = xlWorkbook.Sheets[gbl.SheetType.sTEAM];
            xlRange = xlWorksheet.UsedRange;

            iRowCount = GetLastRow(1, gbl.SheetType.sTEAM);
            int iNoTeams = (iRowCount - 2) / 13;
            int i = 0;
            while ((i < iNoTeams) && (bNotFound))
            {
                iTeamRow = (i * 13) + 13;
                iLastCol = GetLastCol(iTeamRow, gbl.SheetType.sTEAM);
                j = 3;
                while ((j <= iLastCol) && (bNotFound))
                {
                    sCurrentGame = xlRange.Cells[iTeamRow, j]?.Value?.ToString();
                    aSplitString1 = sCurrentGame.Split(')');
                    sCurrentGameNo = aSplitString1[0].ToString().Trim();
                    sOtherTeam = aSplitString1[1].Trim();
                    string sGameNoThisSeason = iGameNoThisSeason.ToString().Trim();
                    if (sCurrentGameNo == sGameNoThisSeason)
                    {
                        mtch.GameNumber = Convert.ToInt32(sCurrentGameNo);
                        sThisTeam = GetTeamName(iTeamRow);
                        sThisTeamCode = GetTeamCode(iTeamRow);
                        xlWorksheet = xlWorkbook.Sheets["Teams"];
                        xlRange = xlWorksheet.UsedRange;
                        sHomeAwayGround = xlRange.Cells[iTeamRow + 1, j].Value.ToString().Trim();
                        aSplitString2 = sHomeAwayGround.Split('-');
                        sGround = aSplitString2[0].ToString().Trim();
                        mtch.GroundName = sGround;
                        sGameDate = aSplitString2[1].Trim();
                        mtch.MatchDate = StringToDate(sGameDate);
                        sHomeAway = aSplitString2[2].ToString().Trim();
                        iMatchRow = iTeamRow - 10;
                        mtch = PopulateTeam(sThisTeam, sHomeAway, iMatchRow, j, mtch);
                        if (mtch == null)
                        {
                            return null;
                        }
                        if (sHomeAway == gbl.sHOMETEAM)
                        {
                            mtch.TeamHome = sThisTeamCode;
                            mtch.TeamAway = sOtherTeam;
                        }
                        else
                        {
                            mtch.TeamAway = sThisTeamCode;
                            mtch.TeamHome = sOtherTeam;
                        }
                        int iCurrGameNo = Convert.ToInt32(sCurrentGameNo);
                        aRowCol = GetOtherTeam(iCurrGameNo, sOtherTeam);
                        sFullOtherTeam = TeamShortToLong(sOtherTeam);
                        if (sHomeAway == gbl.sAWAYTEAM)
                        {
                            sHomeAway = gbl.sHOMETEAM;
                        }
                        else
                        {
                            sHomeAway = gbl.sAWAYTEAM;
                        }
                        mtch = PopulateTeam(sFullOtherTeam, sHomeAway, aRowCol[0] - 10, aRowCol[1], mtch);

                        return mtch;
                    }
                    j++;
                }
                i++;
            }
            Match mtchError = new Match();
            return mtchError;
        }

        public List<TeamInnings> GetTeamSeason(CricketForm fm)
        {
            List<TeamInnings> teaminngs = new List<TeamInnings>();

            string[] aAll;
            string[] aAll1;
            string[] aAll2;
            string[] aAll3;

            xlWorksheet = xlWorkbook.Sheets[gbl.SheetType.sTEAM];
            xlRange = xlWorksheet.UsedRange;
            int iLastRow = GetLastRow(1, gbl.SheetType.sTEAM);
            int iNoTeams = (iLastRow - 1) / 13;
            if (iNoTeams == 0)
            {
                return null;
            }
            for (int i = 0; i < iNoTeams; i++)
            {
                int iLastCol = GetLastCol(((i * 13) + 3), gbl.SheetType.sTEAM);
                for (int j = 3; j <= iLastCol; j++)
                {
                    //string sData = xlRange.Cells[(i * 13) + 3, j].Value.ToString().Trim();
                    //if (xlRange.Cells[(i * 13) + 3, j].Value != "")
                    //{
                    string sData = xlRange.Cells[(i * 13) + 3, j].Value.ToString().Trim();
                    if (sData != "")
                    {
                        TeamInnings onetmssn = new TeamInnings();

                        string sGameNoAndOpp = xlRange.Cells[(i * 13) + 13, j].Value.ToString().Trim();
                        string sGrndDateHomeOrAway = xlRange.Cells[(i * 13) + 14, j].Value.ToString().Trim();
                        onetmssn.Cmpttn = Comp;
                        string sTeamCode = xlRange.Cells[(i * 13) + 3, 1].Value.ToString().Trim();
                        aAll = sTeamCode.Split('.');
                        onetmssn.MyTeam = aAll[1];
                        aAll = sData.Split('=');


                        ResultsMatch ResMtch = new ResultsMatch();
                        ResMtch.Four = aAll[0];
                        ResMtch.Six = aAll[1];
                        ResMtch.XOverScore = aAll[2];
                        ResMtch.OpeningPartnership = aAll[3];

                        string sBatRes = aAll[5];
                        aAll1 = sBatRes.Split('-');

                        ResMtch.WicketsDown = aAll1[0];
                        ResMtch.TotalScore = aAll1[1];
                        ResMtch.OversFaced = aAll1[2];

                        aAll2 = sGameNoAndOpp.Split(')');
                        onetmssn.GameNumber = Convert.ToInt32(aAll2[0].Trim());
                        onetmssn.OppTeam = aAll2[1].Trim();

                        string sBatFrstOrScnd = aAll[4];
                        if (sBatFrstOrScnd == "f")
                        {
                            onetmssn.TeamBatFirst = onetmssn.MyTeam;
                            onetmssn.TeamBatSecond = onetmssn.OppTeam;
                        }
                        else
                        {
                            onetmssn.TeamBatFirst = onetmssn.OppTeam;
                            onetmssn.TeamBatSecond = onetmssn.MyTeam;
                        }

                        aAll3 = sGrndDateHomeOrAway.Split('-');
                        onetmssn.Ground = aAll3[0].Trim();
                        string sGameDate = aAll3[1].Trim();
                        DateTime dGameDate = StringToDate(sGameDate);
                        onetmssn.GameDate = dGameDate;

                        string sHomeOrAway = aAll3[2].Trim();
                        if (sHomeOrAway == "Home")
                        {
                            onetmssn.TeamHome = onetmssn.MyTeam;
                            onetmssn.TeamAway = onetmssn.OppTeam;
                        }
                        else
                        {
                            onetmssn.TeamHome = onetmssn.OppTeam;
                            onetmssn.TeamAway = onetmssn.MyTeam;
                        }

                        string sWinOrLose = aAll[6];
                        onetmssn.BattingWinOrLose = sWinOrLose;

                        onetmssn.ResBatting = ResMtch;
                        teaminngs.Add(onetmssn);

                    }
                }
            }
            return teaminngs;
        }


        private TeamGame GetBattingInnings(int iGameNo, List<TeamInnings> tminngs)
        {
            TeamGame onegame = new TeamGame();
            //TeamInnings oneings = new TeamInnings();
            if (tminngs == null)
            {
                return null;
            }
            if (tminngs.Count == 0)
            {
                return null;
            }
            for (int i = 0; i < tminngs.Count; i++)
            {
                TeamInnings oneings = new TeamInnings();
                oneings = tminngs[i];
                if (oneings.GameNumber == iGameNo)
                {
                    if (oneings.MyTeam == oneings.TeamHome)
                    {
                        onegame.HomeTeam = oneings;
                    }
                    else
                    {
                        onegame.AwayTeam = oneings;
                    }
                }
            }
            return onegame;
        }



        public List<Match> MergeTeamSeasons(List<TeamInnings> tminngs, CricketForm fm)
        {
            List<Match> mtchs = new List<Match>();

            if (tminngs == null)
            {
                return null;
            }
            if (tminngs.Count == 0)
            {
                return null;
            }
            int iNoGames = tminngs.Count / 2;
            for (int i = 0; i < iNoGames; i++)
            {
                int iGameNo = i + 1;
                TeamGame tmgm = new TeamGame();
                tmgm = GetBattingInnings(iGameNo, tminngs);

                Match onematch = new Match();

                if (tmgm.HomeTeam.BattingWinOrLose == "W")
                {
                    onematch.TeamWinner = tmgm.HomeTeam.MyTeam;
                    onematch.TeamLoser = tmgm.AwayTeam.MyTeam;
                }
                else if (tmgm.HomeTeam.BattingWinOrLose == "L")
                {
                    onematch.TeamWinner = tmgm.AwayTeam.MyTeam;
                    onematch.TeamLoser = tmgm.HomeTeam.MyTeam;
                }
                else
                {
                    onematch.TeamWinner = "N";
                    onematch.TeamLoser = "N";
                }
                onematch.Cmpttn = tmgm.HomeTeam.Cmpttn;
                onematch.GameNumber = iGameNo;
                onematch.MatchDate = tmgm.HomeTeam.GameDate;
                onematch.TeamHome = tmgm.HomeTeam.MyTeam;
                onematch.TeamAway = tmgm.AwayTeam.MyTeam;
                onematch.GroundName = tmgm.HomeTeam.Ground;
                onematch.BatFirst = tmgm.HomeTeam.TeamBatFirst;
                onematch.BatSecond = tmgm.HomeTeam.TeamBatSecond;
                onematch.HomeTeamData = tmgm.HomeTeam.ResBatting;
                onematch.AwayTeamData = tmgm.AwayTeam.ResBatting;
                mtchs.Add(onematch);
            }

            return mtchs;
        }
    }



    public class ExcePlayerListBatBowl : ExcelWorkbook
    {
        public List<Player> LoadPlayers(string sSheetName, CricketForm fm)
        {
            List<Player> plysts = new List<Player>();
            bool bAddPlayer = true;
            int iNoPlayers = -1;
            string[] aAll;


            xlWorksheet = (Excel.Worksheet)xlWorkbook.Worksheets[sSheetName];
            int iLastRow = GetLastRow(1, sSheetName);
            if (iLastRow < 9)
            {
                iNoPlayers = 0;
                return plysts;
            }
            else
            {
                iNoPlayers = (iLastRow - 3) / 6;
            }
            for (int i = 0; i < iNoPlayers; i++)
            {
                bAddPlayer = true;
                int iFirstRow = (i * 6) + 4;
                Player CurrPlyr = new Player();


                string sLastName = xlWorksheet.Cells[iFirstRow, 2].Text as string;

                if (sLastName.Contains("."))
                {
                    MessageBox.Show("Error in GetPlayerStatsList: There is a Fullstop in LastName at Cell[" + iFirstRow.ToString() + ", 2]");
                    bAddPlayer = false;
                }
                else
                {
                    CurrPlyr.LastName = sLastName.Trim();
                }
                string sFirstNames = xlWorksheet.Cells[iFirstRow + 1, 2].Text as string;
                if ((sFirstNames == "") || (sFirstNames.Contains("=") == false))
                {
                    MessageBox.Show("Error in GetPlayerStatsList: There is an Error in Firstnames at Cell[" + (iFirstRow + 1).ToString() + ", 2]");
                    bAddPlayer = false;
                }
                else
                {
                    aAll = sFirstNames.Split('=');
                    CurrPlyr.FirstName = aAll[0].Trim();
                    CurrPlyr.MiddleName = aAll[1].Trim();
                }

                string sDOB = xlWorksheet.Cells[iFirstRow + 3, 2].Text as string;
                if ((sDOB == "") || (sDOB.Contains("=") == false))
                {
                    MessageBox.Show("Error in GetPlayerStatsList: There is an Error in DOB at Cell[" + (iFirstRow + 3).ToString() + ", 2]");
                    bAddPlayer = false;
                }
                else
                {
                    CurrPlyr.DOB = StringToDate(sDOB);
                }
                string sCountry = xlWorksheet.Cells[iFirstRow + 4, 2].Text as string;

                if ((sCountry == "") || (sCountry == "Country"))
                {
                    MessageBox.Show("Error in GetPlayerStatsList: There is an Error in Country at Cell[" + (iFirstRow + 4).ToString() + ", 2]");
                    bAddPlayer = false;
                }
                else
                {
                    CurrPlyr.Country = sCountry.Trim();
                }

                if (bAddPlayer)
                {
                    plysts.Add(CurrPlyr);
                }
            }

            return plysts;
        }

        public List<Player> LoadPlayerList(CricketForm fm)
        {
            List<Player> plyrs = new List<Player>();
            int iNoPlayers = -1;

            //InitExcel();
            xlWrkshtPlyrStts = (Excel.Worksheet)xlWorkbook.Worksheets[gbl.SheetType.sPLAYERLIST];
            int iLastRow = GetLastRow(2, gbl.SheetType.sPLAYERLIST);
            if (iLastRow < 4)
            {
                return null;
            }
            else
            {
                iNoPlayers = (iLastRow - 2) / 6;
            }

            for (int i = 0; i < iNoPlayers; i++)
            {
                Player OnePlayer = new Player();
                string[] aAll;

                OnePlayer.LastName = xlWrkshtPlyrStts.Cells[(i * 6) + 4, 2].Value.ToString().Trim();
                string sFirstNames = xlWrkshtPlyrStts.Cells[(i * 6) + 5, 2].Value.ToString().Trim();
                aAll = sFirstNames.Split('=');
                OnePlayer.FirstName = aAll[0];
                OnePlayer.MiddleName = aAll[1];
                xlWrkshtPlyrStts.Cells[(i * 6) + 6, 2].Value = "";

                string sDOB = xlWrkshtPlyrStts.Cells[(i * 6) + 7, 2].Value.ToString().Trim();

                OnePlayer.DOB = StringToDate(sDOB);
                OnePlayer.Country = xlWrkshtPlyrStts.Cells[(i * 6) + 8, 2].Value.ToString().Trim();
                plyrs.Add(OnePlayer);
            }
            //CleanUpExcel();
            return plyrs;
        }

        private void AddPlayerToPlayerList(Player OnePlayer)
        {
            xlWrkshtPlyrStts = (Excel.Worksheet)xlWorkbook.Worksheets[gbl.SheetType.sPLAYERLIST];
            int iLastRow = GetLastRow(2, gbl.SheetType.sPLAYERLIST);
            int iFirstRow = -1;
            if (iLastRow < 4)
            {
                iFirstRow = 4;

            }
            else
            {
                iFirstRow = iLastRow + 2;
            }
            xlWrkshtPlyrStts.Cells[iFirstRow + 0, 2] = OnePlayer.LastName.Trim();
            xlWrkshtPlyrStts.Cells[iFirstRow + 1, 2] = OnePlayer.FirstName + "=" + OnePlayer.MiddleName.Trim();
            xlWrkshtPlyrStts.Cells[iFirstRow + 3, 2] = DateToString(OnePlayer.DOB).Trim();
            xlWrkshtPlyrStts.Cells[iFirstRow + 4, 2] = OnePlayer.Country.Trim();
        }

        public void AddToPlayerListIfNotThere(List<Player> plyrs, List<Player> plyrsinplyrlist)
        {

            if (plyrs == null)
            {
                return;
            }
            bool bPlayerIsInPlayerList = false;
            for (int i = 0; i < plyrs.Count; i++)
            {
                bPlayerIsInPlayerList = false;
                for (int j = 0; j < plyrsinplyrlist.Count; j++)
                {
                    if (IsSamePlayer(plyrs[i], plyrsinplyrlist[j]))
                    {
                        bPlayerIsInPlayerList = true;
                    }
                }
                if (!bPlayerIsInPlayerList)
                {
                    AddPlayerToPlayerList(plyrs[i]);
                }

            }
            xlApp.DisplayAlerts = false;
            xlWorkbook.SaveAs(gbl.sEXCELPATH, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing,
                 Type.Missing, Type.Missing);
        }
    }

    public class ExceWebScraping : ExcelWorkbook
    {
        public DOBchanged UpdatePlayerDOB(string sOldDOB)
        {
            DOBchanged dobc = new DOBchanged();
            string[] aAll;

            bool bDOBchanged = false;
            aAll = sOldDOB.Split('=');
            string sDay = aAll[0];
            string sMonth = aAll[1];
            string sYear = aAll[2];

            if (sDay.Length == 1)
            {
                sDay = "0" + sDay;
                bDOBchanged = true;
            }
            if (sMonth.Length == 1)
            {
                sMonth = "0" + sMonth;
                bDOBchanged = true;
            }
            if (sYear.Length == 2)
            {
                int iYear = Convert.ToInt32(sYear);
                if (iYear < 60)
                {
                    sYear = "20" + sYear;
                }
                else
                {
                    sYear = "19" + sYear;
                }
                bDOBchanged = true;
            }
            dobc.bDOBchanged = bDOBchanged;
            dobc.NewDOB = sDay.Trim() + "=" + sMonth.Trim() + "=" + sYear.Trim();
            return dobc;
        }

        private void UpdatePlayerDOBForSheet(string sSheetName)
        {
            DOBchanged dc = new DOBchanged();
            int iNoPlayers = -1;

            xlWorksheet = (Excel.Worksheet)xlWorkbook.Worksheets[sSheetName];
            int iLastRow = GetLastRowFast(2, xlWorksheet);
            if (iLastRow < 9)
            {
                iNoPlayers = 0;
            }
            else
            {
                iNoPlayers = (iLastRow - 2) / 6;
            }
            for (int i = 0; i < iNoPlayers; i++)
            {
                string sOldDOB = xlWorksheet.Cells[(i * 6) + 7, 2]?.Value?.ToString();
                if (sOldDOB == null)
                {
                    MessageBox.Show("Error in UpdatePlayerDOBForSheet: Date at [" + ((i * 6) + 7).ToString() + ", 2] in " + sSheetName + " is an empty string");
                    return;
                }
                else
                {
                    dc = UpdatePlayerDOB(sOldDOB);
                    if (dc.bDOBchanged)
                    {
                        xlWorksheet.Cells[(i * 6) + 7, 2].Value = dc.NewDOB;
                    }
                }
            }
        }

        public void UpdateAllPlayersDOB()
        {
            DOBchanged dc = new DOBchanged();
            InitExcel();

            UpdatePlayerDOBForSheet(gbl.SheetType.sPLAYERLIST);
            UpdatePlayerDOBForSheet(gbl.SheetType.sBATSMEN);
            UpdatePlayerDOBForSheet(gbl.SheetType.sBOWLER);

            xlApp.DisplayAlerts = false;
            xlWorkbook.SaveAs(gbl.sEXCELPATH, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing,
                 Type.Missing, Type.Missing);
            CleanUpExcel();
        }

        public void UpdateBowler(int iMatchNo, Player plyr, ResultsBowler rbwl, CricketForm fm)
        {
            bool bPlayerFound = false;
            string[] aAll;
            Player CurrPlyr = new Player();
            int iDataRow = -1;
            int iLastRow = -1;
            int iNoPlayers = -1;

            InitExcel();
            xlWorksheet = (Excel.Worksheet)xlWorkbook.Worksheets[gbl.SheetType.sBOWLER];
            iLastRow = GetLastRowFast(1, xlWorksheet);
            if (iLastRow == 0)
            {
                iNoPlayers = 0;
            }
            else
            {
                iNoPlayers = ((iLastRow - 3) / 6) + 1;
            }
            int i = 0;

            bPlayerFound = false;

            while (!bPlayerFound && i < iNoPlayers)
            {
                string sDB = xlWorksheet.Cells[(i * 6) + 7, 2].Text as string;

                if (sDB != "")
                {
                    CurrPlyr.LastName = xlWorksheet.Cells[(i * 6) + 4, 2].Value.ToString().Trim();
                    sDB = xlWorksheet.Cells[(i * 6) + 7, 2].Value().Trim();
                    CurrPlyr.DOB = StringToDate(sDB);
                    string sFirstNames = xlWorksheet.Cells[(i * 6) + 5, 2].Value.ToString().Trim();
                    aAll = sFirstNames.Split('=');
                    CurrPlyr.FirstName = aAll[0].Trim();
                    if (IsSamePlayer(CurrPlyr, plyr))
                    {
                        fm.AddData("Bowler " + plyr.FirstName + " " + plyr.LastName + " HAS been found in BowlersSheet - " + "Adding Game Data now\r\n\r\n");
                        iDataRow = (i * 6) + 4;
                        EnterBowlerData(iMatchNo, iDataRow, plyr, rbwl);
                        bPlayerFound = true;
                    }
                }
                i++;
            }
            if (bPlayerFound == false)
            {
                fm.AddData("Bowler " + plyr.FirstName + " " + plyr.LastName + " HAS NOT been found in BowlerSheet - Creating Player and adding Game Data now" + "\r\n\r\n");

                iLastRow = GetLastRow(1, gbl.SheetType.sBOWLER);
                if (iLastRow == 0)
                {
                    iDataRow = 4;
                }
                else
                {
                    iDataRow = iLastRow + 1;
                }

                bool bIsBatsman = false;

                AddNewPlayer(bIsBatsman, iDataRow, plyr);
                EnterBowlerData(iMatchNo, iDataRow, plyr, rbwl);
                xlApp.Run("ByOrder", "Bowler");
            }

            xlApp.DisplayAlerts = false;
            xlWorkbook.SaveAs(gbl.sEXCELPATH, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing,
                 Type.Missing, Type.Missing);
            CleanUpExcel();
        }

        private void AddNewPlayer(bool bIsBatter, int iDataRow, Player plyr)
        {
            //int iNewBatsman = -1;

            if (plyr.LastName == null)
            {
                MessageBox.Show("plyr is null");
            }

            if (bIsBatter)
            {
                xlWorksheet = (Excel.Worksheet)xlWorkbook.Worksheets[gbl.SheetType.sBATSMEN];
                xlApp.Run("CreateBatter", plyr.TeamCode);
            }
            else
            {
                xlWorksheet = (Excel.Worksheet)xlWorkbook.Worksheets[gbl.SheetType.sBOWLER];
                xlApp.Run("CreateBwlr", plyr.TeamCode);
            }

            xlWorksheet.Cells[iDataRow, 2].Value = plyr.LastName;
            xlWorksheet.Cells[iDataRow + 1, 2].Value = plyr.FirstName + "=" + plyr.MiddleName;
            xlWorksheet.Cells[iDataRow + 2, 2].Value = plyr.TeamCode;
            xlWorksheet.Cells[iDataRow + 3, 2].Value = DateToString(plyr.DOB);
            xlWorksheet.Cells[iDataRow + 4, 2].Value = plyr.Country;
            //int x = 10;
        }

        public void UpdateMatch(int iHighestMatch, int iMatchFours, int iMatchSixes, ResultsMatch rm, string sTeam, CricketForm fm)
        {
            string[] aAll;

            if (rm == null)
            {
                return;
            }
            InitExcel();
            xlWorksheet = (Excel.Worksheet)xlWorkbook.Worksheets[gbl.SheetType.sTEAM];
            int iLastRow = GetLastRowFast(1, xlWorksheet);
            int iNoTeams = (iLastRow - 2) / 13;


            // change to while loop
            for (int i = 0; i < iNoTeams; i++)
            {
                string sTeamString = xlWorksheet.Cells[(i * 13) + 3, 1].Value.ToString().Trim();
                aAll = sTeamString.Split('.');
                string sCurrTeamCode = aAll[1].ToString().Trim();
                if (sCurrTeamCode == sTeam)
                {
                    int iOppRow = (i * 13) + 13;
                    int iLastCol = GetLastCol(iOppRow, gbl.SheetType.sTEAM);
                    //int j = 3;
                    for (int j = 3; j <= iLastCol; j++)
                    //while ((j <= iLastCol) && (bFound == false))
                    {
                        string sGameNoOpp = xlWorksheet.Cells[iOppRow, j].Value.ToString().Trim();
                        aAll = sGameNoOpp.Split(')');
                        string sMatchNo = aAll[0].ToString().Trim();
                        int iMatchNo = Convert.ToInt32(sMatchNo);
                        if (iMatchNo == iHighestMatch)
                        {
                            // Match Found - Now enter data
                            string sMatchData = iMatchFours.ToString().Trim() + "=" + iMatchSixes.ToString().Trim() + "=" + rm.XOverScore + "=" + rm.OpeningPartnership + "="
                                                + rm.BattedFirstOrSecond + "=" + rm.WicketsDown + "-" + rm.TotalScore + "-" + rm.OversFaced;
                            int iDataRow = iOppRow - 10;
                            xlWorksheet.Cells[iDataRow, j].Value = sMatchData;
                            //bFound = true;
                            xlApp.Run("TeamFormat");
                            xlApp.DisplayAlerts = false;
                            xlWorkbook.SaveAs(gbl.sEXCELPATH, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing,
                                    Type.Missing, Type.Missing);
                            CleanUpExcel();
                            return;
                        }
                        //j++;
                    }

                }
            }

            MessageBox.Show("Error in UpdateMatch: Cannot enter data because match number not found");
            CleanUpExcel();
        }

        public void UpdateMatchAbandoned(int iInningsNumber, int iHighestMatch, string sTeam, CricketForm fm)
        {
            string[] aAll;

            //if (rm == null)
            //{
            //    return;
            //}
            InitExcel();
            xlWorksheet = (Excel.Worksheet)xlWorkbook.Worksheets[gbl.SheetType.sTEAM];
            int iLastRow = GetLastRowFast(1, xlWorksheet);
            int iNoTeams = (iLastRow - 2) / 13;


            // change to while loop
            for (int i = 0; i < iNoTeams; i++)
            {
                string sTeamString = xlWorksheet.Cells[(i * 13) + 3, 1].Value.ToString().Trim();
                aAll = sTeamString.Split('.');
                string sCurrTeamCode = aAll[1].ToString().Trim();
                if (sCurrTeamCode == sTeam)
                {
                    int iOppRow = (i * 13) + 13;
                    int iLastCol = GetLastCol(iOppRow, gbl.SheetType.sTEAM);
                    //int j = 3;
                    for (int j = 3; j <= iLastCol; j++)
                    //while ((j <= iLastCol) && (bFound == false))
                    {
                        string sGameNoOpp = xlWorksheet.Cells[iOppRow, j].Value.ToString().Trim();
                        aAll = sGameNoOpp.Split(')');
                        string sMatchNo = aAll[0].ToString().Trim();
                        int iMatchNo = Convert.ToInt32(sMatchNo);
                        if (iMatchNo == iHighestMatch)
                        {
                            // Match Found - Now enter data
                            //string sMatchData = iMatchFours.ToString().Trim() + "=" + iMatchSixes.ToString().Trim() + "=" + rm.XOverScore + "=" + rm.OpeningPartnership + "="
                            //                    + rm.BattedFirstOrSecond + "=" + rm.WicketsDown + "-" + rm.TotalScore + "-" + rm.OversFaced;
                            string sMatchData;
                            if (iInningsNumber == 0)
                            {
                                sMatchData = "0=0=0=0=f=0-0-0=N";
                            }
                            else
                            {
                                sMatchData = "0=0=0=0=s=0-0-0=N";
                            }
                            int iDataRow = iOppRow - 10;
                            xlWorksheet.Cells[iDataRow, j].Value = sMatchData;
                            //bFound = true;
                            xlApp.DisplayAlerts = false;
                            xlWorkbook.SaveAs(gbl.sEXCELPATH, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing,
                                    Type.Missing, Type.Missing);
                            CleanUpExcel();
                            return;
                        }
                        //j++;
                    }

                }
            }

            MessageBox.Show("Error in UpdateMatch: Cannot enter data because match number not found");
            CleanUpExcel();
        }

        public List<Player> LdPlyrLst()
        {
            List<Player> plyrs = new List<Player>();
            int iNoPlayers = -1;

            InitExcel();
            xlWorksheet = (Excel.Worksheet)xlWorkbook.Worksheets[gbl.SheetType.sPLAYERLIST];
            int iLastRow = GetLastRow(2, gbl.SheetType.sPLAYERLIST);
            if (iLastRow < 4)
            {
                iNoPlayers = 0;
                CleanUpExcel();
                return null;
            }
            else
            {
                iNoPlayers = (iLastRow + 1) / 6;
            }

            for (int i = 0; i < iNoPlayers; i++)
            {
                Player OnePlayer = new Player();
                string[] aAll;

                OnePlayer.LastName = xlWorksheet.Cells[(i * 6) + 4, 2].Value.ToString().Trim();
                string sFirstNames = xlWorksheet.Cells[(i * 6) + 5, 2].Value.ToString().Trim();
                aAll = sFirstNames.Split('=');
                OnePlayer.FirstName = aAll[0];
                OnePlayer.MiddleName = aAll[1];
                //xlWorksheet.Cells[(i * 6) + 6, 2].Value = "";

                string sDOB = xlWorksheet.Cells[(i * 6) + 7, 2].Value.ToString().Trim();

                OnePlayer.DOB = StringToDate(sDOB);
                OnePlayer.Country = xlWorksheet.Cells[(i * 6) + 8, 2].Value.ToString().Trim();
                plyrs.Add(OnePlayer);
            }

            CleanUpExcel();
            return plyrs;
        }

        public void UpdatePlyrLst(List<Player> NewPlayers, List<Player> ExistingPlayers)
        {
            int iExistingPlayers = -1;
            bool bThisPlayerIsAlreadyInPlayerList = false;
            //ExcelWorkbook ew = new ExcelWorkbook();

            if (NewPlayers == null)
            {
                return;
            }
            if (ExistingPlayers == null)
            {
                iExistingPlayers = 0;
            }
            else
            {
                iExistingPlayers = ExistingPlayers.Count;
            }

            InitExcel();
            xlWorksheet = (Excel.Worksheet)xlWorkbook.Worksheets[gbl.SheetType.sPLAYERLIST];
            int iLastRow = GetLastRowFast(2, xlWorksheet);
            if (iLastRow == 0)
            {
                iLastRow = 2;
            }
            for (int i = 0; i < NewPlayers.Count; i++)
            {
                bThisPlayerIsAlreadyInPlayerList = false;
                int j = 0;
                while (j < iExistingPlayers)
                {
                    if (IsSamePlayer(NewPlayers[i], ExistingPlayers[j]))
                    {
                        bThisPlayerIsAlreadyInPlayerList = true;
                    }
                    j++;
                }
                if (bThisPlayerIsAlreadyInPlayerList == false)
                {
                    int iFirstRow = iLastRow + 2;
                    xlWorksheet.Cells[iFirstRow, 2].Value = NewPlayers[i].LastName.ToString();
                    xlWorksheet.Cells[iFirstRow, 2].Font.Underline = true;
                    xlWorksheet.Cells[iFirstRow, 2].Font.Bold = true;

                    xlWorksheet.Cells[iFirstRow + 1, 2].Value = NewPlayers[i].FirstName + "=" + NewPlayers[i].MiddleName.ToString();
                    xlWorksheet.Cells[iFirstRow + 1, 2].Font.Bold = true;
                    xlWorksheet.Cells[iFirstRow + 3, 2].Value = DateToString(NewPlayers[i].DOB).ToString();
                    xlWorksheet.Cells[iFirstRow + 3, 2].Font.Bold = true;


                    xlWorksheet.Cells[iFirstRow + 4, 2].Value = NewPlayers[i].Country.ToString();
                    //xlWorksheet.Cells[iFirstRow + 4, 2].Value = "Country";

                    xlWorksheet.Cells[iFirstRow + 4, 2].Font.Bold = true;

                    xlWorksheet.Columns[2].HorizontalAlignment = XlHAlign.xlHAlignCenter;

                    iLastRow = iLastRow + 6;
                    iExistingPlayers = iExistingPlayers++;
                }
            }
            xlApp.DisplayAlerts = false;
            xlWorkbook.SaveAs(gbl.sEXCELPATH, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing,
                    Type.Missing, Type.Missing);
            CleanUpExcel();
        }

        private void EnterBowlerData(int iMatchNo, int iDataRowNo, Player Plyr, ResultsBowler rbwlr)
        {
            string[] aAll;
            int iOppRw = -1;

            xlWorksheet = (Excel.Worksheet)xlWorkbook.Worksheets[gbl.SheetType.sBOWLER];
            iOppRw = iDataRowNo + 4;
            int iLastCol = GetLastCol(iOppRw, gbl.SheetType.sBOWLER);
            for (int i = 3; i <= iLastCol; i++)
            {
                string sMatchNoOpp = xlWorksheet.Cells[iOppRw, i].Value.ToString().Trim();
                aAll = sMatchNoOpp.Split(')');
                string sCurrMatch = aAll[0].Trim();
                int iCurrMatch = Convert.ToInt32(sCurrMatch);
                if (iCurrMatch == iMatchNo)
                {
                    if (rbwlr.OversBowled == null)
                    {
                        return;
                    }
                    else
                    {
                        xlWorksheet.Cells[iDataRowNo, i].Value = rbwlr.Wickets.Trim() + "=" + rbwlr.RunsConceded.Trim() + "=" + rbwlr.OversBowled.Trim();
                        return;
                    }
                }
            }
            return;
        }

        public void CalcNext()
        {
            InitExcel();
            xlApp.Run("TeamFormat");
            xlApp.Run("CalcNext");
            xlApp.Run("Format");
            xlApp.Run("FormatBowl");

            xlApp.DisplayAlerts = false;
            xlWorkbook.SaveAs(gbl.sEXCELPATH, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing,
                              Type.Missing, Type.Missing);

            CleanUpExcel();
        }

        public void FormatTeam()
        {
            InitExcel();
            xlApp.Run("TeamFormat");
            //xlApp.Run("CalcNext");
            //xlApp.Run("Format");
            //xlApp.Run("FormatBowl");

            xlApp.DisplayAlerts = false;
            xlWorkbook.SaveAs(gbl.sEXCELPATH, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing,
                              Type.Missing, Type.Missing);

            CleanUpExcel();
        }

        public void CreateBackups(int iMatchNo)
        {
            InitExcel();

            xlApp.DisplayAlerts = false;
            xlWorkbook.SaveAs(gbl.sEXCELPATH, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing,
                              Type.Missing, Type.Missing);

            string sNewBUDir = gbl.sBUPATH + "\\" + iMatchNo.ToString();
            if (!Directory.Exists(sNewBUDir))
            {
                Directory.CreateDirectory(sNewBUDir);
            }
            string sNewBUFile = sNewBUDir + "\\Backup.xlsm";
            xlApp.DisplayAlerts = false;
            xlWorkbook.SaveAs(sNewBUFile, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing,
                              Type.Missing, Type.Missing);
            xlApp.DisplayAlerts = false;
            xlWorkbook.SaveAs(gbl.sMAINBU, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing,
                              Type.Missing, Type.Missing);
            CleanUpExcel();
        }

        public int GetMatch(int iHighestSoFar, string sHomeTeam, string sAwayTeam)
        {

            InitExcel();
            ExceWebScrapingDraw ewsd = new ExceWebScrapingDraw();
            xlWorksheet = (Excel.Worksheet)xlWorkbook.Worksheets[gbl.SheetType.sDRAWCURRENT];
            int iLastCol = GetLastCol(1, gbl.SheetType.sDRAWCURRENT);

            for (int i = 1; i < iLastCol; i = i + 3)
            {
                string sCurrTeam = xlWorksheet.Cells[1, i].Value.ToString().Trim();
                if (sCurrTeam == sHomeTeam)
                {
                    string sOppAbbrev = ewsd.GetTeamNameAbbrev(sAwayTeam);
                    int iLastRow = GetLastRow(i, gbl.SheetType.sDRAWCURRENT);
                    for (int j = 2; j <= iLastRow; j++)
                    {
                        string sGameNo = xlWorksheet.Cells[j, i].Value.ToString().Trim();
                        int iGameNo = Convert.ToInt32(sGameNo);

                        if (iGameNo >= iHighestSoFar)
                        {
                            string sCurrOpp = xlWorksheet.Cells[j, i + 1].Value.ToString().Trim();
                            if (sCurrOpp == sOppAbbrev)
                            {
                                //string sFoundGame = sGameNo;
                                return Convert.ToInt32(sGameNo);
                            }
                        }
                    }
                }
            }

            CleanUpExcel();
            MessageBox.Show("Error in GetMatch: No Game found in SHEET:DrawCurrent");
            return -1;
        }

        public HighestGameAndLastGame GetLastMatchNo(CricketForm fm)
        {
            HighestGameAndLastGame hglg = new HighestGameAndLastGame();
            string[] aAll;
            bool bIsFirstGame = true;

            fm.AddData("\r\nFetching last game recorded\r\n");
            InitExcel();

            bool bIsReversed = xlApp.Run("IsReversed");

            if (bIsReversed)
            {
                xlApp.Run("Reverse");
                xlApp.DisplayAlerts = false;
                xlWorkbook.SaveAs(gbl.sEXCELPATH, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing,
                        Type.Missing, Type.Missing);
            }
            xlWorksheet = (Excel.Worksheet)xlWorkbook.Worksheets[gbl.SheetType.sTEAM];
            int iLastRow = GetLastRowFast(1, xlWorksheet);
            int iNoTeams = (iLastRow - 2) / 13;
            int iHighestLastGameSoFar = 0;
            int iLowestUnplayedGameSoFar = 1000;
            for (int i = 0; i < iNoTeams; i++)
            {
                int iLastCol = GetLastCol((i * 13) + 3, gbl.SheetType.sTEAM);
                int iLastGame = GetLastCol((i * 13) + 13, gbl.SheetType.sTEAM);

                if (iLastCol > 2)
                {
                    bIsFirstGame = false;
                }
                int iLstCl = -1;
                if ((iLastCol + 1) <= iLastGame)
                {
                    iLstCl = iLastCol + 1;
                    string sFirstUnplayedMatch = xlWorksheet.Cells[(i * 13) + 13, iLstCl].Value.ToString().Trim();
                    aAll = sFirstUnplayedMatch.Split(')');
                    string sFirstUnplayedMatchNo = aAll[0].ToString().Trim();
                    int iFirstUnplayedMatchNo = Convert.ToInt32(sFirstUnplayedMatchNo);
                    string sOpp = aAll[1].ToString().Trim();
                    if (iFirstUnplayedMatchNo <= iLowestUnplayedGameSoFar)
                    {
                        iLowestUnplayedGameSoFar = iFirstUnplayedMatchNo;
                        string sGrndDate = xlWorksheet.Cells[(i * 13) + 14, iLastCol + 1].Value.ToString().Trim();
                        //string[] aGndDate;


                        string[] sSeparator = { " - " };
                        //string[] aAll = sSlug.Split(sSeparator, StringSplitOptions.None);

                        string[] aGndDate = sGrndDate.Split(sSeparator, StringSplitOptions.None);

                        //aGndDate = sGrndDate.Split('-');
                        string sStart = aGndDate[1].ToString().Trim();
                        hglg.sStart = sStart;
                        hglg.dStart = DateTime.ParseExact(sStart, "dd=MM=yyyy", System.Globalization.CultureInfo.InvariantCulture);
                        string sHomeAway = aGndDate[2].ToString().Trim();
                        if (sHomeAway == "Home")
                        {
                            hglg.sTeamAway = aAll[1].ToString().Trim();
                            hglg.sTeamHome = xlWorksheet.Cells[(i * 13) + 3, 2].Value.ToString().Trim();
                        }
                        else
                        {
                            hglg.sTeamHome = aAll[1].ToString().Trim();
                            hglg.sTeamAway = xlWorksheet.Cells[(i * 13) + 3, 2].Value.ToString().Trim();
                        }
                    }
                }
                //else
                //{
                //iLstCl = iLastCol;
                //}
            }

            CleanUpExcel();

            if (iLowestUnplayedGameSoFar != 1000)
            {
                hglg.iHighestGame = iLowestUnplayedGameSoFar;
            }
            else
            {
                hglg.iHighestGame = iHighestLastGameSoFar;
            }

            hglg.iLastGame = iHighestLastGameSoFar;
            hglg.bIsFirstGame = bIsFirstGame;

            if (hglg.iHighestGame == hglg.iLastGame)
            {
                hglg.bIsLastGame = true;
            }

            if (hglg.bIsFirstGame)
            {
                //hglg.iHighestGame = iLowestFirstMatchGameSoFar;
            }

            if (hglg.sTeamHome.Count() > 3)
            {
                hglg.sTeamHome = GetTeamNameAbbrev(hglg.sTeamHome);
            }
            else
            {
                hglg.sTeamAway = GetTeamNameAbbrev(hglg.sTeamAway);
            }

            fm.AddData("Next Game is " + hglg.iHighestGame.ToString() + "\r\n");

            return hglg;
        }

        public void UpdateBatsman(int iMatchNo, Player plyr, ResultsBatsman rb, CricketForm fm)
        {
            bool bPlayerFound = false;
            string[] aAll;
            Player CurrPlyr = new Player();
            int iDataRow = -1;
            int iLastRow = -1;
            int iNoPlayers = -1;

            InitExcel();
            xlWorksheet = (Excel.Worksheet)xlWorkbook.Worksheets[gbl.SheetType.sBATSMEN];
            iLastRow = GetLastRowFast(1, xlWorksheet);
            if (iLastRow < 9)
            {
                iNoPlayers = 0;
            }
            else
            {
                iNoPlayers = ((iLastRow - 3) / 6);
            }
            int i = 0;
            bPlayerFound = false;
            while (!bPlayerFound && i < iNoPlayers)
            {
                string sDB = xlWorksheet.Cells[(i * 6) + 7, 2].Text as string;
                if (sDB != "")
                {
                    CurrPlyr.LastName = xlWorksheet.Cells[(i * 6) + 4, 2].Value.ToString().Trim();
                    sDB = xlWorksheet.Cells[(i * 6) + 7, 2].Value().Trim();
                    CurrPlyr.DOB = StringToDate(sDB);
                    string sFirstNames = xlWorksheet.Cells[(i * 6) + 5, 2].Value.ToString().Trim();
                    aAll = sFirstNames.Split('=');
                    CurrPlyr.FirstName = aAll[0].Trim();
                    if (IsSamePlayer(CurrPlyr, plyr))
                    {
                        fm.AddData("Batsman " + plyr.FirstName + " " + plyr.LastName + " HAS been found in BatsmenSheet - " + "Adding Game Data now\r\n");

                        iDataRow = (i * 6) + 4;
                        //if (rb.IsSignificant == true)
                        //{
                        EnterBatsmanData(iMatchNo, iDataRow, plyr, rb);
                        //}
                        bPlayerFound = true;
                    }
                }
                i++;
            }
            if (bPlayerFound == false)
            {
                fm.AddData("Batsman " + plyr.FirstName + " " + plyr.LastName + " HAS NOT been found in BatsmenSheet - Creating Player and adding Game Data now" + "\r\n");
                xlWorksheet = (Excel.Worksheet)xlWorkbook.Worksheets[gbl.SheetType.sBATSMEN];
                iLastRow = GetLastRowFast(1, xlWorksheet);
                if (iLastRow < 9)
                {
                    iDataRow = 4;
                }
                else
                {
                    iDataRow = iLastRow + 1;
                }

                bool bIsBatsman = true;

                AddNewPlayer(bIsBatsman, iDataRow, plyr);
                //if (rb.IsSignificant == true)
                //{
                EnterBatsmanData(iMatchNo, iDataRow, plyr, rb);
                //}
                xlApp.Run("ByOrder", "Batsmen");

            }
            xlApp.DisplayAlerts = false;
            xlWorkbook.SaveAs(gbl.sEXCELPATH, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing,
                 Type.Missing, Type.Missing);
            CleanUpExcel();
        }

        private void EnterBatsmanData(int iMatchNo, int iDataRowNo, Player Plyr, ResultsBatsman rb)
        {
            string[] aAll;

            int iOppRow = iDataRowNo + 4;
            xlWorksheet = (Excel.Worksheet)xlWorkbook.Worksheets[gbl.SheetType.sBATSMEN];
            int iLastCol = GetLastCol(iOppRow, gbl.SheetType.sBATSMEN);
            for (int i = 3; i <= iLastCol; i++)
            {
                string sMatchNoOpp = xlWorksheet.Cells[iOppRow, i].Value.ToString().Trim();
                aAll = sMatchNoOpp.Split(')');
                string sCurrMatch = aAll[0].Trim();
                int iCurrMatch = Convert.ToInt32(sCurrMatch);
                if (iCurrMatch == iMatchNo)
                {
                    if (rb.Runs == null)
                    {
                        return;
                    }
                    else
                    {
                        if (rb.IsSignificant == true)
                        {
                            xlWorksheet.Cells[iDataRowNo, i].Value = rb.Runs.Trim() + "=" + rb.BallsFaced.Trim() + "=" + rb.Four.Trim() + "=" + rb.Six.Trim();
                        }
                        else
                        {
                            string sData = rb.Runs.Trim() + "=" + rb.BallsFaced.Trim() + "=" + rb.Four.Trim() + "=" + rb.Six.Trim() + "=*";
                            xlWorksheet.Cells[iDataRowNo, i].Value = sData;
                        }
                        if (rb.NotOut == true)
                        {
                            xlWorksheet.Cells[iDataRowNo, i].Font.Color = 12611584;
                        }
                        else
                        {
                            xlWorksheet.Cells[iDataRowNo, i].Font.Color = 0;
                        }
                        return;
                    }
                }
            }
            return;
        }
    }

    public class ExceWebScrapingDraw : ExcelWorkbook
    {
        private int GetTeamCol(string sSlug)
        {
            xlWorksheet = (Excel.Worksheet)xlWorkbook.Worksheets[gbl.SheetType.sDRAWCURRENT];
            int iLastCol = GetLastCol(1, gbl.SheetType.sDRAWCURRENT);

            for (int i = 2; i < iLastCol; i = i + 3)
            {
                string sCurrTeam = xlWorksheet.Cells[1, i].Value.ToString().Trim();
                if (sCurrTeam == sSlug)
                {
                    return i - 1;
                }
            }
            return -1;
        }
        public void AddGameInDraw(GameInDraw gid)
        {
            InitExcel();
            xlWorksheet = (Excel.Worksheet)xlWorkbook.Worksheets[gbl.SheetType.sDRAWCURRENT];

            int iCol = -1;
            int iLastRow = -1;

            iCol = GetTeamCol(gid.sTeamHome);
            iLastRow = GetLastRow(iCol, gbl.SheetType.sDRAWCURRENT);
            xlWorksheet.Cells[iLastRow + 1, iCol].Value = gid.iGameNo.ToString();
            xlWorksheet.Cells[iLastRow + 1, iCol + 1].Value = gid.sTeamAway;
            xlWorksheet.Cells[iLastRow + 1, iCol + 2].Value = gid.sGround + " - " + gid.sStart + " - " + "Home";

            iCol = GetTeamCol(gid.sTeamAway);
            iLastRow = GetLastRow(iCol, gbl.SheetType.sDRAWCURRENT);
            xlWorksheet.Cells[iLastRow + 1, iCol].Value = gid.iGameNo.ToString();
            xlWorksheet.Cells[iLastRow + 1, iCol + 1].Value = gid.sTeamHome;
            xlWorksheet.Cells[iLastRow + 1, iCol + 2].Value = gid.sGround + " - " + gid.sStart + " - " + "Away";

            xlApp.DisplayAlerts = false;
            xlWorkbook.SaveAs(gbl.sEXCELPATH, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing,
                              Type.Missing, Type.Missing);
            CleanUpExcel();
        }

        public string GetGroundNameAbbrev(string sGroundNameFull)
        {
            string sAbbrev = "";

            InitExcel();
            xlWorksheet = (Excel.Worksheet)xlWorkbook.Worksheets[gbl.SheetType.sGROUNDSIZES];
            int iLastRow = GetLastRow(6, gbl.SheetType.sGROUNDSIZES);
            for (int i = 2; i <= iLastRow; i++)
            {
                string sCurrGroundFull = xlWorksheet.Cells[i, 6].Value.ToString().Trim();
                if (sCurrGroundFull == sGroundNameFull)
                {
                    sAbbrev = xlWorksheet.Cells[i, 3].Value.ToString().Trim();
                    CleanUpExcel();
                    return sAbbrev;
                }
            }
            CleanUpExcel();
            return "Error: " + sGroundNameFull + " is NOT in Sheet:GroundSizes";
        }

        public int GetLastInDrawCurrent(CricketForm fm)
        {
            int iLastMatch = 0;
            int iCurrMatch = 0;
            InitExcel();
            xlWorksheet = (Excel.Worksheet)xlWorkbook.Worksheets[gbl.SheetType.sDRAWCURRENT];
            //int iLastRow = GetLastRow(1, gbl.SheetType.sDRAWCURRENT);
            //if (iLastRow == 1)
            //{
            //    fm.AddData("\r\nNo Data in Sheet:DrawCurrent\r\n");
            //}
            //else
            //{
            int iLastCol = GetLastCol(1, gbl.SheetType.sDRAWCURRENT);
            for (int i = 1; i <= iLastCol - 2; i = i + 3)
            {
                int iLastRow = GetLastRow(i, gbl.SheetType.sDRAWCURRENT);
                if (iLastRow > 1)
                {
                    string sCurrMatch = xlWorksheet.Cells[iLastRow, i].Value.ToString().Trim();
                    iCurrMatch = Convert.ToInt32(sCurrMatch);
                    if (iCurrMatch >= iLastMatch)
                    {
                        iLastMatch = iCurrMatch;
                    }
                }
            }
            //}
            CleanUpExcel();
            return iLastMatch;
        }
    }
}