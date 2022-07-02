namespace Cricket
{
    class gbl
    {
        public const int iMAXCOLUMNS = 150;
        public const int iMAXROWS = 1500;
        public const string sBATFIRST = "BatFirst";
        public const string sBATSECOND = "BatSecond";
        public const string sHOMETEAM = "Home";
        public const string sAWAYTEAM = "Away";
        public const string sDRAW = "Draw";
        public const string sTIE = "Tie";
        public const string sNORESULT = "No Result";
        public const int iXOvers = 6;

        public const string sEXCELPATH = "C:\\zBrendan\\Cricket\\Cricket-2022-T20Blast-SouthGroup.xlsm";
        public const string sBUPATH = "C:\\zBrendan\\Cricket\\Backups\\SouthGroup";
        public const string sMAINBU = "C:\\zBrendan\\Cricket\\Cricket-2022-T20Blast-SouthGroup-bu.xlsm";
        //public const string sBELLPATH = "C:\\zBrendan\\Cricket\\Projects\\Cricket\\Sounds\\Bell.mp3";


        public class CntnrType
        {
            public const string sCNTNRSEASON = "CntnrSeason";
            public const string sCNTNRMATCH = "CntnrMatch";
        }

        public class ButtonType
        {

            public const string sDELETESEASON = "DeleteSeason";         
            public const string sDELETEMATCHES = "DeleteMatches";
            public const string sUPDATEPLYRLISTBATBOWL = "UpdatePlyrListBatBowl";
            public const string sLOADMATCHES = "LoadPlyrs";
            public const string sLOADPLYRSTATS = "LoadPlyrStats";


            //public const string sCNTNRPLAYERTEAM = "CntnrPlayerTeam";
            //public const string sPERFORMANCEBAT = "PerformanceBat";
            //public const string sPERFORMANCEBOWL = "PerformanceBowl";
            //public const string sDELETEPERFORMANCES = "DeletePerformances";
            //public const string sDELETEPLAYERTEAMS = "DeletePlayerTeams";


            //public const string sLOADTEAMS = "LoadTeams";
            //public const string sSavePlrTm = "SavePlyrTm";
        }

        

        public class SheetType
        {
            // Load previous 2 seasons for this comp
            public const string sPLAYERSTATS = "PlayerStats";
            public const string sPLAYERLIST = "PlayerList";
            public const string sBATSMEN = "Batsmen";
            public const string sBOWLER = "Bowler";
            public const string sTEAM = "Teams";
            public const string sDRAWCURRENT = "DrawCurrent";
            public const string sCOMPETITION = "Competition";
            public const string sTESTBATSMEN = "TestBatsmen";
            public const string sTESTTEAM = "TestTeam";
            public const string sTESTBOWLER = "TestBowler";
            public const string sMATCHES = "Matches";
            public const string sGROUNDSIZES = "GroundSizes";
        }

        public class ColorType
        {
            public const string sRED = "Red";
            public const string sGREEN = "Green";
            public const string sBLACK = "Black";
            public const string sBLUE = "Blue";
        }
    }
}
