using Newtonsoft.Json;
using System.Collections.Generic;
using System;


namespace Cricket
{
    public class Match
    {
        [JsonProperty(PropertyName = "id")]
        public string MtchID { get; set; }
        public string MatchID { get; set; }
        public string TeamWinner { get; set; }
        public string TeamLoser { get; set; }
        public Competition Cmpttn { get; set; }
        public int GameNumber { get; set; }
        public System.DateTime MatchDate { get; set; }
        public string TeamHome { get; set; }
        public string TeamAway { get; set; }
        public string GroundName { get; set; }
        public string BatFirst { get; set; }
        public string BatSecond { get; set; }
        public ResultsMatch HomeTeamData { get; set; }
        public ResultsMatch AwayTeamData { get; set; }
        public override string ToString()
        {
            return JsonConvert.SerializeObject(this);
        }
    }
    public class TeamGame
    {
        public TeamInnings HomeTeam;
        public TeamInnings AwayTeam;
    }

    public class TeamInnings
    {
        public Competition Cmpttn;
        public string MyTeam;
        public ResultsMatch ResBatting;
        public int GameNumber;
        public string OppTeam;
        public DateTime GameDate;
        public string Ground;
        public string TeamHome;
        public string TeamAway;
        public string BattingWinOrLose;
        public string TeamBatFirst;
        public string TeamBatSecond;
    }

    public class ResultsMatch
    {
        public string Four { get; set; }
        public string Six { get; set; }
        public string XOverScore { get; set; }
        public string OpeningPartnership { get; set; }
        public string WicketsDown { get; set; }
        public string TotalScore { get; set; }
        public string OversFaced { get; set; }

        // "f" or "s"
        public string BattedFirstOrSecond { get; set; }
    }


    public class AllInnFrSsn
    {
        //public string MatchID { get; set; }
        public string GameNumber { get; set; }
        //public string TeamMine { get; set; }
        public string GameNoAndOpp { get; set; }
        public string TeamOpposition { get; set; }
        public ResultsBatsman ResBat { get; set; }
        public ResultsBowler ResBowl { get; set; }
    }

    public class InngsSsnBat
    {
        //public List<ResultsBatsman> AllResBat;
        public int Inngs;
        public int Runs;
        public float HS;
        public int NOs;
        public int Total4s;
        public int Total6s;
        public int BallsFaced;
    }



    public class ResultsBatsman
    {
        public string Runs { get; set; }
        public string BallsFaced { get; set; }
        public string Four { get; set; }
        public string Six { get; set; }
        //public string Average { get; set; }
        public bool NotOut { get; set; }
        public bool IsSignificant { get; set; }
    }

    public class ResultsBowler
    {
        public string Wickets { get; set; }
        public string RunsConceded { get; set; }
        public string OversBowled { get; set; }
        public bool IsSignificant { get; set; }
        public int IntegerFour { get; set; }
        public int IntegerSix { get; set; }
    }



    public class TeamSeason
    {
        public string TmName { get; set; }
        public List<TeamMatches> AllMatches { get; set; }
    }
    public class TeamMatches
    {
        public string TeamName { get; set; }
        public double WinCount { get; set; }
        //public string Opposition { get; set; }
        public double WcktsDwn { get; set; }
        public double Scr { get; set; }
        public double OvrsFcd { get; set; }
        public double Frs { get; set; }
        public double Sxs { get; set; }
        public double OP { get; set; }
        public double PwrPly { get; set; }
        //public double Wns { get; set; }
        public double GmsPlyd { get; set; }
        public double WinByFifteenPerc { get; set; }
        public double WinByTwentyfivePerc { get; set; }
    }

    public class SumAve
    {
        public double Sum;
        public double Ave;
    }
    public class PlyrSts
    {
        public string Comp;
        public string ThisLastSsn;
        public string BeforeThisLastSsn;
        public StatsBat SttsBt;
        public StatsBowl SttsBwl;
    }

    public class StatsBat
    {
        public string Inngs;
        public string Runs;
        public string Ave;
        public string SR;
        public string HS;
        public string NOs;
        public string Ave4s;
        public string Ave6s;
        public string BllFcd;
    }

    public class StatsBatDbl
    {
        public double Inngs;
        public double Runs;
        public double Ave;
        public double SR;
        public double HS;
        public double NOs;
        public double Ave4s;
        public double Ave6s;
        public double BllFcd;
    }

    public class StatsBowl
    {
        public string Inngs;
        public string Ave;
        public string Wkts;
        public string Runs;
        public string Ovrs;
        public string RPO;
        public string SR;
        public string BllBwl;
    }

    public class StatsBowlDbl
    {
        public double Inngs;
        public double Ave;
        public double Wkts;
        public double Runs;
        public double Ovrs;
        public double RPO;
        public double SR;
        public double BllBwl;
    }

    public class Competition
    {
        public string CompetitionCode { get; set; }
        public string Season { get; set; }
    }
    public class Season
    {
        [JsonProperty(PropertyName = "id")]

        public string SsnID { get; set; }
        public string SeasonID { get; set; }

        // Unique on Playr / Competition
        public Player Playr { get; set; }
        public Competition Comp { get; set; }
        public string TeamMine { get; set; }

        // array based on game number 
        public List<AllInnFrSsn> AllInnings { get; set; }
        public override string ToString()
        {
            return JsonConvert.SerializeObject(this);
        }
    }

    public class InngsSsnBowl
    {
        public decimal NoInn;
        public decimal Wkts;
        public decimal RunsConc;
        public decimal OversBwld;
        public decimal BallsBwld;
    }

    public class Player
    {
        public string FirstName { get; set; }
        public string MiddleName { get; set; }
        public string LastName { get; set; }
        public System.DateTime DOB { get; set; }
        public string Country { get; set; }
        public string CountryID { get; set; }
        public string TeamCode { get; set; }
    }

    public class Sounds
    {
        //WMPLib.WindowsMediaPlayer player = new WMPLib.WindowsMediaPlayer();

        public void PlayFile(String url)
        {
#if false
            player.URL = url;
            player.controls.play();
#endif
        }
    }
}


