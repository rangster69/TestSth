using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Cricket
{

    public class SlugObjctID
    {
        public int iGameNo;
        public string sSlug;
        public int iArrNo;
        public string sObjctID;
    }

    public class MatchStrings
    {
        public string sBallByBall;
        public string sMatchOversComparison;
        public string sFullScoreCard;
        public string sMatchResults;
        public string sMatchScheduleFixtures;
        public int iMatchNumber;
        public string sSlug;
    }

    public class DOBchanged
    {
        public string NewDOB;
        public bool bDOBchanged;
    }

    public class HighestGameAndLastGame
    {
        public int iHighestGame;
        public int iLastGame;
        public bool bIsFirstGame;
        public bool bIsLastGame;
        public string sTeamHome;
        public string sTeamAway;
        public DateTime dStart;
        public string sStart;
    }

    public class GameInDraw
    {
        public int iGameNo;
        public string sTeamHome;
        public string sTeamAway;
        public string sGround;
        public DateTime dtStart;
        public string sStart;
    }
}
