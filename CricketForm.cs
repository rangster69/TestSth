using System;
using System.Windows;
using System.Windows.Forms;

namespace Cricket
{
    public partial class CricketForm : Form
    {
        CricketDB cdb = new CricketDB();
        WebScraping ws = new WebScraping();
        public CricketForm()
        {
            InitializeComponent();
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void bCancel_Click(object sender, EventArgs e)
        {
            Close();
        }


        private void bMatches_Click(object sender, EventArgs e)
        {
            cdb.MainEntry(gbl.CntnrType.sCNTNRMATCH, this);
        }

        public void AddData(string s)
        {
            this.tbMain.AppendText(s);
        }

        private void bPerformance_Click(object sender, EventArgs e)
        {
            cdb.MainEntry(gbl.CntnrType.sCNTNRSEASON, this);
        }

        private void DeletePerf_Click(object sender, EventArgs e)
        {
            cdb.MainEntry(gbl.ButtonType.sDELETESEASON, this);
        }

        private void LoadPlyrs_Click(object sender, EventArgs e)
        {
            cdb.MainEntry(gbl.ButtonType.sLOADPLYRSTATS, this);
        }

        private void LoadTeams_Click(object sender, EventArgs e)
        {
            //cdb.MainEntry(gbl.ButtonType.sLOADPLYRLIST, this);            
        }

        private void bDeleteMtch_Click(object sender, EventArgs e)
        {
            cdb.MainEntry(gbl.ButtonType.sDELETEMATCHES, this);
        }

        private void DeletePlrTm_Click(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            cdb.MainEntry(gbl.ButtonType.sLOADMATCHES, this);
        }

        private void LoadPlyrList_Click(object sender, EventArgs e)
        {

        }

        private void CricketForm_Load(object sender, EventArgs e)
        {

        }

        private void LoadPlayerListBatBowl_Click(object sender, EventArgs e)
        {
            cdb.MainEntry(gbl.ButtonType.sUPDATEPLYRLISTBATBOWL, this);
        }

        private void UpdateResults_Click(object sender, EventArgs e)
        {
            bool bFinished = false;

            bFinished = ws.Main(this);
        }

        private void bUpdateAllGames_Click(object sender, EventArgs e)
        {
            bool bFinished = false;

            while (!bFinished)
            {
              bFinished = ws.Main(this);
            }
        }

        private void bLoadFinals_Click(object sender, EventArgs e)
        {
            bool bFinished = false;

            bFinished = ws.LoadMatchSchedule(this);
        }

        private void bLoadSeason_Click(object sender, EventArgs e)
        {
            bool bFinished = false;

            bFinished = ws.LoadAllInDraw(this);
        }
    }
}
