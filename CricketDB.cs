using Microsoft.Azure.Cosmos;
using Microsoft.Azure.Documents;
using Microsoft.Azure.Documents.Client;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Cricket
{
    public class CricketDB
    {
        List<Match> matchesInDB = new List<Match>();
        public List<Season> seasonsInDB = new List<Season>();
        List<Season> aBatsmenCrrnt = new List<Season>();
        List<Season> aBowlerCrrnt = new List<Season>();
        List<Season> aMerged = new List<Season>();

        private const string sDBCRICKET = "DBCricket";
        
        
        // The Azure Cosmos DB endpoint for running this sample.
        //private static readonly string EndpointUri = "https://acccountname1.documents.azure.com:443/";
        private static readonly string EndpointUri = "https://accountname3.documents.azure.com:443/";


        // The primary key for the Azure Cosmos account.
        //private static readonly string PrimaryKey = "0Uwe70RFf6NEZg7UbmAFMjljBdAufy6SFgnwRHBg67K0wSWMWezhxS9wDCauQmQYxajNzTGeVb180klEOLB9Ew==";
        private static readonly string PrimaryKey = "FVLB29ezyxK7dbG5XP6PfcddpXkrBVrdDVbkUkAkY5jtAc8tPNprciV6rQIFjlR0hCaqZ8LnVSjDoOxbFpKT9A==";

        // The Cosmos client instance
        private CosmosClient cosmosClient;
        // The database we will create
        private Microsoft.Azure.Cosmos.Database database;
        // The container we will create.
        private Container container;

        public async Task GetStartedDemoAsync(string sTableName, CricketForm fm)
        {
            // Create a new instance of the Cosmos Client
            this.cosmosClient = new CosmosClient(EndpointUri, PrimaryKey);
            await this.CreateDatabaseAsync(fm);

            if (sTableName == gbl.CntnrType.sCNTNRMATCH)
            {
                await this.AddItemsToCntnrMatchAsync(fm);
            }
            else if (sTableName == gbl.CntnrType.sCNTNRSEASON)
            {
                await this.AddItemsToCntnrSeasonAsync(gbl.SheetType.sBATSMEN, fm);
                //await this.AddItemsToCntnrPerformanceAsync(gbl.SheetType.sBATSMEN, fm);
            }
            else if (sTableName == gbl.ButtonType.sDELETEMATCHES)
            {
                await this.DeleteAllCntnrItemsAsync(gbl.CntnrType.sCNTNRMATCH, "/MatchID", fm);
            }
            else if (sTableName == gbl.ButtonType.sDELETESEASON)
            {
                await this.DeleteAllCntnrItemsAsync(gbl.CntnrType.sCNTNRSEASON, "/SeasonID", fm);
                //await this.DeleteAllCntnrItemsAsync(gbl.CntnrType.sCNTNRPERFORMANCE, "/PerformanceID", fm);
            }
            else if (sTableName == gbl.ButtonType.sLOADPLYRSTATS)
            {
                await this.LoadPlyrStatsAsync(fm);
            }
            else if (sTableName == gbl.ButtonType.sLOADMATCHES)
            {
                await this.LoadMatchesAsync(fm);
            }
            else if (sTableName == gbl.ButtonType.sUPDATEPLYRLISTBATBOWL)
            {
                await this.PlyrListBatBowlAsync(fm);
            }

            //else if (sTableName == gbl.CntnrType.sDELETEPLAYERTEAMS)
            //{
            //    await this.DeleteAllCntnrItemsAsync(gbl.CntnrType.sCNTNRPLAYERTEAM, "/PlayerTeamID", fm);
            //}
            //else if (sTableName == gbl.CntnrType.sLOADTEAMS)
            //{
            //    await this.LoadTeamsAsync(fm);
            //}
            //else if (sTableName == gbl.CntnrType.sSavePlrTm)
            //{
            //    await this.SavePlrTmAsync(fm);
            //}
            //await this.QueryItemsAsync(fm);
            //await this.ReplaceFamilyItemAsync(fm);
            //await this.DeleteFamilyItemAsync(fm);
            //await this.DeleteDatabaseAndCleanupAsync(fm);
        }

        public void MainEntry(string sTablename, CricketForm frm)
        {
            Task t;
            t = Start(sTablename, frm);
        }

        public static async Task Start(string sTableName, CricketForm fm)
        {
            try
            {
                Console.WriteLine("Beginning operations...\n");
                fm.AddData("Beginning operations...\r\n");
                CricketDB p = new CricketDB();
                await p.GetStartedDemoAsync(sTableName, fm);

            }
            catch (CosmosException de)
            {
                Exception baseException = de.GetBaseException();
                Console.WriteLine("{0} error occurred: {1}", de.StatusCode, de);
                fm.AddData("\r\n" + de.StatusCode + " error occurred: " + de + "\r\n");
            }
            catch (Exception e)
            {
                Console.WriteLine("Error: {0}", e);
                fm.AddData("\r\n" + "Error: " + e);
            }
            finally
            {
                Console.WriteLine("Operation has completed successfully!");
                fm.AddData("\r\n" + "Operation has completed successfully!\r\n\r\n");
                Console.ReadKey();
            }
        }

        /// <summary>
        /// Create the database if it does not exist
        /// </summary>
        private async Task CreateDatabaseAsync(CricketForm fm)
        {
            // Create a new database
            this.database = await this.cosmosClient.CreateDatabaseIfNotExistsAsync(sDBCRICKET);
            Console.WriteLine("Created Database: {0}\n", this.database.Id);
            fm.AddData("\r\n" + "Created Database: " + this.database.Id + "\r\n");
        }

        /// <summary>
        /// Delete the database and dispose of the Cosmos Client instance
        /// </summary>
        private async Task DeleteDatabaseAndCleanupAsync(CricketForm fm)
        {
            DatabaseResponse databaseResourceResponse = await this.database.DeleteAsync();
            // Also valid: await this.cosmosClient.Databases["FamilyDatabase"].DeleteAsync();

            //Console.WriteLine("Deleted Database: {0}\n", this.databaseId);
            fm.AddData("\r\n" + "Deleted Database: " + sDBCRICKET + "\r\n");

            //Dispose of CosmosClient
            this.cosmosClient.Dispose();
        }



        /// <summary>
        /// Create the container if it does not exist. 
        /// Specifiy "/LastName" as the partition key since we're storing family information, to ensure good distribution of requests and storage.
        /// </summary>
        /// <returns></returns>
        //private async Task CreateContainerAsync(String sCntrName, string sPartitionPath, CricketForm fm)
        private async Task CreateContainerAsync(String sCntrName, string sPartitionPath, CricketForm fm)
        {
            // Create a new container
            //this.container = await this.database.CreateContainerIfNotExistsAsync(sCntrName, sPartitionPath);
            this.container = await this.database.CreateContainerIfNotExistsAsync(sCntrName, sPartitionPath);

            Console.WriteLine("Created Container: {0}\n", this.container.Id);
            fm.AddData("\r\n" + "Created Container: " + this.container.Id + "\r\n");
        }




        private async Task CreateListAsync(string sCntnr, string sPartitrionKey, CricketForm fm)
        {
            int i = 0;
            var sqlQueryText = "SELECT * FROM c";

            await this.CreateContainerAsync(sCntnr, sPartitrionKey, fm);
            fm.AddData("\r\n" + "Running query: " + sqlQueryText + "\r\n");
            QueryDefinition queryDefinition = new QueryDefinition(sqlQueryText);
            FeedIterator<object> queryResultSetIterator = this.container.GetItemQueryIterator<object>(queryDefinition);
            while (queryResultSetIterator.HasMoreResults)
            {
                Microsoft.Azure.Cosmos.FeedResponse<object> currentResultSet = await queryResultSetIterator.ReadNextAsync();
                foreach (object item in currentResultSet)
                {
                    if (sCntnr == gbl.CntnrType.sCNTNRMATCH)
                    {
                        var result = Newtonsoft.Json.JsonConvert.DeserializeObject<Match>(item.ToString());
                        matchesInDB.Add(result);
                    }
                    else if (sCntnr == gbl.CntnrType.sCNTNRSEASON)
                    {
                        var result = Newtonsoft.Json.JsonConvert.DeserializeObject<Season>(item.ToString());
                        seasonsInDB.Add(result);
                    }
                    i++;
                }
            }
            if (i == 0)
            {
                fm.AddData("\r\n" + "No Items in " + sCntnr + " Container\r\n");
            }

            //Check Partition Key Name
            //ContainerProperties properties = await container.ReadContainerAsync();
            //Console.WriteLine(properties.PartitionKeyPath);
            //string s = properties.PartitionKeyPath;
        }

        private async Task DeleteAllCntnrItemsAsync(string sCntnr, string sPartitrionKey, CricketForm fm)
        {
            int iCount = -1;
            string sCount = "";
            var sqlQueryText = "SELECT VALUE COUNT(1) FROM c";

            await this.CreateContainerAsync(sCntnr, sPartitrionKey, fm);
            fm.AddData("\r\n" + "Running query: " + sqlQueryText + "\r\n");
            QueryDefinition queryDefinition = new QueryDefinition(sqlQueryText);
            FeedIterator<object> queryResultSetIterator = this.container.GetItemQueryIterator<object>(queryDefinition);
            Microsoft.Azure.Cosmos.FeedResponse<object> currentResultSet = await queryResultSetIterator.ReadNextAsync();
            object[] aCount = new object[1];
            aCount = currentResultSet.Resource.ToArray();
            sCount = aCount[0].ToString();
            iCount = Convert.ToInt32(sCount);


            for (int i = 1; i <= iCount; i++)
            {
                // Delete an item. Note we must provide the partition key value and id of the item to delete
                ItemResponse<Season> PerformanceResponse = await this.container.DeleteItemAsync<Season>(i.ToString(), new Microsoft.Azure.Cosmos.PartitionKey(i.ToString()));
                fm.AddData("\r\n" + "Deleted Item from " + sCntnr + ": Item Number [" + i.ToString() + "]\r\n");
            }
        }



        //private async Task LoadTeamsAsync(CricketForm fm)
        //{
            //ExcelLoadPlyrs lp = new ExcelLoadPlyrs();
            //await this.CreateListAsync(gbl.CntnrType.sCNTNRMATCH, "/MatchID", fm);
            //lp.LoadTeams(matchesInDB, fm);
            //lp.CleanUpExcel();
        //}

        private async Task AddToCntnrSeasonAsync(Season ssn, CricketForm fm)
        {

            ssn.SsnID = (seasonsInDB.Count + 1).ToString().Trim();
            ssn.SeasonID = (seasonsInDB.Count + 1).ToString().Trim();

            fm.AddData("\r\n" + ssn.Comp.CompetitionCode + " - " + ssn.Comp.Season + " : "  + ssn.Playr.FirstName + " " + ssn.Playr.LastName + " :  IS NOT in the database");
            try
            {
                // Read the item to see if it exists.  
                ItemResponse<Season> ssnResponse = await this.container.ReadItemAsync<Season>(ssn.SsnID, new Microsoft.Azure.Cosmos.PartitionKey(ssn.SeasonID));
                fm.AddData("\r\n" + "Item in Season container with id: " + ssnResponse.Resource.SsnID + " already exists\r\n");
            }
            catch (CosmosException ex) when (ex.StatusCode == HttpStatusCode.NotFound)
            {
                // Create an item in the container representing the Andersen family. Note we provide the value of the partition key for this item, which is "Andersen"
                ItemResponse<Season> ssnResponse = await this.container.CreateItemAsync<Season>(ssn, new Microsoft.Azure.Cosmos.PartitionKey(ssn.SeasonID));
                // Note that after creating the item, we can access the body of the item with the Resource property off the ItemResponse. We can also access the RequestCharge property to see the amount of RUs consumed on this request.
                fm.AddData("\r\n" + "Created item in Season container with id: " + ssnResponse.Resource.SsnID + " Operation consumed " + ssnResponse.RequestCharge + " RUs.\r\n");
            }
            seasonsInDB.Add(ssn);

        }

        private async Task AddItemsToCntnrSeasonAsync(string sPlayerType, CricketForm fm)
        {
            Season ssn = new Season();
            ExcelSeason es = new ExcelSeason();


            await this.CreateListAsync(gbl.CntnrType.sCNTNRSEASON, "/SeasonID", fm);

            

            aBatsmenCrrnt = es.LoadExcelBatBowl(gbl.SheetType.sBATSMEN, seasonsInDB, fm);
            aBowlerCrrnt = es.LoadExcelBatBowl(gbl.SheetType.sBOWLER, seasonsInDB, fm);
            aMerged = es.MergeBatBowl(aBatsmenCrrnt, aBowlerCrrnt, seasonsInDB, fm);

            if (aMerged != null)
            {
                for (int i = 0; i < aMerged.Count; i++)
                {
                    await this.AddToCntnrSeasonAsync(aMerged[i], fm);
                }
            }
            //es.CleanUpExcel();
        }

        private async Task LoadPlyrStatsAsync(CricketForm fm)
        {
            ExcelLoadPlyrs lp = new ExcelLoadPlyrs();

            await this.CreateListAsync(gbl.CntnrType.sCNTNRSEASON, "/SeasonID", fm);

            lp.LoadPlyrStatsThisLast(seasonsInDB, fm);

            //lp.CleanUpExcel();
        }

        private async Task AddItemsToCntnrMatchAsync(CricketForm fm)
        {

            bool bMatchISInDB = false;
            ExcelMatch em = new ExcelMatch();
            em.InitExcel();
            //int a = 0;
            int x = 0;
            Match mtch = new Match();
            List<Match> mtchs = new List<Match>();
            List<TeamInnings> tminngs = new List<TeamInnings>();

            em.Comp = em.GetCompetition(gbl.SheetType.sTEAM, fm);
            tminngs = em.GetTeamSeason(fm);
            mtchs = em.MergeTeamSeasons(tminngs, fm);

            // matchesInDB
            await this.CreateListAsync(gbl.CntnrType.sCNTNRMATCH, "/MatchID", fm);
            int iNoGames = mtchs.Count;
            for (int i = 0; i < iNoGames; i++)
            {
                //mtch = em.GetNextMatch(i + 1, fm);

                mtch = mtchs[i];

                if (mtch != null)
                {
                    x = 0;
                    bMatchISInDB = false;
                    while ((x <= matchesInDB.Count - 1) && (bMatchISInDB == false))
                    {
                        if (mtch.Cmpttn.CompetitionCode == matchesInDB[x].Cmpttn.CompetitionCode && mtch.Cmpttn.Season == matchesInDB[x].Cmpttn.Season
                            && mtch.GameNumber == matchesInDB[x].GameNumber)
                        {
                            bMatchISInDB = true;
                        }
                        x++;
                    }
                    if (bMatchISInDB)
                    {
                        fm.AddData("\r\n" + "Match:: COMPETITION " + mtch.Cmpttn.CompetitionCode + " SEASON: " + mtch.Cmpttn.Season + " GAME NUMBER: " + mtch.GameNumber + "  IS in the database");
                    }
                    else
                    {
                        mtch.MtchID = (matchesInDB.Count + 1).ToString().Trim();
                        mtch.MatchID = (matchesInDB.Count + 1).ToString().Trim();

                        fm.AddData("\r\n" + "Match:: COMPETITION " + mtch.Cmpttn.CompetitionCode + " SEASON: " + mtch.Cmpttn.Season + " GAME NUMBER: " + mtch.GameNumber + "  IS NOT in the database");
                        try
                        {
                            // Read the item to see if it exists.  
                            ItemResponse<Match> mtchResponse = await this.container.ReadItemAsync<Match>(mtch.MtchID, new Microsoft.Azure.Cosmos.PartitionKey(mtch.MatchID));
                            fm.AddData("\r\n" + "Item in Match container with id: " + mtchResponse.Resource.MtchID + " already exists\r\n");
                        }
                        catch (CosmosException ex) when (ex.StatusCode == HttpStatusCode.NotFound)
                        {
                            // Create an item in the container representing the Andersen family. Note we provide the value of the partition key for this item, which is "Andersen"
                            ItemResponse<Match> mtchResponse = await this.container.CreateItemAsync<Match>(mtch, new Microsoft.Azure.Cosmos.PartitionKey(mtch.MatchID));
                            // Note that after creating the item, we can access the body of the item with the Resource property off the ItemResponse. We can also access the RequestCharge property to see the amount of RUs consumed on this request.
                            fm.AddData("\r\n" + "Created item in Match container with id: " + mtchResponse.Resource.MtchID + " Operation consumed " + mtchResponse.RequestCharge + " RUs.\r\n");
                        }
                        matchesInDB.Add(mtch);
                    }
                }
                else
                {
                    fm.AddData("\r\n" + "This Match has no Data");
                }
            }
            em.CleanUpExcel();
        }

        private async Task LoadMatchesAsync(CricketForm fm)
        {

            ExcelLoadMatches lm = new ExcelLoadMatches();
            //lm.InitExcel();

            await this.CreateListAsync(gbl.CntnrType.sCNTNRMATCH, "/MatchID", fm);
            lm.LoadMatches(matchesInDB, fm);
            //lm.CleanUpExcel();



        }

        private async Task PlyrListBatBowlAsync(CricketForm fm)
        {
            ExcePlayerListBatBowl eplbb = new ExcePlayerListBatBowl();
            List<Player> plyrlist = new List<Player>();
            List<Player> batsmen = new List<Player>();
            List<Player> bowlers = new List<Player>();

            eplbb.InitExcel();
            await this.CreateListAsync(gbl.CntnrType.sCNTNRSEASON, "/SeasonID", fm);

            plyrlist = eplbb.LoadPlayerList(fm);

            batsmen = eplbb.LoadPlayers(gbl.SheetType.sBATSMEN, fm);

            eplbb.AddToPlayerListIfNotThere(batsmen, plyrlist);

            bowlers = eplbb.LoadPlayers(gbl.SheetType.sBOWLER, fm);

            eplbb.AddToPlayerListIfNotThere(bowlers, plyrlist);

            //RemoveDuplicates();

            eplbb.CleanUpExcel();
        }
    }
}
