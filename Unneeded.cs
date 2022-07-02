namespace Cricket
{
    class Unneeded
    {
        /*
        private async Task QueryItemsAsync(CricketForm fm)
        {
            var sqlQueryText = "SELECT * FROM c WHERE c.LastName = 'Andersen'";

            Console.WriteLine("Running query: {0}\n", sqlQueryText);
            fm.AddData("\r\n" + "Running query: " + sqlQueryText + "\r\n");

            QueryDefinition queryDefinition = new QueryDefinition(sqlQueryText);
            FeedIterator<Family> queryResultSetIterator = this.container.GetItemQueryIterator<Family>(queryDefinition);

            List<Family> families = new List<Family>();

            while (queryResultSetIterator.HasMoreResults)
            {
                FeedResponse<Family> currentResultSet = await queryResultSetIterator.ReadNextAsync();
                foreach (Family family in currentResultSet)
                {
                    families.Add(family);
                    Console.WriteLine("\tRead {0}\n", family);
                    fm.AddData("\r\n" + "\tRead " + family + "\r\n");
                }
            }
        }

        /// <summary>
        /// Replace an item in the container
        /// </summary>
        private async Task ReplaceFamilyItemAsync(CricketForm fm)
        {
            ItemResponse<Family> wakefieldFamilyResponse = await this.container.ReadItemAsync<Family>("Wakefield.7", new PartitionKey("Wakefield"));
            var itemBody = wakefieldFamilyResponse.Resource;

            // update registration status from false to true
            itemBody.IsRegistered = true;
            // update grade of child
            itemBody.Children[0].Grade = 6;

            // replace the item with the updated content
            wakefieldFamilyResponse = await this.container.ReplaceItemAsync<Family>(itemBody, itemBody.Id, new PartitionKey(itemBody.LastName));
            Console.WriteLine("Updated Family [{0},{1}].\n \tBody is now: {2}\n", itemBody.LastName, itemBody.Id, wakefieldFamilyResponse.Resource);
            fm.AddData("\r\n" + "Updated Family [" + itemBody.LastName + "," + itemBody.Id + "].\r\n \tBody is now: " + wakefieldFamilyResponse.Resource + "\r\n");
        }

        /// <summary>
        /// Delete an item in the container
        /// </summary>
        private async Task DeleteFamilyItemAsync(CricketForm fm)
        {
            var partitionKeyValue = "Wakefield";
            var familyId = "Wakefield.7";

            // Delete an item. Note we must provide the partition key value and id of the item to delete
            ItemResponse<Family> wakefieldFamilyResponse = await this.container.DeleteItemAsync<Family>(familyId, new PartitionKey(partitionKeyValue));
            Console.WriteLine("Deleted Family [{0},{1}]\n", partitionKeyValue, familyId);
            fm.AddData("\r\n" + "Deleted Family [" + partitionKeyValue + "," + familyId + "]\r\n");
        }

        private async Task AddItemsToPlayerContainerAsync(string sPlayerType, CricketForm fm)
        {
            bool bPlayerISInDB = false;
            int x = 0;
            int iRowCount = -1;
            ExcelPlayer ep = new ExcelPlayer();
            Player plyr = new Player();

            await this.CreateListAsync(gbl.CntnrType.sCNTNRPLAYER, "/LastName", fm);
            if (sPlayerType == gbl.SheetType.sBATSMEN)
            {
                iRowCount = ep.GetLastRow(1, gbl.SheetType.sBATSMEN);
                iRowCount = (iRowCount - 1) / 6;
            }
            else if (sPlayerType == gbl.SheetType.sBOWLER)
            {
                iRowCount = ep.GetLastRow(1, gbl.SheetType.sBOWLER);
                iRowCount = (iRowCount - 1) / 5;
            }
            for (int i = 0; i < iRowCount; i++)
            {
                x = 0;
                bPlayerISInDB = false;
                if (sPlayerType == gbl.SheetType.sBATSMEN)
                {
                    plyr = ep.GetNextPlayerBatsman(i, fm);
                }
                else if (sPlayerType == gbl.SheetType.sBOWLER)
                {
                    plyr = ep.GetNextPlayerBowler(i, fm);
                }
                while ((x <= players.Count - 1) && (bPlayerISInDB == false))
                {
                    if (plyr.FirstName == players[x].FirstName && plyr.MiddleName == players[x].MiddleName && plyr.LastName == players[x].LastName && plyr.DOB == players[x].DOB && plyr.Country == players[x].Country)
                    {
                        bPlayerISInDB = true;
                    }
                    x++;
                }
                if (bPlayerISInDB)
                {
                    fm.AddData("\r\n" + "Player " + plyr.LastName + "  IS in the database");
                }

                else
                {
                    await this.AddToPlayerCntnrAsync(plyr, fm);
                }
            }
        }
        private Player GetPlayerBtsmn(ref bool bPlayerIsInDB, int i, CricketForm fm)
        {
            ExcelPlayer eplyr = new ExcelPlayer();
            Player NewPlyr = new Player();

            bool bGlobalPlayerISInDB = false;
            NewPlyr = eplyr.GetNextPlayerBatsman(i, fm);
            int a = 0;
            while ((a < players.Count) && (bGlobalPlayerISInDB == false))
            {

                if (NewPlyr.FirstName == players[a].FirstName && NewPlyr.MiddleName == players[a].MiddleName && NewPlyr.LastName == players[a].LastName && NewPlyr.DOB == players[a].DOB && NewPlyr.Country == players[a].Country)
                {
                    bPlayerIsInDB = true;
                    NewPlyr.PlayerID = players[a].PlayerID;
                    NewPlyr.PlyrID = players[a].PlayerID;
                }
                a++;
            }
            return NewPlyr;
        }

        private Player GetPlayerBwlr(ref bool bPlayerIsInDB, int i, CricketForm fm)
        {
            ExcelPlayer eplyr = new ExcelPlayer();
            Player NewPlyr = new Player();

            bool bGlobalPlayerISInDB = false;
            NewPlyr = eplyr.GetNextPlayerBowler(i, fm);
            int a = 0;
            while ((a < players.Count) && (bGlobalPlayerISInDB == false))
            {

                if (NewPlyr.FirstName == players[a].FirstName && NewPlyr.MiddleName == players[a].MiddleName && NewPlyr.LastName == players[a].LastName && NewPlyr.DOB == players[a].DOB && NewPlyr.Country == players[a].Country)
                {
                    bPlayerIsInDB = true;
                    NewPlyr.PlayerID = players[a].PlayerID;
                    NewPlyr.PlyrID = players[a].PlayerID;
                }
                a++;
            }
            return NewPlyr;
        }

        private async Task AddToPlayerCntnrAsync(Player plyr, CricketForm fm)
        {
            plyr.PlyrID = (players.Count + 1).ToString().Trim();
            plyr.PlayerID = (players.Count + 1).ToString().Trim();

            fm.AddData("\r\n" + "Player " + plyr.LastName + "  IS NOT in the database");
            try
            {
                // Read the item to see if it exists.  
                ItemResponse<Player> plyrResponse = await this.container.ReadItemAsync<Player>(plyr.PlyrID, new PartitionKey(plyr.LastName));
                fm.AddData("\r\n" + "Item in Player container with id: " + plyrResponse.Resource.PlyrID + " already exists\r\n");
            }
            catch (CosmosException ex) when (ex.StatusCode == HttpStatusCode.NotFound)
            {
                // Create an item in the container representing the Andersen family. Note we provide the value of the partition key for this item, which is "Andersen"
                ItemResponse<Player> plyrResponse = await this.container.CreateItemAsync<Player>(plyr, new PartitionKey(plyr.LastName));
                // Note that after creating the item, we can access the body of the item with the Resource property off the ItemResponse. We can also access the RequestCharge property to see the amount of RUs consumed on this request.
                ;
                //fm.AddData("\r\n" + "Created item in database with id: " + iItemCount.ToString().Trim() + " Operation consumed " + plyrResponse.RequestCharge + " RUs.\r\n");
                fm.AddData("\r\n" + "Created item in Player container with id: " + plyrResponse.Resource.PlyrID + " Operation consumed " + plyrResponse.RequestCharge + " RUs.\r\n");
            }
            players.Add(plyr);

            private async Task AddItemsToPerformanceBowlerContainerAsync(string sPlayerType, CricketForm fm)
            {
                Performance prfrmnc = new Performance();
                Player plyr = new Player();
                ExcelPlayer epl = new ExcelPlayer();
                Match mtch = new Match();
                ExcelMatch em = new ExcelMatch();
                ExcelPerformance ep = new ExcelPerformance();
                int iGameNum = -1;
                string sGameNum = "";
                //bool bBatsmanDataIsInDB = false;
                bool bPlyrISInDB = false;
                string[] aSplitString = new string[2];

                await this.CreateListAsync(gbl.CntnrType.sCNTNRPLAYER, "/LastName", fm);
                await this.CreateListAsync(gbl.CntnrType.sCNTNRMATCH, "/CompetitionName", fm);
                await this.CreateListAsync(gbl.CntnrType.sCNTNRPERFORMANCE, "/PerformanceID", fm);


                int iRowCount = ep.GetLastRow(1, gbl.SheetType.sBOWLER);
                int iNoBowlers = (iRowCount - 1) / 5;
                for (int i = 0; i < iNoBowlers; i++)
                {
                    //plyr = GetPlayerBtsmn(ref bPlyrISInDB, i, fm);
                    plyr = GetPlayerBwlr(ref bPlyrISInDB, i, fm);

                    if (bPlyrISInDB)
                    {
                        // Get all games for this batsman
                        fm.AddData("\r\n" + "Player " + plyr.LastName + "  IS in the database");


                        //int iGameRow = (i * 6) + 6;
                        int iGameRow = (i * 5) + 5;

                        //int iLastCol = epl.GetLastCol(iGameRow, gbl.SheetType.sBATSMEN);
                        int iLastCol = epl.GetLastCol(iGameRow, gbl.SheetType.sBOWLER);

                        for (int j = 3; j <= iLastCol; j++)
                        {
                            int b = 0;
                            bool bMatchISInDB = false;

                            aSplitString = ep.GetGameNo(iGameRow, j, gbl.SheetType.sBOWLER);
                            sGameNum = aSplitString[0].Trim();
                            iGameNum = Int32.Parse(sGameNum);


                            while ((b < matches.Count) && (bMatchISInDB == false))
                            {
                                mtch = em.GetNextMatch(iGameNum, fm);
                                if (mtch.CompetitionName == matches[b].CompetitionName && mtch.Season == matches[b].Season && mtch.GameNumber == matches[b].GameNumber)
                                {
                                    bMatchISInDB = true;
                                    mtch.MatchID = matches[b].MatchID;
                                    mtch.MtchID = matches[b].MatchID;
                                }
                                b++;
                            }
                            if (bMatchISInDB)
                            {
                                fm.AddData("\r\n" + "Match " + mtch.CompetitionName + " :" + mtch.Season + " :" + mtch.GameNumber + " :" + mtch.GameNumber + "  IS in the database");
                                bool bPerformanceIsInDB = false;
                                int c = 0;
                                while ((c < performances.Count) && (bPerformanceIsInDB == false))
                                {
                                    prfrmnc = ep.GetNextPerformance(plyr.PlayerID, plyr.LastName, mtch.MatchID, mtch.GameNumber, gbl.SheetType.sBOWLER, iGameRow, j, fm);

                                    if (mtch.MatchID == performances[c].MatchID && plyr.PlayerID == performances[c].Playr.PlayerID)
                                    {
                                        bPerformanceIsInDB = true;
                                        prfrmnc.MatchID = mtch.MatchID;
                                    }
                                    c++;
                                }
                                if (bPerformanceIsInDB)
                                {
                                    if (prfrmnc.BowlerRes == null)
                                    {

                                    }
                                }
                                else
                                {
                                    if (performances.Count == 0)
                                    {
                                        prfrmnc = ep.GetNextPerformance(plyr.PlayerID, plyr.LastName, mtch.MatchID, mtch.GameNumber, gbl.SheetType.sBOWLER, iGameRow, j, fm);
                                    }
                                    await this.AddToPerformnceCntnrAsync(prfrmnc, plyr, mtch, fm);
                                }
                            }
                            else
                            {
                                MessageBox.Show("Error in AddItemsToPerformanceBatsmanContainerAsync: Match in NOT in the Database");
                            }
                        }
                    }
                    else
                    {
                        MessageBox.Show("Error in AddItemsToPerformanceBatsmanContainerAsync: Player in NOT in the Database");
                    }
                }
            }
        }

        public string GetPlayerNo(Player plyr, List<Player> players)
        {
            //bool bNotFound = true;
            int x = 0;
            while (true)
            {
                if (plyr.FirstName == players[x].FirstName && plyr.MiddleName == players[x].MiddleName && plyr.LastName == players[x].LastName && plyr.DOB == players[x].DOB && plyr.Country == players[x].Country)
                {
                    //bNotFound = false;
                    return players[x].PlayerID;
                }
                x++;
            }
            //return "Error in GetPlayerNo";
        }

        public int GetMemberCount(string sPlayerType)
        {
            int iRowCount = 0;
            if (sPlayerType == gbl.SheetType.sBATSMEN)
            {
                xlWorksheet = xlWorkbook.Sheets[gbl.SheetType.sBATSMEN];
                xlRange = xlWorksheet.UsedRange;

                iRowCount = GetLastRow(1, gbl.SheetType.sBATSMEN);

                return (iRowCount - 1) / 6;
            }
            else if (sPlayerType == gbl.SheetType.sBOWLER)
            {
                xlWorksheet = xlWorkbook.Sheets[gbl.SheetType.sBOWLER];
                xlRange = xlWorksheet.UsedRange;
                iRowCount = xlRange.Rows.Count;
                return (iRowCount - 1) / 5;
            }
            else
            {
                MessageBox.Show("Error in GetMemberCount: CntnrType must be Batsmen or Bowler ");
                return -1;
            }
        }








using System;
using System.Threading.Tasks;
using System.Configuration;
using System.Collections.Generic;
using System.Net;
using Microsoft.Azure.Cosmos;
using System.Windows.Forms;

namespace Cricket
{
    public class CosmosDB
    {
        // ADD THIS PART TO YOUR CODE

        // The Azure Cosmos DB endpoint for running this sample.
        private static readonly string EndpointUri = "https://acccountname1.documents.azure.com:443/";
        // The primary key for the Azure Cosmos account.
        private static readonly string PrimaryKey = "0Uwe70RFf6NEZg7UbmAFMjljBdAufy6SFgnwRHBg67K0wSWMWezhxS9wDCauQmQYxajNzTGeVb180klEOLB9Ew==";

        // The Cosmos client instance
        private CosmosClient cosmosClient;

        // The database we will create
        private Database database;

        // The container we will create.
        private Container container;

        // The name of the database and container we will create
        private string databaseId = "FamilyDatabase";
        private string containerId = "FamilyContainer";

        //public CricketDB;

        //public CricketForm f;

        public CricketForm f;

        public void MainEntry(CricketForm frm)
        {
            //f = frm;
            //f.AddData("yessss");
            Task t;
            t = Main1(frm);
        }

        public static async Task Main1(CricketForm fm)
        {
            try
            {
                Console.WriteLine("Beginning operations...\n");
                fm.AddData("Beginning operations...\r\n");
                CosmosDB p = new CosmosDB();
                await p.GetStartedDemoAsync(fm);
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
                Console.WriteLine("End of demo, press any key to exit.");
                fm.AddData("\r\n" + "End of demo\r\n");
                Console.ReadKey();
            }
        }

        // ADD THIS PART TO YOUR CODE
        /*
            Entry point to call methods that operate on Azure Cosmos DB resources in this sample

        public async Task GetStartedDemoAsync(CricketForm fm)
        {
            // Create a new instance of the Cosmos Client
            this.cosmosClient = new CosmosClient(EndpointUri, PrimaryKey);
            //ADD THIS PART TO YOUR CODE
            await this.CreateDatabaseAsync(fm);
            //ADD THIS PART TO YOUR CODE
            await this.CreateContainerAsync(fm);
            //ADD THIS PART TO YOUR CODE
            await this.AddItemsToContainerAsync(fm);
            //ADD THIS PART TO YOUR CODE
            await this.QueryItemsAsync(fm);
            //ADD THIS PART TO YOUR CODE
            await this.ReplaceFamilyItemAsync(fm);
            //ADD THIS PART TO YOUR CODE
            await this.DeleteFamilyItemAsync(fm);
            await this.DeleteDatabaseAndCleanupAsync(fm);
        }

        /// <summary>
        /// Create the database if it does not exist
        /// </summary>
        private async Task CreateDatabaseAsync(CricketForm fm)
        {
            // Create a new database
            this.database = await this.cosmosClient.CreateDatabaseIfNotExistsAsync(databaseId);
            Console.WriteLine("Created Database: {0}\n", this.database.Id);
            fm.AddData("\r\n" + "Created Database: " + this.database.Id + "\r\n");
        }

        /// <summary>
        /// Create the container if it does not exist. 
        /// Specifiy "/LastName" as the partition key since we're storing family information, to ensure good distribution of requests and storage.
        /// </summary>
        /// <returns></returns>
        private async Task CreateContainerAsync(CricketForm fm)
        {
            // Create a new container
            this.container = await this.database.CreateContainerIfNotExistsAsync(containerId, "/LastName");
            Console.WriteLine("Created Container: {0}\n", this.container.Id);
            fm.AddData("\r\n" + "Created Container: " + this.container.Id + "\r\n");
        }

        /// <summary>
        /// Add Family items to the container
        /// </summary>
        private async Task AddItemsToContainerAsync(CricketForm fm)
        {
            // Create a family object for the Andersen family
            Family andersenFamily = new Family
            {
                Id = "Andersen.1",
                LastName = "Andersen",
                Parents = new Parent[]
                {
            new Parent { FirstName = "Thomas" },
            new Parent { FirstName = "Mary Kay" }
                },
                Children = new Child[]
                {
            new Child
            {
                FirstName = "Henriette Thaulow",
                Gender = "female",
                Grade = 5,
                Pets = new Pet[]
                {
                    new Pet { GivenName = "Fluffy" }
                }
            }
                },
                Address = new Address { State = "WA", County = "King", City = "Seattle" },
                IsRegistered = false
            };

            try
            {
                // Read the item to see if it exists.  
                ItemResponse<Family> andersenFamilyResponse = await this.container.ReadItemAsync<Family>(andersenFamily.Id, new PartitionKey(andersenFamily.LastName));
                Console.WriteLine("Item in database with id: {0} already exists\n", andersenFamilyResponse.Resource.Id);
                fm.AddData("\r\n" + "Item in database with id: " + andersenFamilyResponse.Resource.Id + " already exists\r\n");
            }
            catch (CosmosException ex) when (ex.StatusCode == HttpStatusCode.NotFound)
            {
                // Create an item in the container representing the Andersen family. Note we provide the value of the partition key for this item, which is "Andersen"
                ItemResponse<Family> andersenFamilyResponse = await this.container.CreateItemAsync<Family>(andersenFamily, new PartitionKey(andersenFamily.LastName));


                // Note that after creating the item, we can access the body of the item with the Resource property off the ItemResponse. We can also access the RequestCharge property to see the amount of RUs consumed on this request.
                Console.WriteLine("Created item in database with id: {0} Operation consumed {1} RUs.\n", andersenFamilyResponse.Resource.Id, andersenFamilyResponse.RequestCharge);
                fm.AddData("\r\n" + "Created item in database with id: " + andersenFamilyResponse.Resource.Id + " Operation consumed" + andersenFamilyResponse.RequestCharge + " RUs.\r\n");
            }

            // Create a family object for the Wakefield family
            Family wakefieldFamily = new Family
            {
                Id = "Wakefield.7",
                LastName = "Wakefield",
                Parents = new Parent[]
                {
            new Parent { FamilyName = "Wakefield", FirstName = "Robin" },
            new Parent { FamilyName = "Miller", FirstName = "Ben" }
                },
                Children = new Child[]
                {
            new Child
            {
                FamilyName = "Merriam",
                FirstName = "Jesse",
                Gender = "female",
                Grade = 8,
                Pets = new Pet[]
                {
                    new Pet { GivenName = "Goofy" },
                    new Pet { GivenName = "Shadow" }
                }
            },
            new Child
            {
                FamilyName = "Miller",
                FirstName = "Lisa",
                Gender = "female",
                Grade = 1
            }
                },
                Address = new Address { State = "NY", County = "Manhattan", City = "NY" },
                IsRegistered = true
            };

            try
            {
                // Read the item to see if it exists
                ItemResponse<Family> wakefieldFamilyResponse = await this.container.ReadItemAsync<Family>(wakefieldFamily.Id, new PartitionKey(wakefieldFamily.LastName));
                Console.WriteLine("Item in database with id: {0} already exists\n", wakefieldFamilyResponse.Resource.Id);
                fm.AddData("\r\n" + "Item in database with id: " + wakefieldFamilyResponse.Resource.Id + " already exists\r\n");
            }
            catch (CosmosException ex) when (ex.StatusCode == HttpStatusCode.NotFound)
            {
                // Create an item in the container representing the Wakefield family. Note we provide the value of the partition key for this item, which is "Wakefield"
                ItemResponse<Family> wakefieldFamilyResponse = await this.container.CreateItemAsync<Family>(wakefieldFamily, new PartitionKey(wakefieldFamily.LastName));

                // Note that after creating the item, we can access the body of the item with the Resource property off the ItemResponse. We can also access the RequestCharge property to see the amount of RUs consumed on this request.
                Console.WriteLine("Created item in database with id: {0} Operation consumed {1} RUs.\n", wakefieldFamilyResponse.Resource.Id, wakefieldFamilyResponse.RequestCharge);
                fm.AddData("\r\n" + "Created item in database with id: " + wakefieldFamilyResponse.Resource.Id + " Operation consumed " + wakefieldFamilyResponse.RequestCharge + " RUs.\r\n");
            }
        }

        /// <summary>
        /// Run a query (using Azure Cosmos DB SQL syntax) against the container
        /// </summary>
        private async Task QueryItemsAsync(CricketForm fm)
        {
            var sqlQueryText = "SELECT * FROM c WHERE c.LastName = 'Andersen'";

            Console.WriteLine("Running query: {0}\n", sqlQueryText);
            fm.AddData("\r\n" + "Running query: " + sqlQueryText + "\r\n");

            QueryDefinition queryDefinition = new QueryDefinition(sqlQueryText);
            FeedIterator<Family> queryResultSetIterator = this.container.GetItemQueryIterator<Family>(queryDefinition);

            List<Family> families = new List<Family>();

            while (queryResultSetIterator.HasMoreResults)
            {
                FeedResponse<Family> currentResultSet = await queryResultSetIterator.ReadNextAsync();
                foreach (Family family in currentResultSet)
                {
                    families.Add(family);
                    Console.WriteLine("\tRead {0}\n", family);
                    fm.AddData("\r\n" + "\tRead " + family + "\r\n");
                }
            }
        }

        /// <summary>
        /// Replace an item in the container
        /// </summary>
        private async Task ReplaceFamilyItemAsync(CricketForm fm)
        {
            ItemResponse<Family> wakefieldFamilyResponse = await this.container.ReadItemAsync<Family>("Wakefield.7", new PartitionKey("Wakefield"));
            var itemBody = wakefieldFamilyResponse.Resource;

            // update registration status from false to true
            itemBody.IsRegistered = true;
            // update grade of child
            itemBody.Children[0].Grade = 6;

            // replace the item with the updated content
            wakefieldFamilyResponse = await this.container.ReplaceItemAsync<Family>(itemBody, itemBody.Id, new PartitionKey(itemBody.LastName));
            Console.WriteLine("Updated Family [{0},{1}].\n \tBody is now: {2}\n", itemBody.LastName, itemBody.Id, wakefieldFamilyResponse.Resource);
            fm.AddData("\r\n" + "Updated Family [" + itemBody.LastName + "," + itemBody.Id + "].\r\n \tBody is now: " + wakefieldFamilyResponse.Resource + "\r\n");
        }

        /// <summary>
        /// Delete an item in the container
        /// </summary>
        private async Task DeleteFamilyItemAsync(CricketForm fm)
        {
            var partitionKeyValue = "Wakefield";
            var familyId = "Wakefield.7";

            // Delete an item. Note we must provide the partition key value and id of the item to delete
            ItemResponse<Family> wakefieldFamilyResponse = await this.container.DeleteItemAsync<Family>(familyId, new PartitionKey(partitionKeyValue));
            Console.WriteLine("Deleted Family [{0},{1}]\n", partitionKeyValue, familyId);
            fm.AddData("\r\n" + "Deleted Family [" + partitionKeyValue + "," + familyId + "]\r\n");
        }

        /// <summary>
        /// Delete the database and dispose of the Cosmos Client instance
        /// </summary>
        private async Task DeleteDatabaseAndCleanupAsync(CricketForm fm)
        {
            DatabaseResponse databaseResourceResponse = await this.database.DeleteAsync();
            // Also valid: await this.cosmosClient.Databases["FamilyDatabase"].DeleteAsync();

            Console.WriteLine("Deleted Database: {0}\n", this.databaseId);
            fm.AddData("\r\n" + "Deleted Database: " + this.databaseId + "\r\n");

            //Dispose of CosmosClient
            this.cosmosClient.Dispose();
        }
    }
}











using Newtonsoft.Json;

namespace Cricket
{
    class Family
    {
        [JsonProperty(PropertyName = "id")]
        public string Id { get; set; }
        public string LastName { get; set; }
        public Parent[] Parents { get; set; }
        public Child[] Children { get; set; }
        public Address Address { get; set; }
        public bool IsRegistered { get; set; }
        public override string ToString()
        {
            return JsonConvert.SerializeObject(this);
        }
    }

    public class Parent
    {
        public string FamilyName { get; set; }
        public string FirstName { get; set; }
    }

    public class Child
    {
        public string FamilyName { get; set; }
        public string FirstName { get; set; }
        public string Gender { get; set; }
        public int Grade { get; set; }
        public Pet[] Pets { get; set; }
    }

    public class Pet
    {
        public string GivenName { get; set; }
    }

    public class Address
    {
        public string State { get; set; }
        public string County { get; set; }
        public string City { get; set; }
    }
}










 public Player GetNextPlayerBowler(int iPlyrNo, CricketForm fm)
 {
     string sNames = "";
     string[] aSplitString;
     string sFirstName = "";
     string sMiddleName = "";
     string sDate = "";

     //Console.WriteLine("Excel is Working...\n");
     fm.AddData("\r\n" + "Excel has Loaded in Player...\r\n");

     xlWorksheet = xlWorkbook.Sheets["Bowler"];
     xlRange = xlWorksheet.UsedRange;

     Player plyrCurrent = new Player();
     int iFirstRow;
     iFirstRow = (iPlyrNo * 5) + 2;
     //plyrCurrent.PlayerID = ((iPlyrNo + 1).ToString()).Trim();
     plyrCurrent.LastName = xlRange.Cells[iFirstRow, 2].Value2.ToString().Trim();

     sNames = xlRange.Cells[iFirstRow + 1, 2].Value2.ToString().Trim();
     aSplitString = sNames.Split(' ');
     sFirstName = aSplitString[0];
     sMiddleName = aSplitString[1];

     plyrCurrent.FirstName = sFirstName.Trim();
     plyrCurrent.MiddleName = sMiddleName.Trim();


     sDate = xlRange.Cells[iFirstRow + 2, 2].Value2.ToString().Trim();
     sDate = FrmtDate(sDate);
     DateTime date = DateTime.ParseExact(sDate, "dd/MM/yyyy HH:mm:ss", CultureInfo.InvariantCulture);

     plyrCurrent.DOB = date;

     plyrCurrent.Country = xlRange.Cells[iFirstRow + 3, 2].Value.ToString().Trim();
     return plyrCurrent;
 }

*/

        // Archived 7-3-22
        /*
        private async Task SavePlrTmAsync(CricketForm fm)
        {
            ExcelSavePlrTm esp = new ExcelSavePlrTm();
            PlayerTeam plyrtm = new PlayerTeam();

            bool bPlayerTeamISInDB = false;
            //ExcelMatch em = new ExcelMatch();
            //int a = 0;
            int x = 0;
            Match mtch = new Match();


            await this.CreateListAsync(gbl.CntnrType.sCNTNRPLAYERTEAM, "/PlayerTeamID", fm);

            //int iNoGames = em.GetNoGamesForSeason(fm);
            int iLastRow = esp.GetLastRow(1, gbl.SheetType.sBATSMEN);
            int iNoPlayers = (iLastRow - 1) / 6;
            int iCounter = playerteams.Count;
            for (int i = 0; i < iNoPlayers; i++)
            {
                 plyrtm = esp.GetNextPlayer(i, fm);
                if (plyrtm != null)
                {
                    x = 0;
                    bPlayerTeamISInDB = false;
                    while ((x <= playerteams.Count - 1) && (bPlayerTeamISInDB == false))
                    {
                        if (plyrtm.Playr.FirstName == playerteams[x].Playr.FirstName && plyrtm.Playr.MiddleName == playerteams[x].Playr.MiddleName && plyrtm.Playr.LastName == playerteams[x].Playr.LastName && plyrtm.Playr.DOB == playerteams[x].Playr.DOB && plyrtm.Playr.Country == playerteams[x].Playr.Country)
                        {
                            if (plyrtm.Cmpttn.CompetitionCode == playerteams[x].Cmpttn.CompetitionCode && plyrtm.Cmpttn.Season == playerteams[x].Cmpttn.Season)
                            {
                                bPlayerTeamISInDB = true;
                            }

                        }
                        x++;
                    }
                    if (bPlayerTeamISInDB)
                    {
                        fm.AddData("\r\n" + "Player/Team:: PLAYER: " + plyrtm.Playr.LastName + " TEAM: " + plyrtm.Tm.TeamName + " COMPETITION: " + plyrtm.Cmpttn.CompetitionName +" SEASON: " + plyrtm.Cmpttn.Season + " ID:" + playerteams[x - 1].PlayerTeamID + "  IS in the database\r\n");
                    }
                    else
                    {
                        plyrtm.PlyrTmID = (playerteams.Count + 1).ToString().Trim();
                        plyrtm.PlayerTeamID = (playerteams.Count + 1).ToString().Trim();

                        //mtch.MtchID = (matches.Count + 1).ToString().Trim();
                        //mtch.MatchID = (matches.Count + 1).ToString().Trim();

                        fm.AddData("\r\n" + "Player/Team:: PLAYER: " + plyrtm.Playr.LastName + " TEAM: " + plyrtm.Tm.TeamName + " COMPETITION: " + plyrtm.Cmpttn.CompetitionName + " SEASON: " + plyrtm.Cmpttn.Season + "  IS NOT in the database");
                        try
                        {
                            // Read the item to see if it exists.  
                            ItemResponse<PlayerTeam> plyrtmResponse = await this.container.ReadItemAsync<PlayerTeam>(plyrtm.PlyrTmID, new Microsoft.Azure.Cosmos.PartitionKey(plyrtm.PlayerTeamID));
                            fm.AddData("\r\n" + "Item in PlayerTeam container with id: " + plyrtmResponse.Resource.PlyrTmID + " already exists\r\n");
                        }
                        catch (CosmosException ex) when (ex.StatusCode == HttpStatusCode.NotFound)
                        {
                            // Create an item in the container representing the Andersen family. Note we provide the value of the partition key for this item, which is "Andersen"
                            ItemResponse<PlayerTeam> plyrtmResponse = await this.container.CreateItemAsync<PlayerTeam>(plyrtm, new Microsoft.Azure.Cosmos.PartitionKey(plyrtm.PlayerTeamID));
                            // Note that after creating the item, we can access the body of the item with the Resource property off the ItemResponse. We can also access the RequestCharge property to see the amount of RUs consumed on this request.
                            //fm.AddData("\r\n" + "Created item in database with id: " + iItemCount.ToString().Trim() + " Operation consumed " + plyrResponse.RequestCharge + " RUs.\r\n");
                            fm.AddData("\r\n" + "Created item in PlayerTeam container with id: " + plyrtmResponse.Resource.PlyrTmID + " Operation consumed " + plyrtmResponse.RequestCharge + " RUs.\r\n");
                        }
                        playerteams.Add(plyrtm);
                    }
                }
                else
                {
                    fm.AddData("\r\n" + "This Match has no Data");
                }
            }
        }

        */

        //Archived 7-3-22
        /*
        public class ExcelPerformance : ExcelWorkbook
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
                //DateTime date = DateTime.ParseExact(sDate, "dd/MM/yyyy HH:mm:ss", CultureInfo.InvariantCulture);
                //plyrCurrent.DOB = date;



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
                        ResultsBatsman resbat = new ResultsBatsman();
                        sData = xlRange.Cells[iRow - 4, iCol].Value.ToString().Trim();
                        double dColor = xlRange.Cells[iRow - 4, iCol].Font.Color;
                        sColor = GetColorType(dColor);
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
                        //prf.BatsmanRes = resbat;
                    }
                    aSplitString1 = GetGameNo(iRow, iCol, gbl.SheetType.sBATSMEN);
                    if (aSplitString1 != null)
                    {
                        sOppTeamCode = aSplitString1[1].ToString().Trim();
                        //prf.TeamOppositionAbbrv = sOppTeamCode;
                        //prf.TeamOpposition = TeamShortToLong(sOppTeamCode);
                        xlWorksheet = xlWorkbook.Sheets[gbl.SheetType.sBATSMEN];
                        xlRange = xlWorksheet.UsedRange;
                        sTeamCode = xlRange.Cells[iRow, 1].Value.ToString().Trim();
                        aSplitString = sTeamCode.Split('.');
                        sTeamCode = aSplitString[0];
                        sTeam = TeamShortToLong(sTeamCode);
                        ssn.TeamMineAbbrv = sTeamCode;
                    }
                }
                //prf.MatchID = aSplitString1[0].ToString().Trim();
                //prf.TeamMine = sTeam;
                //prf.GameNumber = sGameNo;
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
                                if (sGameNo == mtch.GameNumber)
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
                iFirstRow = (iPlyrNo * 6) + 2;
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
        }
        
    }
}
        */
        /*
using CefSharp;
using ControlzEx.Standard;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows;

namespace Cricket
    {
        //[DllImport("user32.dll")]
        public class MoveWindow
        {
            //you also need to define this struct
            public struct WINDOWINFO
            {
                public uint cbSize;
                public RECT rcWindow; //holds the coords of the window
                public RECT rcClient;
                public uint dwStyle;
                public uint dwExStyle;
                public uint dwWindowStatus;
                public uint cxWindowBorders;
                public uint cyWindowBorders;
                public ushort atomWindowType;
                public ushort wCreatorVersion;
            }

            public struct RECT
            {
                public int Left;    // Specifies the x-coordinate of the upper-left corner of the rectangle. 
                public int Top;        // Specifies the y-coordinate of the upper-left corner of the rectangle. 
                public int Right;    // Specifies the x-coordinate of the lower-right corner of the rectangle.
                public int Bottom;    // Specifies the y-coordinate of the lower-right corner of the rectangle. 

            }
            /*
                    BOOL SetWindowPos(
                HWND hWnd,
                HWND hWndInsertAfter,
                int X,
                int Y,
                int cx,
                int cy,
                UINT uFlags
            );

                    public int NextScreen(int curr, int max)
                    {
                        if (curr < max)
                            curr++;
                        else
                            curr = 0;
                        return curr;
                    }
                    public Point MoveToNextScreen(Point OrigLocation, int OrigScreen, int NumOfScreens)
                    {

                        OrigLocation.X -= System.Windows.Forms.Screen.AllScreens[OrigScreen].Bounds.Location.X;
                        OrigLocation.X += System.Windows.Forms.Screen.AllScreens[NextScreen(OrigScreen, NumOfScreens)].Bounds.Location.X;
                        OrigLocation.Y -= System.Windows.Forms.Screen.AllScreens[OrigScreen].Bounds.Location.Y;
                        OrigLocation.Y += System.Windows.Forms.Screen.AllScreens[NextScreen(OrigScreen, NumOfScreens)].Bounds.Location.Y;

                        return OrigLocation;
                    }
                    /*
                    public void MoveWindow(Process p)
                    {
                        Point WindowSize = new Point();
                        Point WindowLocation = new Point();
                        int max = 4;
                        double scrnNum;
                        IntPtr h1 = p.MainWindowHandle;
                        System.Windows.Forms.Screen scrn = System.Windows.Forms.Screen.FromHandle(h1);
                        WINDOWINFO winfo = new WINDOWINFO();

                        GetWindowInfo(h1, ref winfo);
                        WindowSize.X = winfo.cxWindowBorders;
                        WindowSize.Y = winfo.cyWindowBorders;
                        WindowLocation.X = winfo.rcWindow.Left;
                        WindowLocation.Y = winfo.rcWindow.Right;
                        scrnNum = WindowLocation.X % scrn.Bounds.Width;

                        WindowLocation = MoveToNextScreen(WindowLocation, Convert.ToInt32(scrnNum), max);


                        SetWindowPos(h1, -1, WindowLocation.X, WindowLocation.Y, WindowSize.X, WindowSize.Y, SWP_NOZORDER | SWP_SHOWWINDOW);
                    }


            public void MoveWin()
            {
                string msg = "";
                int monId = 1;
                foreach (System.Windows.Forms.Screen screen in System.Windows.Forms.Screen.AllScreens)
                {
                    string str = String.Format("Monitor {0}: {1} x {2} @ {3},{4}\n", monId, screen.Bounds.Width,
                        screen.Bounds.Height, screen.Bounds.X, screen.Bounds.Y);
                    msg += str;
                    monId++;
                }

                MessageBox.Show(msg, "EnumDisp");

            }
        }
    }

2022-04-09
        public void GetContentOfNode()
        {
            XmlDocument xmldoc = new XmlDocument();
            //XmlNodeList xmlnode;

            //TestXML2();

            WebClient WbClnt = new WebClient();
            string sBaseURLSourceCode = WbClnt.DownloadString(sBaseURL);

            string sBaseURLxmlString = GetBaseURLxmlString(sBaseURLSourceCode);

            //FileStream fsBaseURLSourceCode = new FileStream(sBaseURLxmlString, FileMode.Open, FileAccess.Read);

            // convert string to stream
            //byte[] byteArray = Encoding.UTF8.GetBytes(sBaseURLxmlString);
            //byte[] byteArray = Encoding.ASCII.GetBytes(contents);
            //MemoryStream stream = new MemoryStream(byteArray);



            //FileStream fs = new FileStream("C:\\zBrendan\\Cricket\\WebScraping\\BaseUrl-String.xml", FileMode.Open, FileAccess.Read);
            //FileStream fs = new FileStream("C:\\zBrendan\\Cricket\\WebScraping\\ScoreCardString.xml", FileMode.Open, FileAccess.Read);
            FileStream fs = new FileStream("C:\\zBrendan\\Cricket\\WebScraping\\ball-by-ball-commentary-String.xml", FileMode.Open, FileAccess.Read);
            //FileStream fs = new FileStream("C:\\zBrendan\\Cricket\\WebScraping\\match-overs-comparison-String.xml", FileMode.Open, FileAccess.Read);

            xmldoc.Load(fs);


            XmlNode root = xmldoc.FirstChild;
            string sText = root.InnerText;
            string sXML = root.InnerXml;

            //string jsonText = JsonConvert.SerializeXmlNode(sText);
            //string sJSON = JsonConvert.SerializeObject(sXML, Newtonsoft.Json.Formatting.Indented);
            //string jsonString = Newtonsoft.Json.JsonConvert.SerializeObject(sXML, Newtonsoft.Json.Formatting.None);




            var jsonScoreCard = Newtonsoft.Json.JsonConvert.DeserializeObject(sXML);

            string indentedJsonString = JsonConvert.SerializeObject(jsonScoreCard, Newtonsoft.Json.Formatting.Indented);
            string JsonString = JsonConvert.SerializeObject(jsonScoreCard, Newtonsoft.Json.Formatting.None);
            Console.WriteLine(indentedJsonString);


            var jArray = JToken.Parse(JsonString); //It's actually a JArray, not a JObject

            var j1 = JObject.Parse(JsonString);

            object s1 = j1["Root"];


            var jTitle = jArray.SelectToken("$Count");
            var title = (string)jTitle;



            //JObject JSONObject;

            JObject obj = JObject.Parse(JsonString);



            JToken sec = obj["props"]["appPageProps"]["data"]["content"]["inningsScore"][0]["inningBatsmen"][0]["runs"];
            //JToken sec = obj["results"]["records"];

            //JToken jtokBatsman = obj["props"]["appPageProps"]["data"]["content"]["inningsScore"][0]["inningBatsmen"][0]["runs"];
            string s4 = sec.Value<string>();


            //var outerJObject = JObject.Parse(JsonString);
            //var dataJson = (string)outerJObject["props"];
            //var dataJObject = JObject.Parse(dataJson);
            //var userName = (string)dataJObject["null"];

            //string sJSONPath = @"$[0]";

            //var sTest = jsonScoreCard.SelectToken("Manufacturers[0].Name");


            //var jsonScoreCard1 = Newtonsoft.Json.JsonConvert.SerializeObject(xmldoc);

            int x = 10;

            // To convert JSON text contained in string json into an XML node
            //XmlDocument doc = JsonConvert.DeserializeXmlNode(json);

            //Display the contents of the child nodes.
            //if (root.HasChildNodes)
            //{
            //    for (int i = 0; i < root.ChildNodes.Count; i++)
            //    {
            //        MessageBox.Show(root.ChildNodes[i].InnerText);
            //    }
            //}

            //xmlnode = xmldoc.GetElementsByTagName("Product");
            //for (i = 0; i <= xmlnode.Count - 1; i++)
            //{
            //    xmlnode[i].ChildNodes.Item(0).InnerText.Trim();
            //    str = xmlnode[i].ChildNodes.Item(0).InnerText.Trim() + "  " + xmlnode[i].ChildNodes.Item(1).InnerText.Trim() + "  " + xmlnode[i].ChildNodes.Item(2).InnerText.Trim();
            //    MessageBox.Show(str);
            //}
        }
*/

        /* 2022-04-15
        byte[] byteArray = Encoding.UTF8.GetBytes(sMatchString);
        MemoryStream sMatchStream = new MemoryStream(byteArray);
        xmlDocMatch.Load(sMatchStream);
        XmlNode root = xmlDocMatch.FirstChild;
        string sXML = root.InnerXml;
        object jsonBatsman = Newtonsoft.Json.JsonConvert.DeserializeObject(sXML);
        string jsonString = JsonConvert.SerializeObject(jsonBatsman, Newtonsoft.Json.Formatting.None);
        JObject obj = JObject.Parse(jsonString);


        JToken jtokPlayerName = obj["props"]["appPageProps"]["data"]["content"]["inningsScore"][0]["inningBatsmen"][0]["player"]["longName"];
        string sFullName = jtokPlayerName.Value<string>();
        jtokPlayerName = obj["props"]["appPageProps"]["data"]["content"]["inningsScore"][0]["inningBatsmen"][0]["player"]["dateOfBirth"]["year"];
        string sDOByear = jtokPlayerName.Value<string>();
        jtokPlayerName = obj["props"]["appPageProps"]["data"]["content"]["inningsScore"][0]["inningBatsmen"][0]["player"]["dateOfBirth"]["month"];
        string sDOBmonth = jtokPlayerName.Value<string>();
        jtokPlayerName = obj["props"]["appPageProps"]["data"]["content"]["inningsScore"][0]["inningBatsmen"][0]["player"]["dateOfBirth"]["date"];
        string sDOBday = jtokPlayerName.Value<string>();

        JToken jtokBatsman = obj["props"]["appPageProps"]["data"]["content"]["inningsScore"][0]["inningBatsmen"][0]["runs"];
        string sRuns = jtokBatsman.Value<string>();
        jtokBatsman = obj["props"]["appPageProps"]["data"]["content"]["inningsScore"][0]["inningBatsmen"][0]["balls"];
        string sBallsFaced = jtokBatsman.Value<string>();
        jtokBatsman = obj["props"]["appPageProps"]["data"]["content"]["inningsScore"][0]["inningBatsmen"][0]["fours"];
        string sFours = jtokBatsman.Value<string>();
        jtokBatsman = obj["props"]["appPageProps"]["data"]["content"]["inningsScore"][0]["inningBatsmen"][0]["sixes"];
        string sSixes = jtokBatsman.Value<string>();


        JToken jtokBowler = obj["props"]["appPageProps"]["data"]["content"]["inningsScore"][0]["inningBowlers"][0]["wickets"];
        string sWickets = jtokBowler.Value<string>();
        jtokBowler = obj["props"]["appPageProps"]["data"]["content"]["inningsScore"][0]["inningBowlers"][0]["conceded"];
        string sConceded = jtokBowler.Value<string>();
        jtokBowler = obj["props"]["appPageProps"]["data"]["content"]["inningsScore"][0]["inningBowlers"][0]["overs"];
        string sOvers = jtokBowler.Value<string>();
        jtokBowler = obj["props"]["appPageProps"]["data"]["content"]["inningsScore"][0]["inningBowlers"][0]["fours"];
        string sBowlerFours = jtokBowler.Value<string>();
        jtokBowler = obj["props"]["appPageProps"]["data"]["content"]["inningsScore"][0]["inningBowlers"][0]["sixes"];
        string sBowlerSixes = jtokBowler.Value<string>();


        string sMatchCompSource = WbClnt.DownloadString(ms.sMatchOversComparison);
        string sMatchCompString = GetMatchResultString(sMatchSource);
        byte[] byteArrayComp = Encoding.UTF8.GetBytes(sMatchCompString);
        MemoryStream sMatchCompStream = new MemoryStream(byteArrayComp);
        XmlDocument xmlDocMatchComp = new XmlDocument();
        xmlDocMatchComp.Load(sMatchCompStream);
        XmlNode rootComp = xmlDocMatchComp.FirstChild;
        string sXMLComp = rootComp.InnerXml;
        object jsonComp = Newtonsoft.Json.JsonConvert.DeserializeObject(sXMLComp);
        string jsonCompString = JsonConvert.SerializeObject(jsonComp, Newtonsoft.Json.Formatting.None);
        JObject objComp = JObject.Parse(jsonCompString);


        // Match Result - need to do 4's and 6's by adding up bowlers
        JToken jtokComp = objComp["props"]["appPageProps"]["data"]["content"]["inningsScore"][0]["wickets"];
        string sWktsDown = jtokComp.Value<string>();
        jtokComp = objComp["props"]["appPageProps"]["data"]["content"]["inningsScore"][0]["runs"];
        string sMatchScore = jtokComp.Value<string>();
        jtokComp = objComp["props"]["appPageProps"]["data"]["content"]["inningsScore"][0]["overs"];
        string sMatchOvers = jtokComp.Value<string>();


        jtokComp = objComp["props"]["appPageProps"]["data"]["content"];

        */

        /*
        FileStream fs = new FileStream("C:\\zBrendan\\Cricket\\WebScraping\\Players-Source.xml", FileMode.Open, FileAccess.Read);
        XmlDocument xmldoc = new XmlDocument();
        xmldoc.Load(fs);
        XmlNode root = xmldoc.FirstChild;
        string sText = root.InnerText;
        string sXML = root.InnerXml;

        //string jsonText = JsonConvert.SerializeXmlNode(sText);
        string sJSON = JsonConvert.SerializeObject(sXML, Newtonsoft.Json.Formatting.Indented);
        string jsonString = Newtonsoft.Json.JsonConvert.SerializeObject(sXML, Newtonsoft.Json.Formatting.None);
        JObject obj = JObject.Parse(jsonString);
        JToken jtokPlayerName = obj;
        */

        /*
                private MatchStrings GetMatchURLs(int iHighestSoFar, CricketForm fm)
                {
                    MatchStrings mstr = new MatchStrings();
                    WebClient WbClnt = new WebClient();
                    XmlDocument xmlDocBaseURL = new XmlDocument();
                    JToken jtokBase;
                    string sSlug = "";


                    string sBaseSource = WbClnt.DownloadString(sBaseURL);
                    string sBaseString = GetMatchResultString(sBaseSource);


                    byte[] byteArray = Encoding.UTF8.GetBytes(sBaseString);
                    MemoryStream sBaseStream = new MemoryStream(byteArray);
                    xmlDocBaseURL.Load(sBaseStream);

                    XmlNode rootBase = xmlDocBaseURL.FirstChild;
                    string sXMLbase = rootBase.InnerXml;
                    object jsonBase = Newtonsoft.Json.JsonConvert.DeserializeObject(sXMLbase);
                    string jsonBaseString = JsonConvert.SerializeObject(jsonBase, Newtonsoft.Json.Formatting.None);
                    JObject objBase = JObject.Parse(jsonBaseString);

                    bool bNotFound = true;
                    jtokBase = objBase["props"]["appPageProps"]["data"]["content"]["recentResults"][0]["slug"];
                    sSlug = jtokBase.Value<string>();
                    int iHighestMatchNo = GetMatchNo(sSlug);

                    jtokBase = objBase["props"]["appPageProps"]["data"]["content"]["recentResults"][9]["slug"];
                    sSlug = jtokBase.Value<string>();
                    int iLowestMatchNo = GetMatchNo(sSlug);

                    if (!(iHighestSoFar >= iLowestMatchNo - 1))
                    {
                        fm.AddData("\r\nError in GetMatchURLs: Last match recorded (" + iHighestSoFar.ToString() + ") is TWO OR MORE games LESS than the lowest game completed (" + iLowestMatchNo.ToString() + ")\r\n");
                        return null;
                    }

                    if (!(iHighestSoFar < iHighestMatchNo))
                    {
                        fm.AddData("\r\nError in GetMatchURLs: Last match recorded (" + iHighestSoFar.ToString() + ") is NOT less than the highest game completed (" + iHighestMatchNo.ToString() + ")\r\n");
                        return null;
                    }

                    int iCtr = 9;

                    int iCurrMatchNo = iLowestMatchNo;

                    string sCurrSlug = "";

                    while ((bNotFound) && (iCurrMatchNo <= iHighestMatchNo) && (iCtr >= 0))
                    //while ((bNotFound) && (iCurrMatchNo <= iHighestMatchResult) && (iCtr > 0))
                    {

                        //kolkata-knight-riders-vs-punjab-kings-8th-match

                        jtokBase = objBase["props"]["appPageProps"]["data"]["content"]["recentResults"][iCtr]["slug"];
                        sCurrSlug = jtokBase.Value<string>();
                        iCurrMatchNo = GetMatchNo(sCurrSlug);

                        if (iCurrMatchNo > iHighestSoFar)
                        {
                            bNotFound = false;
                        }
                        else
                        {
                            iCtr = iCtr - 1;
                        }
                    }
                    jtokBase = objBase["props"]["appPageProps"]["data"]["content"]["recentResults"][iCtr]["objectId"];
                    string sObjectID = jtokBase.Value<string>();

                    mstr.sBallByBall = sBaseURL + "/" + sCurrSlug + "-" + sObjectID + "/ball-by-ball-commentary";
                    mstr.sMatchOversComparison = sBaseURL + "/" + sCurrSlug + "-" + sObjectID + "/match-overs-comparison";
                    mstr.sFullScoreCard = sBaseURL + "/" + sCurrSlug + "-" + sObjectID + "/full-scorecard";
                    mstr.sMatchResults = sBaseURL + "/match-results";
                    mstr.iMatchNumber = iCurrMatchNo;
                    return mstr;
                }
        */

        /*
                private string GetTeamCodeFromComp(string sTeamName, string sComp)
                {
                    switch (sTeamName)
                    {

                        case "IPL":
                            return GetTeamCodeForIPL(sTeamName);
                        case "T20Blast-South":
                            return GetTeamCodeForT20BlastSouth(sTeamName);
                        case "T20Blast-North":
                            return GetTeamCodeForT20BlastNorth(sTeamName);
                        default:
                            MessageBox.Show("Error in GetTeamCodeFromComp: No Competition found for " + sComp);
                            return null;
                    }
                }

                private string GetTeamCodeForT20BlastNorth(string sTeamName)
                {
                    switch (sTeamName)
                    {
                        case "derbyshire":
                            return "der";
                        case "durham":
                            return "dur";
                        case "lancashire":
                            return "lan";
                        case "leicestershire":
                            return "lei";
                        case "northamptonshire":
                            return "nor";
                        case "nottinghamshire":
                            return "not";
                        case "warwickshire":
                            return "war";
                        case "worcestershire":
                            return "wor";
                        case "yorkshire":
                            return "yor";
                        default:
                            MessageBox.Show("Error in GetTeamCodeForT20BlastNorth: No TeamCode for " + sTeamName);
                            return null;
                    }
                }

                private string GetTeamCodeForT20BlastSouth(string sTeamName)
                {
                    switch (sTeamName)
                    {
                        case "essex":
                            return "ess";
                        case "glamorgan":
                            return "gla";
                        case "gloucestershire":
                            return "glo";
                        case "hampshire":
                            return "ham";
                        case "kent":
                            return "ken";
                        case "middlesex":
                            return "mid";
                        case "somerset":
                            return "som";
                        case "surrey":
                            return "sur";
                        case "sussex":
                            return "sus";
                        default:
                            MessageBox.Show("Error in GetTeamCodeForT20BlastSouth: No TeamCode for " + sTeamName);
                            return null;
                    }
                }

                private string GetTeamCodeForIPL(string sTeamName)
                {
                    switch (sTeamName)
                    {
                        // IPL
                        case "royal-challengers-bangalore":
                            return "ban";
                        case "chennai-super-kings":
                            return "che";
                        case "delhi-capitals":
                            return "del";
                        case "gujarat-titans":
                            return "guj";
                        case "sunrisers-hyderabad":
                            return "hyd";
                        case "kolkata-knight-riders":
                            return "kol";
                        case "lucknow-super-giants":
                            return "luc";
                        case "mumbai-indians":
                            return "mum";
                        case "punjab-kings":
                            return "pun";
                        case "rajasthan-royals":
                            return "raj";
                        default:
                            MessageBox.Show("Error in GetTeamCodeForIPL: No TeamCode for " + sTeamName);
                            return null;
                    }
                }
        */
/*
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
                        //TimeSpan a2 = tzi1.BaseUtcOffset + tsTimeDiifBrisMinusLocal;
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
                        //}
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
*/
    }
}
