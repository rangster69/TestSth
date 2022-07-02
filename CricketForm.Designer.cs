
namespace Cricket
{
    partial class CricketForm
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.tbMain = new System.Windows.Forms.TextBox();
            this.bCancel = new System.Windows.Forms.Button();
            this.bPerfBat = new System.Windows.Forms.Button();
            this.bMatches = new System.Windows.Forms.Button();
            this.DeletePerf = new System.Windows.Forms.Button();
            this.LoadPlyrs = new System.Windows.Forms.Button();
            this.bDeleteMtch = new System.Windows.Forms.Button();
            this.LoadMtch = new System.Windows.Forms.Button();
            this.LoadPlyrList = new System.Windows.Forms.Button();
            this.LoadPlayerListBatBowl = new System.Windows.Forms.Button();
            this.UpdateResults = new System.Windows.Forms.Button();
            this.bUpdateAllGames = new System.Windows.Forms.Button();
            this.bLoadSeason = new System.Windows.Forms.Button();
            this.bLoadFinals = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // tbMain
            // 
            this.tbMain.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.tbMain.Location = new System.Drawing.Point(-1, 0);
            this.tbMain.Multiline = true;
            this.tbMain.Name = "tbMain";
            this.tbMain.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.tbMain.Size = new System.Drawing.Size(799, 378);
            this.tbMain.TabIndex = 0;
            this.tbMain.TextChanged += new System.EventHandler(this.textBox1_TextChanged);
            // 
            // bCancel
            // 
            this.bCancel.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.bCancel.Location = new System.Drawing.Point(713, 415);
            this.bCancel.Name = "bCancel";
            this.bCancel.Size = new System.Drawing.Size(75, 23);
            this.bCancel.TabIndex = 2;
            this.bCancel.Text = "Cancel";
            this.bCancel.UseVisualStyleBackColor = true;
            this.bCancel.Click += new System.EventHandler(this.bCancel_Click);
            // 
            // bPerfBat
            // 
            this.bPerfBat.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.bPerfBat.Location = new System.Drawing.Point(375, 415);
            this.bPerfBat.Name = "bPerfBat";
            this.bPerfBat.Size = new System.Drawing.Size(105, 23);
            this.bPerfBat.TabIndex = 3;
            this.bPerfBat.Text = "Save SeasonPlyr";
            this.bPerfBat.UseVisualStyleBackColor = true;
            this.bPerfBat.Click += new System.EventHandler(this.bPerformance_Click);
            // 
            // bMatches
            // 
            this.bMatches.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.bMatches.Location = new System.Drawing.Point(12, 415);
            this.bMatches.Name = "bMatches";
            this.bMatches.Size = new System.Drawing.Size(75, 23);
            this.bMatches.TabIndex = 4;
            this.bMatches.Text = "Save Mtch";
            this.bMatches.UseVisualStyleBackColor = true;
            this.bMatches.Click += new System.EventHandler(this.bMatches_Click);
            // 
            // DeletePerf
            // 
            this.DeletePerf.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.DeletePerf.Location = new System.Drawing.Point(603, 415);
            this.DeletePerf.Name = "DeletePerf";
            this.DeletePerf.Size = new System.Drawing.Size(104, 23);
            this.DeletePerf.TabIndex = 5;
            this.DeletePerf.Text = "Delete SeasonPlyr";
            this.DeletePerf.UseVisualStyleBackColor = true;
            this.DeletePerf.Click += new System.EventHandler(this.DeletePerf_Click);
            // 
            // LoadPlyrs
            // 
            this.LoadPlyrs.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.LoadPlyrs.Location = new System.Drawing.Point(486, 415);
            this.LoadPlyrs.Name = "LoadPlyrs";
            this.LoadPlyrs.Size = new System.Drawing.Size(111, 23);
            this.LoadPlyrs.TabIndex = 6;
            this.LoadPlyrs.Text = "Load Plyr/PlyrStats";
            this.LoadPlyrs.UseVisualStyleBackColor = true;
            this.LoadPlyrs.Click += new System.EventHandler(this.LoadPlyrs_Click);
            // 
            // bDeleteMtch
            // 
            this.bDeleteMtch.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.bDeleteMtch.Location = new System.Drawing.Point(175, 415);
            this.bDeleteMtch.Name = "bDeleteMtch";
            this.bDeleteMtch.Size = new System.Drawing.Size(75, 23);
            this.bDeleteMtch.TabIndex = 10;
            this.bDeleteMtch.Text = "Delete Mtch";
            this.bDeleteMtch.UseVisualStyleBackColor = true;
            this.bDeleteMtch.Click += new System.EventHandler(this.bDeleteMtch_Click);
            // 
            // LoadMtch
            // 
            this.LoadMtch.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.LoadMtch.Location = new System.Drawing.Point(94, 416);
            this.LoadMtch.Name = "LoadMtch";
            this.LoadMtch.Size = new System.Drawing.Size(75, 23);
            this.LoadMtch.TabIndex = 11;
            this.LoadMtch.Text = "Load Mtch";
            this.LoadMtch.UseVisualStyleBackColor = true;
            this.LoadMtch.Click += new System.EventHandler(this.button1_Click);
            // 
            // LoadPlyrList
            // 
            this.LoadPlyrList.Location = new System.Drawing.Point(-100, -100);
            this.LoadPlyrList.Name = "LoadPlyrList";
            this.LoadPlyrList.Size = new System.Drawing.Size(113, 23);
            this.LoadPlyrList.TabIndex = 12;
            this.LoadPlyrList.Text = "Load BatBwl Plyrs";
            this.LoadPlyrList.UseVisualStyleBackColor = true;
            this.LoadPlyrList.Click += new System.EventHandler(this.LoadPlyrList_Click);
            // 
            // LoadPlayerListBatBowl
            // 
            this.LoadPlayerListBatBowl.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.LoadPlayerListBatBowl.Location = new System.Drawing.Point(256, 416);
            this.LoadPlayerListBatBowl.Name = "LoadPlayerListBatBowl";
            this.LoadPlayerListBatBowl.Size = new System.Drawing.Size(113, 23);
            this.LoadPlayerListBatBowl.TabIndex = 13;
            this.LoadPlayerListBatBowl.Text = "UpdatePlyrLstBatBowl";
            this.LoadPlayerListBatBowl.UseVisualStyleBackColor = true;
            this.LoadPlayerListBatBowl.Click += new System.EventHandler(this.LoadPlayerListBatBowl_Click);
            // 
            // UpdateResults
            // 
            this.UpdateResults.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.UpdateResults.Location = new System.Drawing.Point(12, 384);
            this.UpdateResults.Name = "UpdateResults";
            this.UpdateResults.Size = new System.Drawing.Size(104, 23);
            this.UpdateResults.TabIndex = 14;
            this.UpdateResults.Text = "UpdateOneGame";
            this.UpdateResults.UseVisualStyleBackColor = true;
            this.UpdateResults.Click += new System.EventHandler(this.UpdateResults_Click);
            // 
            // bUpdateAllGames
            // 
            this.bUpdateAllGames.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.bUpdateAllGames.Location = new System.Drawing.Point(123, 383);
            this.bUpdateAllGames.Name = "bUpdateAllGames";
            this.bUpdateAllGames.Size = new System.Drawing.Size(99, 23);
            this.bUpdateAllGames.TabIndex = 15;
            this.bUpdateAllGames.Text = "UpdateAllGames";
            this.bUpdateAllGames.UseVisualStyleBackColor = true;
            this.bUpdateAllGames.Click += new System.EventHandler(this.bUpdateAllGames_Click);
            // 
            // bLoadSeason
            // 
            this.bLoadSeason.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.bLoadSeason.Location = new System.Drawing.Point(228, 383);
            this.bLoadSeason.Name = "bLoadSeason";
            this.bLoadSeason.Size = new System.Drawing.Size(75, 23);
            this.bLoadSeason.TabIndex = 16;
            this.bLoadSeason.Text = "LoadSeason";
            this.bLoadSeason.UseVisualStyleBackColor = true;
            this.bLoadSeason.Click += new System.EventHandler(this.bLoadSeason_Click);
            // 
            // bLoadFinals
            // 
            this.bLoadFinals.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.bLoadFinals.Location = new System.Drawing.Point(309, 384);
            this.bLoadFinals.Name = "bLoadFinals";
            this.bLoadFinals.Size = new System.Drawing.Size(75, 23);
            this.bLoadFinals.TabIndex = 17;
            this.bLoadFinals.Text = "LoadFinals";
            this.bLoadFinals.UseVisualStyleBackColor = true;
            this.bLoadFinals.Click += new System.EventHandler(this.bLoadFinals_Click);
            // 
            // CricketForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(800, 450);
            this.Controls.Add(this.bLoadFinals);
            this.Controls.Add(this.bLoadSeason);
            this.Controls.Add(this.bUpdateAllGames);
            this.Controls.Add(this.UpdateResults);
            this.Controls.Add(this.LoadPlayerListBatBowl);
            this.Controls.Add(this.LoadPlyrList);
            this.Controls.Add(this.LoadMtch);
            this.Controls.Add(this.bDeleteMtch);
            this.Controls.Add(this.LoadPlyrs);
            this.Controls.Add(this.DeletePerf);
            this.Controls.Add(this.bMatches);
            this.Controls.Add(this.bPerfBat);
            this.Controls.Add(this.bCancel);
            this.Controls.Add(this.tbMain);
            this.Location = new System.Drawing.Point(2250, 150);
            this.Name = "CricketForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.Manual;
            this.Text = "Cricket Database";
            this.Load += new System.EventHandler(this.CricketForm_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TextBox tbMain;
        private System.Windows.Forms.Button bCancel;
        private System.Windows.Forms.Button bPerfBat;
        private System.Windows.Forms.Button bMatches;
        private System.Windows.Forms.Button DeletePerf;
        private System.Windows.Forms.Button LoadPlyrs;
        private System.Windows.Forms.Button bDeleteMtch;
        private System.Windows.Forms.Button LoadMtch;
        private System.Windows.Forms.Button LoadPlyrList;
        private System.Windows.Forms.Button LoadPlayerListBatBowl;
        private System.Windows.Forms.Button UpdateResults;
        private System.Windows.Forms.Button bUpdateAllGames;
        private System.Windows.Forms.Button bLoadSeason;
        private System.Windows.Forms.Button bLoadFinals;
    }
}

