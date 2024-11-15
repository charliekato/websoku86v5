#define USBBOOT
using System.Net;
using System.Data.OleDb;
using Renci.SshNet;
using System.Windows.Forms;


namespace websoku86v5
{
    public partial class Form1 : Form
    {
        readonly static string myName = System.Diagnostics.Process.GetCurrentProcess().ProcessName;
#if USBBOOT
        readonly static string workDir = "..\\html\\";
        readonly static string iniFile =  myName + ".ini";
#else
        readonly static string workDir = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData) +
             "\\" + myName + "\\";
        readonly static string iniFile = workDir + myName + ".ini";
#endif
        string mdbFile = "", htmlPath = "",
                indexFile = "", prgResult = "", rankingFile = "", scoreFile = "",
                secKeyFile = "";
        string hostName = "", port = "22", userName = "";

        public Form1()
        {
            InitializeComponent();
            this.Width = 1500;
            this.Height = 1500;


            Misc.ReadIniFile(iniFile, ref mdbFile,
                ref htmlPath, ref indexFile, ref prgResult, ref rankingFile,
                ref scoreFile, ref hostName, ref port, ref userName, ref secKeyFile);
            txtBoxMDBFile.Text = mdbFile;
            txtBoxIndexFile.Text = indexFile;
            txtBoxPrgResult.Text = prgResult;
            txtBoxScoreFile.Text = scoreFile;
            txtBoxRanking.Text = rankingFile;
            txtBoxHtmlPath.Text = htmlPath;
            txtBoxHostName.Text = hostName;
            if (port == "")
            {
                port = "22";
            }
            txtBoxPort.Text = port;
            txtBoxUserName.Text = userName;
            txtBoxKeyFile.Text = secKeyFile;
            InitTimer();


        }

        public static System.Windows.Forms.Timer timer;

        private void btnBrowse_Click(object sender, EventArgs e)
        {
            var proc = new System.Diagnostics.Process();
            proc.StartInfo.FileName = workDir + indexFile;
            proc.StartInfo.UseShellExecute = true;
            proc.Start();
        }

        public static EventHandler ev1;
        private void InitTimer()
        {
            timer = new System.Windows.Forms.Timer();
            ev1 = new EventHandler(CreateRun);
            timer.Tick += ev1;
            timer.Enabled = false;

        }

        private void CreateRun(object sender, EventArgs ev)
        {
            Cursor.Current = Cursors.WaitCursor;
            Misc.WriteIniFile(iniFile, txtBoxMDBFile.Text,
                txtBoxHtmlPath.Text, txtBoxIndexFile.Text, txtBoxPrgResult.Text, txtBoxRanking.Text,
                txtBoxScoreFile.Text, txtBoxHostName.Text, txtBoxPort.Text,
                txtBoxUserName.Text, txtBoxKeyFile.Text);
            Html.CreateHTML(
                txtBoxMDBFile.Text,
                workDir,
                txtBoxIndexFile.Text,
                txtBoxRanking.Text,
                txtBoxPrgResult.Text,
                txtBoxScoreFile.Text,
                txtBoxHtmlPath.Text,
                txtBoxHostName.Text,
                Convert.ToInt32(txtBoxPort.Text),
                txtBoxUserName.Text,
                txtBoxKeyFile.Text,
                txtBoxYouTube.Text);
            Cursor.Current = Cursors.Default;

        }
        private void BtnRunClick(object sender, EventArgs e)
        {
            CreateRun(sender, e);
        }

        private void BtnAutoRunClick(object sender, EventArgs e)
        {
            int interval;
            try
            {
                if (txtBoxInterval.Text == "") interval = 300000;
                else interval = Convert.ToInt32(txtBoxInterval.Text) * 60000;
            }
            catch
            {
                interval = 300000;
            }
            if (interval == 0) interval = 300000;
            if (!timer.Enabled)
            {
                timer.Enabled = true;
                timer.Interval = interval;
                btnAutoRun.Text = "停止";
            }
            else
            {
                timer.Enabled = false;
                btnAutoRun.Text = "開始";
            }
        }

        private void btnQuit_Click(object sender, EventArgs e)
        {
            Misc.WriteIniFile(iniFile, txtBoxMDBFile.Text,
                txtBoxHtmlPath.Text, txtBoxIndexFile.Text, txtBoxPrgResult.Text, txtBoxRanking.Text,
                txtBoxScoreFile.Text, txtBoxHostName.Text, txtBoxPort.Text,
                txtBoxUserName.Text, txtBoxKeyFile.Text);
            this.Close();
        }
        private void btnMDBFile_Click(object sender, EventArgs e)
        {
            try
            {
                mdbFile = Misc.GetFilePath("MDBを選択してください", Path.GetDirectoryName(txtBoxMDBFile.Text),
                    "MDB File (*.mdb)|*.mdb"); /**placeH**/
            }
            catch
            {
                mdbFile = Misc.GetFilePath("MDBを選択してください", "C:\\", "MDB File (*.mdb)|*.mdb");
            }
            txtBoxMDBFile.Text = mdbFile;
        }
        private void btnKeyFile_Click(object sender, EventArgs e)
        {
            try
            {
                secKeyFile = Misc.GetFilePath("秘密鍵Fileを選択してください",
                    Path.GetDirectoryName(txtBoxKeyFile.Text), "秘密鍵ファイル(*.pem)|*.pem|すべてのファイル|*.*");
            }
            catch
            {
                secKeyFile = Misc.GetFilePath("秘密鍵Fileを選択してください",
                    "C:\\", "秘密鍵ファイル(*.pem)|*.pem");
            }
            txtBoxKeyFile.Text = secKeyFile;
        }

    }
    public class Result
    {
        public int uid;
        public int kumi;
        public int goalTime;
        public string lapString;
        public int swimmerID;
        public int[] rswimmer = new int[4];
        public int laneNo;
        public int reasonCode;
        public int rank;
    }
    public static class Html
    {
        static string thisYouTubeURL;
        public static void CreateHTML(
            string mdbFile,    // Seiko Result System database.
            string workDir,    // local PC work directory
            string indexFile,  //indexFile   is something like ossSpring2024.Html
            string rankingFile,//rankingFile is something like ossSpring2024r.Html
            string kanproFile, //kanproFile  is something like ossSpring2024p.Html
            string scoreFile,  //scoreFile   is something like ossSpring2024s.Html
            string htmlDir,    //htmlDir     is something like rFlash
            string hostName,
            int port,
            string userName,
            string keyFile,
            string youTubeURL)
        {
            string indexFilePath;
            string rankingFilePath;
            string kanproFilePath;
            string teamScoreFilePath;
            string distDir = htmlDir + "/";
            MDBInterface mdb2Html;
            thisYouTubeURL = youTubeURL;
            if (!File.Exists(mdbFile))
            {
                MessageBox.Show("MDB File (" + mdbFile + ")が見つかりません。");
                return;
            }
            if (!Directory.Exists(workDir))
            {
                MessageBox.Show("作業フォルダー(" + workDir + ")が見つかりません。");
                return;
            }
            mdb2Html = new MDBInterface(mdbFile);

            ///Call init_machin_specific_variables
            indexFilePath = distDir + indexFile;
            rankingFilePath = distDir + rankingFile;
            kanproFilePath = distDir + kanproFile; //from txtBox
            teamScoreFilePath = distDir + scoreFile;
            if (indexFile != string.Empty)
            {
                string srcFile = workDir + "\\" + indexFile;
                CreateIndexHTML(mdb2Html, srcFile, rankingFile, kanproFile);
                if (keyFile != "")
                    Misc.SendFile(srcFile, indexFilePath, hostName, port, userName, keyFile);
            }
            if (rankingFile != string.Empty)
            {
                string srcFile = workDir + "\\" + rankingFile;
                CreateRankingFile(mdb2Html, srcFile, indexFile, kanproFile);
                if (keyFile != "")
                    Misc.SendFile(srcFile, rankingFilePath, hostName, port, userName, keyFile);
            }

            if (kanproFile != string.Empty)
            {
                string srcFile = workDir + "\\" + kanproFile;
                CreateHTMLProgramFormat(mdb2Html, srcFile, indexFile, rankingFile);
                if (keyFile != "")
                    Misc.SendFile(srcFile, kanproFilePath, hostName, port, userName, keyFile);
            }
            if (teamScoreFilePath != string.Empty)
            {

                //read_score_rule();
                //gen_team_score_html(teamScoreFilePath);

            }

        }



        static void PrintShumoku(MDBInterface mdb, StreamWriter sw, int uid)
        {
            sw.WriteLine("<hr id=\"PRGH" + mdb.GetProgramNoFromUID(uid) + "\">");
            sw.WriteLine("<table width=\"95%\">");
            sw.WriteLine("  <tr>");
            sw.WriteLine("    <td>  No. " + mdb.GetProgramNoFromUID(uid) + "</td> <td>" +
                         mdb.GetGenderFromUID(uid) + "</td><td>" +
                         mdb.GetClassFromUID(uid) + "</td>" +
                         "<td align=\"right\">" + mdb.GetDistanceFromUID(uid) + "</td>" +
                         "<td align=\"left\">" + mdb.GetStyleFromUID(uid) + "</td>" +
                         "<td align=\"right\">" + mdb.GetPhaseFromUID(uid) + "&nbsp;&nbsp;");

            if (mdb.GameRecordAvailable)
            {
                sw.WriteLine("大会記録:" + Misc.TimeIntToStr(mdb.GetGameRecord(uid)));
            }
            sw.WriteLine("</td></tr></table>");
            sw.WriteLine("<hr>");
        }

        static void CreateRankingFile(MDBInterface mdb, string srcFile, string indexFile, string prgFile)
        {
            int uid;
            int prgNo;
            int position;
            int numberOfLap;
            string thisLap;
            string prevLap;
            string[] splitTime = new string[5];
            int splitCounter;
            int ithLap;

            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
            using (StreamWriter sw = new StreamWriter(srcFile, false, System.Text.Encoding.GetEncoding("shift_jis")))
            {
                PrintHTMLHead(mdb, sw, 2);

                for (prgNo = 1; prgNo <= mdb.MaxProgramNo; prgNo++)
                {
                    uid = mdb.GetUIDFromPrgNo(prgNo);
                    if (uid > 0)
                    {
                        numberOfLap = Misc.HowManyLapTimes(mdb.GetDistanceFromUID(uid));

                        if (mdb.RaceExist(uid))
                        {
                            PrintShumoku(mdb, sw, uid);
                            sw.WriteLine("<div align=\"right\"> <a href=\"" + prgFile + "#PRGH" + mdb.GetProgramNoFromUID(uid) + "\">レーン順の結果</a>&nbsp;");
                            sw.WriteLine("<a href=\"" + indexFile + "\">種目選択に戻る</a></div>");
                            sw.WriteLine("<br><br>");
                            sw.WriteLine("<table border=\"0\" width=\"100%\">");
                            sw.WriteLine("<tr><th align=\"left\" width=\"8%\">順位</th>" +
                                "<th align=\"left\" width=\"29%\">氏名</th>" +
                                "<th align=\"left\" width=\"34%\">チーム名</th>" +
                                "<th align=\"left\" width=\"17%\">タイム</th>" +
                                "<th align=\"left\" width=\"12%\">新記録</th></tr>");
                        }

                        List<int> records = new List<int>();
                        int numSwimmers = mdb.GetHowManySwimmers(uid);

                        for (position = 1; position <= numSwimmers; position++) ////2024/6/18 bug fix using HowMany...
                        {
                            records.Clear();
                            mdb.GetResultNo(ref records, uid, position);
                            //if (records.Count == 0) break; //<--!!bug
                            for (int rn = 0; rn < records.Count; rn++)
                            {
                                Result result = mdb.GetResult(records[rn]);
                                if (result.swimmerID > 0)
                                {

                                    if (Misc.IsDQorDNS(result.reasonCode))
                                    {
                                        sw.WriteLine("<tr><td valign=\"top\">    </td>");
                                    }
                                    else if (result.laneNo >= 50)
                                    {
                                        sw.WriteLine("<tr><td align=\"right\" valign=\"top\" style=\"padding-right: 2px\">補欠" +
                                            (result.laneNo - 49) + "</td>");
                                    }
                                    else
                                    {
                                        sw.WriteLine("<tr><td align=\"right\" valign=\"top\" style=\"padding-right: 10px\">" + position + "</td>");
                                    }

                                    if (Misc.IsRelay(mdb.GetStyleFromUID(uid)))
                                    {
                                        sw.WriteLine("<td valign=\"top\">" + HtmlName4Relay(mdb, result.rswimmer) + "</td>");
                                        sw.WriteLine("<td valign=\"top\">" + mdb.GetRelayTeamName(result.swimmerID) + "</td>");
                                    }
                                    else
                                    {
                                        sw.WriteLine("<td>" + mdb.GetSwimmerName(result.swimmerID) + "</td>");
                                        sw.WriteLine("<td valign=\"top\">" + mdb.GetTeamName(result.swimmerID) + "</td>");
                                    }

                                    sw.WriteLine("<td valign=\"top\">");

                                    if (result.reasonCode > 0)
                                    {
                                        sw.WriteLine(CONSTANTS.reason[result.reasonCode] + "</td></tr>");
                                    }

                                    string timeStr = Misc.TimeIntToStr(result.goalTime);
                                    if (timeStr != "")
                                    {
                                        sw.WriteLine(timeStr);
                                        sw.WriteLine("</td>");

                                        if (mdb.GameRecordAvailable)
                                        {
                                            if (mdb.GetGameRecord(uid) > result.goalTime)
                                            {
                                                sw.WriteLine("<td valign=\"top\">大会新</td>");
                                            }
                                        }

                                        sw.WriteLine("</tr>");

                                        if (numberOfLap > 1)
                                        {
                                            thisLap = Misc.ParseLap(result.lapString);
                                            prevLap = "";
                                            splitCounter = 1;

                                            for (ithLap = 1; ithLap <= numberOfLap; ithLap++)
                                            {
                                                if (ithLap % 4 == 1)
                                                {
                                                    sw.WriteLine("<tr> <td colspan=4 align=\"center\"> <div class=\"lap_container\">");
                                                }

                                                sw.WriteLine("<div class=\"lap_time\">" + thisLap + "</div>");

                                                if (ithLap == 1)
                                                {
                                                    splitTime[splitCounter] = "";
                                                }
                                                else
                                                {
                                                    splitTime[splitCounter] = "(" + Misc.TimeSubtract(thisLap, prevLap) + ")";
                                                }

                                                splitCounter++;

                                                prevLap = thisLap;
                                                thisLap = Misc.ParseLap("");

                                                if (ithLap % 4 == 0)
                                                {
                                                    sw.WriteLine("</div></td></tr>");
                                                    splitCounter = 1;
                                                    Misc.PrintSplitTime(sw, splitTime[1], splitTime[2], splitTime[3], splitTime[4]);
                                                }
                                            }

                                            if (ithLap % 4 == 3)
                                            {
                                                sw.WriteLine("<div class=\"lap_time\">  </div><div class=\"lap_time\"> </div></td></tr>");
                                                Misc.PrintSplitTime(sw, splitTime[1], splitTime[2], "", "");
                                            }
                                        }
                                    }
                                }
                            }
                        }
                        sw.WriteLine("</table><br><br>");
                    }
                }

                PrintTailAndClose(sw);
            }
        }
        static string HtmlName4Relay(MDBInterface mdb, int[] rswimmer)
        {
            return mdb.GetSwimmerName(rswimmer[0]) + "<br>"
                + mdb.GetSwimmerName(rswimmer[1]) + "<br>"
                + mdb.GetSwimmerName(rswimmer[2]) + "<br>"
                + mdb.GetSwimmerName(rswimmer[3]);

        }
        static void CreateHTMLProgramFormat(MDBInterface mdb, string srcFile, string indexFile, string rankingFile)
        {

            int maxProgramNo = mdb.MaxProgramNo; // You need to implement GetMaxProgramNo() function

            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
            using (StreamWriter writer = new StreamWriter(srcFile, false, System.Text.Encoding.GetEncoding("shift_jis")))
            {
                PrintHTMLHead(mdb, writer, 2);

                for (int prgNo = 1; prgNo <= maxProgramNo; prgNo++)
                {
                    int uid = mdb.GetUIDFromPrgNo(prgNo);

                    if (uid > 0)
                    {
                        PrintShumoku(mdb, writer, uid);
                        writer.WriteLine("<div align=\"right\"> <a href=\"" + rankingFile + "#PRGH" + prgNo + "\">ランキング</a>&nbsp;");
                        writer.WriteLine("<a href=\"" + indexFile + "\">種目選択に戻る</a></div>");
                        PrintRaceResult(mdb, writer, uid);
                    }
                }

                PrintTailAndClose(writer);
            }
        }

        static void PrintRaceResult(MDBInterface mdb, StreamWriter writer, int uid)
        {

            int kumi = 0;

            List<Result> results = new List<Result>();
            mdb.GetResultList(uid, ref results);

            foreach (Result record in results)
            {
                {
                    if (kumi < record.kumi)
                    {
                        if (kumi > 0)
                        {
                            writer.WriteLine("</table>");
                        }

                        kumi = record.kumi;

                        writer.WriteLine("<div class=\"kumi\">" + kumi + "組</div>");
                        writer.WriteLine("<table border=\"0\" width=\"100%\">");
                    }

                    string laneStr = (record.laneNo >= 50) ? "補欠" + (record.laneNo - 49).ToString() : record.laneNo.ToString();

                    writer.WriteLine("<tr><td align=\"right\">" + laneStr + "</td>");

                    if (record.swimmerID > 0)
                    {
                        writer.WriteLine("<td align=\"left\">&nbsp;");

                        if (Misc.IsRelay(mdb.GetStyleFromUID(uid)))
                        {
                            writer.WriteLine(mdb.GetRelayTeamName(record.swimmerID) + "</td><td align=\"left\">" +
                                             mdb.GetSwimmerName(record.rswimmer[0]) + "&nbsp;&nbsp; " +
                                             mdb.GetSwimmerName(record.rswimmer[1]) + "<br>" +
                                             mdb.GetSwimmerName(record.rswimmer[2]) + "&nbsp; &nbsp;" +
                                             mdb.GetSwimmerName(record.rswimmer[3]) + "</td>");
                        }
                        else
                        {
                            writer.Write(mdb.GetSwimmerName(record.swimmerID) + "</td>" +
                                             "<td align=\"left\"> (" + mdb.GetTeamName(record.swimmerID) + ")</td>");
                        }

                        if (record.reasonCode == 0)
                            writer.WriteLine("<td align=\"right\">" + Misc.TimeIntToStr(record.goalTime) + "</td>");
                        else if (record.reasonCode == 4)
                            writer.WriteLine("<td align=\"right\">" + Misc.TimeIntToStr(record.goalTime) + "(op)</td>");
                        else
                            writer.WriteLine("<td align=\"right\">" + CONSTANTS.reason[record.reasonCode] + "</td>");


                        if (record.goalTime > 0)
                        {
                            if (mdb.GameRecordAvailable && mdb.GetGameRecord(uid) > record.goalTime)
                            {
                                writer.WriteLine("<td>大会新</td>");
                            }
                            if (mdb.GameRecordAvailable && mdb.GetGameRecord(uid) == record.goalTime)
                            {
                                writer.WriteLine("<td>大会タイ</td>");
                            }

                        }

                        writer.WriteLine("</tr>");
                    }
                    else
                    {
                        writer.WriteLine("<td>   </td><td> </td></tr>");
                    }
                }
            }

            writer.WriteLine("</table>");
        }
        static void PrintHTMLHead(MDBInterface mdb, StreamWriter writer, int fType)
        {
            writer.WriteLine("<?php");
            writer.WriteLine(" header(\"Content-Type: text/html; charset=Shift-JIS\");");
            writer.WriteLine("?>");
            writer.WriteLine("<!DOCTYPE Html><Html>");
            writer.WriteLine("<?php");
            writer.WriteLine(" $qarray = explode(\"&\", $_SERVER['QUERY_STRING']);");
            writer.WriteLine(" list($vname1,$value1) = explode(\"=\",$qarray[0]);");
            writer.WriteLine(" list($vname2,$value2) = explode(\"=\",$qarray[1]);");
            writer.WriteLine(" if (strcmp($vname1,\"prgNo\")==0) {");
            writer.WriteLine("     $prgNo=$value1;");
            writer.WriteLine("     $kumiNo=$value2;");
            writer.WriteLine(" } else {");
            writer.WriteLine("     $kumiNo=$value1;");
            writer.WriteLine("     $prgNo=$value2;");
            writer.WriteLine(" }");
            writer.WriteLine("?>");
            writer.WriteLine("<head> ");
            writer.WriteLine($"<meta charset=\"Shift_JIS\"><title>{mdb.GetEventName()} </title>");
            if (fType == 1)
            {
                writer.WriteLine("<link rel=\"stylesheet\" media=\"all\" href=\"css/cman.css\">");
                writer.WriteLine("<script type=\"text/javascript\" src=\"cman.js\"></script>");
            }
            if (fType == 2)
            {
                writer.WriteLine("<link rel=\"stylesheet\" media=\"all\" href=\"css/swim.css\">");
            }
            if (fType == 3)
            {
                writer.WriteLine("<link rel=\"stylesheet\" media=\"all\" href=\"css/swimcall.css\">");
            }
            writer.WriteLine("</head>");
            writer.WriteLine("<body>");
            if (fType < 3)
            {
                writer.WriteLine($"<h2>{mdb.GetEventName()} &nbsp;&nbsp;開催地 : {mdb.GetEventVenue()} &nbsp;&nbsp;期日 : {mdb.GetEventDate()}</h2>");
            }
            if (thisYouTubeURL != "")
            {
                writer.WriteLine($"<h1><a href=\"{thisYouTubeURL} \">YouTube ライブ配信はこちら</a></h1>");
            }

        }
        static void CreateIndexHTML(MDBInterface mdb, string myName, string rankingFile, string kanproFile)
        {
            int uid;
            int prgNo;
            int maxPrgNo;

            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
            using (StreamWriter writer = new StreamWriter(myName, false, System.Text.Encoding.GetEncoding("shift_jis")))
            {
                PrintHTMLHead(mdb, writer, 1);
                writer.WriteLine("<table id=\"sampleTable\" border=\"0\" width=\"95%\">");
                writer.WriteLine("<tr><th cmanFilterBtn>競技番号</th><th cmanFilterBtn>クラス</th>" +
                                 "<th cmanFilterBtn>性別</th><th cmanFilterBtn>距離</th>" +
                                 "<th cmanFilterBtn>種目</th><th cmanFilterBtn>予/決</th><th>  </th><th>  </th></tr>");
                maxPrgNo = mdb.MaxProgramNo;

                for (prgNo = 1; prgNo <= maxPrgNo; prgNo++)
                {
                    uid = mdb.GetUIDFromPrgNo(prgNo);
                    if (uid > 0)
                    {
                        writer.WriteLine($"<tr><td align=\"right\">{prgNo}</td>");
                        writer.WriteLine($"<td align=\"left\">{mdb.GetClassFromUID(uid)}</td>");
                        writer.WriteLine($"<td align=\"center\">{mdb.GetGenderFromUID(uid)}</td> ");
                        writer.WriteLine($"<td align=\"right\">{mdb.GetDistanceFromUID(uid)}</td>");
                        writer.WriteLine($"<td align=\"left\">{mdb.GetStyleFromUID(uid)}</td>");
                        writer.WriteLine($"<td align=\"left\">{mdb.GetPhaseFromUID(uid)}</td>");
                        writer.WriteLine($"<td> <a href=\"{kanproFile}#PRGH{prgNo}\"> レーン順</a></td>");
                        writer.WriteLine($"<td> <a href=\"{rankingFile}#PRGH{prgNo}\"> 順位表</a></td></tr>");
                    }
                }

                writer.WriteLine("</table>");
                PrintTailAndClose(writer);
            }
        }

        static void PrintTailAndClose(StreamWriter writer)
        {
            // Implement your logic for printing HTML tail and closing the writer
            writer.WriteLine("<br><br><br><br><br>");
            writer.WriteLine($"<div class=\"footer\" align=\"right\"> updated by {Dns.GetHostName()} at {DateTime.Now} </div>");
            writer.WriteLine("</body></Html>");
            writer.Close();
        }

    }



    public class MDBInterface
    {
        const int NORECORDYET = 0;
        private List<Result> resultList;



        public Result GetResult(int rn) { return resultList[rn]; }
        public void GetResultList(int uid, ref List<Result> extracted)
        {
            foreach (Result result in resultList)
            {
                if (result.uid == uid)
                {
                    extracted.Add(result);
                }
            }
        }
        private readonly string mdbFile;

        private readonly string[] ShumokuTable = new string[8];
        public string GetShumoku(int id) { return ShumokuTable[id]; }
        private readonly string[] DistanceTable = new string[8];

        private const string magicWord = "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=";
        private string[] genderStr;
        public bool GameRecordAvailable;
        private string[] swimmerName;
        public string GetSwimmerName(int id) { return swimmerName[id]; }
        private string[] kana;
        private int[] belongsTo;
        public string GetTeamName(int swimmerID) { return clubName[belongsTo[swimmerID]]; }
        private string[] clubName;
        private readonly int maxProgramNo = 0;
        public int MaxProgramNo
        {
            get { return maxProgramNo; }
        }
        private readonly int maxUID = 0;
        public int MaxUID
        {
            get { return maxUID; }
        }
        private int[] UIDFromProgramNo;
        int[] ClassNoByUID;
        int[] genderByUID;　　// 男子, 女子, 混合
        string[] styleByUID;
        string[] distanceByUID;
        string[] phaseByUID; // 予選/決勝/タイム決勝
        int[] gameRecord4UID;
        public int GetGameRecord(int uid) { return gameRecord4UID[uid]; }
        int[] programNo;
        int[] numSwimmers4UID;
        public int GetHowManySwimmers(int uid) { return numSwimmers4UID[uid]; }

        public int GetUIDFromPrgNo(int prgNo) { return UIDFromProgramNo[prgNo]; }
        public string GetClassFromUID(int uid) { return className[ClassNoByUID[uid]]; }
        public string GetGenderFromUID(int uid) { return genderStr[genderByUID[uid]]; }
        public string GetStyleFromUID(int uid) { return styleByUID[uid]; }
        public string GetDistanceFromUID(int uid) { return distanceByUID[uid]; }
        public string GetPhaseFromUID(int uid) { return phaseByUID[uid]; }
        public int GetProgramNoFromUID(int uid) { return programNo[uid]; }


        private readonly string[] TeamName4Relay;
        public string GetRelayTeamName(int id)
        {
            return TeamName4Relay[id];
        }
        public MDBInterface(string mdbFilePath)
        {
            InitTables();
            resultList = new List<Result>();
            mdbFile = mdbFilePath;
            InitProgramDB(ref maxUID, ref maxProgramNo);
            RedimProgramDBArrays(maxUID, maxProgramNo);
            InitClassDB();
            clubName = new string[10];
            InitClubName();
            RedimSwimmerDBArrays(GetNumSwimmers());
            TeamName4Relay = new string[GetNumRelayTeams() + 1];
            ReadMDB();
        }
        public void GetResultNo(ref List<int> recordNums, int uid, int rank)
        {
            for (int i = 0; i < resultList.Count; i++)
            {
                if ((resultList[i].uid == uid) && (resultList[i].rank == rank))
                {
                    recordNums.Add(i);
                }
            }
        }
        void ReadMDB()
        {
            ReadEventDB();
            ReadTeamDB();
            ReadClassDB();
            ReadSwimmerDB();
            ReadProgramDB();
            ReadResultDB();
            ReadRecordDB();

            AnalyzeResult();
        }
        void InitStyleTable()
        {
            ShumokuTable[0] = "";
            ShumokuTable[1] = "自由形";
            ShumokuTable[2] = "背泳ぎ";
            ShumokuTable[3] = "平泳ぎ";
            ShumokuTable[4] = "バタフライ";
            ShumokuTable[5] = "個人メドレー";
            ShumokuTable[6] = "リレー";
            ShumokuTable[7] = "メドレーリレー";
        }
        void InitDistanceTable()
        {
            DistanceTable[0] = "";
            DistanceTable[1] = "  25m";
            DistanceTable[2] = "  50m";
            DistanceTable[3] = " 100m";
            DistanceTable[4] = " 200m";
            DistanceTable[5] = " 400m";
            DistanceTable[6] = " 800m";
            DistanceTable[7] = "1500m";

        }

        int LocateDistanceID(string distance)
        {
            for (int cnt = 1; cnt < 8; cnt++)
            {
                if (DistanceTable[cnt] == distance) return cnt;
            }
            return 0; //error
        }
        void InitTables()
        {
            genderStr = new string[4] { "", "男子", "女子", "混合" };
            InitStyleTable();
            InitDistanceTable();
        }
        public bool RaceExist(int uid)
        {
            for (int cnt = 0; cnt < resultList.Count; cnt++)
            {
                if (resultList[cnt].uid == uid) return true;
            }
            return false;
        }
        string[] className;
        void InitClassDB()
        {

            int numClasses = 0;
            OleDbConnection conn = new OleDbConnection(magicWord + mdbFile);
            using (conn)
            {
                string myQuery = "SELECT MAX(番号) as maxClass FROM クラス;";

                OleDbCommand comm = new OleDbCommand(myQuery, conn);
                try
                {
                    conn.Open();
                    using (var dr = comm.ExecuteReader())
                    {
                        dr.Read();
                        numClasses = Misc.Obj2Int(dr["maxClass"]);
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
            className = new string[numClasses + 1];
        }


        void ReadRecordDB()
        {
            using (OleDbConnection connection = new OleDbConnection(magicWord + mdbFile))
            {
                try
                {

                    connection.Open();

                    string query = "SELECT プログラム.UID, プログラム.種目, プログラム.距離, 新記録.記録 FROM プログラム " +
                                   "INNER JOIN 新記録 ON プログラム.種目 = 新記録.種目 " +
                                   "AND プログラム.距離 = 新記録.距離 " +
                                   "AND プログラム.クラス番号 = 新記録.記録区分番号 " +
                                   "AND プログラム.性別 = 新記録.性別;";

                    using (OleDbCommand command = new OleDbCommand(query, connection))
                    {
                        using (OleDbDataReader reader = command.ExecuteReader())
                        {
                            if (reader.HasRows)
                            {
                                GameRecordAvailable = true;
                                while (reader.Read())
                                {
                                    int uid = Convert.ToInt32(reader["UID"]);
                                    object recordValue = reader["記録"];
                                    if (recordValue == DBNull.Value)
                                    {
                                        gameRecord4UID[uid] = NORECORDYET;
                                    }
                                    else
                                    {
                                        gameRecord4UID[uid] = Misc.TimeStrToInt(recordValue.ToString().Trim());
                                    }

                                }
                            }
                            else
                            {
                                GameRecordAvailable = false;
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }

            }
        }

        void ReadClassDB()
        {
            OleDbConnection conn = new OleDbConnection(magicWord + mdbFile);
            using (conn)
            {
                string myQuery = "SELECT 番号,クラス名称 FROM クラス;";

                OleDbCommand comm = new OleDbCommand(myQuery, conn);
                try
                {
                    conn.Open();
                    using (var dr = comm.ExecuteReader())
                    {
                        while (dr.Read())
                        {
                            className[Misc.Obj2Int(dr["番号"])] = Misc.Obj2String(dr["クラス名称"]).Trim();
                        }
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
        }

        void InitProgramDB(ref int maxUID, ref int maxProgramNo)
        {
            OleDbConnection conn = new OleDbConnection(magicWord + mdbFile);
            using (conn)
            {
                string myQuery = "SELECT MAX(競技番号) AS MAXPRGNO, MAX(UID) AS MAXUID FROM プログラム;";
                OleDbCommand comm = new OleDbCommand(myQuery, conn);
                try
                {
                    conn.Open();
                    using (var dr = comm.ExecuteReader())
                    {
                        dr.Read();
                        maxProgramNo = Misc.Obj2Int(dr["MAXPRGNO"]);
                        maxUID = Misc.Obj2Int(dr["MAXUID"]);
                    }
                }
                catch
                {
                    MessageBox.Show("指定されたMDB (" + mdbFile + ") が見つかりません。\n ");
                }
            }
        }
        void RedimProgramDBArrays(int maxUID, int maxPrgNo)
        {
            maxPrgNo++;
            maxUID++;
            UIDFromProgramNo = new int[maxPrgNo];
            genderByUID = new int[maxUID];
            styleByUID = new string[maxUID];
            distanceByUID = new string[maxUID];
            phaseByUID = new string[maxUID];
            ClassNoByUID = new int[maxUID];
            gameRecord4UID = new int[maxUID];
            programNo = new int[maxUID];
            numSwimmers4UID = new int[maxUID];
        }
        void ReadProgramDB()
        {
            int uid;
            int prgNo;

            OleDbConnection conn = new OleDbConnection(magicWord + mdbFile);
            using (conn)
            {
                string myQuery = "SELECT UID, 競技番号 , 種目, 距離, " +
                    "性別, 予決, クラス番号, Point FROM プログラム ;";
                OleDbCommand comm = new OleDbCommand(myQuery, conn);
                try
                {
                    conn.Open();
                    using (var dr = comm.ExecuteReader())
                    {
                        while (dr.Read())
                        {
                            uid = Misc.Obj2Int(dr["UID"]);
                            prgNo = Misc.Obj2Int(dr["競技番号"]);
                            programNo[uid] = prgNo;
                            UIDFromProgramNo[prgNo] = uid;
                            ClassNoByUID[uid] = Misc.Obj2Int(dr["クラス番号"]);
                            genderByUID[uid] = Misc.Obj2Int(dr["性別"]);
                            styleByUID[uid] = Misc.Obj2String(dr["種目"]);
                            distanceByUID[uid] = Misc.Obj2String(dr["距離"]);
                            phaseByUID[uid] = Misc.Obj2String(dr["予決"]);
                        }
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
        }
        private string eventName;
        private string eventDate;
        private string eventVenue;
        public string GetEventName() { return eventName; }
        public string GetEventDate() { return eventDate; }
        public string GetEventVenue() { return eventVenue; }
        void ReadEventDB()
        {
            OleDbConnection conn = new OleDbConnection(magicWord + mdbFile);
            using (conn)
            {
                string myQuery = "SELECT 大会名１,開催地,始期間,終期間 FROM 大会設定;";
                OleDbCommand comm = new OleDbCommand(myQuery, conn);
                try
                {
                    conn.Open();
                    using (var dr = comm.ExecuteReader())
                    {
                        dr.Read();
                        eventName = Misc.Obj2String(dr["大会名1"]);
                        if ((Misc.Obj2String(dr["始期間"])).Equals(Misc.Obj2String(dr["終期間"])))
                        {
                            eventDate = Misc.Obj2String(dr["始期間"]);
                        }
                        else
                        {
                            eventDate = Misc.Obj2String(dr["始期間"]) + "～" + Misc.Obj2String(dr["終期間"]);
                        }
                        eventVenue = Misc.Obj2String(dr["開催地"]);

                    }
                }
                catch (Exception ex) { MessageBox.Show(ex.Message); }
            }
        }

        int GetNumRelayTeams()
        {
            int numRelayTeams = 0;
            OleDbConnection conn = new OleDbConnection(magicWord + mdbFile);
            using (conn)
            {
                string myQuery = "SELECT MAX(チーム番号) AS MAXRTEAMNUM FROM チームマスター;";
                OleDbCommand comm = new OleDbCommand(myQuery, conn);

                try
                {
                    conn.Open();
                    using (var dr = comm.ExecuteReader())
                    {
                        dr.Read();
                        numRelayTeams = Misc.Obj2Int(dr["MAXRTEAMNUM"]);
                    }
                }
                catch (Exception ex) { MessageBox.Show(ex.Message); }
            }
            return numRelayTeams;
        }
        void ReadTeamDB()
        {
            OleDbConnection conn = new OleDbConnection(magicWord + mdbFile);
            using (conn)
            {
                string myQuery = "SELECT チーム番号,チーム名 FROM チームマスター;";
                OleDbCommand comm = new OleDbCommand(myQuery, conn);
                try
                {
                    conn.Open();
                    using (var dr = comm.ExecuteReader())
                    {
                        while (dr.Read())
                        {
                            TeamName4Relay[Misc.Obj2Int(dr["チーム番号"])] = Misc.Obj2String(dr["チーム名"]);
                        }
                    }
                }
                catch (Exception ex) { MessageBox.Show(ex.Message); }
            }
        }
        int GetNumSwimmers()
        {
            int numSwimmers = 0;
            OleDbConnection conn = new OleDbConnection(magicWord + mdbFile);
            using (conn)
            {
                string myQuery = "SELECT MAX(選手番号) AS MAXSNUM FROM 選手マスター;";
                OleDbCommand comm = new OleDbCommand(myQuery, conn);

                try
                {
                    conn.Open();
                    using (var dr = comm.ExecuteReader())
                    {
                        dr.Read();
                        numSwimmers = Misc.Obj2Int(dr["MAXSNUM"]);
                    }
                }
                catch (Exception ex) { MessageBox.Show(ex.Message); }
            }
            return numSwimmers;
        }

        void RedimSwimmerDBArrays(int maxnum)
        {
            maxnum++;
            swimmerName = new string[maxnum];
            kana = new string[maxnum];
            belongsTo = new int[maxnum];
        }
        int numTeams = 0;
        int LocateTeamID(string teamName)
        {
            int team_id;
            for (team_id = 1; team_id <= numTeams; team_id++)
            {
                if (clubName[team_id] == teamName) return team_id;
            }
            numTeams = team_id;
            throw new ArgumentException("error in LocateTeamID. InitCoubName has a kind of bug.");

        }
        int InitClubName()
        {
            OleDbConnection conn = new OleDbConnection(magicWord + mdbFile);
            using (conn)
            {
                string myQuery = "SELECT DISTINCT 所属名称１ AS CLUBNAME FROM 選手マスター;";
                OleDbCommand comm = new OleDbCommand(myQuery, conn);
                try
                {
                    conn.Open();
                    using (var dr = comm.ExecuteReader())
                    {
                        while (dr.Read())
                        {
                            numTeams++;
                            if (numTeams >= clubName.Length) Array.Resize(ref clubName, numTeams + 1);
                            clubName[numTeams] = dr.GetString(0).Trim();
                        }
                    }
                }
                catch (Exception ex) { MessageBox.Show(ex.Message); }
            }
            return (numTeams);
        }
        void ReadSwimmerDB()
        {
            OleDbConnection conn = new OleDbConnection(magicWord + mdbFile);
            using (conn)
            {
                int clubNo;
                int swimmerID = 0;
                string myQuery = "SELECT 選手番号, カナ, 氏名, 所属１, 所属名称１ FROM 選手マスター;";
                OleDbCommand comm = new OleDbCommand(myQuery, conn);
                try
                {
                    conn.Open();
                    using (var dr = comm.ExecuteReader())
                    {
                        while (dr.Read())
                        {
                            clubNo = LocateTeamID(Misc.Obj2String(dr["所属名称1"]).Trim());
                            swimmerID = Misc.Obj2Int(dr["選手番号"]);
                            swimmerName[swimmerID] =
                                Misc.Obj2String(dr["氏名"]).TrimEnd();
                            kana[swimmerID] = Misc.Obj2String(dr["カナ"]).TrimEnd();
                            belongsTo[swimmerID] = clubNo;
                        }
                    }
                }
                catch (Exception ex) { MessageBox.Show(ex.Message); }
            }
        }

        class Swimmer
        {
            public int swimmerID { get; set; }
            public int goalTime { get; set; }
            public string lapStr { get; set; }
            public int rank { get; set; }
        }




        void ReadResultDB()
        {
            OleDbConnection conn = new OleDbConnection(magicWord + mdbFile);
            using (conn)
            {
                Result result;
                string myQuery = "SELECT UID,ゴール, 選手番号,  " +
                  "第１泳者, 第２泳者, 第３泳者, 第４泳者  " +
                  ", 組, 水路, 新記録印刷マーク, 事由入力ステータス, " +
                  "ラップ１, ラップ２, ラップ３ " +
                  "FROM 記録マスター ORDER BY UID, 組, 水路; ";
                OleDbCommand comm = new OleDbCommand(myQuery, conn);
                try
                {
                    conn.Open();
                    using (var dr = comm.ExecuteReader())
                    {
                        while (dr.Read())
                        {
                            result = new Result();
                            result.uid = Misc.Obj2Int(dr["UID"]);
                            numSwimmers4UID[result.uid]++;
                            result.kumi = Misc.Obj2Int(dr["組"]);
                            result.lapString = Misc.RecordConcatenate(dr["ラップ１"], dr["ラップ" +
                                "２"], dr["ラップ３"]);
                            result.swimmerID = Misc.Obj2Int(dr["選手番号"]);
                            result.rswimmer[0] = Misc.Obj2Int(dr["第１泳者"]);
                            result.rswimmer[1] = Misc.Obj2Int(dr["第２泳者"]);
                            result.rswimmer[2] = Misc.Obj2Int(dr["第３泳者"]);
                            result.rswimmer[3] = Misc.Obj2Int(dr["第４泳者"]);
                            result.laneNo = Misc.Obj2Int(dr["水路"]);
                            result.reasonCode = Misc.Obj2Int(dr["事由入力ステータス"]);
                            if ((result.reasonCode == 0) || (result.reasonCode == 4))
                            {
                                result.goalTime = Misc.TimeStrToInt(Misc.Obj2String(dr["ゴール"]));
                            }
                            if (result.reasonCode == CONSTANTS.DQ)
                            {
                                result.goalTime = CONSTANTS.TIME4DQ;
                            }
                            if (result.reasonCode == CONSTANTS.DNS)
                            {
                                result.goalTime = CONSTANTS.TIME4DNS;
                            }
                            if (result.reasonCode == CONSTANTS.DNSM)
                            {
                                result.goalTime = CONSTANTS.TIME4DNSM;
                            }
                            result.rank = 1;
                            resultList.Add(result);
                        }
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
        }


        void AnalyzeResult()
        {
            int uid;
            for (uid = 1; uid <= MaxUID; uid++)
            {
                int myStart = 1, myEnd = resultList.Count, flag;
                flag = 0;
                for (int j = 0; j < resultList.Count; j++)
                {
                    if (flag == 0)
                    {
                        if (resultList[j].uid == uid)
                        {
                            myStart = j;
                            flag = 1;
                        }
                    }
                    if (resultList[j].uid > uid)
                    {
                        myEnd = j;
                        break;
                    }
                }
                if (flag == 1)
                {
                    for (int j = myStart; j < myEnd; j++)
                    {
                        int myTime = resultList[j].goalTime;
                        for (int k = myStart; k < myEnd; k++)
                        {
                            if (resultList[k].goalTime < myTime) resultList[j].rank++;
                        }
                    }
                }
            }
        }

    }

    public static class Misc
    {
        static string orgstrSave = string.Empty;
        const int ELEMENTLENGTH = 18;

        public static string ParseLap(string orgstr)
        {
            if (!string.IsNullOrEmpty(orgstr))
            {
                orgstrSave = orgstr;
            }

            string parsedLap = orgstrSave.Substring(0, Math.Min(ELEMENTLENGTH, orgstrSave.Length)).Trim();
            orgstrSave = orgstrSave.Substring(Math.Min(ELEMENTLENGTH, orgstrSave.Length));

            return parsedLap;
        }

        public static bool IsDQorDNS(int code) { return (code >= 1); }
        public static string TimeSubtract(string time1, string time2)
        {
            long intTime1 = Str2Milliseconds(time1);
            long intTime2 = Str2Milliseconds(time2);
            long answer = intTime1 - intTime2;
            return Milliseconds2Str(answer);
        }

        static string Milliseconds2Str(long milliseconds)
        {
            long myMinutes = milliseconds / 6000;
            string minutesStr = (myMinutes > 0) ? $"{myMinutes:D2}:" : "";
            long rest = milliseconds % 6000;
            string secondStr = $"{rest / 100:D2}";
            string millisecondStr = $"{rest % 100:D2}";
            return $"{minutesStr}{secondStr}.{millisecondStr}";
        }

        static long Str2Milliseconds(string myTime)
        {
            if (myTime == "") return 0;
            long myMinutes;
            int colonPos = myTime.IndexOf(":");
            if (colonPos > 0)
            {
                myMinutes = long.Parse(myTime.Substring(0, colonPos));
            }
            else
            {
                myMinutes = 0;
            }

            long millisecond = long.Parse(myTime.Substring(colonPos + 1, 5).Replace(".", ""));
            return 6000 * myMinutes + millisecond;
        }


        public static void PrintSplitTime(StreamWriter sw, string time1, string time2, string time3, string time4)
        {
            sw.WriteLine("<tr><td colspan=4 align=\"center\"><div class=\"lapcontainer\">");
            sw.WriteLine("<div class=\"lap_time\">" + time1 + "</div>");
            sw.WriteLine("<div class=\"lap_time\">" + time2 + "</div>");
            sw.WriteLine("<div class=\"lap_time\">" + time3 + "</div>");
            sw.WriteLine("<div class=\"lap_time\">" + time4 + "</div>");
            sw.WriteLine("</div></td></tr>");
        }
        public static bool IsRelay(string style)
        {
            return style.Contains("リレー");
        }
        public static string RecordConcatenate(object stra, object strb, object strc)
        {
            string concatenatedString = "";
            if (stra != null && stra != DBNull.Value)
            {
                concatenatedString = stra.ToString();
            }
            if (strb != null && strb != DBNull.Value)
            {
                concatenatedString += strb.ToString();
            }
            if (strc != null && strc != DBNull.Value)
            {
                concatenatedString += strc.ToString();
            }
            return concatenatedString;
        }
        public static int TimeStrToInt(string timeStr)
        {
            if (timeStr == "") return CONSTANTS.TIME4DQ;
            string workStr = timeStr.Replace(":", "");
            workStr = workStr.Replace(".", "");
            return Convert.ToInt32(workStr);
        }


        public static string TimeIntToStr(int mytime)
        {
            string minutes;
            string seconds;
            string centiSecond;
            int mytimeCP;

            if ((mytime == 0) || (mytime > 995999))
            {
                return "";
            }

            mytimeCP = mytime;

            minutes = Right(" " + Convert.ToString((int)(mytimeCP / 10000)), 2);
            mytimeCP = mytimeCP % 10000;
            seconds = Right("0" + Convert.ToString((int)(mytimeCP / 100)), 2);
            mytimeCP = mytimeCP % 100;
            centiSecond = Right("0" + Convert.ToString(mytimeCP), 2);

            if (mytime >= 10000)
            {
                return minutes + ":" + seconds + "." + centiSecond;
            }
            else
            {
                return "   " + seconds + "." + centiSecond;
            }
        }
        static string Right(string value, int length)
        {
            if (value.Length <= length)
            {
                return value;
            }
            else
            {
                return value.Substring(value.Length - length);
            }
        }
        public static void ReadIniFile(string filename,
            ref string mdbFilePath,
             ref string htmlFilePath,
            ref string indexFile, ref string kanproFile,
            ref string rankingFile, ref string scoreFile,
            ref string hostName, ref string port,
            ref string userName, ref string keyFile)
        {

            try
            {

                System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
                using (StreamReader reader = new StreamReader(filename, System.Text.Encoding.GetEncoding("shift_jis")))
                {
                    string line = "";
                    while ((line = reader.ReadLine()) != null)
                    {
                        if (line == "") continue;
                        if (line.Substring(0, 1) == "#") continue;
                        string[] words = line.Split('>');
                        if (words[0] == "mdbFilePath") mdbFilePath = words[1];
                        if (words[0] == "htmlFilePath") htmlFilePath = words[1];
                        if (words[0] == "indexFile") indexFile = words[1];
                        if (words[0] == "kanproFile") kanproFile = words[1];
                        if (words[0] == "rankingFile") rankingFile = words[1];
                        if (words[0] == "scoreFile") scoreFile = words[1];
                        if (words[0] == "hostName") hostName = words[1];
                        if (words[0] == "port") port = words[1];
                        if (words[0] == "userName") userName = words[1];
                        if (words[0] == "keyFile") keyFile = words[1];
                    }
                }
            }
            catch
            {
                MessageBox.Show("cannot find " + filename);
            }
        }
        public static void WriteIniFile(string filename,     // CreateWebReport.ini
                                          string mdbFilePath,  // C:Users\ykato\OneDrive\SwimDB\DB2024\Swim32.mdb 
                                          string htmlFilePath, // usually rFlash/xxxx
                                          string indexFile,    // xxxx.html
                                          string kanproFile,   // xxxxp.html
                                          string rankingFile,  // xxxxr.html
                                          string scoreFile,
                                          string hostName,
                                          string port,
                                          string userName,
                                          string keyFile)
        {
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
            using (StreamWriter sw = new StreamWriter(filename, false, System.Text.Encoding.GetEncoding("shift_jis")))
            {
                sw.WriteLine($"mdbFilePath>{mdbFilePath}");
                sw.WriteLine($"htmlFilePath>{htmlFilePath}");
                sw.WriteLine($"indexFile>{indexFile}");
                sw.WriteLine($"kanproFile>{kanproFile}");
                sw.WriteLine($"rankingFile>{rankingFile}");
                sw.WriteLine($"scoreFile>{scoreFile}");
                sw.WriteLine($"hostName>{hostName}");
                sw.WriteLine($"port>{port}");
                sw.WriteLine($"userName>{userName}");
                sw.WriteLine($"keyFile>{keyFile}");
            }
        }
        //
        /// <summary>
        /// SendFile sends source file to webserver.
        /// </summary>
        /// <param name="source">Source html file which is to be sent to the server(host).</param>
        /// <param name="destination">server directory in which source html file is stored</param>
        /// <param name="hostName">web server host name or ip address. Default: otsuswim.ddns.net</param>
        /// <param name="port">web server ssh port number (usually 22) </param>
        /// <param name="userName">user name of hte server. Default: www-data</param>
        /// <param name="keyFile">Secret file that corresponds to the server public key.</param>
        // 
        public static void SendFile(string source, string destination, string hostName, int port, string userName,
            string keyFile)
        {

            try
            {
                var connectionInfo = new PrivateKeyConnectionInfo(hostName, port, userName, new PrivateKeyFile(keyFile));

                FileInfo fi = new FileInfo(source);
                using (var client = new ScpClient(connectionInfo))
                {
                    client.Connect();

                    using (var fileStream = fi.OpenRead())
                    {
                        client.Upload(fileStream, destination);

                    }

                    client.Disconnect();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + "\n" + "Web serverにはsend しませんでした。");
            }
        }
        public static string Obj2String(object obj)
        {
            if (obj == DBNull.Value) return string.Empty;
            return (string)obj;
        }
        public static int Obj2Int(object obj)
        {
            if (obj != DBNull.Value) return Convert.ToInt32(obj);
            return 0;
        }
        public static int HowManyLapTimes(string distance)
        {
            switch (distance)
            {
                case "  25m": // 25m
                    return 0;
                case "  50m": // 50m
                    return 1;
                case " 100m": // 100m
                    return 2;
                case " 200m": // 200m
                    return 4;
                case " 400m": // 400m
                    return 8;
                case " 800m": // 800m
                    return 16;
                case "1500m": // 1500m
                    return 30;
                default:
                    return 0; // Default to 0 if distance is not recognized
            }
        }
        public static string GetFilePath(string title = "ファイルを選択してください",
                            string initFolder = @"C:\",
                            string option = "テキストファイル (*.txt)|*.txt|すべてのファイル (*.*)|*.*"
                            )
        {
            string selectedFilePath = "";
            OpenFileDialog openFileDialog = new OpenFileDialog();

            // ダイアログのタイトルを設定
            openFileDialog.Title = title;

            // 初期ディレクトリを設定（オプション）
            openFileDialog.InitialDirectory = initFolder;
            // initFolder=@"C:\"

            // ファイルの種類を指定（オプション）
            openFileDialog.Filter = option;

            // ユーザーが選択したファイルを取得
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                selectedFilePath = openFileDialog.FileName;
            }
            return selectedFilePath;
        }
        public static string GetFolder(string title = "フォルダーを選択してください",
                        string initFolder = @"C:\")
        {
            string selectedFolderPath = "";
            FolderBrowserDialog folderBrowserDialog = new FolderBrowserDialog();
            // ダイアログのタイトルを設定
            folderBrowserDialog.Description = title;
            // 初期フォルダーを設定（オプション）
            folderBrowserDialog.SelectedPath = initFolder;
            if (folderBrowserDialog.ShowDialog() == DialogResult.OK)
            {
                selectedFolderPath = folderBrowserDialog.SelectedPath;
            }
            return selectedFolderPath;
        }
    }
    public class CONSTANTS
    {
        public static readonly int TIME4DNS = 999998;
        public static readonly int TIME4DNSM = 999997;
        public static readonly int TIME4DQ = 999999;
        public static readonly int DNS = 1;
        public static readonly int DQ = 2;
        public static readonly int DNSM = 3;
        public static readonly string[] reason = new string[] { "", "棄権", "失格", "途中退水","オープン",
            "OP(失格)","OP(棄権)" };
    }
}
