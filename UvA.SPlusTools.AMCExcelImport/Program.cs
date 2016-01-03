using System;
using System.Xml;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.IO;
using System.Data;
using System.Data.OleDb;
using UvA.SPlusTools.Data;
using UvA.SPlusTools.Data.Entities;
using UvA.SPlusTools.Data.Tasks;
using System.Globalization;

namespace AMC_Excel_SyllabusPlus
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Welkom to AMC_Excel_To_SPlus\n");

            List<MyActivity> AllActivities = ReadExcel("CleanedInput_5052PRMM6Y.xls");
            AddAMCDataToSPlus(AllActivities, "5052PRMM6Y",
                              "Practicum Medische Microbiologie", "P3", 40, 10);

            Console.WriteLine("\nPress any key to exit\n");
            Console.ReadKey();
        }

        /* https://stackoverflow.com/questions/15828/reading-excel-files-from-c-sharp */
        static List<MyActivity> ReadExcel(string FileName)
        {
            Console.WriteLine("Now Running ReadExcel\n");

            // List of Activity objects
            List<MyActivity> AllActivities = new List<MyActivity>();

            // string TheDirectory = Directory.GetCurrentDirectory();
            string TheDirectory = "C:\\Users\\timohalbesma\\Desktop";

            // Read from Excel File
            var fileName = string.Format("{0}\\{1}",
                                         TheDirectory, FileName);

            if (!File.Exists(fileName))
            {
                Console.WriteLine("File not found: {0}", fileName);
                // Environment.Exit(1);
            }

            var connectionString = string.Format(
                    "Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};" +
                    "Extended Properties=Excel 12.0;", fileName);

            Console.WriteLine("Filename = {0}\nconnectionString = {1}\n",
                    fileName, connectionString);

            var WeekDict = new Dictionary<string, string>()
            {
                // FIXME: buggy ambiguous notatie!! Elke file is aangeleverd met andere notatie *_*
                // Per3, notatie in Excel is week in jaarkalender!
                { "1", "Aca wk 19 Kal wk 01" },
                { "2", "Aca wk 20 Kal wk 02" },
                { "3", "Aca wk 21 Kal wk 03" },
                { "4", "Aca wk 22 Kal wk 04" },
                // Per5, notatie in Excel is week in jaarkalender!
                { "13", "Aca wk 31 Kal wk 13" },
                { "14", "Aca wk 32 Kal wk 14" },
                { "15", "Aca wk 33 Kal wk 15" },
                { "16", "Aca wk 34 Kal wk 16" },
                { "17", "Aca wk 35 Kal wk 17" },
                // Per6, notatie in Excel is week in academisch jaar!
                { "40", "Aca wk 40 Kal wk 22" },
                { "41", "Aca wk 41 Kal wk 23" },
                { "42", "Aca wk 42 Kal wk 24" },
                { "43", "Aca wk 43 Kal wk 25" }
            };

            try
            {
                DataSet ds = new DataSet();

                // Open connection to Exel file, read data from Sheet1
                OleDbConnection conn = new OleDbConnection(connectionString);
                conn.Open();
                // NB `Blad1' might have to be altered if the datasheet is
                // named differently
                var adapter = new OleDbDataAdapter("SELECT * FROM [Blad1$]", connectionString);
                adapter.Fill(ds, "Sheet1Data");
                var rows = ds.Tables[0].Rows.Cast<DataRow>();

                Console.WriteLine("Data in Excel file");
                foreach (var row in rows)
                {
                    printRawExcelDataRow(row);

                    // Datum looks like 29-3-2016 00:00:00
                    // NB in Excel it could look like 29-Mar, 1-Apr, 22-Apr, etc depending on cell format options
                    string Datum = row["Date"].ToString().Split()[0];
                    DateTime CleanedDatum = DateTime.ParseExact(
                        Datum, "d-M-yyyy", CultureInfo.InvariantCulture);

                    string Week = row["Week"].ToString();
                    string CleanedWeek = WeekDict[Week];

                    // Somehow raw looks like 1-1-1974 hh:mm:ss. Is that the beginning on UNIX time plus four years? O_o
                    // NB in Excel it looks like 09:00-12:00, 13:00-16:00
                    // This could be different in other files, change if needed
                    // NB this could be different on a different machine too?
                    string CleanedStartTime = row["Time"].ToString().Split('-')[0];

                    // Clean ActivityType
                    string CleanedActivityType = row["ActivityType"].ToString();
                    // NB if other types are expection, add them
                    string[] ValidActivityTypes = { "HC", "WC", "Prac", "Tent", "ZS", "KC", "WG" };
                    List<string> searchFor = new List<string>(ValidActivityTypes);
                    bool IsValidActivityType = searchFor.Any
                        (word => CleanedActivityType.Contains(word));
                    if (!IsValidActivityType)
                    {
                        throw new InvalidDataException(
                            "Incorrect activity type specified in Type column");
                    }
                    if (CleanedActivityType == "HC")
                    {
                        // Because that's what S+ calls it.
                        CleanedActivityType = "H";
                    }

                    string CleanedGroupIdentifier = row["Group"].ToString();

                    // Clean Duration, which could be of decimal form
                    // (e.g. 0.75 hours --> 3 quarters of an hour).
                    float MyDurationFloat;
                    string RawDuration = row["Duration"].ToString().Split(' ')[0].Replace(',', '.');
                    // Console.WriteLine("Value in Duration column = {0}", RawDuration);
                    // Globalization because otherwise usage of '.' or ','
                    // matters and f****s up the TryParse output.
                    float.TryParse(RawDuration, NumberStyles.AllowDecimalPoint,
                        CultureInfo.InvariantCulture, out MyDurationFloat);
                    // Console.WriteLine("MyDurationFloat = {0}", MyDurationFloat);
                    int CleanedDuration = (int)(4 * MyDurationFloat);
                    // Console.WriteLine("Cleaned int of Duration column = {0}", CleanedDuration);
                    if (CleanedDuration == 0)
                    {
                        throw new InvalidDataException(
                            "Incorrect duration specified in Duration column");
                    }

                    string CleanedDescription = row["Description"].ToString();

                    // Clean Location, here specifically insert locations as known in S+
                    string CleanedLocatie = row["Location"].ToString();
                    if (CleanedLocatie == "L2-N")
                    {
                        CleanedLocatie = "Extern: FNWI: AMC L2-N";
                    }
                    else if (CleanedLocatie == "L2-242")
                    {
                        CleanedLocatie = "Extern: FNWI: AMC L2-242";
                    }
                    else
                    {
                        CleanedLocatie = "";
                    }
                    string CleanedStaff = row["Staff"].ToString();

                    string CleanedActivityCounter = row["ActivityCounter"].ToString();

                    // Add cleaned activity to Activity class
                    MyActivity CleanedActivity = new MyActivity
                    {
                        Date = CleanedDatum,
                        Week = CleanedWeek,
                        StartTime = CleanedStartTime,
                        ActivityType = CleanedActivityType,
                        ActivityCounter = CleanedActivityCounter,
                        GroupIdentifier = CleanedGroupIdentifier,
                        Duration = CleanedDuration,
                        Description = CleanedDescription,
                        Location = CleanedLocatie,
                        Staff = CleanedStaff
                    };

                    Console.Write("Clean:\t");
                    CleanedActivity.print();
                    Console.Write("\n");
                    // Console.ReadKey();

                    AllActivities.Add(CleanedActivity);

                }

                conn.Close();
                conn.Dispose();
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error reading Excel file: {0}", ex.Message);
                // Environment.Exit(1);
            }

            return AllActivities;
        }

        static void AddAMCDataToSPlus(List<MyActivity> AllActivities, string moduleName,
            string moduleDescription, string Period, int size, int NumberOfGroups)
        {
            Console.WriteLine("Now running AddAMCDataToSPlus\n");


            // College is the S+ institution. This connects to S+ standalone.
            // Can be set in COM menu in S+ ?
            var col = new College("Splus");

            // Just a quick test to see if we have a connection (should print 61000ish).
            Console.WriteLine("Activitycount = {0}", col.Activities.Count);
            bool test_connection = false;
            if (test_connection)
            {
                Console.WriteLine("Exiting after check if SPlus connection works");
                Console.ReadKey();
                Environment.Exit(0);
            }

            var mod = col.Modules.FindByName(moduleName);
            if (mod == null)
            {
                Console.WriteLine("Module does not exist, creating it!");
                mod = new Module(col);
                mod.Name = moduleName;
                Console.WriteLine("Name = {0}", mod.Name);
                mod.Description = moduleDescription;
                Console.WriteLine("Description = {0}", mod.Description);
                mod.Department = col.Departments.FindByName("UvA/FNWI/OWI/CoS");
                Console.WriteLine("Department = {0}", mod.Department.Name);
                // Mandatory-, and Optional Programma cannot be set.
                // No student sets generated and no student sets linked.
                // No usertext about DataNose page set, no tags added (like "FNWI_EldersGegeven")
            }

            // Dictionary like object counter[HC] --> 01, 02, 03. Also counter[ZS] --> 01, 02, 03, etc
            var Counter = new Dictionary<string, int>()
            {
                { "H", 0 },
                { "Tent", 0 },
                { "Prac", 0 },
                { "KC", 0 },
                { "WC", 0 },
                { "ZS", 0 },
                { "WG", 0 }
            };

            foreach (var xlsAct in AllActivities)
            {
                // Available data from Excel file
                // Datum, begintijd, eindtijd, type, duur, description, locatie, docent

                Counter[xlsAct.ActivityType] += 1;
                int myCounterInt = 0;
                if (xlsAct.ActivityCounter != "")
                {
                    int.TryParse(xlsAct.ActivityCounter, out myCounterInt);
                    Counter[xlsAct.ActivityType] = myCounterInt;
                }

                var CounterString = Counter[xlsAct.ActivityType].ToString().PadLeft(2, '0');

                string ActivityName = moduleName + "/" + xlsAct.ActivityType +
                    "/" + Period + "/" + CounterString + xlsAct.GroupIdentifier;

                TimeSpan StartTime;
                TimeSpan.TryParse(xlsAct.StartTime, out StartTime);

                var act = col.Activities.FindByName(ActivityName);
                if (act == null)
                {
                    Console.WriteLine("Activity '{0}' does not exist. Creating it.",
                        ActivityName);

                    act = new UvA.SPlusTools.Data.Entities.Activity(col);
                    act.Module = mod;
                    act.Name = ActivityName;
                    act.StartDate = xlsAct.Date;
                    act.SuggestedStartTime = StartTime;
                    act.ActivityType = col.ActivityTypes.FindByName(xlsAct.ActivityType);
                    act.Duration = xlsAct.Duration;
                    act.Description = xlsAct.Description;

                    var loc = col.Locations.FindByName(xlsAct.Location);
                    if (loc != null)
                    {
                        act.LocationRequirement.Resources.Add(loc);
                        act.LocationRequirement.Set(ResourceRequirementType.Preset);
                        Console.WriteLine("LocationRequirement added. Locations.name = {0}", loc.Name);
                    }

                    var doc = col.StaffMembers.FindByName(xlsAct.Staff);
                    if (doc != null)
                    {
                        act.StaffRequirement.Resources.Add(doc);
                        act.StaffRequirement.Set(ResourceRequirementType.Preset);
                        Console.WriteLine("Staffrequirement added. StaffMembers.name = {0}", doc.Name);
                    }

                    if (xlsAct.GroupIdentifier != "")
                    {
                        int PlannedSize = (int)(xlsAct.GroupIdentifier.Length *
                                (float)size / NumberOfGroups);
                        act.PlannedSize = PlannedSize;

                    }
                    else
                    {
                        act.PlannedSize = size;
                    }

                    act.Zone = col.Zones.FindByName("UvA-AMC");

                    act.NamedAvailability = col.AvailabilityPatterns.FindByName(xlsAct.Week);

                    var Suit = col.Suitabilities.FindByName("01 LocType/Onderwijs");
                    act.LocationSuitabilities.Add(Suit);
                    Suit = col.Suitabilities.FindByName("02 LocFacil/Beamer");
                    act.LocationSuitabilities.Add(Suit);
                    Suit = col.Suitabilities.FindByName("02 LocFacil/Bord");
                    act.LocationSuitabilities.Add(Suit);
                    act.SaveSuitabilities(ResourceType.Location);

                    act.Schedule();
                }
                else
                {
                    Console.WriteLine("Activity '{0}' exists.", ActivityName);
                }

            }
            return;

        }

        static void printRawExcelDataRow(DataRow row)
        {
            Console.Write("Raw:\t");
            Console.Write("Date = {0}, ", row["Date"].ToString());
            Console.Write("Week = {0}, ", row["Week"].ToString());
            Console.Write("Time = {0}, ", row["Time"].ToString());
            Console.Write("ActivityType = {0}, ", row["ActivityType"].ToString());
            Console.Write("GroupIdentifier = {0}, ", row["Group"].ToString());
            Console.Write("Duration = {0}, ", row["Duration"].ToString());
            Console.Write("Description = {0}, ", row["Description"].ToString());
            Console.Write("Location = {0}, ", row["Location"].ToString());
            Console.Write("Staff = {0}\n", row["Staff"].ToString());
        }

    }

    /* https://stackoverflow.com/questions/12148340/
     * how-do-i-read-excel-columns-a-b-and-c-into-a-c-sharp-data-structure-vectors
     */
    public class MyActivity
    {
        public DateTime Date { get; set; }
        public string Week { get; set; }
        public string StartTime { get; set; }
        public string ActivityType { get; set; }
        public string ActivityCounter { get; set; } // todelete ?
        public string GroupIdentifier { get; set; }
        public int Duration { get; set; }
        public string Description { get; set; }
        public string Location { get; set; }
        public string Staff { get; set; }

        public void print()
        {
            Console.Write("Date = {0}, ", this.Date.ToString());
            Console.Write("Week = {0}, ", this.Week);
            Console.Write("Starttime = {0}, ", this.StartTime);
            Console.Write("ActivityType = {0}, ", this.ActivityType);
            Console.Write("GroupIdentifier = {0}, ", this.GroupIdentifier);
            Console.Write("Duration = {0}, ", this.Duration);
            Console.Write("Description = {0}, ", this.Description);
            Console.Write("Location = {0}, ", this.Location);
            Console.Write("Staff = {0}\n", this.Staff);
        }
    }

}

