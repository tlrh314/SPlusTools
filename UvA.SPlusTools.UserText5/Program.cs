using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.IO;
using UvA.SPlusTools.Data;
using UvA.SPlusTools.Data.Entities;
using UvA.SPlusTools.Data.Tasks;
using System.Globalization;
using System.Diagnostics;

namespace UvA.SPlusTools.UserText5
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Welcome to UserText5");

            // College is the S+ institution. This connects to S+ standalone.
            var col = new College("Splus"); // Splus

            // Just a quick test to see if we have a connection (should print 61000ish).
            Console.WriteLine("Activitycount = {0}", col.Activities.Count);
            bool DEBUG = true;
            if (DEBUG) {
                Console.WriteLine("Exiting after check if SPlus connection works");
                Console.ReadKey();
                Environment.Exit(0);
            }

            // Get all activities
            //int i = 0;
            var AllActivities = col.Activities;
            foreach (var act in AllActivities)
            {
                // Some departments happen to be empty. Prevent those from crashing the script.
                var dept = act.Department;
                string deptName = "";
                if (dept != null && dept.Name.Contains('/'))
                {
                    deptName = dept.Name.Split('/')[1];
                }
                else
                {
                    Console.WriteLine("Empty department detected. Continuing");
                }

                // Filter activities such that FNWI modules only
                if (deptName == "FNWI" && act.UserText5 != "")
                {
                        // For debugging purposes, only do first ten iterations.
                        //i = i + 1;
                        //if (i > 10)
                        //{
                        //    Console.WriteLine("Ten iterations passed: break!");
                        //    break;
                        //}

                        string UserText5 = act.UserText5;
                        string UserText4 = act.UserText4;

                        Console.WriteLine("Activity = {0}\nUserText5 = {1}\nUserText4 = {2}", act.Name, UserText5, UserText4);

                        if (UserText4 == "")
                        {
                            Console.WriteLine("Success: inserted UT5 into UT4, emptied UT5\n");
                            act.UserText4 = UserText5;

                        }
                        else
                        {
                            Console.WriteLine("Warning: UT4 is not empty, appending UT5 to it!\n");
                            act.UserText4 += " " + UserText5;
                        }
                        act.UserText5 = "";
                    }
            }

            char key = ' ';
            while (key != 'q')
            {
                Console.WriteLine("Press 'q' to exit");
                key = Console.ReadKey().KeyChar;
            }

        }
    }
}
