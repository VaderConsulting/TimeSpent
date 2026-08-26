//
// Copyright © Microsoft Corporation.  All Rights Reserved.
// This code released under the terms of the 
// Microsoft Public License (MS-PL, http://opensource.org/licenses/ms-pl.html)
//
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using TimeSpentLib;

namespace TimeSpent
{
    class CmdLine
    {
        
        // Command line switch constants for parsing.
        private const string StartSwitch = "/start:";
        private const string EndSwitch = "/end:";
        private const string OutfileSwitch = "/out:";
        private const string MailPathSwitch = "/mailpath:";
        private const string DomainSwitch = "/domain:";
        private const string UsernameSwitch = "/username:";
        private const string PasswordSwitch = "/password:";
        private const string VersionSwitch = "/version:";

        // Options as set by the user.
        private DateTime _StartDate { get; set; }
        private DateTime _EndDate { get; set; }
        private string _OutputFile { get; set; }
        private string _MailPath { get; set; }
        private string _Domain { get; set; }
        private string _Username { get; set; }
        private string _Password { get; set; }
        private string _Version { get; set; }

        /// <summary>
        /// Parse all input arguments and validate all data included.
        /// </summary>
        /// <param name="args">Arguments as passed to Main</param>
        /// <returns>true if all data is ok; prints usage and returns false otherwise</returns>
        public bool Parse(string[] args)
        {
            // Parsing logic is very simple:  check each argument to see if it is
            // a switch and if so grab the value.
            foreach (string str in args)
            {
                if (str.StartsWith(StartSwitch))
                {
                    DateTime date;
                    string sdate = str.Substring(StartSwitch.Length);
                    if ((DateTime.TryParse(sdate, out date)) != true)
                        return Usage();
                    _StartDate = date;
                }
                else if (str.StartsWith(EndSwitch))
                {
                    DateTime date;
                    string sdate = str.Substring(EndSwitch.Length);
                    if ((DateTime.TryParse(sdate, out date)) != true)
                        return Usage();
                    _EndDate = date;
                }
                else if (str.StartsWith(OutfileSwitch))
                {
                    _OutputFile = str.Substring(OutfileSwitch.Length);
                }
                else if (str.StartsWith(MailPathSwitch))
                {
                    _MailPath = str.Substring(MailPathSwitch.Length);
                }
                else if (str.StartsWith(DomainSwitch))
                {
                    _Domain = str.Substring(DomainSwitch.Length);
                }
                else if (str.StartsWith(UsernameSwitch))
                {
                    _Username = str.Substring(UsernameSwitch.Length);
                }
                else if (str.StartsWith(PasswordSwitch))
                {
                    _Password = str.Substring(PasswordSwitch.Length);
                }
                else if (str.StartsWith(VersionSwitch))
                {
                    _Version = str.Substring(VersionSwitch.Length);
                }
                else
                {
                    Console.WriteLine(String.Format("Unrecognized option: {0}"), str);
                    return Usage();
                }
            }

            // Validate the dates.  Date range must be specified and be in proper 
            // chronilogial order.
            if (_StartDate.ToString() == DateTime.MinValue.ToString() ||
                _EndDate.ToString() == DateTime.MinValue.ToString() ||
                _StartDate > _EndDate)
            {
                Console.WriteLine("Invalid dates");
                return Usage();
            }

            // Path to EWS must be specified or we cannot retrieve any data.
            if (string.IsNullOrEmpty(_MailPath))
            {
                Console.WriteLine("Invalid Exchange path");
                return Usage();
            }

            return true;
        }

        /// <summary>
        /// Prints out the valid usage patterns and for convience returns false.
        /// </summary>
        /// <returns>Always false</returns>
        public bool Usage()
        {
            Console.WriteLine("timespent /domain:<domain> /username:<username> /password:<password> /start:<date> /end:<date> /mailpath:<path> /version: <exchange version> /out:<location>");
            Console.WriteLine("example: /start:01/01/2013 /end:31/12/2013 /mailpath:https://exchange.example/EWS/Exchange.asmx /out:timespent.csv /domain:mydomainname /username:myusername /password:itsasecret /version:2010");
            return false;
        }

        /// <summary>
        /// Executes the main logic of the console application.
        /// </summary>
        public void Run()
        {
            // Print banner information.
            Console.WriteLine("timespent:");
            Console.WriteLine(String.Format("\tFrom:      {0}", _StartDate));
            Console.WriteLine(String.Format("\tTo:        {0}", _EndDate));
            Console.WriteLine(String.Format("\tOutput:    {0}", _OutputFile));
            Console.WriteLine(String.Format("\tExchange:  {0}", _MailPath));
            Console.WriteLine(String.Format("\tVersion:   {0}", _Version));
            if (_Domain != "")
            {
                Console.WriteLine(String.Format("\tDomain:    {0}", _Domain));
                Console.WriteLine(String.Format("\tUsername:  {0}", _Username));
                //Console.WriteLine(String.Format("\tPassword:  {0}", _Password));
            }

            try
            {
                CalendarItemData item;
                TextWriter writer;

                // Allocate our own helper object and give it the top level data.
                ServerData server = new ServerData();

                Console.WriteLine(String.Format("Connecting to {0}", _MailPath));

                // Connect to EWS.
                if (_Version == "2007" || _Version == "") // Exchange 2007
                {
                    if (_Domain == "")
                    {
                        server.ConnectToServer2007(_MailPath);
                    }
                    else
                    {
                        server.ConnectToServer2007(_MailPath, _Domain, _Username, _Password);
                    }
                }

                if (_Version == "2007_SP1") // Exchange 2007SP1
                {
                    if (_Domain == "")
                    {
                        server.ConnectToServer2007_SP1(_MailPath);
                    }
                    else
                    {
                        server.ConnectToServer2007_SP1(_MailPath, _Domain, _Username, _Password);
                    }
                }

                if (_Version == "2010") // Exchange 2010
                {
                    if (_Domain == "")
                    {
                        server.ConnectToServer2010(_MailPath);
                    }
                    else
                    {
                        server.ConnectToServer2010(_MailPath, _Domain, _Username, _Password);
                    }
                }

                if (_Version == "2010_SP1") // Exchange 2010SP1
                {
                    if (_Domain == "")
                    {
                        server.ConnectToServer2010_SP1(_MailPath);
                    }
                    else
                    {
                        server.ConnectToServer2010_SP1(_MailPath, _Domain, _Username, _Password);
                    }
                }

                if (_Version == "2010_SP2") // Exchange 2010SP2
                {
                    if (_Domain == "")
                    {
                        server.ConnectToServer2010_SP2(_MailPath);
                    }
                    else
                    {
                        server.ConnectToServer2010_SP2(_MailPath, _Domain, _Username, _Password);
                    }
                }

                // Ask for the calendar items.  This will be the most expensive
                // part of the logic.
                Console.WriteLine("Retrieving calendar data from server");
                CalendarItemList list = server.GetCalendarItems(_StartDate, _EndDate);
                
                // If there is no output file specified, just write to the console.
                if (String.IsNullOrEmpty(_OutputFile))
                {
                    writer = Console.Out;
                }
                // Otherwise create the new output file.
                else
                {
                    Console.WriteLine(String.Format("Creating file {0}", _OutputFile));
                    writer = new StreamWriter(_OutputFile);
                }

                // Write the header for the .csv file
                writer.WriteLine("Subject,Date,Start,End");

                // Iterate every calendar item and write the output.
                while ((item = list.GetNextItem()) != null)
                {
                    
                    // Categories can be an array, we want just a simple string item in single column.
                    string strCat = GetCategoryString(item.Categories);

                    // Get a simple date string for that column.
                    string strDate = GetDateString(item.StartDate);

                    // Get a simple date string for that column.
                    string strStart = GetTimeString(item.StartDate);

                    // Get a simple date string for that column.
                    string strEnd = GetTimeString(item.EndDate);

                    // Dump te data row to the output stream.
                    string str = String.Format("\"{0}\",\"{1}\",\"{2}\",\"{3}\"",
                                item.Subject, strDate, strStart, strEnd);
                    writer.WriteLine(str);
                }

                writer.Flush();
                writer.Close();

                Console.WriteLine("done");
            }
            catch (Exception ex)
            {
                Console.WriteLine("Fatal error: ");
                Console.WriteLine(ex.ToString());
            }
        }

        /// <summary>
        /// Turn a category array into a single string item.
        /// </summary>
        /// <param name="item">Calendar item with categories</param>
        /// <returns>Combined string</returns>
        private static string GetCategoryString(string[] Categories)
        {
            string strCat = "";
            if (Categories != null)
            {
                foreach (string cat in Categories)
                    strCat += cat;
            }
            if (strCat.Length == 0)
                strCat = "<empty>";
            return strCat;
        }

        /// <summary>
        /// Give a short form of the data.  
        /// </summary>
        /// <param name="itemDate">Date to parse</param>
        /// <returns>Simple formatted date</returns>
        private static string GetDateString(DateTime itemDate)
        {
            string str = String.Format("{0}/{1}/{2}", itemDate.Day, itemDate.Month, itemDate.Year);
            return str;
        }

        /// <summary>
        /// Give a short form of the data.  
        /// </summary>
        /// <param name="itemDate">Time to parse</param>
        /// <returns>Simple formatted time</returns>
        private static string GetTimeString(DateTime itemDate)
        {
            string str = String.Format("{0}:{1}:{2}", itemDate.Hour, itemDate.Minute, itemDate.Second);
            return str;
        }

    }
}
