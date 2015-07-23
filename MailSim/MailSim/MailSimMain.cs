//--------------------------------------------------------------------------------------
// Copyright (c) 2015 Microsoft Corporation. All rights reserved.
// Use of this sample source code is subject to the terms of the Microsoft license
// agreement under which you licensed this sample source code and is provided AS-IS.
// If you did not accept the terms of the license agreement, you are not authorized
// to use this sample source code. For the terms of the license, please see the
// license agreement between you and Microsoft.
//--------------------------------------------------------------------------------------

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;


namespace MailSim
{
    class MailSimMain
    {
        /// <summary>
        /// Main program
        /// </summary>
        /// <param name="args">Command line arguments</param>
        static void Main(string[] args)
        {
            AppDomain.CurrentDomain.ProcessExit += new EventHandler(OnProcessExit);
            
            if (args.Length > 0)
            {
                if (args[0] == "/t")
                {
                    MailSimTest testClass = new MailSimTest();
                    testClass.Execute(args);
                    Log.Out(Log.Severity.Info, "", "Press any key to quit");
                    Console.Read();
                    return;
                }
            }

            if (args.Length != 1)
            {
                Log.Out(Log.Severity.Error, "", "Invalid parameter!");
                PrintUsage();
                return;
            }

            if (!File.Exists(args[0]))
            {
                Log.Out(Log.Severity.Error, "", "Invalid parameter, file does not exist");
                PrintUsage();
                return;
            }

            ProcessArgs(args);
        }

        static void OnProcessExit(object sender, EventArgs e)
        {
            Log.CloseLogFileElement();
        }

        /// <summary>
        /// Starting main execution engine
        /// </summary>
        /// <param name="args"></param>
        private static void ProcessArgs(string[] args)
        {
            try
            {
                Log.Initialize(args[0]);
                MailSimSequence seq = ConfigurationFile.LoadSequenceFile(args[0]);

                if (seq == null)
                {
                    Log.Out(Log.Severity.Error, Process.GetCurrentProcess().ProcessName, "Unable to load sequence XML file {0}", args[0]);
                    return;
                }

                ExecuteSequence exeSeq = new ExecuteSequence(seq);

                // initializes logging
                Log.LogFileLocation(seq.LogFileLocation);

                exeSeq.Execute();
            }
            catch (Exception ex)
            {
                Log.Out(Log.Severity.Error, Process.GetCurrentProcess().ProcessName, "Error encountered\n" + ex.ToString());
            }
        }


        /// <summary>
        /// This method prints the usage of the program
        /// </summary>
        public static void PrintUsage()
        {
            string binName = Process.GetCurrentProcess().ProcessName;

            Log.Out(Log.Severity.Info, binName, "{0} connects with Outlook and runs operations described in the input XML file", binName);
            Log.Out(Log.Severity.Info, binName, "Usage: {0} Sequence.xml", binName);
            Log.Out(Log.Severity.Info, binName, "   Sequence.xml: an XML file that specifies the sequence of an operation. Refer to the Sequence.xsd for its structure");
        }
    }
}
