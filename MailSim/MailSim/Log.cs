//--------------------------------------------------------------------------------------
// Copyright (c) 2015 Microsoft Corporation. All rights reserved.
// Use of this sample source code is subject to the terms of the Microsoft license
// agreement under which you licensed this sample source code and is provided AS-IS.
// If you did not accept the terms of the license agreement, you are not authorized
// to use this sample source code. For the terms of the license, please see the
// license agreement between you and Microsoft.
//--------------------------------------------------------------------------------------
using System;
using System.IO;
using System.Collections.Generic;
using System.Text;
using System.Xml;
using System.Xml.Linq;
using System.Diagnostics;
using System.Threading.Tasks;


namespace MailSim
{
    static class Log
    {
        private static string logFileName = null;
        private static string seqFileName;
        private static StreamWriter logWriter;

        /// <summary>
        /// Severity of the log
        /// </summary>
        public enum Severity
        {
            Info,
            Warning,
            Error
        };


        /// <summary>
        /// Initialize logging
        /// </summary>
        /// <param name="sequeceFile">sequence file name</param>
        public static void Initialize(string sequeceFile)
        {
            seqFileName = sequeceFile;
        }


        /// <summary>
        /// This method configures the log file name and location, and also write the root element to the log file
        /// </summary>
        /// <param name="fileLocation">location of the log file</param>
        public static void LogFileLocation(string fileLocation)
        {
            // use the local directory by default
            string logFileLocation = "";

            if (!string.IsNullOrEmpty(fileLocation))
            {
                // verify the directory exists
                if (!Directory.Exists(fileLocation))
                {
                    Out(Severity.Error, "Log", "Log file directory {0} doesn't exist", fileLocation);
                    throw new ArgumentException("Log file directory in the sequence XML file does not exist");
                }

                logFileLocation = fileLocation;
            }

            logFileName = logFileLocation + System.DateTime.Now.ToString("yyyy-MM-dd HH-mm-ss") + " " + Environment.MachineName + " " + Path.GetFileName(seqFileName);

            // append the initial element to the log file
            logWriter = new StreamWriter(logFileName, true, Encoding.UTF8);
            logWriter.WriteLine("<" + Process.GetCurrentProcess().ProcessName + ">");        
        }


        /// <summary>
        /// This method write the closing element for the log file
        /// </summary>
        public static void CloseLogFileElement()
        {
            if (!string.IsNullOrEmpty(logFileName))
            {
                logWriter.WriteLine("</" + Process.GetCurrentProcess().ProcessName + ">");
                logWriter.Close();
            }
        }


        /// <summary>
        /// This method displays the message to the console and also writes it to the log file
        /// </summary>
        /// <param name="type">Severity of the message</param>
        /// <param name="name">name of the task/process</param>
        /// <param name="format">format of the message</param>
        /// <param name="args">variable to use with the format</param>
        public static void Out(Severity type, string name, string format, params object[] args)
        {
            
            switch (type)
            {
                case Severity.Error:
                    Console.ForegroundColor = ConsoleColor.Red;
                    break;
                case Severity.Warning:
                    Console.ForegroundColor = ConsoleColor.Yellow;
                    break;
                default:
                    if(name == ExecuteSequence.eventString)
                    {
                        Console.ForegroundColor = ConsoleColor.Green;
                    }
                    break;
            }

            // writes to the console
            Console.WriteLine(System.DateTime.Now + " " + type.ToString() + "\t: " + 
                name + "\t: " + format, args);

            Console.ResetColor();

            // writes to the log file
            if (!string.IsNullOrEmpty(logFileName))
            {
                XElement element = new XElement(type.ToString(),
                    new XAttribute("Name", name),
                    new XAttribute("Time", System.DateTime.Now.ToString()),
                    new XElement("Detail", String.Format(format, args)));
                logWriter.WriteLine(element.ToString());
            }
        }
    }
}
