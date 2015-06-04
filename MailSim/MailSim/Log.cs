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
using System.Linq;
using System.Text;
using System.Xml;
using System.Diagnostics;
using System.Threading.Tasks;


namespace MailSim
{
    static class Log
    {
        private static string logFileName = null;
        private static string seqFileName;

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
                    throw new ArgumentException("Log file directory in sequence file does not exist");
                }

                logFileLocation = fileLocation;
            }

            logFileName = logFileLocation + System.DateTime.Now.ToString("yyyy-MM-dd HH-mm-ss") + " " + Environment.MachineName + " " + Path.GetFileName(seqFileName);

            // append the initial element to the log file
            StreamWriter sWriter = new StreamWriter(logFileName, true, Encoding.UTF8);
            sWriter.WriteLine("<" + Process.GetCurrentProcess().ProcessName + ">");
            sWriter.Close();            
        }


        /// <summary>
        /// This method write the closing element for the log file
        /// </summary>
        public static void CloseLogFileElement()
        {
            if (!string.IsNullOrEmpty(logFileName))
            {
                StreamWriter sWriter = new StreamWriter(logFileName, true, Encoding.UTF8);
                sWriter.WriteLine("</" + Process.GetCurrentProcess().ProcessName + ">");
                sWriter.Close();
            }
        }


        /// <summary>
        /// This method 
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
            }

            // write to the console
            Console.WriteLine(System.DateTime.Now + " " + type.ToString() + "\t: " + 
                name + "\t: " + format, args);

            Console.ResetColor();

            // write to the log file
            if (!string.IsNullOrEmpty(logFileName))
            {
                StringBuilder sBuilder = new StringBuilder();
                using (StringWriter sWriter = new StringWriter(sBuilder))
                {
                    using (XmlTextWriter xmlWriter = new XmlTextWriter(sWriter))
                    {
                        xmlWriter.WriteStartElement(type.ToString());
                        xmlWriter.WriteAttributeString("Name", name);
                        xmlWriter.WriteAttributeString("Time", System.DateTime.Now.ToString());
                        xmlWriter.WriteElementString("Detail", String.Format(format, args));
                        xmlWriter.WriteEndElement();
                    }
                }
                using (StreamWriter sWriter = new StreamWriter(logFileName, true, Encoding.UTF8))
                {
                    sWriter.WriteLine(sBuilder.ToString());
                }
            }
        }
        
    }
}
