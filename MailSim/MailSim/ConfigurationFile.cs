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
using System.IO;
using System.Xml;
using System.Xml.Serialization;
using System.Xml.Schema;
using MailSim.Common;

namespace MailSim
{
    static class ConfigurationFile
    {
        public const string SequenceSchema = "Sequence.xsd";
        public const string OperationSchema = "Operations.xsd";
        public const string XMLProcessing = "XML Processing";


        /// <summary>
        /// This method loads the sequence XML file, validates it with the schema and deserializes it
        /// </summary>
        /// <param name="sequenceFile">Path and file name of the sequence file</param>
        /// <returns>Returns MailSimSequence if successful, otherwise returns null</returns>
        public static MailSimSequence LoadSequenceFile(string sequenceFile)
        {
            return LoadXml<MailSimSequence>(sequenceFile, SequenceSchema);
        }


        /// <summary>
        /// This method loads the operation XML file, validates it with the schema and deserializes it
        /// </summary>
        /// <param name="opFile">Path and file name of the operation file</param>
        /// <param name="opXML">XmlDocument of the operation file</param>
        /// <returns>Returns MailSimOperations if successful, otherwise returns null </returns>
        public static MailSimOperations LoadOperationFile(string opFile)
        {
            return LoadXml<MailSimOperations>(opFile, OperationSchema);
        }

        public static T LoadXml<T>(string xmlFile, string schemaFile)
        {
            XmlDocument opXML;
            return LoadXml<T>(xmlFile, schemaFile, out opXML);
        }

        private static T LoadXml<T>(string xmlFile, string schemaFile, out XmlDocument outXML)
        {
            outXML = null;

            if (/*!xmlFile.EndsWith(".xml", StringComparison.InvariantCultureIgnoreCase) ||*/ !File.Exists(xmlFile))
            {
                Log.Out(Log.Severity.Error, XMLProcessing, "Specified xml file {0} does not exist", xmlFile);
                return default(T);
            }

            if (!File.Exists(schemaFile))
            {
                Log.Out(Log.Severity.Error, XMLProcessing, "Unable to locate schema file {0}", schemaFile);
                return default(T);
            }

            try
            {
                outXML = ValidateXML(xmlFile, schemaFile);

                // Validates the operation file.
                if (outXML == null)
                {
                    Log.Out(Log.Severity.Error, XMLProcessing, "Unable to process file {0}", xmlFile);
                    return default(T);
                }

                var serializer = new XmlSerializer(typeof(T));
                using (var reader = XmlReader.Create(xmlFile))
                {
                    return (T)serializer.Deserialize(reader);
                }
            }
            catch (Exception ex)
            {
                Log.Out(Log.Severity.Error, XMLProcessing, "LoadXml exception\n" + ex.ToString());
                return default(T);
            }
        }

        /// <summary>
        /// This method validates the XML file with the schema
        /// </summary>
        /// <param name="xmlFile">XML file to validate</param>
        /// <param name="xmlSchema">XML Schema file to use for validation</param>
        /// <returns>Returns XmlDocument if the validation is successful, otherwise returns null </returns>
        private static XmlDocument ValidateXML(string xmlFile, string xmlSchema)
        {
            if (string.IsNullOrEmpty(xmlFile) || !File.Exists(xmlFile))
            {
                Log.Out(Log.Severity.Error, "XML file {0} doesn't exist!", xmlFile);
                return null;
            }

            if (string.IsNullOrEmpty(xmlSchema) || !File.Exists(xmlSchema))
            {
                Log.Out(Log.Severity.Error, "Schema file {0} doesn't exist!", xmlSchema);
                return null;
            }

            XmlReaderSettings settings = new XmlReaderSettings();
            settings.Schemas.Add(null, xmlSchema);
            settings.ValidationType = ValidationType.Schema;

            XmlReader reader = XmlReader.Create(xmlFile, settings);
            XmlDocument testConfig = new XmlDocument();

            try
            {
                testConfig.Load(reader);
                ValidationEventHandler eventHandler = new ValidationEventHandler(ValidationEventHandler);
                testConfig.Validate(eventHandler);
            }
            catch (Exception ex)
            {
                Log.Out(Log.Severity.Error, XMLProcessing,
                    "{0} schema validation exception encountered\n" + ex.ToString(), xmlFile);
                return null;
            }

            return testConfig;
        }
        
        /// <summary>
        /// Handler for XML schema validation
        /// </summary>
        /// <param name="sender">caller of this handler</param>
        /// <param name="e">validation event argument</param>
        static void ValidationEventHandler(object sender, ValidationEventArgs e)
        {
            switch (e.Severity)
            {
                case XmlSeverityType.Error:
                    Log.Out(Log.Severity.Error, XMLProcessing, "Schema Validation Error: {0}", e.Message);
                    break;
                case XmlSeverityType.Warning:
                    Log.Out(Log.Severity.Warning, XMLProcessing, "Schema Validation Warning {0}", e.Message);
                    break;
                default:
                    Log.Out(Log.Severity.Error, XMLProcessing, "Schema Validation: {0}", e.Message);
                    break;
            }
        }
    }
}
