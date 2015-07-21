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
            MailSimSequence sequence = null;

            if (!sequenceFile.EndsWith(".xml", StringComparison.InvariantCultureIgnoreCase) || !File.Exists(sequenceFile))
            {
                Log.Out(Log.Severity.Error, XMLProcessing, "Sequence file {0} does not exist", sequenceFile);
                return null;
            }

            if (!File.Exists(SequenceSchema))
            {
                Log.Out(Log.Severity.Error, XMLProcessing, "Unable to locate schema file {0}", SequenceSchema);
                return null;
            }

            try
            {
                if (ValidateXML(sequenceFile, SequenceSchema) == null)
                {
                    Log.Out(Log.Severity.Error, XMLProcessing, "Unable to process the sequence file {0}", sequenceFile);
                    return null;
                }

                // Deserializes the sequence XML file
                XmlSerializer seqSer = new XmlSerializer(typeof(MailSimSequence));
                using (XmlReader seqReader = XmlReader.Create(sequenceFile))
                {
                    sequence = (MailSimSequence)seqSer.Deserialize(seqReader);
                }
            }
            catch (Exception ex)
            {
                Log.Out(Log.Severity.Error, XMLProcessing, "Run exception\n" + ex.ToString());
                return null;
            }

            return sequence;
        }


        /// <summary>
        /// This method loads the operation XML file, validates it with the schema and deserializes it
        /// </summary>
        /// <param name="opFile">Path and file name of the operation file</param>
        /// <param name="opXML">XmlDocument of the operation file</param>
        /// <returns>Returns MailSimOperations if successful, otherwise returns null </returns>
        public static MailSimOperations LoadOperationFile(string opFile, out XmlDocument opXML)
        {
            MailSimOperations operations = null;
            opXML = null;

            if (!opFile.EndsWith(".xml", StringComparison.InvariantCultureIgnoreCase) || !File.Exists(opFile))
            {
                Log.Out(Log.Severity.Error, XMLProcessing, "Specified operation file {0} does not exist", opFile);
                return null;
            }

            if (!File.Exists(OperationSchema))
            {
                Log.Out(Log.Severity.Error, XMLProcessing, "Unable to locate schema file {0}", OperationSchema);
                return null;
            }
            try
            {
                opXML = ValidateXML(opFile, OperationSchema);

                // Validates the operation file.
                if (opXML == null)
                {
                    Log.Out(Log.Severity.Error, XMLProcessing, "Unable to process the operation file {0}", opFile);
                    return null;
                }

                // Loads each referenced operations file from the sequence file.
                XmlSerializer opSer = new XmlSerializer(typeof(MailSimOperations));
                using (XmlReader opReader = XmlReader.Create(opFile))
                {
                    operations = (MailSimOperations)opSer.Deserialize(opReader);
                }

            }
            catch (Exception ex)
            {
                Log.Out(Log.Severity.Error, XMLProcessing, "LoadOperationFile exception\n" + ex.ToString());
                return null;
            }

            return operations;
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
