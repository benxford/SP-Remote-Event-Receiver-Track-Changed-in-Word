using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.EventReceivers;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace SP_Remote_Event_Receiver___Track_Changed_in_WordWeb.Services
{
    public class WordEventReceiver : IRemoteEventService
    {
        /// <summary>
        /// Not used.
        /// </summary>
        /// <param name="properties">Holds information about the event.</param>
        /// <returns>Holds information returned from the event.</returns>
        public SPRemoteEventResult ProcessEvent(SPRemoteEventProperties properties)
        {
            return new SPRemoteEventResult();
        }

        /// <summary>
        /// Enables tracked changes in the Word document.
        /// </summary>
        /// <param name="properties">Unused.</param>
        public void ProcessOneWayEvent(SPRemoteEventProperties properties)
        {
            using (var hostContext = new SPAppContext(properties.ItemEventProperties.WebUrl))
            {
                // Get the item based on the event properties
                var item = hostContext.Web.Lists.GetById(properties.ItemEventProperties.ListId).GetItemById(properties.ItemEventProperties.ListItemId);
                
                // The file tied to the list item
                var file = item.File;

                // Load the file properties
                hostContext.Load(file);
                hostContext.ExecuteQuery();

                // Verify it is .docx file
                if (!file.Name.EndsWith(".docx"))
                {
                    return;
                }

                // Open the file
                var content = file.OpenBinaryStream();
                hostContext.ExecuteQuery();

                // Memory stream
                using (var stream = new System.IO.MemoryStream())
                {
                    // Copy the file into the memory stream
                    content.Value.CopyTo(stream);

                    // Open the file as a ZipArchive
                    using (var zip = new System.IO.Compression.ZipArchive(stream, System.IO.Compression.ZipArchiveMode.Update))
                    {
                        // XmlDocument for processing word/settings.xml
                        var doc = new System.Xml.XmlDocument();

                        // Load word/settings.xml, then delete from the zip.
                        var entry = zip.GetEntry("word/settings.xml");
                        using (var settingsStream = entry.Open())
                        {
                            doc.Load(settingsStream);
                        }
                        entry.Delete();

                        // Namespace required for enabling tracked changes.
                        var nsMgr = new System.Xml.XmlNamespaceManager(doc.NameTable);
                        nsMgr.AddNamespace("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");

                        // Check if tracked changes already enabled
                        if (doc.SelectSingleNode("/w:settings/w:trackRevisions", nsMgr) != null)
                        {
                            return;
                        }
                        
                        // Enable tracked changes
                        doc.DocumentElement.AppendChild(doc.CreateElement("w:trackRevisions", "http://schemas.openxmlformats.org/wordprocessingml/2006/main"));

                        // Save word/settings.xml
                        entry = zip.CreateEntry("word/settings.xml");
                        using (var settingsStream = entry.Open())
                        {
                            doc.Save(settingsStream);
                        }
                    }
                    stream.Flush();

                    // Save the document with tracked changes enabled
                    file.SaveBinary(new Microsoft.SharePoint.Client.FileSaveBinaryInformation
                    {
                        Content = stream.ToArray()
                    });

                    hostContext.ExecuteQuery();
                }
            }
        }
    }
}
