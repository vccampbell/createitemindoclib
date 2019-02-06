using System;
using Microsoft.SharePoint;
using System.Collections;

namespace CreateItemInDocLib
{
    class Program
    {
        const string siteUrl = "";
        const string docLibTitle = "Shared Documents";
        const string folderContentTypeId = "0x012000AE42FBBCB4E27944B1C558B2E5602B6D";
        const string documentContentTypeId = "0x010100D00ADB6ED139F44195A971E927767D7300945F742540F8C8438445F85E2918E2B5";
        static void Main(string[] args)
        {
            using (SPSite site = new SPSite(siteUrl))
            {
                using (SPWeb web = site.OpenWeb())
                {
                    SPList list = web.Lists.TryGetList(docLibTitle);
                    // Create Folder in Document Library
                    Hashtable folderProperties = new Hashtable();
                    folderProperties.Add("Content Type ID", folderContentTypeId);
                    folderProperties.Add("FileLeafRef", "Mike");
                    SPListItem spFolderItem = CreateItem.CreateDocument(web, list.ID, -1, folderProperties, null);

                    // Create a document based on the document template and set the Title column
                    Hashtable documentProperties = new Hashtable();
                    documentProperties.Add("Content Type ID", documentContentTypeId);
                    documentProperties.Add("FileLeafRef", String.Format("{0}/MyNewDocument.docx", list.RootFolder.ServerRelativeUrl));
                    documentProperties.Add("Title", "This is the title");
                    SPListItem spDocumentItem = CreateItem.CreateDocument(web, list.ID, -1, documentProperties, null);
                }
            }
        }
        
    }
}
