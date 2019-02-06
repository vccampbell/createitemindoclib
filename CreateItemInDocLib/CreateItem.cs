using Microsoft.SharePoint;
using System;
using System.Collections;
using System.Globalization;
using System.IO;
using System.Text;

namespace CreateItemInDocLib
{
    public class CreateItem
    {
        /// <summary>
        /// Creates a document or folder in the Doc Lib specified. 
        /// </summary>
        /// <param name="web"></param>
        /// <param name="listId"></param>
        /// <param name="itemId"></param>
        /// <param name="itemProperties"></param>
        /// <param name="fileContents"></param>
        /// <returns></returns>
        public static SPListItem CreateDocument(SPWeb web, Guid listId, int itemId, Hashtable itemProperties, byte[] fileContents)
        {
            return CreateDocumentInternal(web, listId, itemId, itemProperties, fileContents);
        }
        private static SPListItem CreateDocumentInternal(
            SPWeb web,
            Guid listId,
            int itemId,
            Hashtable itemProperties,
            byte[] fileContents)
        {
            SPList targetList = web.Lists[listId];
            SPContentTypeId typeIdIfAvailable = GetContentTypeIdIfAvailable(targetList, itemProperties);
            string templateUrl = DetermineTemplateUrl(targetList, itemProperties);
            byte[] documentBytes = DetermineBytesForTemplate(web, templateUrl);
            string filePath = (string)itemProperties[(object)"FileLeafRef"];
            itemProperties.Remove((object)"FileLeafRef");
            if (!filePath.StartsWith("/", StringComparison.Ordinal))
                filePath = targetList.RootFolder.Url + "/" + filePath;
            bool isFolderType = false;
            if (typeIdIfAvailable.IsChildOf(SPBuiltInContentTypeId.Folder))
            {
                isFolderType = true;
                int length = filePath.LastIndexOf('.');
                if (length > -1)
                    filePath = filePath.Substring(0, length);
            }
            bool flag2 = false;
            foreach (SPContentType contentType in (SPBaseCollection)targetList.ContentTypes)
            {
                if (contentType.Id.Parent == typeIdIfAvailable || contentType.Id == typeIdIfAvailable)
                {
                    flag2 = true;
                    break;
                }
            }
            if (!flag2)
            {
                SPContentTypeId spContentTypeId = targetList.ContentTypes.BestMatch(typeIdIfAvailable);
                if (typeIdIfAvailable.Parent == spContentTypeId.Parent)
                    itemProperties[(object)"Content Type ID"] = (object)spContentTypeId.ToString();
            }
            SPListItem spListItem;
            if (!isFolderType)
                spListItem = CommitCreateListItemInDocLib(filePath, itemProperties, documentBytes, targetList);
            else
                spListItem = CommitCreateFolderInDocLib(filePath, targetList);           
            
            spListItem = SetFieldsInDocument(spListItem, targetList, itemProperties);
            spListItem.Update();
            return spListItem;
        }
        /// <summary>
        /// Creates Document in Document Library
        /// </summary>
        /// <param name="filePath"></param>
        /// <param name="itemProperties"></param>
        /// <param name="documentBytes"></param>
        /// <param name="targetList"></param>
        /// <returns></returns>
        private static SPListItem CommitCreateListItemInDocLib(string filePath, Hashtable itemProperties, byte[] documentBytes, SPList targetList)
        {
            string documentName = Path.GetFileName(filePath);
            string path = filePath.Replace(documentName, "");
            if(targetList.ParentWeb.GetFolder(path).Exists)
            {
                SPFolder target = targetList.ParentWeb.GetFolder(path);

                SPFileCollection files = target.Files;
                SPFile spFile;
                try
                {
                    spFile = files.Add(Path.GetFileName(filePath), documentBytes);
                    return spFile.Item;
                }
                catch (SPException ex)
                {
                    if (ex.ErrorCode == -2130575257) //file already exists
                    {
                        string uniqueName = CreateUniqueName(Path.GetFileName(filePath));
                        spFile = files.Add(uniqueName, documentBytes);
                        return spFile.Item;
                    }
                    else
                        throw;
                }
            }
            else
            {
                throw new FileNotFoundException(String.Format("Folder does not exist ({0}).", path));
            }             
        }

        /// <summary>
        /// Creates Folder in Document Library
        /// </summary>
        /// <param name="folder"></param>
        /// <param name="targetList"></param>
        /// <returns></returns>
        private static SPListItem CommitCreateFolderInDocLib(string folder, SPList targetList)
        {
            try
            {
                if (targetList.ParentWeb.GetFolder(folder).Exists)
                    folder = CreateUniqueName(folder);
                int length = folder.LastIndexOf('/');
                string folderUrl = targetList.RootFolder.Url;
                if (length > -1)
                {
                    folderUrl = folder.Substring(0, length);
                    folder = folder.Substring(length + 1);
                }

                SPListItem spListItem = targetList.AddItem(folderUrl, SPFileSystemObjectType.Folder, folder);
                return spListItem;
            }
            catch(Exception)
            {
                throw;
            }            
        }

        /// <summary>
        /// Creates a unique name for document by appending date and time file name
        /// </summary>
        /// <param name="original"></param>
        /// <returns></returns>
        private static string CreateUniqueName(string original)
        {
            StringBuilder sb = new StringBuilder();
            int num = original.LastIndexOf('.');
            if (num >= 0)
                sb.Append(original.Substring(0, num));
            else
                sb.Append(original);
            DateTime now = DateTime.Now;
            sb.Append("(");
            sb.Append(now.ToString("yyyy-MM-dd_H-mm-ss", (IFormatProvider)CultureInfo.InvariantCulture));
            sb.Append(")");
            if (num >= 0)
                sb.Append(original.Substring(num, original.Length - num));
            return sb.ToString();
        }

        /// <summary>
        /// Gets the Content Type Id if it is available on the list.
        /// </summary>
        /// <param name="list"></param>
        /// <param name="itemProperties"></param>
        /// <returns></returns>
        private static SPContentTypeId GetContentTypeIdIfAvailable(
        SPList list,
        Hashtable itemProperties)
        {
            SPContentTypeId contentTypeId = SPContentTypeId.Empty;
            object index;
            if (itemProperties.ContainsKey((object)"Content Type ID"))
            {
                index = (object)"Content Type ID";
            }
            else
            {
                if (!itemProperties.ContainsKey((object)SPBuiltInFieldId.ContentTypeId))
                    return contentTypeId;
                index = (object)SPBuiltInFieldId.ContentTypeId;
            }
            contentTypeId = new SPContentTypeId((string)itemProperties[index]);
            SPContentTypeId spContentTypeId = list.ContentTypes.BestMatch(contentTypeId);
            itemProperties[index] = (object)spContentTypeId;
            return spContentTypeId;
        }

        /// <summary>
        /// Determines the URL for the document template
        /// </summary>
        /// <param name="list"></param>
        /// <param name="itemProperties"></param>
        /// <returns></returns>
        private static string DetermineTemplateUrl(SPList list, Hashtable itemProperties)
        {
            SPDocumentLibrary spDocumentLibrary = list as SPDocumentLibrary;
            if (spDocumentLibrary == null)
                throw new ArgumentException(nameof(list));
            string documentTemplateUrl = spDocumentLibrary.DocumentTemplateUrl;
            if (spDocumentLibrary.AllowContentTypes)
            {
                string id = "";
                if (itemProperties.Contains((object)SPBuiltInFieldId.ContentTypeId) && itemProperties[(object)SPBuiltInFieldId.ContentTypeId] != null)
                    id = itemProperties[(object)SPBuiltInFieldId.ContentTypeId].ToString();
                else if (itemProperties.Contains((object)"Content Type ID") && itemProperties[(object)"Content Type ID"] != null)
                    id = itemProperties[(object)"Content Type ID"].ToString();
                if (!string.IsNullOrEmpty(id))
                {
                    SPContentTypeId index = new SPContentTypeId(id);
                    SPContentType contentType = spDocumentLibrary.ContentTypes[index];
                    if (contentType != null && !string.IsNullOrEmpty(contentType.DocumentTemplateUrl))
                        documentTemplateUrl = contentType.DocumentTemplateUrl;
                }
            }
            return documentTemplateUrl;
        }

        /// <summary>
        /// Gets the bytes for the template document
        /// </summary>
        /// <param name="web"></param>
        /// <param name="docTemplateUrl"></param>
        /// <returns></returns>
        private static byte[] DetermineBytesForTemplate(SPWeb web, string docTemplateUrl)
        {
            if (docTemplateUrl.Length > 0)
                return web.GetFile(docTemplateUrl).OpenBinary();
            return new Byte[1] { (byte)0 };
        }

        /// <summary>
        /// Applies the values to the metadata columns
        /// </summary>
        /// <param name="theRecord"></param>
        /// <param name="theList"></param>
        /// <param name="itemProperties"></param>
        /// <returns></returns>
        private static SPListItem SetFieldsInDocument(SPListItem theRecord, SPList theList, Hashtable itemProperties)
        {
            if (itemProperties == null || theRecord == null)
                return null;
            IDictionaryEnumerator enumerator = itemProperties.GetEnumerator();
            while (enumerator.MoveNext())
            {
                object key = enumerator.Key;
                object obj = enumerator.Value;
                if (key is string)
                    theRecord[(string)key] = obj;
                else if (key is Guid)
                    theRecord[(Guid)key] = obj;
            }
            return theRecord;
        }
    }
}
