using Microsoft.SharePoint.Client;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security;
using System.Text;
using System.Threading.Tasks;
using File = Microsoft.SharePoint.Client.File;

namespace Backup
{
    public class Utils
    {
        

        public static bool AddJSON(Folder parentFolder, ListItem item, string targetListTitle, ClientContext context)
        {
            try
            {
                using (context)
                {
                    ItemVersions itemVersions = GetItemVersions(item, context);
                    Console.WriteLine($"total Version : {itemVersions.versions.Count}");

                    string JsonFile = JsonConvert.SerializeObject(itemVersions);
                    string fileName = $"{targetListTitle}_ID_{item.Id}.json";

                    FileCreationInformation fileInfo = new FileCreationInformation();
                    byte[] bytes = ASCIIEncoding.ASCII.GetBytes(JsonFile);
                    using (Stream stream = new MemoryStream(bytes))
                    {
                        fileInfo.ContentStream = stream;
                        fileInfo.Url = fileName;
                        fileInfo.Overwrite = true;
                        File CreatedJsonFile = parentFolder.Files.Add(fileInfo);
                        context.Load(CreatedJsonFile);
                        context.ExecuteQuery();
                    }

                    return true;
                }
            }
            catch (Exception e)
            {
                printError(e.Message); return false;

            }
        }
        public static int CopyAttachments(Folder ParentFolder, ListItem item, ClientContext context)
        {
            try
            {
                using (context)
                {
                    Folder attachmentFolder = createListFolder(ParentFolder, "Attachments", context);
                    if (attachmentFolder == null) return -1;

                    AttachmentCollection attachments = item.AttachmentFiles;
                    context.Load(attachments);
                    context.ExecuteQuery();
                    int totalAttachments = 0;
                    foreach (Attachment attachment in attachments)
                    {

                        context.Load(attachment);
                        context.ExecuteQuery();

                        string serverRelativeUrl = attachment.ServerRelativeUrl;


                        File currFile = context.Web.GetFileByServerRelativeUrl(serverRelativeUrl);
                        context.Load(currFile);
                        context.ExecuteQuery();
                        ClientResult<Stream> clientResult = currFile.OpenBinaryStream();
                        context.ExecuteQuery();
                        using (Stream clientResultStream = clientResult.Value)
                        {

                            byte[] buffer = new byte[clientResultStream.Length];
                            int res = clientResultStream.Read(buffer, 0, buffer.Length);

                            if (res == -1) continue;


                            using (Stream stream = new MemoryStream(buffer))
                            {
                                FileCreationInformation creationInfo = new FileCreationInformation();
                                creationInfo.ContentStream = stream;
                                creationInfo.Url = attachment.FileName;
                                creationInfo.Overwrite = true;
                                File CreatedFile = attachmentFolder.Files.Add(creationInfo);

                                context.Load(CreatedFile);
                                context.ExecuteQuery();
                                totalAttachments++;
                                printSuccess("File created");

                            }

                        }


                    }

                    return totalAttachments;
                }

            }
            catch (Exception e)
            {
                printError(e.Message);
                return -1;
            }
        }
        public static ItemVersions GetItemVersions(ListItem item, ClientContext context)
        {
            try
            {
                using (context)
                {
                    int itemID = item.Id;
                    string createdDateTime = (string)item["Created_x0020_Date"];
                    ItemVersions currItemVersion = new ItemVersions(createdDateTime, itemID);


                    ListItemVersionCollection currItemVersionColl = item.Versions;
                    context.Load(currItemVersionColl);
                    context.ExecuteQuery();



                    for (int i = 0; i < currItemVersionColl.Count - 1; i++)
                    {
                        ListItemVersion currVersion = currItemVersionColl[i];
                        ListItemVersion prevVersion = currItemVersionColl[i + 1];

                        Dictionary<string, object> changedFieldValues = getChangedVersionItems(context, currVersion, prevVersion);

                        currItemVersion.versions.Add(changedFieldValues);

                    }
                    ListItemVersion initialV = currItemVersionColl[currItemVersionColl.Count - 1];

                    currItemVersion.versions.Add(initialV.FieldValues);
                    currItemVersion.currentItems = initialV.FieldValues;


                    return currItemVersion;
                }
            }
            catch (Exception e)
            {
                ItemVersions error = new ItemVersions(e.Message, -1);
                return error;
            }
        }
        public static Dictionary<string, object> getChangedVersionItems(ClientContext context, ListItemVersion currVersion, ListItemVersion prevVersion)
        {
            try
            {
                using (context)
                {

                    Dictionary<string, object> result = new Dictionary<string, object>();

                    Dictionary<string, object> currFieldValues = currVersion.FieldValues;
                    Dictionary<string, object> prevFieldValues = prevVersion.FieldValues;


                    foreach (KeyValuePair<string, object> pair in currFieldValues)
                    {
                        string key = pair.Key;
                        object value = pair.Value;

                        string CurrJSON = JsonConvert.SerializeObject(value);
                        string PrevJSON = JsonConvert.SerializeObject(prevFieldValues[key]);

                        if (CurrJSON.Equals(PrevJSON) == false || (key == "Editor" || key == "Modified"))
                        {
                            result[key] = value;

                        }
                    }

                    return result;
                }
            }
            catch (Exception e)
            {

                Dictionary<string, object> error = new Dictionary<string, object>();
                error["Error"] = e;
                return error;
            }

        }


        public static Folder createListFolder(Folder rootFolder, string currFolderName, ClientContext context)
        {
            try
            {
                using (context)
                {
                    if (rootFolder.FolderExists(currFolderName))
                    {
                        Console.WriteLine(currFolderName);
                        FolderCollection folders = rootFolder.Folders;
                        context.Load(folders); context.ExecuteQuery();

                        foreach (Folder folder in folders)
                            if (folder.Name == currFolderName) return folder;

                        return null;
                    }
                    else
                    {
                        Folder result = rootFolder.CreateFolder(currFolderName);
                        context.Load(result);
                        context.ExecuteQuery();
                        return result;
                    }

                }
            }
            catch (Exception e)
            {
                printError(e.Message);
                return null;

            }
        }


        public static SecureString getSecuredPass(string pass)
        {
            SecureString securedPass = new SecureString();
            foreach (char ch in pass) securedPass.AppendChar(ch);
            return securedPass;
        }


        public static void printError(string message)
        {
            Console.ForegroundColor = ConsoleColor.Red;
            Console.WriteLine(message);
            Console.ForegroundColor = ConsoleColor.White;

        }

        public static void printSuccess(string message)
        {
            Console.ForegroundColor = ConsoleColor.Green;
            Console.WriteLine(message);
            Console.ForegroundColor = ConsoleColor.White;

        }

    }



}



