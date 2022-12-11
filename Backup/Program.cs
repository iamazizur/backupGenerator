using Microsoft.SharePoint.Client;
using Newtonsoft.Json;
using PnP.Framework.Modernization.Cache;
using System;
using System.CodeDom;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Runtime.Remoting.Contexts;
using System.Text;
using System.Threading.Tasks;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.TrackBar;
using File = Microsoft.SharePoint.Client.File;

namespace Backup
{

   
    internal class Program
    {
        protected const string url = "";
        protected const string username = "";
        protected const string pass = "";
        protected const string targetedtListTitle = "";
        protected const string backupListTitle = "Backups";

        
       
       public static void Main(string[] args)
        {
            BackupGenerator MyBackup = new BackupGenerator(url, username, pass, targetedtListTitle, backupListTitle);
            MyBackup.CreateBackup();
            Console.ReadLine();
        }
        
       
        
        
        
        
        //main ends here
        
        
        
        
        
        
        
        static void ForSubFolders(string targetListTitle, string baseListTitle)
        {
            using (ClientContext context = new ClientContext(url))
            {
                context.Credentials = new SharePointOnlineCredentials(username, Utils.getSecuredPass(pass));
                List targetList = context.Web.Lists.GetByTitle(targetListTitle);
                CamlQuery query = CamlQuery.CreateAllItemsQuery();
                ListItemCollection items = targetList.GetItems(query);
                context.Load(items); context.Load(targetList); context.ExecuteQuery();

                foreach (ListItem item in items)
                {
                    context.Load(item); context.ExecuteQuery();

                    Dictionary<string, object> fields = item.FieldValues;
                    FieldLookupValue fieldLookupValue = (FieldLookupValue)fields[baseListTitle];
                    int id = fieldLookupValue.LookupId;

                    List baseList = context.Web.Lists.GetByTitle(baseListTitle);
                    ListItem baseListItem = baseList.GetItemById(id);

                    context.Load(baseListItem);
                    context.Load(baseList);
                    context.ExecuteQuery();

                    if (baseListItem == null) continue;

                    string baseFolderName = $"{baseList.Title}_ID_{id}";   //Tickets_ID_1

                    //getting backup list
                    List backupList = context.Web.Lists.GetByTitle("Backups");

                    if (backupList.RootFolder.FolderExists(baseFolderName) == false) continue;

                    FolderCollection backupListFolders = backupList.RootFolder.Folders;

                    Folder baseFolder = null; //Tickets_ID_1
                    foreach (Folder folder in backupListFolders)
                    {
                        if (folder.Name == baseFolderName)
                        {
                            baseFolder = folder;
                            break;
                        }

                    }

                    Folder parentFolder = Utils.createListFolder(baseFolder, $"{targetList.Title}", context);
                    context.Load(parentFolder);
                    context.ExecuteQuery();

                    string currFolderName = $"{targetList.Title}_ID_{item.Id}";
                    Folder currFolder = Utils.createListFolder(parentFolder, currFolderName, context);
                    context.Load(currFolder); context.ExecuteQuery();

                    //creating attachments
                    Folder attachmentFolder = Utils.createListFolder(currFolder, "Attachments", context);
                    if (attachmentFolder != null)
                    {
                        AttachmentCollection attachments = item.AttachmentFiles;
                        context.Load(attachments);
                        context.ExecuteQuery();
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

                                if (res == -1)
                                {
                                    clientResultStream.Close();
                                    continue;

                                }
                                using (Stream stream = new MemoryStream(buffer))
                                {
                                    FileCreationInformation creationInfo = new FileCreationInformation();
                                    creationInfo.ContentStream = stream;
                                    creationInfo.Url = attachment.FileName;
                                    creationInfo.Overwrite = true;
                                    File isFileCreated = attachmentFolder.Files.Add(creationInfo);

                                    context.Load(isFileCreated);
                                    context.ExecuteQuery();
                                }
                                Utils.printSuccess("File created");
                            }


                        }
                    }


                    //attachment code ends here

                    //adding JSON File in root folder

                    ItemVersions currItemVersion = Utils.GetItemVersions(item, context);

                    Console.WriteLine($"total Version : {currItemVersion.versions.Count}");


                    string JsonFile = JsonConvert.SerializeObject(currItemVersion);
                    string fileName = $" {targetList.Title}_ID_{item.Id}.json";

                    FileCreationInformation fileInfo = new FileCreationInformation();
                    byte[] bytes = ASCIIEncoding.ASCII.GetBytes(JsonFile);
                    using (Stream stream = new MemoryStream(bytes))
                    {
                        fileInfo.ContentStream = stream;
                        fileInfo.Url = fileName;
                        fileInfo.Overwrite = true;
                        File isJsonFileCreated = currFolder.Files.Add(fileInfo);
                        context.Load(isJsonFileCreated);
                        context.ExecuteQuery();
                    }

                    //adding json code ends here

                }
            }
        }
    }
}


