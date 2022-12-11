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
    internal class BackupGenerator
    {
        protected string url;
        protected string username;
        protected string password;
        protected string targetListTitle;
        protected string backupDocTitle;
        protected SecureString securedPassword;
        public BackupGenerator(string url, string username, string password, string targetListTitle, string backupDocTitle)
        {
            this.url = url;
            this.username = username;
            this.password = password;
            this.targetListTitle = targetListTitle;
            this.backupDocTitle = backupDocTitle;

            securedPassword = Utils.getSecuredPass(password);
        }

        //starts
        public bool CreateBackup()
        {
            using (ClientContext context = new ClientContext(url))
            {

                context.Credentials = new SharePointOnlineCredentials(username, securedPassword);
                List targetedList = context.Web.Lists.GetByTitle(targetListTitle);
                CamlQuery query = CamlQuery.CreateAllItemsQuery();
                ListItemCollection items = targetedList.GetItems(query);
                context.Load(items);
                context.ExecuteQuery();

                if (items == null) return false;

                foreach (ListItem item in items)
                {
                    int itemID = item.Id;
                    string createdDateTime = (string)item["Created_x0020_Date"];

                    string folderName = $"{targetListTitle}_ID_{itemID}";

                    //getting and creating backup folder
                    List backupList = context.Web.Lists.GetByTitle(backupDocTitle);
                    Folder parentFolder = Utils.createListFolder(backupList.RootFolder, folderName, context);

                    if (parentFolder == null) continue;





                    //adding all the attachments files in attachment folder
                    int CreatedAttachments = Utils.CopyAttachments(parentFolder, item, context);
                    Utils.printSuccess($"total attachments created : {CreatedAttachments}");


                    //adding JSON File in root folder

                    bool jsonCreated = Utils.AddJSON(parentFolder, item, targetListTitle, context);
                    if (!jsonCreated) continue;
                    
                }

            }

            /*
            ForSubFolders("Tasks", "Tickets");
            ForSubFolders("Email", "Tickets");
            Console.WriteLine("Function ends here");
            Console.ReadKey();
            return;
            */


            return false;
        }//creatbackup function

        //ends



































    }//class


}//namespace


/*
                    Folder attachmentFolder = Utils.createListFolder(parentFolder, "Attachments", context);
                    if (attachmentFolder == null) continue;
                    
                    
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

                    */





/*
                    ItemVersions currItemVersion = Utils.GetItemVersions(item, context);

                    Console.WriteLine($"total Version : {currItemVersion.versions.Count}");


                    string JsonFile = JsonConvert.SerializeObject(currItemVersion);
                    string fileName = $" {targetListTitle}_ID_{itemID}.json";

                    FileCreationInformation fileInfo = new FileCreationInformation();
                    byte[] bytes = ASCIIEncoding.ASCII.GetBytes(JsonFile);
                    using (Stream stream = new MemoryStream(bytes))
                    {
                        fileInfo.ContentStream = stream;
                        fileInfo.Url = fileName;
                        fileInfo.Overwrite = true;
                        File isJsonFileCreated = parentFolder.Files.Add(fileInfo);
                        context.Load(isJsonFileCreated);
                        context.ExecuteQuery();
                    }
                    */
