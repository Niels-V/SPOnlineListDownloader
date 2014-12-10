using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security;
using System.Text;
using System.Threading.Tasks;
using log4net;
using Microsoft.SharePoint.Client;

namespace SPOnlineListDownloader
{
    class Program
    {
        static string SiteUrl { get; set; }
        static string UserName { get; set; }
        static SecureString Password { get; set; }
        static string LocalRootFolder { get; set; }

        private static readonly ILog log = LogManager.GetLogger(typeof(Program));

        static int Main(string[] args)
        {
            try
            {
                if (!ReadArguments(args))
                {
                    return 1;
                }
                // Starting with ClientContext, the constructor requires a URL to the 
                // server running SharePoint. 
                using (ClientContext context = new ClientContext(SiteUrl))
                {
                    context.Credentials = new SharePointOnlineCredentials(UserName, Password);

                    //Load Libraries from SharePoint
                    context.Load(context.Web.Lists);
                    context.ExecuteQuery();
                    log.InfoFormat("Found {0} lists", context.Web.Lists.Count);
                    foreach (List list in context.Web.Lists)
                    {
                        log.InfoFormat("Processing list {0}", list.Title);
                        bool docLib = (list.BaseType == BaseType.DocumentLibrary);

                        LocalList localList = new LocalList();
                        localList.Title = list.Title;

                        FillLocalList(context, list, localList);
                        log.InfoFormat("Downloaded listdata. Found {0} list items", localList.Items.Count);

                        string localFilePath = LocalFileLocation(list.ParentWebUrl, false);

                        localFilePath = Path.Combine(localFilePath, localList.Title + ".csv");

                        SaveLocalListToCSV(localList, true, localFilePath);
                        log.InfoFormat("Saved CSV {0}", localFilePath);

                        if (docLib)
                        {
                            log.Info("Downloading doclib items");
                            SaveFilesFromList(context, localList, LocalRootFolder);
                            log.Info("Downloading doclib items done!");
                        }
                    }
                }
                log.InfoFormat("Finished successful for site {0}", SiteUrl);
                return 0;
            }
            catch (Exception ex)
            {
                log.Error("Terminated with error for site " + SiteUrl, ex);
                return 2;
            }
        }

        /// <summary>
        /// Determines a local file storage location for a server relative file or directory, and ensures that local directory is created
        /// </summary>
        /// <param name="serverRelativeUrl"></param>
        /// <returns></returns>
        private static string LocalFileLocation(string serverRelativeUrl, bool isFile)
        {
            string localFilePath = Path.Combine(LocalRootFolder, serverRelativeUrl.Replace('/', Path.DirectorySeparatorChar).TrimStart('\\'));
            string directoryPath = localFilePath;
            if (isFile)
            {
                directoryPath = Path.GetDirectoryName(localFilePath);
            }
            Directory.CreateDirectory(directoryPath);
            return localFilePath;
        }

        private static bool ReadArguments(string[] args)
        {
            if (args.Length != 4)
            {
                Console.WriteLine("Usage: SPOnlineListDownloader.exe [siteUrl] [username] [password] [localFolderStore]");
                log.Warn("Started with invalid arguments");
                return false;
            }
            SiteUrl = args[0];
            UserName = args[1];
            Password = new SecureString();
            foreach (char c in args[2].ToCharArray()) Password.AppendChar(c);
            Password.MakeReadOnly();
            LocalRootFolder = args[3];
            log.InfoFormat("Started. Downloading site {0} to {1}", SiteUrl, LocalRootFolder);
            return true;
        }

        private static void SaveFilesFromList(ClientContext context, LocalList list, string localRootPath)
        {
            if (list.Items.Count == 0) { return; }
            if (list.Items.First().Fields.ContainsKey("FileLeafRef") && list.Items.First().Fields.ContainsKey("FileRef")) {
                foreach (var listItem in list.Items)
                {
                    if (!string.IsNullOrEmpty(listItem.Fields["FileLeafRef"]) && listItem.Fields["ContentTypeId"].StartsWith("0x0101"))
                    {
                        //Document ContentTypeID
                        log.InfoFormat("Downloading {0} ", listItem.Fields["FileLeafRef"]);
                        string localFilePath = LocalFileLocation(listItem.Fields["FileRef"],true);

                        if (!System.IO.File.Exists(localFilePath))
                        {
                            Microsoft.SharePoint.Client.FileInformation f = Microsoft.SharePoint.Client.File.OpenBinaryDirect(context, listItem.Fields["FileRef"]);

                            using (var fileStream = new FileStream(localFilePath, FileMode.Create))
                            {
                                f.Stream.CopyTo(fileStream);
                            }
                            log.Info(" Completed!");
                        }
                        else
                        {
                            log.Info(" already existed!");
                        }
                    }
                }
            }
        }

        private static void FillLocalList(ClientContext context, List list, LocalList localList)
        {
            //Create a itempos
            ListItemCollectionPosition itemPosition = null;

            while (true)
            {
                // This creates a CamlQuery that has a RowLimit of 100, and also specifies Scope="RecursiveAll" 
                // so that it grabs all list items, regardless of the folder they are in. 
                CamlQuery query = CamlQuery.CreateAllItemsQuery(100);
                query.ListItemCollectionPosition = itemPosition;
                ListItemCollection items = list.GetItems(query);

                // Retrieve all items in the ListItemCollection from List.GetItems(Query). 
                context.Load(items, itema => itema.Include(i => i.FieldValuesAsText), itemb => itemb.ListItemCollectionPosition);
                context.ExecuteQuery();

                itemPosition = items.ListItemCollectionPosition;

                foreach (ListItem listItem in items)
                {
                    localList.Items.AddLast(new LocalListItem(listItem.FieldValuesAsText.FieldValues));
                }
                if (itemPosition == null)
                {
                    break;
                }
            }
        }

        public static bool SaveLocalListToCSV(LocalList list, bool includeHeader, string fileName)
        {
            if (list == null || list.Items.Count==0 || System.IO.File.Exists(fileName)) return false;


            using (StreamWriter writer = new StreamWriter(fileName, false))
            {
                if (writer == null) return false;

                if (includeHeader)
                {
                    string[] columnNames = list.Items.First().Fields.Keys.Select(column => column == null ? string.Empty : "\"" + column.Replace("\n", string.Empty).Replace("\"", "\"\"") + "\"").ToArray<string>();
                    writer.WriteLine(String.Join(",", columnNames));
                    writer.Flush();
                }

                foreach (LocalListItem row in list.Items)
                {
                    string[] fields = row.Fields.Values.Select(field => field==null? string.Empty : "\"" + field.ToString().Replace("\n",string.Empty).Replace("\"", "\"\"") + "\"").ToArray<string>();
                    writer.WriteLine(String.Join(",", fields));
                    writer.Flush();
                }
            }
            return true;
        }
    }
}
