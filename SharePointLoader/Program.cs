using Microsoft.SharePoint.Client;
using NLog;
using System;
using System.IO;
using System.Linq;
using System.Security; 
using SPClient = Microsoft.SharePoint.Client;
namespace SharePointLoader
{
    class Program
    {
        public object ServerURL { get; private set; }
        private static ConnectionConfiguration Configuration;
        private static Logger logger = LogManager.GetCurrentClassLogger();

        static void Main(string[] args)
        {


            AppDomain.CurrentDomain.UnhandledException += (s, e) => logger.Fatal(e.ExceptionObject.ToString(), "Uhandled exception in application");

            try
            {
                Configuration = JsonSettings.Get<ConnectionConfiguration>();
                logger.Trace("Configuration loaded succesfull");

            }
            catch (Exception ex)
            {
                logger.Fatal(ex, "Error when loading configuration");
                Environment.Exit(-1);
            }

            if (args.Count() == 2 &&
                !string.IsNullOrWhiteSpace(args[0]) &&
                !string.IsNullOrWhiteSpace(args[1]))
            {
                Configuration.UserName = args[0];
                Configuration.Password = args[1];
            }


            if (!Configuration.Sites.Any() ||
                string.IsNullOrWhiteSpace(Configuration.UserName) ||
                string.IsNullOrWhiteSpace(Configuration.Password) ||
                !Configuration.Sites.SelectMany(x => x.FileLinks).Any()
            )
            {

                logger.Error("Configuration uncorrect, please provide next data:" + Environment.NewLine
                    + (!Configuration.Sites.Any() ? "SiteName" + Environment.NewLine : string.Empty)
                    + (string.IsNullOrWhiteSpace(Configuration.UserName) ? "UserName" + Environment.NewLine : string.Empty)
                    + (string.IsNullOrWhiteSpace(Configuration.Password) ? "Password" + Environment.NewLine : string.Empty)
                    + (!Configuration.Sites.SelectMany(x => x.FileLinks).Any() ? "File links" : string.Empty)
                    );
                Environment.Exit(-2);

            }

            //create directory for files, 
            //will create absolute or relative directory
            Directory.CreateDirectory(Configuration.DestinationFolder);
            var creds = GetSharePointCreds();
            int errorsCount = 0;
            foreach (var site in Configuration.Sites)
            {

                foreach (var fileLink in site.FileLinks)
                {

                    try
                    {
                        logger.Trace($"Try to load file {fileLink}");

                        DownloadFilesFromSharePoint(site.SiteName,fileLink, Configuration.DestinationFolder,creds);
                        logger.Trace($"File {Path.GetFileName(fileLink)} loaded from share");
                    }
                    catch (Exception ex)
                    {
                        logger.Error(ex, $"Cannot load file {Path.GetFileName(fileLink)}");
                        errorsCount++;
                    }
                }
            }
            if (errorsCount>0)
            {
                logger.Fatal($"Errors occurred when downloading {errorsCount} of {Configuration.Sites.SelectMany(x => x.FileLinks)} files!");
                Environment.Exit(-3);
            }
            logger.Trace($"All files downloaded successfully!");

        }

        private static SharePointOnlineCredentials GetSharePointCreds()
        {
            SecureString securePassword = new SecureString();
            Configuration.Password.ToList().ForEach(securePassword.AppendChar);
            return new SharePointOnlineCredentials(Configuration.UserName, securePassword);
        }

        static void DownloadFilesFromSharePoint(string siteUrl, string fileLink,
            string localTempLocation, SharePointOnlineCredentials credentials)
        {
            SecureString securePassword = new SecureString();
            for (int i = 0; i < Configuration.Password.Length; i++)
            {
                securePassword.AppendChar(Configuration.Password[i]);
            }

            using (ClientContext ctx = new ClientContext(siteUrl))
            {
                ctx.Credentials = credentials;

                try
                {
                    using (FileInformation fInfo = SPClient.File.OpenBinaryDirect(ctx, fileLink))
                    {
                        using (var sReader = new BinaryReader(fInfo.Stream))
                        {
                            var destPath = Path.Combine(localTempLocation, Path.GetFileName(fileLink));
                            using (var sWriter = new FileStream(destPath, FileMode.OpenOrCreate, FileAccess.Write))
                            {
                                byte[] buffer = new byte[16 * 1024];
                                using (MemoryStream ms = new MemoryStream())
                                {
                                    int read;
                                    while ((read = sReader.Read(buffer, 0, buffer.Length)) > 0)
                                    {
                                        ms.Write(buffer, 0, read);
                                    }
                                    ms.Position = 0;
                                    ms.CopyTo(sWriter);
                                    ms.Flush();
                                }
                                sWriter.Flush();
                            }
                            fInfo.Stream.Flush();
                        }
                    }
                }
                catch (Exception ex)
                {
                    logger.Fatal(ex, "Can't read file");
                    throw;
                }
            }
        }
    }
}
