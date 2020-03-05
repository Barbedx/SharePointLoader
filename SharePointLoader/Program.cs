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
            AppDomain.CurrentDomain.UnhandledException += (s, e) => logger.Error( e.ExceptionObject.ToString(), "Uhandled exception in application");

            try
            {
                Configuration = JsonSettings.Get<ConnectionConfiguration>();
                logger.Trace("Configuration loaded succesfull");

            }
            catch (Exception ex)
            {
                logger.Error(ex, "Error when loading configuration");
                return;
            }
            if (string.IsNullOrWhiteSpace(Configuration.SiteName) ||
                string.IsNullOrWhiteSpace(Configuration.UserName) ||
                string.IsNullOrWhiteSpace(Configuration.Password) ||
                !Configuration.FileLinks.Any()
            )
            {
                logger.Error("Configuration uncorrect, please provide next data:" + Environment.NewLine
                    + (string.IsNullOrWhiteSpace(Configuration.SiteName) ? "SiteName" + Environment.NewLine : string.Empty)
                    + (string.IsNullOrWhiteSpace(Configuration.UserName) ? "UserName" + Environment.NewLine : string.Empty)
                    + (string.IsNullOrWhiteSpace(Configuration.Password) ? "Password" + Environment.NewLine : string.Empty)
                    + (!Configuration.FileLinks.Any() ? "File links" : string.Empty)
                    );
                return;
            }

            //create directory for files, 
            //will create absolute or relative directory
            Directory.CreateDirectory(Configuration.DestinationFolder);
            var creds = GetSharePointCreds();
            foreach (var fileLink in Configuration.FileLinks)
            {
                try
                {
                    DownloadFilesFromSharePoint(Configuration.SiteName,
                        fileLink, Configuration.DestinationFolder,
                       creds
                        );
                    logger.Trace($"File {Path.GetFileName(fileLink)} loaded from share");
                }
                catch (Exception ex)
                {
                    logger.Error(ex, $"Cannot load file {Path.GetFileName(fileLink)}");
                }

            }
        }

        private static SharePointOnlineCredentials GetSharePointCreds()
        {
            SecureString securePassword = new SecureString();
            for (int i = 0; i < Configuration.Password.Length; i++)
            {
                securePassword.AppendChar(Configuration.Password[i]);
            }
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
                FileInformation fInfo = SPClient.File.OpenBinaryDirect(ctx, fileLink);
                //Console.WriteLine(web.Title);Directory.GetCurrentDirectory()

                var destPath = Path.Combine(localTempLocation, Path.GetFileName(fileLink));
                using (var sReader = new StreamReader(fInfo.Stream))
                {
                    using (var sWriter = new StreamWriter(destPath))
                    {
                        sWriter.Write(sReader.ReadToEnd());
                    }
                }
            };
        }
    }

}
