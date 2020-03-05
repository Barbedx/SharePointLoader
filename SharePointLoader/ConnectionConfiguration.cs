using System.Collections.Generic;
namespace SharePointLoader
{
    class ConnectionConfiguration
    {
        public string UserName { get; set; }
        public string Password { get; set; }
        public string SiteName { get; set; }
        public List<string> FileLinks { get; set; }
        public string DestinationFolder { get; set; }
    }

}
