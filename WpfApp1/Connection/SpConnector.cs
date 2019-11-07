using System.Security;
using SPClient = Microsoft.SharePoint.Client;

namespace WpfApp1.Connection
{
    public class SpConnector
    {
        private SPClient.ClientContext ctx;
        
        /// <summary>
        /// Constructor with username and securepassword
        /// </summary>
        /// <param name="username"></param>
        /// <param name="password"></param>
        public SpConnector(string username, SecureString password)
        {
            string siteUrl = "https://meyer74labor.sharepoint.com/sites/Testsite/";

            ctx = new SPClient.ClientContext(siteUrl)
            {
                Credentials = new SPClient.SharePointOnlineCredentials(username, password)
            };

            SPClient.Web web = ctx.Web;
            ctx.Load(web);

            ctx.ExecuteQuery();

        }

        /// <summary>
        /// Property SPContext
        /// </summary>
        public SPClient.ClientContext SPContext
        {
            get { return ctx; }
        }
    }
}
