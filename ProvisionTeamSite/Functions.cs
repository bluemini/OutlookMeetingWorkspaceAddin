using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Azure.WebJobs;
using Microsoft.SharePoint.Client;
using Microsoft.Online.SharePoint.TenantAdministration;
using System.Configuration;
using ProvisionTeamSite.extensions;
using OfficeDevPnP.Core.Extensions;

namespace ProvisionTeamSite
{
	public class Functions
	{
		private static string RootUrl = "https://sponlinesite.sharepoint.com/sites/";
		private static string AdminUrl = "https://sponlinesite-admin.sharepoint.com";

		// This function will get triggered/executed when a new message is written 
		// on an Azure Queue called queue.
		public static void ProcessQueueMessage(
			[QueueTrigger("meetingworkspace")] MeetingWorkspaceMessage message, 
			TextWriter log)
		{
			log.WriteLine(message);

			if (!string.IsNullOrEmpty(message.WorkspaceName))
			{
				var creds = GetCredsFromConfig();
				log.WriteLine("Username: " + creds.Item1);
				log.WriteLine("Password: " + creds.Item2);

				string newWorkspaceUrl = RootUrl + message.WorkspaceName;
				log.WriteLine(newWorkspaceUrl);

				try {
					DeployNewSite(newWorkspaceUrl, creds);
				}
				catch (Exception e)
				{
					log.WriteLine("Error: unable to deploy new site. " + e.Message);
				}
			}
		}

		private static Tuple<string, string> GetCredsFromConfig()
		{
			var username = ConfigurationManager.AppSettings["username"];
			var password = ConfigurationManager.AppSettings["password"];
			return Tuple.Create<string, string>(username, password);
		}

		private static void DeployNewSite(string newSiteUrl, Tuple<string, string> creds)
		{
			using (ClientContext context = new ClientContext(AdminUrl))
			{
				SharePointOnlineCredentials adminCredentials = new SharePointOnlineCredentials(creds.Item1, creds.Item2.ToSecureString());
				context.Credentials = adminCredentials;

				var tenant = new Tenant(context);

				// Create new site collection with storage limits and settings from the form
				tenant.CreateSiteCollection(newSiteUrl,
											"Outlook Meeting Workspace",
											creds.Item1,
											"STS#0",
											(int) 20,
											(int) 10,
											3, // 3 is the timezone for Stockholm
											0,
											0,
											1033);

			}
		}
	}
}
