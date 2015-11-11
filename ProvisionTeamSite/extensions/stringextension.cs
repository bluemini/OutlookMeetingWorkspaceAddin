using System;
using System.Collections.Generic;
using System.Linq;
using System.Security;
using System.Text;
using System.Threading.Tasks;

namespace ProvisionTeamSite.extensions
{
	public static class stringextension
	{
		public static SecureString ToSecureString(this string normalString)
		{
			SecureString ss = new SecureString();
			foreach (char c in normalString.ToCharArray()) ss.AppendChar(c);
			return ss;
		}
	}
}
