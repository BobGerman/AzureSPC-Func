using Microsoft.IdentityModel.Clients.ActiveDirectory;
using Microsoft.SharePoint.Client;
using System;
using System.IO;
using System.Security.Cryptography.X509Certificates;
using System.Threading.Tasks;

namespace AddCommentCS
{
    public static class CSOMHelper
    {
        private static string ClientId = Environment.GetEnvironmentVariable("ClientId");
        private static string Cert = "AddComment.pfx";
        private static string CertPassword = Environment.GetEnvironmentVariable("CertPassword");
        private static string TenantName = Environment.GetEnvironmentVariable("TenantName");
        private static string Authority = "https://login.windows.net/" + TenantName + ".onmicrosoft.com/";
        private static string Resource = "https://" + TenantName + ".sharepoint.com";

        public async static Task<ClientContext> GetClientContext(string siteUrl)
        {
            var authenticationContext = new AuthenticationContext(Authority, false);

            var certPath = Environment.GetEnvironmentVariable("LOCALHOME");
            if (string.IsNullOrEmpty(certPath))
            {
                certPath = Path.Combine(Environment.GetEnvironmentVariable("HOME"), "site\\wwwroot\\", Cert);
            }
            var cert = new X509Certificate2(System.IO.File.ReadAllBytes(certPath),
            CertPassword,
            X509KeyStorageFlags.Exportable |
            X509KeyStorageFlags.MachineKeySet |
            X509KeyStorageFlags.PersistKeySet);

            var authenticationResult = await authenticationContext.AcquireTokenAsync(Resource, new ClientAssertionCertificate(ClientId, cert));
            var token = authenticationResult.AccessToken;

            var ctx = new ClientContext(siteUrl);
            ctx.ExecutingWebRequest += (s, e) =>
            {
                e.WebRequestExecutor.RequestHeaders["Authorization"] = "Bearer " + authenticationResult.AccessToken;
            };

            return ctx;
        }
    }
}
