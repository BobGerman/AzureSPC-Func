using System;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.Azure.WebJobs.Host;
using Newtonsoft.Json;

namespace AddCommentCS
{
    public static class GetProductName
    {
        #region Product Names

        private static readonly string[] FIRST_NAMES = { "Azure", "Info", "Live", "Office", "One", "Power", "Response", "Share", "Visual" };
        private static readonly string[] LAST_NAMES = { " Analytics", " Apps", "Drive", "Flow", " Manager", "Office", "Path", "Point", "Shell", " Studio", "View" };
        private static readonly string[] SUFFIX = { "", "", "2018", "2019", "Essentials", "for Business", "for Office", "Framework", "Foundation", "Professional", "Server", "Services", "Ultimate", "Update" };

        #endregion

        [FunctionName("GetProductName")]
        public static async Task<HttpResponseMessage> Run([HttpTrigger(AuthorizationLevel.Function, "get", Route = null)]HttpRequestMessage req, TraceWriter log)
        {
            // Randomly generate the name
            Random r = new Random();
            string firstName = FIRST_NAMES[r.Next(FIRST_NAMES.Length)];
            string lastName = LAST_NAMES[r.Next(FIRST_NAMES.Length)];
            string suffix = SUFFIX[r.Next(FIRST_NAMES.Length)];
            string fullName = (firstName + lastName + " " + suffix).Trim();

            var result = new productNameResult { name= fullName };
            var resultJson = JsonConvert.SerializeObject(
                result, Formatting.Indented);

            return new HttpResponseMessage(HttpStatusCode.OK)
            {
                Content = new StringContent(resultJson, Encoding.UTF8, "application/json")
            };
        }
    }

    public class productNameResult
    {
        public string name { get; set; }
    }
}
