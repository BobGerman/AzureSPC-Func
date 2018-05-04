using System;
using System.Net;
using System.Net.Http;
using System.Threading.Tasks;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.Azure.WebJobs.Host;
using Microsoft.SharePoint.Client;

namespace AddCommentCS
{
    public static class AddCommentCS
    {
        [FunctionName("AddCommentCS")]
        public static async Task<HttpResponseMessage> Run([HttpTrigger(AuthorizationLevel.Function, "post", Route = null)]HttpRequestMessage req, TraceWriter log)
        {
            string result = null;
            try
            {
                log.Info("Executing function");

                // Get request body
                dynamic data = await req.Content.ReadAsAsync<object>();

                // Set name to query string or body data
                string siteUrl = data?.siteUrl;
                string comment = data?.comment;

                if (!string.IsNullOrEmpty(siteUrl) && !string.IsNullOrEmpty(comment))
                {
                    // Get Office Online (WOPI) URL
                    using (var ctx = await CSOMHelper.GetClientContext(siteUrl))
                    {
                        var web = ctx.Web;
                        ctx.Load(web);

                        var list = ctx.Web.Lists.GetByTitle("Comments");
                        var itemCreateInfo = new ListItemCreationInformation();
                        var item = list.AddItem(itemCreateInfo);
                        item["Title"] = comment;
                        item.Update();

                        ctx.ExecuteQuery();

                        result = "POSTED to " +  web.Title;
                    }
                }
            }
            catch (Exception ex)
            {
                return req.CreateErrorResponse(HttpStatusCode.BadRequest, "Error: " + ex.Message);
            }

            return result == null
                ? req.CreateResponse(HttpStatusCode.BadRequest, "Please pass a site and comment in your request")
                : req.CreateResponse(HttpStatusCode.OK, result);

            //log.Info("C# HTTP trigger function processed a request.");

            //// parse query parameter
            //string name = req.GetQueryNameValuePairs()
            //    .FirstOrDefault(q => string.Compare(q.Key, "name", true) == 0)
            //    .Value;

            //// Get request body
            //dynamic data = await req.Content.ReadAsAsync<object>();

            //// Set name to query string or body data
            //name = name ?? data?.name;

            //return name == null
            //    ? req.CreateResponse(HttpStatusCode.BadRequest, "Please pass a name on the query string or in the request body")
            //    : req.CreateResponse(HttpStatusCode.OK, "Hello " + name);
        }
    }
}
