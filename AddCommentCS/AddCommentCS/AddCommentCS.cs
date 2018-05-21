using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Host;
using Microsoft.SharePoint.Client;
using System;
using System.Threading.Tasks;

namespace AddCommentCS
{
    public static class AddCommentCS
    {
        // Timer - "0 0 */2 * * *" is chron for every 2 hours
        [FunctionName("AddCommentCS")]
        public static void Run([TimerTrigger("0 0 */2 * * *")]TimerInfo myTimer, TraceWriter log)
        {
            string siteUrl = Environment.GetEnvironmentVariable("SiteUrl");
            string comment = Environment.GetEnvironmentVariable("Comment");

            try
            {
                log.Info("Executing function");

                if (!string.IsNullOrEmpty(siteUrl) && !string.IsNullOrEmpty(comment))
                {
                    var t = Task.Run<ClientContext>(async () => await            CSOMHelper.GetClientContext(siteUrl));

                    using (var ctx = t.Result)
                    {
                        var web = ctx.Web;
                        ctx.Load(web);

                        var list = ctx.Web.Lists.GetByTitle("Comments");
                        var itemCreateInfo = new ListItemCreationInformation();
                        var item = list.AddItem(itemCreateInfo);
                        item["Title"] = comment;
                        item.Update();

                        ctx.ExecuteQuery();
                   }
                }
            }
            catch (Exception ex)
            {
                log.Error(ex.Message);
            }

            log.Info("Completed run");

        }
    }
}
