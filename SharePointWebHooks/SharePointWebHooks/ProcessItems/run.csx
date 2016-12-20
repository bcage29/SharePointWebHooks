using System;
using System.Linq;
using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core;
public static async Task Run(WebHookPayload myQueueItem, TraceWriter log)
{
    log.Info($"Subscription Id: {myQueueItem.SubscriptionId}");
    var internalClientState = "myextrasecurity";
    if (!myQueueItem.ClientState.Equals(internalClientState, StringComparison.CurrentCultureIgnoreCase))
    {
        log.Info($"internal client state {internalClientState} does not match supplied client state {myQueueItem.ClientState}");
        return;
    }

    var ctx = GetSharePointContext(myQueueItem.SiteUrl);

    var listId = new Guid(myQueueItem.Resource);
    var list = ctx.Web.Lists.GetById(listId);
    log.Info("Fetching last change token");
    var changeTokenStringValue = list.GetPropertyBagValueString("ChangeToken", string.Empty);

    ChangeToken lastChangeToken = null;
    if (!string.IsNullOrEmpty(changeTokenStringValue))
    {
        lastChangeToken = new ChangeToken { StringValue = changeTokenStringValue };
    }

    var changeQuery = new ChangeQuery(false, false)
    {
        Item = true,
        Add = true,
        RecursiveAll = true,
    };

    // Start pulling down the changes
    bool allChangesRead = false;
    do
    {
        changeQuery.ChangeTokenStart = lastChangeToken;

        // Execute the change query
        var changeCollection = list.GetChanges(changeQuery);
        ctx.Load(changeCollection);
        ctx.ExecuteQuery();

        if (changeCollection.Count > 0)
        {
            log.Info("Processing the new items");
            foreach (Change change in changeCollection)
            {
                lastChangeToken = change.ChangeToken;

                if (change is ChangeItem)
                {
                    // ProcessNewItem with the found change
                    await ProcessNewItem(ctx, list, change);
                }
            }

            // We potentially can have a lot of changes so be prepared to repeat the 
            // change query in batches of 'FetchLimit' untill we've received all changes
            if (changeCollection.Count < changeQuery.FetchLimit)
            {
                allChangesRead = true;
            }
        }
        else
        {
            allChangesRead = true;
        }
        // Are we done?
    } while (allChangesRead == false);

    log.Info("Finished processing items");
    list.SetPropertyBagValue("ChangeToken", lastChangeToken.StringValue);
}

public static ClientContext GetSharePointContext(string siteUrl)
{
    string url = $"https://rpbcage.sharepoint.com{siteUrl}";
    var clientId = System.Environment.GetEnvironmentVariable("ClientId", EnvironmentVariableTarget.Process);
    var clientSecret = System.Environment.GetEnvironmentVariable("ClientSecret", EnvironmentVariableTarget.Process);
    log.Info("authenticating to SharePoint");
    return new AuthenticationManager().GetAppOnlyAuthenticatedContext(url, clientId, clientSecret);
}

public static async Task ProcessNewItem(ClientContext ctx, List list, Change change)
{
    var itemId = (change as ChangeItem).ItemId;
    var item = list.GetItemById(itemId);
    ctx.Load(item, i => i["Title"], i => i["Status"]);
    ctx.ExecuteQueryRetry();

    if (item["Status"].ToString() == "Provisioned")
    return;

    await Task.Run(() =>
    {
        ctx.Web.CreateWeb(item["Title"].ToString(), item["Title"].ToString(), "", "STS#0", 1033);
    });

    item["Status"] = "Provisioned";
    item.Update();
    ctx.ExecuteQueryRetry();
}

public class WebHookPayload
{
    public string SubscriptionId { get; set; }
    public string ClientState { get; set; }
    public string ExpirationDateTime { get; set; }
    public string Resource { get; set; }
    public string TenantId { get; set; }
    public string SiteUrl { get; set; }
    public string WebId { get; set; }
}

/*

If you need to reference a dll in the function's bin folder

#r "Microsoft.IdentityModel.dll"
#r "Microsoft.IdentityModel.Extensions.dll" 

*/
