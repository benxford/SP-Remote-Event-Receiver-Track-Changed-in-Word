# SP Remote Event Receiver - Track Changed in Word

## Setup notes
1. Create a new App Registration in Azure
2. Remove all API permissions for the app registration, and generate a client secret
3. Add the ClientId and ClientSecret to web.config
4. Publish the web project
5. Register the app in SharePoint by navigating to https://contoso.sharepoint.com/_layouts/appinv.aspx
	Permissions XML:
		<AppPermissionRequests AllowAppOnlyPolicy="true" >
		  <AppPermissionRequest Scope="http://sharepoint/content/sitecollection/web/list" Right="Write" />
		</AppPermissionRequests>
6. Register the remote event receiver
	let siteUrl = "/sites/dev";
	let listName = "Documents";
	let azureUrl = "https://APPHOST.azurewebsites.net";

	// Get form digest value
	fetch(siteUrl + "/_api/contextinfo", { headers: { Accept: "application/json; odata=nometadata", "Content-Type": "application/json; odata=nometadata" }, method: "POST" })
	.then((r) => r.json())

	// Add event receiver
	.then((r) => fetch(siteUrl + "/_api/lists/getbytitle('" + listName + "')/eventreceivers", { headers: { Accept: "application/json; odata=nometadata", "Content-Type": "application/json; odata=nometadata", "X-RequestDigest": r.FormDigestValue }, method: "POST", body: JSON.stringify({
    		// Item added
		EventType: 10001,
		ReceiverUrl: azureUrl + "/services/WordEventReceiver.svc",
		ReceiverName: "Word-TrackChanges",
		// Default
		Synchronization: 0
	})}));
