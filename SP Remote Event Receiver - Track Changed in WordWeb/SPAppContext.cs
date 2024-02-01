using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace SP_Remote_Event_Receiver___Track_Changed_in_WordWeb
{
    /// <summary>
    /// ClientContext configured to use app-only authentication. 
    /// </summary>
    public class SPAppContext : ClientContext
    {
        private System.Net.WebProxy Proxy = new System.Net.WebProxy();
        private string AccessToken { get; }

        public SPAppContext(string url) : this(new Uri(url)) { }
        public SPAppContext(Uri uri) : this(uri, TokenHelper.GetAppOnlyAccessToken(TokenHelper.SharePointPrincipal, uri.Authority, TokenHelper.GetRealmFromTargetUrl(uri)).AccessToken) { }
        public SPAppContext(string url, string accessToken) : this(new Uri(url), accessToken) { }
        public SPAppContext(Uri uri, string accessToken) : base(uri)
        {
            AuthenticationMode = ClientAuthenticationMode.Anonymous;
            FormDigestHandlingEnabled = false;
            AccessToken = accessToken;
        }

        protected override void OnExecutingWebRequest(WebRequestEventArgs args)
        {
            args.WebRequestExecutor.WebRequest.Proxy = Proxy;
            args.WebRequestExecutor.RequestHeaders["Authorization"] = "Bearer " + AccessToken;

            base.OnExecutingWebRequest(args);
        }
    }
}