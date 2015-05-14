using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Net;

namespace GoogleTrendsCombinator
{
    public class CookieAwareWebClient : WebClient
    {
        public CookieAwareWebClient()
            : this(new CookieContainer())
        {
        }

        public CookieContainer cookies { get; set; }

        public CookieAwareWebClient(CookieContainer c)
        {
            this.cookies = c;
            this.Headers.Add("User-Agent: Mozilla/5.0 (Windows NT 6.1) AppleWebKit/536.5 (KHTML, like Gecko) Chrome/19.0.1084.52 Safari/536.5");
        }

        protected override WebRequest GetWebRequest(Uri address)
        {
            WebRequest request = base.GetWebRequest(address);

            if (request is HttpWebRequest)
                (request as HttpWebRequest).CookieContainer = this.cookies;

            return request;
        }
    }
}
