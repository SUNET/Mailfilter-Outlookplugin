using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Text;
using System.Text.RegularExpressions;
using System.Web;
using System.Xml.Linq;

namespace Sunet.Mailfilter.OutlookPlugin
{
    class WebClientEx : WebClient
    {
        private CookieContainer _container = new CookieContainer();

        public WebClientEx()
        {
        }

        protected override WebRequest GetWebRequest(Uri address)
        {
            var r = base.GetWebRequest(address);
            var req = r as HttpWebRequest;
            if (null != req)
            {
                req.CookieContainer = _container;
            }

            return r;
        }

        protected override WebResponse GetWebResponse(WebRequest request, IAsyncResult result)
        {
            var r = base.GetWebResponse(request, result);
            ReadCookies(r);
            return r;
        }

        protected override WebResponse GetWebResponse(WebRequest request)
        {
            var r = base.GetWebResponse(request);
            ReadCookies(r);
            return r;
        }

        private void ReadCookies(WebResponse r)
        {
            var rr = r as HttpWebResponse;
            if (null != rr)
            {
                var c = rr.Cookies;
                foreach (Cookie cc in c)
                {
                    cc.Path = "/";
                }

                _container.Add(c);
            }
        }
    }
}
