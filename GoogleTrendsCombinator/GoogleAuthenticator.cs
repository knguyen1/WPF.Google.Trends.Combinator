using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Net.Http;
using System.Net;
using System.IO;
using System.Text.RegularExpressions;
using System.Collections.Specialized;

namespace GoogleTrendsCombinator
{
    public class GoogleAuthenticator
    {
        private string _username = string.Empty;
        private string _password = string.Empty;
        private readonly CookieAwareWebClient _webClient;
        private bool _isLoggedIn = false;

        public GoogleAuthenticator(string username, string password, CookieAwareWebClient client)
        {
            this._username = username;
            this._password = password;
            this._webClient = client;
        }

        public bool Authenticate()
        {
            string galx = GetGalX();
            return Login(galx);
        }

        private string GetGalX()
        {
            string galx = null;
            string response = null;

            try
            {
                response = _webClient.DownloadString(Settings1.Default.loginUrl);

                //from login page
                Match match = Regex.Match(response, "<input.*name=\"GALX\".*[ \n\t]+value=\"([a-zA-Z0-9_-]+)\".*>");
                if (!match.Success)
                    throw new Exception("Cannot parse GALX!");

                galx = match.Groups[0].ToString();
            }
            catch (WebException exc)
            {
                throw exc;
            }
            catch (IOException exc)
            {
                throw exc;
            }

            return galx;
        }

        private bool Login(string galx)
        {
            _isLoggedIn = false;

            var formPost = new NameValueCollection()
            {
                {"Email", _username},
                {"Passwd", _password},
                {"GALX", galx}
            };

            try
            {
                _webClient.UploadValues("", formPost);
            }
            catch (WebException exc)
            {
                throw exc;
            }

            _isLoggedIn = true;

            return _isLoggedIn;
        }
    }
}
