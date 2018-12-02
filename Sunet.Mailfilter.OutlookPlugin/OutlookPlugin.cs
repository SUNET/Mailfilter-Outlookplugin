using Microsoft.Office.Core;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Web;
using System.Windows.Forms;
using Office = Microsoft.Office.Core;

namespace Sunet.Mailfilter.OutlookPlugin
{
    [ComVisible(true)]
    public class OutlookPlugin : Office.IRibbonExtensibility
    {
        private static string DefaultSpamButtonText = "Spam";
        private static string DefaultHamButtonText = "Non-spam";
        private static string DefaultForgetButtonText = "Forget";
        private static string DefaultForwardButtonText = "Forward to support";
        private static string DefaultButtonGroupText = "SUNET Mailfilter";
        private static string RegistryForwardingButtonText = "ForwardingButtonText";
        private static string RegistrySpamButtonText = "SpamButtonText";
        private static string RegistryHamButtonText = "HamButtonText";
        private static string RegistryForgetButtonText = "ForgetButtonText";
        private static string RegistryButtonGroupText = "ButtonGroupText";
        private static string RegistryForwardingAddress = "ForwardingAddress";
        private static string ShortSpamAction = "s";
        private static string ShortHamAction = "n";
        private static string ShortForgetAction = "f";

        private Office.IRibbonUI ribbon;
        private static string TransportMessageHeadersSchema = "http://schemas.microsoft.com/mapi/proptag/0x007D001E";
        private static string MimePropertySchema = "http://schemas.microsoft.com/mapi/string/{00020386-0000-0000-C000-000000000046}";
        private static string MimeHeaderSpam = "X-Antispam-Training-Spam";
        private static string MimeHeaderHam = "X-Antispam-Training-Nonspam";
        private static string MimeHeaderForget = "X-Antispam-Training-Forget";
        private static string MimeHeaderCanItId = "X-Canit-Stats-ID";
        private static string MimeHeaderCanItIdRegExp = @"(\w+)\s+-\s+(\w+)\s+-\s+(\w+).*";

        internal string ForwardingButtonText { get; set; }
        internal string SpamButtonText { get; set; }
        internal string HamButtonText { get; set; }
        internal string ForgetButtonText { get; set; }
        internal string ButtonGroupText { get; set; }

        public OutlookPlugin()
        {
        }

        /// <summary>
        /// Gets all the headers of a mailitem
        /// </summary>
        /// <param name="mailItem"></param>
        /// <returns></returns>
        private string GetAllHeaders(Microsoft.Office.Interop.Outlook.MailItem mailItem)
        {
            if (null != mailItem)
            {
                return (string)mailItem.PropertyAccessor.GetProperty(TransportMessageHeadersSchema);
            }

            return string.Empty;
        }

        /// <summary>
        /// Finds a single mime-header among all the headers
        /// </summary>
        /// <param name="mailItem"></param>
        /// <param name="header"></param>
        /// <returns></returns>
        private string GetHeader(Microsoft.Office.Interop.Outlook.MailItem mailItem, string header)
        {
            string headerValue = string.Empty;

            string allHeaders = GetAllHeaders(mailItem);
            if (!string.IsNullOrWhiteSpace(allHeaders))
            {
                var match = System.Text.RegularExpressions.Regex.Match(allHeaders, string.Format(CultureInfo.CurrentCulture, "{0}:[\\s]*(.*)[\\s]*\r\n", header));
                if (null != match && match.Groups.Count > 1)
                {
                    headerValue = match.Groups[1].Value.Trim(new char[] { ' ', '<', '>' });
                }
            }

            return headerValue;
        }

        /// <summary>
        /// Adds a header to a mailitem
        /// </summary>
        /// <param name="mailItem"></param>
        /// <param name="header"></param>
        /// <param name="value"></param>
        private void AddHeader(Microsoft.Office.Interop.Outlook._MailItem mailItem, string header, string value)
        {
            if (null != mailItem)
            {
                mailItem.PropertyAccessor.SetProperty(string.Format("{0}/{1}", MimePropertySchema, header), value);
            }
        }

        private void ReportMail(string mimeHeader, string action, string shortAction)
        {
            List<Tuple<string, Microsoft.Office.Interop.Outlook.MailItem>> headerItems = new List<Tuple<string, Microsoft.Office.Interop.Outlook.MailItem>>();
            int failedMessages = 0;

            try
            {
                using (var wc = new WebClientEx())
                {
                    // Use API if configured, otherwise call the url anonymously
                    if (null != Globals.ThisAddIn.ApiUrl && !string.IsNullOrEmpty(Globals.ThisAddIn.ApiUrl))
                    {
                        foreach (Microsoft.Office.Interop.Outlook.MailItem mailitem in Globals.ThisAddIn.Application.ActiveExplorer().Selection)
                        {
                            string queryData = string.Empty;

                            // Primarily use the canit mime header if using the api
                            var header = "";// GetHeader(mailitem, MimeHeaderCanItId);
                            if (!string.IsNullOrWhiteSpace(header))
                            {
                                var match = Regex.Match(header, MimeHeaderCanItIdRegExp, RegexOptions.IgnoreCase);
                                if (match.Groups.Count >= 4)
                                {
                                    queryData = string.Format("i={0}&m={1}&t={2}&c={3}", match.Groups[1], match.Groups[2], match.Groups[3], shortAction);
                                }
                            }

                            header = GetHeader(mailitem, mimeHeader);
                            if (string.IsNullOrWhiteSpace(queryData) && !string.IsNullOrWhiteSpace(header))
                            {
                                var uri = new Uri(header);
                                queryData = uri.Query.Substring(1);
                            }

                            if (!string.IsNullOrWhiteSpace(queryData))
                            {
                                headerItems.Add(new Tuple<string, Microsoft.Office.Interop.Outlook.MailItem>(queryData, SpamButtonText == action ? mailitem : null));
                            }
                        }

                        // No messages were selected
                        if (headerItems.Count <= 0)
                        {
                            System.Windows.Forms.MessageBox.Show("No traininginfo was available in the selected items.");
                            return;
                        }

                        var apiUri = new Uri(Globals.ThisAddIn.ApiUrl);

                        // Only log on once per click
                        wc.Headers[HttpRequestHeader.ContentType] = "application/x-www-form-urlencoded";
                        wc.UploadString(string.Format("{0}login", Globals.ThisAddIn.ApiUrl), string.Format("user={0}&password={1}", HttpUtility.UrlEncode(Globals.ThisAddIn.ApiUser), HttpUtility.UrlEncode(Globals.ThisAddIn.ApiPassword)));

                        foreach (var reportTuple in headerItems)
                        {
                            try
                            {
                                wc.Headers[HttpRequestHeader.ContentType] = "application/x-www-form-urlencoded";
                                wc.UploadString(string.Format("{0}vote", Globals.ThisAddIn.ApiUrl), reportTuple.Item1);
                                Globals.ThisAddIn.LogMessage(string.Format("Voting {0} (authenticated): {1}", action, reportTuple.Item1), string.Empty, System.Diagnostics.EventLogEntryType.Information);
                            }
                            catch (WebException we)
                            {
                                // Ignore 404 exceptions (spam entry not found)
                                if (we.Response is HttpWebResponse && (we.Response as HttpWebResponse).StatusCode == HttpStatusCode.NotFound)
                                {
                                    failedMessages++;
                                }
                                else
                                {
                                    throw;
                                }
                            }
                        }
                    }
                    else
                    {
                        foreach (Microsoft.Office.Interop.Outlook.MailItem mailitem in Globals.ThisAddIn.Application.ActiveExplorer().Selection)
                        {
                            var header = GetHeader(mailitem, mimeHeader);
                            if (!string.IsNullOrWhiteSpace(header))
                            {
                                headerItems.Add(new Tuple<string, Microsoft.Office.Interop.Outlook.MailItem>(header, SpamButtonText == action ? mailitem : null));
                            }
                        }

                        // No messages were selected
                        if (headerItems.Count <= 0)
                        {
                            System.Windows.Forms.MessageBox.Show("No traininginfo was available in the selected items.");
                            return;
                        }

                        foreach (var reportTuple in headerItems)
                        {
                            // TODO : rewrite url?
                            //var rewrittenUrl = string.Format("{0}://{1}{2}", apiUri.Scheme, apiUri.Host, uri.PathAndQuery);

                            try
                            {
                                var uri = new Uri(reportTuple.Item1);
                                var httpData = wc.DownloadString(reportTuple.Item1);
                                if (!httpData.ToLower().Contains(Globals.ThisAddIn.AnonymousMatchingString))
                                {
                                    Globals.ThisAddIn.LogMessage(string.Format("Unexpected result while voting anonymously at {0}", uri.Query.Substring(1)), string.Empty, System.Diagnostics.EventLogEntryType.Information);
                                    failedMessages++;
                                }
                                else
                                {
                                    Globals.ThisAddIn.LogMessage(string.Format("Voting {0} (anonymously): {1}", action, uri.Query.Substring(1)), string.Empty, System.Diagnostics.EventLogEntryType.Information);
                                }
                            }
                            catch (WebException we)
                            {
                                // Ignore 404 exceptions (spam entry not found)
                                if (we.Response is HttpWebResponse && (we.Response as HttpWebResponse).StatusCode == HttpStatusCode.NotFound)
                                {
                                    failedMessages++;
                                }
                                else
                                {
                                    throw;
                                }
                            }
                        }
                    }

                    foreach (var reportTuple in headerItems)
                    {
                        if (null != reportTuple.Item2)
                        {
                            try
                            {
                                if (null != Globals.ThisAddIn.JunkFolder)
                                {
                                    reportTuple.Item2.Move(Globals.ThisAddIn.JunkFolder);
                                }
                            }
                            catch (Exception)
                            {
                                // Quench error and continue
                            }
                        }
                    }

                    if (failedMessages > 0 && failedMessages < headerItems.Count)
                    {
                        System.Windows.Forms.MessageBox.Show("Some of the messages could not be voted on.");
                    }
                    else if (failedMessages > 0)
                    {
                        System.Windows.Forms.MessageBox.Show("None of the messages could not be voted on.");
                    }
                    else
                    {
                        System.Windows.Forms.MessageBox.Show(string.Format("The selected message(s) has been voted as {0}", action));
                    }
                }
            }
            catch (Exception ex)
            {
                Globals.ThisAddIn.LogMessage(ex.Message, ex.StackTrace);
                if (Globals.ThisAddIn.ShowPopups)
                {
                    System.Windows.Forms.MessageBox.Show("An error occurred, check the eventlog for detailed information.");
                }
            }
        }

        #region IRibbonExtensibility Members

        public string GetCustomUI(string ribbonID)
        {
            // Only add ribbon to the main window, not in individual mails
            if (ribbonID == "Microsoft.Outlook.Explorer")
            {
                var xmlData = GetResourceText("Sunet.Mailfilter.OutlookPlugin.OutlookPlugin.xml");
                // Workaround, patching the resouce XML with the configured registry texts since Outlook 2010 does not fire the getLabel events properly.
                // If Outlook 2010 compatibility is not required, the xml should define a getLabel callback and provide the texts through that instead.
                xmlData = xmlData.Replace("GroupPlaceholderLabel", ThisAddIn.LoadResourceString(RegistryButtonGroupText, DefaultButtonGroupText));
                this.SpamButtonText = ThisAddIn.LoadResourceString(RegistrySpamButtonText, DefaultSpamButtonText);
                xmlData = xmlData.Replace("SpamButtonPlaceholderLabel", this.SpamButtonText);
                this.HamButtonText = ThisAddIn.LoadResourceString(RegistryHamButtonText, DefaultHamButtonText);
                xmlData = xmlData.Replace("HamButtonPlaceholderLabel", this.HamButtonText);
                this.ForgetButtonText = ThisAddIn.LoadResourceString(RegistryForgetButtonText, DefaultForgetButtonText);
                xmlData = xmlData.Replace("ForgetButtonPlaceholderLabel", this.ForgetButtonText);

                // Remove forwarding button if a forwarding address has not been configured in the registry, otherwise set button text
                var forwardingAddress = ThisAddIn.LoadResourceString(RegistryForwardingAddress, string.Empty);
                var forwardingButtonText = ThisAddIn.LoadResourceString(RegistryForwardingButtonText, DefaultForwardButtonText);
                if (string.Empty == forwardingAddress || string.Empty == forwardingButtonText)
                {
                    // Find button definition in xml
                    var buttonMatch = Regex.Match(xmlData, "<.*\"ForwardButton\".*/>");
                    if (buttonMatch.Groups.Count > 0)
                    {
                        xmlData = xmlData.Replace(buttonMatch.Groups[0].ToString(), string.Empty);
                    }
                }
                else
                {
                    xmlData = xmlData.Replace("ForwardButtonPlaceholderLabel", forwardingButtonText);
                }

                return xmlData;
            }

            return null;
        }

        #endregion

        #region Ribbon Callbacks
        //Create callback methods here. For more information about adding callback methods, visit http://go.microsoft.com/fwlink/?LinkID=271226

        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        {
            this.ribbon = ribbonUI;
        }

        public System.Drawing.Bitmap getOutlookPluginImage(IRibbonControl control)
        {
            switch (control.Id)
            {
                case "SpamButton":
                    return Sunet.Mailfilter.OutlookPlugin.Properties.Resources.spam_icon;
                case "HamButton":
                    return Sunet.Mailfilter.OutlookPlugin.Properties.Resources.ham_icon;
                case "ForgetButton":
                    return Sunet.Mailfilter.OutlookPlugin.Properties.Resources.forget_icon;
                case "ForwardButton":
                    return Sunet.Mailfilter.OutlookPlugin.Properties.Resources.forward_icon;
            }

            return null;
        }

        public void spam_Click(IRibbonControl control)
        {
            this.ReportMail(MimeHeaderSpam, SpamButtonText, ShortSpamAction);
            
        }

        public void ham_Click(IRibbonControl control)
        {
            this.ReportMail(MimeHeaderHam, HamButtonText, ShortHamAction);
        }

        public void forget_Click(IRibbonControl control)
        {
            this.ReportMail(MimeHeaderForget, ForgetButtonText, ShortForgetAction);
        }

        public void forward_Click(IRibbonControl control)
        {
            try
            {
                if (Globals.ThisAddIn.Application.ActiveExplorer().Selection.Count <= 0)
                {
                    return;
                }
                MessageBoxControlBox controlUI = new MessageBoxControlBox();
                controlUI.ShowDialog();
                if (controlUI.DialogResult == true)
                {
                    Microsoft.Office.Interop.Outlook._MailItem mi = (Microsoft.Office.Interop.Outlook._MailItem)Globals.ThisAddIn.Application.CreateItem(Microsoft.Office.Interop.Outlook.OlItemType.olMailItem);
                    mi.Subject = Globals.ThisAddIn.ForwardingSubject;
                    mi.Body = Globals.ThisAddIn.ForwardingBody;
                    mi.Recipients.Add(Globals.ThisAddIn.ForwardingAddress);
                    foreach (Microsoft.Office.Interop.Outlook.MailItem mailitem in Globals.ThisAddIn.Application.ActiveExplorer().Selection)
                    {
                        while (mailitem.Attachments.Count > 0)
                        {
                            mailitem.Attachments.Remove(1);
                        }
                        mi.Attachments.Add(mailitem);
                        mi.Body = controlUI.Reason;
                    }

                    AddHeader(mi, Globals.ThisAddIn.ForwardingMimeHeader, Globals.ThisAddIn.ForwardingMimeValue);
                    mi.Send();

                    System.Windows.Forms.MessageBox.Show("The message(s) has been forwarded.");
                }
                else
                {
                    return;
                }

                
            }
            catch (Exception ex)
            {
                Globals.ThisAddIn.LogMessage(ex.Message, ex.StackTrace);
                if (Globals.ThisAddIn.ShowPopups)
                {
                    System.Windows.Forms.MessageBox.Show("An error occurred, check the eventlog for detailed information.");
                }
            }
        }

        #endregion

        #region Helpers

        private static string GetResourceText(string resourceName)
        {
            Assembly asm = Assembly.GetExecutingAssembly();
            string[] resourceNames = asm.GetManifestResourceNames();
            for (int i = 0; i < resourceNames.Length; ++i)
            {
                if (string.Compare(resourceName, resourceNames[i], StringComparison.OrdinalIgnoreCase) == 0)
                {
                    using (StreamReader resourceReader = new StreamReader(asm.GetManifestResourceStream(resourceNames[i])))
                    {
                        if (resourceReader != null)
                        {
                            return resourceReader.ReadToEnd();
                        }
                    }
                }
            }
            return null;
        }

        #endregion
    }
}
