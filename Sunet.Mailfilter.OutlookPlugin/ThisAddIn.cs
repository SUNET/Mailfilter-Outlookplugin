using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;
using System.Diagnostics;
using System.Windows.Forms;
using System.Drawing;
using Microsoft.Win32;

namespace Sunet.Mailfilter.OutlookPlugin
{
    public partial class ThisAddIn
    {
        private static string RegistryPath = "Software\\Sunet.Mailfilter";
        private static string LogSource = "Mailfilter Plugin";
        private static string DefaultAnonymousMatchingString = "your vote has been registered";
        private static string DefaultForwardingSubject = "Sent via Mailfilter Plugin";
        private static string DefaultForwardingBody = "Sent via SUNET Mailfilter Outlook Plugin.";
        private static string DefaultForwardingPopup = "Message(s) has been sent to Support";
        private static string DefaultForwardingMimeHeader = "X-Antispam-Plugin";
        private static string DefaultForwardingMimeValue = "Outlook-add-in";
        private static string RegistryApiUrl = "ApiUrl";
        private static string RegistryApiUser = "ApiUser";
        private static string RegistryApiPassword = "ApiPassword";
        private static string RegistryAnonymousMatchingString = "AnonymousMatchingString";
        private static string RegistryForwardingSubject = "ForwardingSubject";
        private static string RegistryForwardingBody = "ForwardingBody";
        private static string RegistryForwardingPopup = "ForwardingPopup";
        private static string RegistryForwardingAddress = "ForwardingAddress";
        private static string RegistryForwardingMimeHeader = "ForwardingMimeHeader";
        private static string RegistryForwardingMimeValue = "ForwardingMimeValue";
        private static string RegistryBaseHost = "BaseHost";
        private static string RegistryShowPopups = "ShowPopups";

        internal string ApiUrl { get; set; }
        internal string ApiUser { get; set; }
        internal string ApiPassword { get; set; }
        internal string AnonymousMatchingString { get; set; }
        internal string ForwardingSubject { get; set; }
        internal string ForwardingBody { get; set; }
        internal string ForwardingPopup { get; set; }
        internal string ForwardingAddress { get; set; }
        internal string ForwardingMimeHeader { get; set; }
        internal string ForwardingMimeValue { get; set; }
        internal string BaseHost { get; set; }
        internal bool ShowPopups { get; set; }
        internal Outlook.MAPIFolder JunkFolder { get; set; }
        internal static Dictionary<string, string> registryCache = new Dictionary<string, string>();

        internal static string TryGetRegistryValue(RegistryHive registryHive, RegistryView registryView, string registryPath, string registryKey)
        {
            using (var view = RegistryKey.OpenBaseKey(registryHive, registryView))
            {
                if (null != view)
                {
                    using (var key = view.OpenSubKey(registryPath))
                    {
                        if (null != key)
                        {
                            return (string)key.GetValue(registryKey);
                        }
                    }
                }
            }

            return string.Empty;
        }

        internal static string GetRegistryValue(string key)
        {
            if (registryCache.ContainsKey(key))
            {
                return registryCache[key];
            }

            string value = string.Empty;

            // Try reading the configured value first from current user, and then local machine hives
            // Also try both 32-bit and 64-bit views of the registry. Registry64 will effectivly point to Registry32 if run
            // on a 32-bit platform.
            value = TryGetRegistryValue(RegistryHive.CurrentUser, RegistryView.Registry64, RegistryPath, key);
            if (null == value || string.IsNullOrEmpty(value))
            {
                value = TryGetRegistryValue(RegistryHive.LocalMachine, RegistryView.Registry64, RegistryPath, key);
            }
            if (null == value || string.IsNullOrEmpty(value))
            {
                value = TryGetRegistryValue(RegistryHive.CurrentUser, RegistryView.Registry32, RegistryPath, key);
            }
            if (null == value || string.IsNullOrEmpty(value))
            {
                value = TryGetRegistryValue(RegistryHive.LocalMachine, RegistryView.Registry32, RegistryPath, key);
            }

            if (!string.IsNullOrWhiteSpace(value))
            {
                registryCache[key] = value;
            }

            if (string.Empty == value)
            {
                Globals.ThisAddIn.LogMessage(string.Format("Failed to read registry key: {0}", key), string.Empty, EventLogEntryType.Information);
            }

            return value;
        }

        internal void LogMessage(string data, string callstack, EventLogEntryType eventType = EventLogEntryType.Error)
        {
            try
            {
                if (EventLog.SourceExists(LogSource))
                {
                    EventLog.WriteEntry(LogSource, data + "\r\n" + callstack, eventType);
                }
            }
            catch (Exception)
            {
                // Prevent logging from crashing the plugin
            }
        }

        protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            return new OutlookPlugin();
        }

        public static string LoadResourceString(string registryKey, string defaultValue)
        {
            var str = GetRegistryValue(registryKey);
            if (string.IsNullOrWhiteSpace(str))
                return defaultValue;

            return str;
        }

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            this.ApiUrl = GetRegistryValue(RegistryApiUrl);
            if (null != this.ApiUrl && !string.IsNullOrEmpty(this.ApiUrl))
            {
                if (!this.ApiUrl.EndsWith("/"))
                {
                    this.ApiUrl += "/";
                }
            }

            this.ApiUser = GetRegistryValue(RegistryApiUser);
            this.ApiPassword = GetRegistryValue(RegistryApiPassword);
            this.ForwardingAddress = GetRegistryValue(RegistryForwardingAddress);
            this.BaseHost = GetRegistryValue(RegistryBaseHost);

            this.AnonymousMatchingString = LoadResourceString(RegistryAnonymousMatchingString, DefaultAnonymousMatchingString);
            this.ForwardingSubject = LoadResourceString(RegistryForwardingSubject, DefaultForwardingSubject);
            this.ForwardingBody = LoadResourceString(RegistryForwardingBody, DefaultForwardingBody);
            this.ForwardingPopup = LoadResourceString(RegistryForwardingPopup, DefaultForwardingPopup);
            this.ForwardingMimeHeader = LoadResourceString(RegistryForwardingMimeHeader, DefaultForwardingMimeHeader);
            this.ForwardingMimeValue = LoadResourceString(RegistryForwardingMimeValue, DefaultForwardingMimeValue);
            bool showPopups = true;
            bool.TryParse(LoadResourceString(RegistryShowPopups, "True"), out showPopups);
            this.ShowPopups = showPopups;
            this.JunkFolder = this.Application.GetNamespace("MAPI").GetDefaultFolder(Outlook.OlDefaultFolders.olFolderJunk);

           System.Net.ServicePointManager.Expect100Continue = false;
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }

        #endregion

        internal class PictureConverter : AxHost
        {
            private PictureConverter() : base(String.Empty) { }

            static public stdole.IPictureDisp ImageToPictureDisp(Image image)
            {
                return (stdole.IPictureDisp)GetIPictureDispFromPicture(image);
            }

            static public stdole.IPictureDisp IconToPictureDisp(Icon icon)
            {
                return ImageToPictureDisp(icon.ToBitmap());
            }

            static public Image PictureDispToImage(stdole.IPictureDisp picture)
            {
                return GetPictureFromIPicture(picture);
            }
        }
    }
}
