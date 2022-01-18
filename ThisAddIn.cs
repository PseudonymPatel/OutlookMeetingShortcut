using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;
using System.Windows.Forms;
using System.Text.RegularExpressions;
using Microsoft.Office.Interop.Outlook;

namespace MeetingLinkFinder {
    public partial class ThisAddIn {

        public String foundMeeting = "";

        Outlook.Explorer explorer = null;
        Ribbon1 ribbon = null;

        private void ThisAddIn_Startup(object sender, System.EventArgs e) {
            //get the explorer and then register the selection change event.
            explorer = this.Application.ActiveExplorer();
            explorer.SelectionChange += new Outlook.ExplorerEvents_10_SelectionChangeEventHandler(SelectionChange_Event);
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e) {
            // Note: Outlook no longer raises this event. If you have code that 
            //    must run when Outlook shuts down, see https://go.microsoft.com/fwlink/?LinkId=506785
        }

        private void SelectionChange_Event() {

            String foundURLs = FindMeetingURLs();
            
            ribbon.refreshRibbon();

            if (foundURLs != null && ribbon != null) {
                ribbon.show();
            } else if (foundURLs == null && ribbon != null) {
                ribbon.hide();
            }
        }

        public String FindMeetingURLs() {
            return FindMeetingURLs(out _);
        }
        public String FindMeetingURLs(out String pass) {
            pass = null;
            Outlook.Explorer activeExplorer = this.Application.ActiveExplorer();

            try {
                if (activeExplorer.Selection.Count > 0 && activeExplorer.Selection[1] is Outlook.AppointmentItem) {
                    //selected appointment item, check for zoom, teams, etc. meeting links
                    Outlook.AppointmentItem appointmentItem = (Outlook.AppointmentItem)activeExplorer.Selection[1];

                    String appointementBody = appointmentItem.Body;

                    //regex for a url, use the first one?
                    String[] meetingAppliationDomains = {"teams.microsoft", "zoom.us"}; //representing the domain of meeting applications
                    String rawFoundURL = "";
                    
                    foreach (Match item in Regex.Matches(appointementBody, @"(http|ftp|https):\/\/([\w\-_]+(?:(?:\.[\w\-_]+)+))([\w\-\.,@?^=%&:/~\+#]*[\w\-\@?^=%&/~\+#])?")) {
                        Console.WriteLine(item.Value);
                        foreach (String domain in meetingAppliationDomains) {
                            if (item.Value.Contains(domain)) {
                                rawFoundURL = item.Value;

                                if (rawFoundURL.Contains("zoom")) {
                                    if (appointementBody.ToLower().Contains("pass")) {
                                        pass = appointementBody.Replace(rawFoundURL, "");
                                    }
                                }
                                goto foundURL;
                            }
                        }
                    }

                foundURL:
                    if (rawFoundURL == "") {
                        return null;
                    }

                    bool validURL = true;

                    //TODO: url checking

                    
                    return validURL ? rawFoundURL : null;

                }
            } catch (System.Exception ex) {
                Console.Error.WriteLine(ex.Message);
                MessageBox.Show(ex.Message);
                return null;
            }
            return null;
        }

        // generated code for ribbon element
        protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject() {
            ribbon = new Ribbon1(this);
            return ribbon;
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
    }
}
