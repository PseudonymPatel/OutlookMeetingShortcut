using stdole;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;
using Office = Microsoft.Office.Core;

// TODO:  Follow these steps to enable the Ribbon (XML) item:

// 1: Copy the following code block into the ThisAddin, ThisWorkbook, or ThisDocument class.

//  protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
//  {
//      return new Ribbon1();
//  }

// 2. Create callback methods in the "Ribbon Callbacks" region of this class to handle user
//    actions, such as clicking a button. Note: if you have exported this Ribbon from the Ribbon designer,
//    move your code from the event handlers to the callback methods and modify the code to work with the
//    Ribbon extensibility (RibbonX) programming model.

// 3. Assign attributes to the control tags in the Ribbon XML file to identify the appropriate callback methods in your code.  

// For more information, see the Ribbon XML documentation in the Visual Studio Tools for Office Help.


namespace MeetingLinkFinder {
    [ComVisible(true)]
    public class Ribbon1 : Office.IRibbonExtensibility {
        private Office.IRibbonUI ribbon;

        private ThisAddIn addIn;
        private bool buttonVisibility = false;

        public Ribbon1(ThisAddIn addin) {
            this.addIn = addin;
        }

        public void show() {
            buttonVisibility = true;
        }

        public void hide() {
            buttonVisibility = false;
        }

        #region IRibbonExtensibility Members

        public string GetCustomUI(string ribbonID) {
            return GetResourceText("MeetingLinkFinder.Ribbon1.xml");
        }

        #endregion

        #region Ribbon Callbacks
        //Create callback methods here. For more information about adding callback methods, visit https://go.microsoft.com/fwlink/?LinkID=271226

        public void Ribbon_Load(Office.IRibbonUI ribbonUI) {
            this.ribbon = ribbonUI;
        }

        public void OnOpenMeeting(Office.IRibbonControl control) {
            String zoomPass;
            String meetingURL = addIn.FindMeetingURLs(out zoomPass);

            if (meetingURL != null) {
                System.Diagnostics.Process.Start(meetingURL);
                if (zoomPass != null && zoomPass != "") {
                    MessageBox.Show(zoomPass);
                }
            } else {
                MessageBox.Show("No meeting found :(");
            }
        }

        public bool IsVisibleCallback(Office.IRibbonControl control) {
            return buttonVisibility;
        }

        public IPictureDisp GetImage(String url) {

            //todo: actually get the proper url?

            Image image = Image.FromFile(@"C:\Users\sheen\OneDrive - University of Toronto\Documents\source\MeetingLinkFinder\open.png");
            return AxHostConverter.ImageToPictureDisp(image);
        }


        #endregion

        #region Helpers

        private static string GetResourceText(string resourceName) {
            Assembly asm = Assembly.GetExecutingAssembly();
            string[] resourceNames = asm.GetManifestResourceNames();
            for (int i = 0; i < resourceNames.Length; ++i) {
                if (string.Compare(resourceName, resourceNames[i], StringComparison.OrdinalIgnoreCase) == 0) {
                    using (StreamReader resourceReader = new StreamReader(asm.GetManifestResourceStream(resourceNames[i]))) {
                        if (resourceReader != null) {
                            return resourceReader.ReadToEnd();
                        }
                    }
                }
            }
            return null;
        }

        public void refreshRibbon() {
            if (ribbon != null) {
                ribbon.Invalidate();
            }
        }

        #endregion

        internal class AxHostConverter : AxHost {

            private AxHostConverter() : base("") { }

            static public stdole.IPictureDisp ImageToPictureDisp(Image image) {
                return (stdole.IPictureDisp)GetIPictureDispFromPicture(image);
            }

            static public Image PictureDispToImage(stdole.IPictureDisp pictureDisp) {
                return GetPictureFromIPicture(pictureDisp);
            }

        }
    }
}
