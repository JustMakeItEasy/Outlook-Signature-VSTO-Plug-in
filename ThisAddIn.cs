using Microsoft.Office.Interop.Outlook;
using System;
using System.Collections.Generic;
using System.IO;

namespace Signature
{
    public partial class ThisAddIn
    {

        private void ThisAddIn_Startup(object sender, EventArgs e) => SetupSignature();

        private void ThisAddIn_Shutdown(object sender, EventArgs e) { }

        void SetupSignature()
        {
            string htmlBody = File.ReadAllText("template.html");

            ExchangeUser currentUser = Application.Session.CurrentUser.AddressEntry.GetExchangeUser();

            //Change the following variables to suit your signature's requirements. Anything is valid, including dynamic img URL's depending on user 
            //For a list of properties, refer to https://docs.microsoft.com/en-us/dotnet/api/microsoft.office.interop.outlook.exchangeuser?view=outlook-pia 

            Dictionary<string, string> variableMap = new Dictionary<string, string>();
            variableMap.Add("%%FirstName%%", currentUser.FirstName);
            variableMap.Add("%%LastName%%", currentUser.LastName);
            variableMap.Add("%%Title%%", currentUser.JobTitle);
            variableMap.Add("%%PhoneNumber%%", currentUser.BusinessTelephoneNumber);
            variableMap.Add("%%Email%%", currentUser.PrimarySmtpAddress);
            if (currentUser.FirstName[0] == 'S')
            {
                variableMap.Add("%%ImgURL%%", @"https://emojipedia-us.s3.dualstack.us-west-1.amazonaws.com/thumbs/120/apple/237/smiling-face-with-smiling-eyes_1f60a.png");
            }

            foreach (KeyValuePair<string, string> map in variableMap) htmlBody = htmlBody.Replace(map.Key, map.Value);

            string SignatureName = "Name not set";
            string appDataDir = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + "\\Microsoft\\Signatures";
            string targetFile = Path.Combine(appDataDir, SignatureName + ".htm");

            File.WriteAllText(targetFile, htmlBody);
        }


        #region VSTO generated code

        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }

        #endregion
    }
}
