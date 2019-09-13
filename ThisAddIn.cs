using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Interop.Outlook;
using System.Net.Http;
using System.IO;

namespace Signature
{
    public partial class ThisAddIn
    {   
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            setupSignature();
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e) {}

        async void setupSignature()
        {
            string htmlBody = "";

            string template = @"Z:\Challenge Databases\Signature Management\signature.html";

            using (StreamReader sr = new StreamReader(template))
            {
                htmlBody = await sr.ReadToEndAsync();
            }

            ExchangeUser currentUser = Application.Session.CurrentUser.AddressEntry.GetExchangeUser();
            htmlBody = htmlBody.Replace("%%FirstName%%", currentUser.FirstName);
            htmlBody = htmlBody.Replace("%%LastName%%", currentUser.LastName);
            htmlBody = htmlBody.Replace("%%Title%%", currentUser.JobTitle);
            htmlBody = htmlBody.Replace("%%PhoneNumber%%", currentUser.BusinessTelephoneNumber);
            htmlBody = htmlBody.Replace("%%Email%%", currentUser.PrimarySmtpAddress);

            string appDataDir = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + "\\Microsoft\\Signatures";
            string targetFile = Path.Combine(appDataDir, "Challenge Signature.htm");

            if (File.Exists(targetFile)) File.Delete(targetFile);

            using (FileStream fs = new FileStream(targetFile, FileMode.CreateNew, FileAccess.Write))
            {
                byte[] htmlBytes = Encoding.Default.GetBytes(htmlBody);
                MemoryStream memoryStream = new MemoryStream(htmlBytes);
                memoryStream.CopyTo(fs);
            }

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
