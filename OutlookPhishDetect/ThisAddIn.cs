using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;
using System.Windows.Forms;
using System.IO;
using System.Text.RegularExpressions;

namespace OutlookPhishDetect
{
    public partial class ThisAddIn
    {
        Outlook.Explorer currentExplorer = null;
        List<string> mailItemIDs = new List<string>();

        public List<string> MailItemIDs
        {
            get => mailItemIDs;
            set => mailItemIDs = value;
        }

        public class RuleCount
        {
            public int ruleCountValue;
        }

        //public int totalRuleCount
        //{
        //    get => totalRuleCount;
        //    set => totalRuleCount = value;
        //}

        //public string reason
        //{
        //    get => reason;
        //    set => reason = value;
        //}

        //public string totalReason
        //{
        //    get => totalReason;
        //    set => totalReason = value;
        //}


        private void ThisAddIn_Startup
            (object sender, System.EventArgs e)
        {

            currentExplorer = Application.ActiveExplorer();
            currentExplorer.SelectionChange += new Outlook.ExplorerEvents_10_SelectionChangeEventHandler(CurrentExplorer_Event);

        }

        private void CurrentExplorer_Event()
        {

            if (this.Application.ActiveExplorer().Selection.Count > 0)
            {
                Object selObject = this.Application.ActiveExplorer().Selection[1];

                if (selObject is Outlook.MailItem)
                {
                    Outlook.MailItem mailItem = (selObject as Outlook.MailItem);

                    List<string> phishtankURLs = new List<string>(); //list to hold Phishtank CSV URLs
                    List<string> localEmailURLs = new List<string>(); //list to hold local URLs embedded in the email message
                    List<string> urgencyCues = new List<string>() { "Urgent", "ASAP", "as soon as possible", "account", "immediately", "canceled", "important", "penal", "W2", "SSN", "social security number", "employment", "IRS", "payment", "refund", "billing", "salar", "bank", "invoice", "wire transfer", "tax"}; // list to hold a library of urgency cues
                    List<string> whiteList = new List<string>() { "auctorcorp.sharepoint.com", "auctor.com" };


                    //totalRuleCount = 0; //total rules count against the entire email message
                    //totalReason = ""; //all reasons for the rule-based trigger
                    RuleCount ruleCount = new RuleCount();


                    string bodyHTML = mailItem.HTMLBody; //Get HTML from email body
                    string fromAddress = GetSenderSMTPAddress(mailItem);

                    string GetSenderSMTPAddress(Outlook.MailItem mail)
                    {
                        string PR_SMTP_ADDRESS =
                            @"http://schemas.microsoft.com/mapi/proptag/0x39FE001E";
                        if (mail == null)
                        {
                            throw new ArgumentNullException();
                        }
                        if (mail.SenderEmailType == "EX")
                        {
                            Outlook.AddressEntry sender =
                                mail.Sender;
                            if (sender != null)
                            {
                                //Now we have an AddressEntry representing the Sender
                                if (sender.AddressEntryUserType ==
                                    Outlook.OlAddressEntryUserType.
                                    olExchangeUserAddressEntry
                                    || sender.AddressEntryUserType ==
                                    Outlook.OlAddressEntryUserType.
                                    olExchangeRemoteUserAddressEntry)
                                {
                                    //Use the ExchangeUser object PrimarySMTPAddress
                                    Outlook.ExchangeUser exchUser =
                                        sender.GetExchangeUser();
                                    if (exchUser != null)
                                    {
                                        return exchUser.PrimarySmtpAddress;
                                    }
                                    else
                                    {
                                        return null;
                                    }
                                }
                                else
                                {
                                    return sender.PropertyAccessor.GetProperty(
                                        PR_SMTP_ADDRESS) as string;
                                }
                            }
                            else
                            {
                                return null;
                            }
                        }
                        else
                        {
                            return mail.SenderEmailAddress;
                        }
                    }


                    //Find all instances of URLs embedded in the email message

                    //var linkFinder = new Regex(@"\b(?:https?://|www\.)\S+\b", RegexOptions.Compiled | RegexOptions.IgnoreCase);
                    Regex linkFinder = new Regex(@"(http|ftp|https):\/\/([\w\-_]+(?:(?:\.[\w\-_]+)+))([\w\-\.,@?^=%&amp;:/~\+#]*[\w\-\@?^=%&amp;/~\+#])?", RegexOptions.Compiled | RegexOptions.IgnoreCase);

                    //look for base href="
                    Regex baseHrefFinder = new Regex(@"base href=""(http|ftp|https):\/\/([\w\-_]+(?:(?:\.[\w\-_]+)+))([\w\-\.,@?^=%&amp;:/~\+#]*[\w\-\@?^=%&amp;/~\+#])?""");


                    
                    //string baseURL = "";

                    //foreach (Match m in baseHrefFinder.Matches(bodyHTML))
                    //{
                    //    baseURL = m.Value.Remove(0, 10); // finds base URL and removes the "base href" from the beginning
                    //}




                    foreach (Match m in linkFinder.Matches(bodyHTML))
                    {
                        localEmailURLs.Add(m.Value);
                        //MessageBox.Show(baseURL + m.Value);
                    }

                    //foreach (string url in localEmailURLs)
                    //{
                    //    MessageBox.Show(url);
                    //}

                    phishtankURLs = Check_PhishtankDB2(); //Create a list of all known phishtank entries using the method

                    foreach (var phishtankURL in phishtankURLs)
                    {

                        if (mailItemIDs.Contains(mailItem.EntryID) == false)
                        {
                            if (localEmailURLs.Contains(phishtankURL))
                            {
                                //create message box
                                MessageBox.Show("This email contains the link " + phishtankURL + ", which is a known phishing threat.  Do not continue!");


                                mailItemIDs.Add(mailItem.EntryID);

                                break; //we don't need to keep going if one is found.
                            }

                        }
                    }

                    //The Rule_Based_Checks method will return an int match count and a string of reasons as a Tuple.


                    if (mailItemIDs.Contains(mailItem.EntryID) == false)
                    {

                        foreach (var url in localEmailURLs)
                        {

                            foreach (var whiteListTerm in whiteList)
                            {
                                if (url.ToLower().Contains(whiteListTerm)) //skips whitelisted terms in urls
                                {
                                    //MessageBox.Show(whiteListTerm);
                                    //do nothing and move on to the next url
                                }
                                
                                else
                                {
                                    var ruleTuple = Rule_Based_Checks(url);  //checking the url against the rules method and returns ruleCount and reason
                                    {



                                        var totalCount = ruleTuple.Item1;
                                        string totalReason = ruleTuple.Item2;

                                        //MessageBox.Show(fromAddress);

                                        //Check for external addresses
                                        try
                                        {

                                            if (fromAddress.ToLower().Contains("auctor.com") == false)
                                            {
                                                totalCount += 1;
                                                totalReason += "--The email is from an external sender.\n";
                                            }
                                        }
                                        catch
                                        {
                                            //MessageBox.Show("I tried");
                                            //Move on, the email is probabaly a draft with no sender address.
                                        }

                                        if (baseHrefFinder.IsMatch(bodyHTML)) // check for the base HTML header (baseStriker hack)
                                        {
                                            totalCount += 1;
                                            totalReason += "--The email contains a suspicious HTML header.\n";
                                        }

                                        foreach (string word in urgencyCues)
                                        {
                                            if (bodyHTML.Contains(word))
                                            {
                                                totalCount += 1;
                                                totalReason += "--The email contains the word or phrase: " + word + ".\n";
                                            }
                                        }

                                        //MessageBox.Show(mailItem.EntryID);
                                        if (totalCount > 3) //4 strikes triggers the messagebox.
                                        {
                                            MessageBox.Show("There is at least one link in this email that may be suspicious.\n The reasons are:\n" + totalReason, "Think Before You Link");
                                            MessageBox.Show(url);
                                            mailItemIDs.Add(mailItem.EntryID);

                                            break;  //if one is found, that is enough.
                                                    //totalReason += ruleTuple.Item2;
                                        }

                                        mailItemIDs.Add(mailItem.EntryID);
                                        //MessageBox.Show("Did the code make it here?");





                                    }
                                }
                            }


                        }
                    }


                }
            }

        }

        private bool Check_PhishtankDB(string checkURL)
        {
            var listofURLs = new List<string> { };
            bool isMatch;

            foreach (var item in File.ReadLines(@"C:\temp\phishtank\phishtank.csv"))
            {
                string[] splitItems = item.Split(',');

                //listofURLs.Add(splitItems[1]);
                //isMatch = Regex.IsMatch(checkURL, splitItems[1]);

                if (splitItems[1].Contains(checkURL))
                {
                    return true;
                }  
            }

            return false;
        }

        private List<string> Check_PhishtankDB2()
        {
            var listofURLs = new List<string> { };
            //bool isMatch;

            foreach (var item in File.ReadLines(@"C:\temp\phishtank\phishtank.csv"))
            {
                string[] splitItems = item.Split(',');

                listofURLs.Add(splitItems[1]);
                //isMatch = Regex.IsMatch(checkURL, splitItems[1]);

            }

            listofURLs.RemoveAt(0);  //removes the column header from the Phishtank database CSV.
            return listofURLs;
        }


        //Make the following a class???????
        private Tuple<int, string> Rule_Based_Checks(string localUrl)
        
        {

            RuleCount rule = new RuleCount();
            string reason = "";
            int charDotCount = 0;
            int charSlashCount = 0;

            //The first rule will check if the Hyperlinks are an IP address
            var ipFinder = new Regex(@"\b(?:\d{1,3}\.){3}\d{1,3}\b", RegexOptions.Compiled | RegexOptions.IgnoreCase);
            if (ipFinder.IsMatch(localUrl))
            {
                rule.ruleCountValue += 1;
                reason += "--A Link URL uses an IP address.\n";
                //MessageBox.Show("ipaddress URL");
            }

            //The second rule will check if it uses a non-SSL URL
            if (localUrl.ToLower().Contains("https") == false)
            {
                rule.ruleCountValue += 1;
                reason += "--A link URL is not using the secure HTTPS protocol.\n";
            }

            //The third rule will check address length
            if (localUrl.Length > 50)
            {
                rule.ruleCountValue += 1;
                reason += "--This message contains a link with a URL length that is suspiciously long.\n";
            }

            //The fourth rule will check for multiple subdomains and directories


            foreach (var character in localUrl)
            {
                if (character == '.')
                {
                    charDotCount += 1;
                }
                if (character == '/')
                {
                    charSlashCount += 1;
                }

            }

            if (charDotCount > 4 || charSlashCount > 4)
            {
                rule.ruleCountValue += 1;
                reason += "--The message contains a link with a suspicious amount of subdomains and/or directories.\n";
            }

            //The fifth rule will check for the inclusion of an @ symbol
            if (localUrl.Contains("@"))
            {
                rule.ruleCountValue += 1;
                reason += "--A Link contains an @ symbol.\n";
            }

            return Tuple.Create(rule.ruleCountValue, reason);
        }

        //other checks
        //"<base href" should trigger something or add to the count used to split up the URL "base Striker"
        //Maybe the checker can put them back together somehow

        //check the from address against the actual address (how?)
        //or do the checks against the from address also.




        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            // Note: Outlook no longer raises this event. If you have code that 
            //    must run when Outlook shuts down, see https://go.microsoft.com/fwlink/?LinkId=506785
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
