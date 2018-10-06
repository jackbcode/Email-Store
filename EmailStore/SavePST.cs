using Microsoft.Office.Interop.Outlook;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace EmailStore
{
    static class SavePST
    {

        public static void SavePSTFile(IEnumerable<MailItem> pstfile)
        {


            foreach (object item in pstfile)

            {

                if (item is MailItem)
                {



                    // Retrieve the Object into MailItem

                    MailItem mailItem = item as MailItem;

                    Console.WriteLine("Saving message {0} ....", mailItem.Subject);

                    var sent = mailItem.SentOn;
                    var sentDate = sent.ToString("MMM dd yyyy hh-mmtt");
                    var bodystring = mailItem.Body;
                  string[] bodyarray = new string[0];

                    if (bodystring != null)
                    {
                        bodyarray = bodystring.Split(' ');
                    }
                    
                   

                    if (bodyarray != null)
                    {

                        var stringToCheck = "Reference:";
                        var subject = " ";
                        var fullsubject = " ";

                        for (int i = 0; i < bodyarray.Length; i++)
                        {
                            string check = bodyarray[i];

                            if (check == stringToCheck)
                            {
                                subject = bodyarray[i + 1];
                            }
                        }

          

                        var Sent = mailItem.To;


                        var sentFinal = " ";

                        if (Sent != null)
                        {

                            Regex pattern = new Regex("[<>:?|/\"*]");
                             sentFinal = pattern.Replace(Sent, "!");
                        }


                        var newname =sentFinal;

                        var subjectFinal = TS.TruncateLongString(subject, 10);


                        //fullsubject = subjectFinal + " " + sentDate;


                        if (subject != " ")
                        {
                            string filepath = @"I:\Emails20172018\" + subjectFinal + "-" + sentDate + " " + sentFinal + ".msg";

                            if (!File.Exists(filepath))
                            {
                                mailItem.SaveAs(filepath, OlSaveAsType.olMSG);

                            }



                        }

                        else if (mailItem.To != null)
                        {
                            string filepath2 = @"I:\Emails20172018\Unknown Email -" + sentDate + " " + sentFinal + ".msg";

                            if (!File.Exists(filepath2))
                            {

                                mailItem.SaveAs(filepath2, OlSaveAsType.olMSG);

                            }

                        }

                        else
                        {

                            string filepath3 = @"I:\Emails20172018\Unknown Email -" + sentDate + " NoAddress.msg";
                            if (!File.Exists(filepath3))
                            {

                                mailItem.SaveAs(filepath3, OlSaveAsType.olMSG);

                            }

                        }
                        
                    }
                }










            }
        }

    }
}
