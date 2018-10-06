using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Outlook;

namespace EmailStore
{
    class Program
    {
        static void Main(string[] args)
        {

            Console.WriteLine("Gathering email data......................");


            var pstArray = new string[,] {
                {@"I:\Outlook\DocJournal 2017\Inbox 2017.pst","Personal Folders Sep 17 to Nov 17" },
                //{ @"I:\Outlook\archive.pst","Archive Folders Oct17 to Jan18" },
                //{@"I:\Outlook\DocJournal Oct.pst","DocJournal Aug17 to Oct17" },
                //{@"I:\Outlook\DocJournal Dec.pst","DocJournal Sep17 to Dec17" },
                //{@"I:\Outlook\Docjournal.pst","Docjournal" },
                //{@"I:\Outlook\DocJournalJanuary2018.pst","_DocJournal Jan 2018 to Feb 2018" },
                // {@"I:\Outlook\DocJournal 2017\AdminBackup.pst", "AdminBackup" },
                //{@"I:\Outlook\DocJournal 2017\archive.pst","Archive Folders Oct17 to Jan18" },
                //{@"I:\Outlook\DocJournal 2017\Doc Journal 3.pst","DocJournal Jun17 to Sep17" },
                //{@"I:\Outlook\DocJournal 2017\Doc Journal 3-4.pst","January 2017 to May 2017" },
                //{@"I:\Outlook\DocJournal 2017\DocJournal 11-07-2017.pst","DocJournal May17 to Oct17" },
                //  {@"I:\Outlook\DocJournal 2017\DocJournal01-02.pst","DocJournal Jan17 to Mar17" },
                //    {@"I:\Outlook\DocJournal 2017\Inbox 2017.pst","Personal Folders Sep 17 to Nov 17"},
                //     {@"I:\Outlook\Doc Journal March 2018.pst","Doc Journal March 2018" },

            };





            for (int i = 0; i <= pstArray.Length; i++)
            {

                var filepath = pstArray[i, 0];

                var filename = pstArray[i, 1];

                var outlookItems = PST.readPst(filepath, filename);

                SavePST.SavePSTFile(outlookItems);

            }

            string filePath = @"I:\Emails20172018\";

            Read_Files.ReadFolder(filePath);


        }

      

    }

}














