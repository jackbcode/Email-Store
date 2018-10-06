using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EmailStore
{
    public class Read_Files
    {
        public static void ReadFolder(string filePath)
        {
            string[] array1 = Directory.GetFiles(filePath);

            // Put all bin files in root directory into array.
            // ... This is case-insensitive.
            string[] array2 = Directory.GetFiles(filePath, "*.BIN");

            // Display all files.
            Console.WriteLine("--- Files: ---");

            string text = @"I:\Emails20172018.txt";

            using (TextWriter writer = File.CreateText(text))
            {
                foreach (string name in array1)
                {
                    
                    Console.WriteLine(name);

                    var newString = name.Remove(0, name.IndexOf('-') + 1);

                    var date2  = newString.Substring(0, newString.LastIndexOf(" "));

                    var clientRef2 = name.Substring(0, name.IndexOf("-"));
                    var clientRef3 = clientRef2.Substring(clientRef2.LastIndexOf("\\") +1);


                    var emailAddress = name.Substring(name.LastIndexOf(" ") + 1);
                    var emailAddressFinal = emailAddress.Substring(0, emailAddress.LastIndexOf("."));

                    var fileLocation =  name.Substring(name.LastIndexOf("\\") + 1);

                    var final = (clientRef3 + "," + date2 + "," + emailAddressFinal + "," + fileLocation);
                    Console.WriteLine(final);

                    writer.WriteLine(final);
                }

            }
            
            //System.IO.File.WriteAllLines(@"C:\Users\Jack\Desktop\Email\Emails.txt", array1);
            
        }


    }
}
