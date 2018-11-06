using System;
using System.IO;
using System.Collections;
using System.Collections.Generic;
using Microsoft.Graph;
using LookMeUp.Controllers;
using System.Linq;

// UserSig index 10 indicates location
// 


namespace LookMeUp
{
    class Program
    {
        public static GraphServiceClient graphClient;

        static void Main(string[] args)
        {
            graphClient = Authentication.GetAuthenticatedClient();

            String search = "Id,DisplayName,Mail,JobTitle,Department,StreetAddress,City,State,PostalCode,Country,BusinessPhones,MobilePhone";

            var me = graphClient.Me.Request().Select(search).GetAsync().Result;
            var cre = graphClient.Groups["9dea13e9-e5cf-49bb-b725-f7eb11507c1b"].Members.Request().GetAsync().Result;
            var creUsers = new List<DirectoryObject>();
            creUsers.AddRange(cre.CurrentPage);
            while(cre.NextPageRequest != null)
            {
                cre = cre.NextPageRequest.GetAsync().Result;
                creUsers.AddRange(cre.CurrentPage);
            }
            
            bool isCredo = false;
            bool isExcluded = false;

            foreach(var user in creUsers)
            {
                if (user.Id == me.Id)
                {
                    isCredo = true;
                }
            }
            if (me.Mail.Contains("bluerubicon")) { isExcluded = true; }

            List<String> sb = new List<String>();
            
            // metaData
            sb.Add(me.DisplayName );            // index 0
            sb.Add(me.JobTitle );               // index 1
            sb.Add(me.Department );             // index 2
            // location Info
            sb.Add(me.StreetAddress );
            sb.Add(me.City );
            sb.Add(me.State );
            sb.Add(me.PostalCode );
            sb.Add(me.Country );                // index 7
            sb.Add((me.BusinessPhones).First());
            sb.Add(me.MobilePhone);
            if (isCredo) { sb.Add("1"); }
            else if(isExcluded) { sb.Add("2"); }
            else { sb.Add("0"); }

            if(System.IO.File.Exists("C:/temp/usersig.txt"))
            {
                System.IO.File.Delete("C:/temp/usersig.txt");
            }

            if (System.IO.Directory.Exists("C:/temp"))
            {
                System.IO.File.WriteAllLines(@"C:/temp/usersig.txt", sb);
            }
            else
            {
                System.IO.Directory.CreateDirectory("C:/temp");
                System.IO.File.WriteAllLines(@"C:/temp/usersig.txt", sb);
            }
        }

        public static bool IsMember(User user, List<DirectoryObject> group)
        {
            bool result = false;
            foreach(var member in group)
            {
                if (user.Id == member.Id) { result = true; }
            }
            return result;
        }
    }
}
