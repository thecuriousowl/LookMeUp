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

            Console.WriteLine("1 of 5     Retrieving user details. . . ");
            var me = graphClient.Me.Request().Select(search).GetAsync().Result;

            Console.WriteLine("2 of 5     Retrieving group memberships");
            var cre = ParseGroup("9dea13e9-e5cf-49bb-b725-f7eb11507c1b");
            var tbr = ParseGroup("902c612c-145e-42f5-bf2d-e67ceec05c90");
            var murica = ParseGroup("ea26b5f0-b9ad-4b72-9d77-fc247925dda1");

            String where = "0";                             // base case for A4 templates

            if (IsMember(me, cre)) { where = "1"; Console.WriteLine("3 of 5     Credo Account"); }         // credo templates
            else if (IsMember(me, tbr)) { where = "2"; Console.WriteLine("3 of 5     TBR Account"); }    // tbr no templates
            else if (IsMember(me, murica)) { where = "3"; Console.WriteLine("3 of 5     NA Account"); } // letter templates
            else { where = "0"; Console.WriteLine("3 of 5     OffNet Account"); }                           // base case for A4 templates

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
            sb.Add(where);

            Console.WriteLine("4 of 5     Creating output directory and writing to file.");
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
            Console.WriteLine("5 of 5     Complete");
        }




        // Helper Functions

        public static bool IsMember(User user, List<DirectoryObject> group)
        {
            bool result = false;
            foreach(var member in group)
            {
                if (user.Id == member.Id) { result = true; }
            }
            return result;
        }

        public static List<DirectoryObject> ParseGroup(String groupID)
        {
            GraphServiceClient thisClient = Authentication.GetAuthenticatedClient();

            List<DirectoryObject> result = new List<DirectoryObject>();
            List<DirectoryObject> request = new List<DirectoryObject>();
            var root = thisClient.Groups[groupID].Members.Request().GetAsync().Result;

            // build full member list of member objects
            request.AddRange(root.CurrentPage);
            while(root.NextPageRequest != null)
            {
                root = root.NextPageRequest.GetAsync().Result;
                request.AddRange(root.CurrentPage);
            }

            foreach(var dirObject in request)
            {
                if(dirObject.ODataType == "#microsoft.graph.group")
                {
                    result.AddRange(ParseGroup(dirObject.Id));
                }
                else
                {
                    result.Add(dirObject);
                }
            }
            return result;
        }
    }
}
