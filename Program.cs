using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.DirectoryServices;

namespace CreateADGroups
{
    class Program
    {
        static void Main(string[] args)
        {
            string activeDirPath = args[0];
            string activeDirOU = args[1];
            string usersLocation = args[2];
            string filterIn = args[3];
            string filterOut = args[4];

            var DBConnString = System.Configuration.ConfigurationManager.ConnectionStrings["CreateADGroups.Properties.Settings.SchoolbaseConnectionString"].ToString();
            GetGroupDataDataContext GroupData = new GetGroupDataDataContext(DBConnString);
            List<string> groupNameIncludeFilter = new List<string> { };
            List<string> groupNameExcludeFilter = new List<string> { };
            List<string> groupBeenChecked = new List<string> { };

            string[] filterInLoop = filterIn.Split('|');
            foreach (string filterInWord in filterInLoop)
            {
                groupNameIncludeFilter.Add(filterInWord);
            }

            string[] filterOutLoop = filterOut.Split('|');
            foreach (string filterOutWord in filterOutLoop)
            {
                groupNameExcludeFilter.Add(filterOutWord);
            }


            var groupRecord = from x in GroupData.GroupReps.AsEnumerable()
                              where groupNameIncludeFilter.Any(f => x.Subject.ToLower().Contains(f.ToLower()))
                              && groupNameExcludeFilter.Any(e => !x.Subject.ToLower().Contains(e.ToLower()))
                              select x;

            foreach (var groupName in groupRecord)
            {
                try
                {

                    if (!groupBeenChecked.Contains(groupName.GroupName.ToString()))
                    {
                        RemoveUserFromGroup(groupName.GroupName, activeDirOU, activeDirPath);
                        Create(activeDirOU + "," + activeDirPath, groupName.GroupName);
                        groupBeenChecked.Add(groupName.GroupName.ToString());
                    }
                    AddToGroup(groupName.PupADSname, groupName.GroupName, usersLocation, activeDirOU, activeDirPath);
                }
                catch (Exception e)
                {
                    Console.WriteLine("Group Name Log (" + e.Message.ToString() + ")");
                }

            }

        }


        static void Create(string ouPath, string name)
        {
            if (!DirectoryEntry.Exists("LDAP://CN=" + name + "," + ouPath))
            {
                try
                {
                    // bind to the container, e.g. LDAP://cn=Users,dc=...
                    DirectoryEntry entry = new DirectoryEntry("LDAP://" + ouPath);

                    // create group entry
                    DirectoryEntry group = entry.Children.Add("CN=" + name, "group");

                    // set properties
                    group.Properties["sAmAccountName"].Value = name;

                    // save group
                    group.CommitChanges();
                }
                catch (Exception e)
                {
                    Console.WriteLine("Create Group Log " + ouPath.ToString() + " - " + name.ToString() + "(" + e.Message.ToString() + ")");
                }
            }
            else { Console.WriteLine(ouPath + " already exists"); }
        }



        static void RemoveUserFromGroup(string groupDn, string activeDirOU, string activeDirPath)
        {
            try
            {
                List<string> usersInGroup = new List<string> { };
                DirectoryEntry dirEntry = new DirectoryEntry("LDAP://CN=" + groupDn + "," + activeDirOU + "," + activeDirPath);

                foreach (var dn in dirEntry.Properties["member"])
                {
                    usersInGroup.Add(dn.ToString());
                }

                foreach (var user in usersInGroup)
                {
                    dirEntry.Properties["member"].Remove(user);
                    dirEntry.CommitChanges();
                    dirEntry.Close();
                }

            }
            catch (System.DirectoryServices.DirectoryServicesCOMException e)
            {
                Console.WriteLine("Remove User Log" + " - group probably exists elsewhere on AD (" + e.Message.ToString() + ")");
            }
        }


        static void AddToGroup(string userDn, string groupDn, string usersLocation, string activeDirOU, string activeDirPath)
        {
            try
            {
                DirectoryEntry dirEntry = new DirectoryEntry("LDAP://CN=" + groupDn + "," + activeDirOU + "," + activeDirPath);
                string distinguished = GetUserDn(userDn, usersLocation, activeDirPath);
                dirEntry.Properties["member"].Add(distinguished);
                dirEntry.CommitChanges();
                dirEntry.Close();
            }
            catch (System.DirectoryServices.DirectoryServicesCOMException e)
            {
                Console.WriteLine("Add User Log For " + groupDn.ToString() + " - group probably exists elsewhere on AD (" + e.Message.ToString() + ")");
            }
        }


        static string GetUserDn(string identity, string usersLocation, string activeDirPath)
        {
            using (var rootEntry = new DirectoryEntry("LDAP://" + usersLocation + "," + activeDirPath, null, null, AuthenticationTypes.Secure))
            {
                using (var directorySearcher = new DirectorySearcher(rootEntry, String.Format("(sAMAccountName={0})", identity)))
                {
                    var searchResult = directorySearcher.FindOne();
                    if (searchResult != null)
                    {
                        using (var userEntry = searchResult.GetDirectoryEntry())
                        {
                            return (string)userEntry.Properties["distinguishedName"].Value;
                        }
                    }
                }
            }
            return null;
        }

    }
}
