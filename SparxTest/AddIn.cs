using EA;
using System.Runtime.InteropServices;

namespace SparxTest
{
    [ComVisible(true)]
    public class AddIn
    {
        // Called Before EA starts to check Add-In Exists
        public string EA_Connect(Repository repository)
        {
            // nothing special
            return "SparxTest.AddIn - connected";
        }
        // Called when user Click Add-Ins Menu item.
        public object EA_GetMenuItems(Repository repository,
            string location, string menuName)
        {
            switch (menuName)
            {
                case "":
                    return "My Add-In!";
            }
            return "";
        }

        static bool IsProjectOpen(Repository repository)
        {
            try
            {
                return null != repository.Models;
            }
            catch
            {
                return false;
            }
        }

        public void EA_GetMenuState(Repository repository, string location, string menuName, string itemName, ref bool isEnabled, ref bool isChecked)
        {
            if (IsProjectOpen(repository))
            {
                if (itemName == "My Add-In!")
                    isEnabled = true;
            }
            else
                isEnabled = false;
        }

        public void EA_MenuClick(Repository repository, string location, string menuName, string itemName)
        {
            switch (itemName)
            {
                case "My Add-In!":
                    var count = 0;
                    foreach (Package model in repository.Models)
                        foreach (Package package in model.Packages)
                            count += CountClasses(package);
                    MessageBox.Show("this project contains "
                        + count + " " + (count == 1 ? "class" : "classes"));
                    break;
            }
        }

        private static int CountClasses(Package package)
        {
            var count = 0;
            foreach (Element e in package.Elements)
                if (e.Type == "Class")
                    count++;
            foreach (Package p in package.Packages)
                count += CountClasses(p);
            return count;
        }
    }
}
