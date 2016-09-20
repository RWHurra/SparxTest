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
    }
}
