using System;

namespace MyAddin
{
    public class MyAddinClass
	{
		// define menu constants
		const string menuHeader = "-&GenericFTS";

        /// Called Before EA starts to check Add-In Exist
        public String EA_Connect(EA.Repository Repository)
		{
			//No special processing required.
			return "a string";
		}

		/// Called when user Clicks Add-Ins Menu item from within EA.
		/// Populates the Menu with our desired selections
		public object EA_GetMenuItems(EA.Repository Repository, string Location, string MenuName)
		{

            switch (MenuName)
            {
                case "":
                    return menuHeader;
            }
            return "Open";
        }

		/// returns true if a project is currently opened
		bool IsProjectOpen(EA.Repository Repository)
		{
			try
			{
				EA.Collection c = Repository.Models;
				return true;
			}
			catch
			{
				return false;
			}
		}

		/// Called when user makes a selection in the menu.
		public void EA_MenuClick(EA.Repository Repository, string Location, string MenuName, string ItemName)
		{
            try
            {
                GenericFTS();
            }
            catch (Exception)
            {
            }
            
        }

        // calling form created in separate class
        public void GenericFTS()
        {
            try
            {
                Form1 objForm1 = new Form1();
                objForm1.ShowDialog();
            }
            catch (Exception)
            {
            }
                
        }

        // EA calls this operation when it exists. Can be used to do some cleanup work.
        public void EA_Disconnect()
		{
			GC.Collect();
			GC.WaitForPendingFinalizers();
		}
    }
}