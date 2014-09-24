//==============================================================================
//
//        Filename: Program.cs
//
//        Created by: CENIT AG (Jan Assmann)
//              Version: NX 8.5.2.3 MP1
//              Date: 11-11-2013  (Format: mm-dd-yyyy)
//              Time: 08:30 (Format: hh-mm)
//
//==============================================================================

using System;
using NXOpen;

namespace Daimler.NX.BemiStructure
{
   /// <summary>class Program</summary>
   public class Program
   {
      /// <summary>Main method </summary>
      /// <returns></returns>
      public static void Main()
      {
         StructureCreation theStructureCreation = null;
         try
         {
            Session theSession = Session.GetSession();
            Part workPart = theSession.Parts.Work;
            if ( null != workPart )
            {
               theStructureCreation = new StructureCreation(theSession);
               theStructureCreation.Show();
            }
            else
            {
               UI.GetUI().NXMessageBox.Show("Main", NXMessageBox.DialogType.Information, "Current work part is empty. Please create/open one and try again.");
            }        
         }
         catch (Exception ex)
         {
            UI.GetUI().NXMessageBox.Show("Block Styler", NXMessageBox.DialogType.Error, ex.ToString());
         }
         finally
         {
            if (theStructureCreation != null)
               theStructureCreation.Dispose();
            theStructureCreation = null;
         }

      }

      public static int GetUnloadOption(string arg)
      {
         //return System.Convert.ToInt32(Session.LibraryUnloadOption.Explicitly);
         return System.Convert.ToInt32(Session.LibraryUnloadOption.Immediately);
         //return System.Convert.ToInt32(Session.LibraryUnloadOption.AtTermination);
      }

      public static void UnloadLibrary(string arg)
      {
         try
         {
         }
         catch (Exception ex)
         {
            UI.GetUI().NXMessageBox.Show("Main Function", NXMessageBox.DialogType.Error, ex.ToString());
         }
      }

   }
}
