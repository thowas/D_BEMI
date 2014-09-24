//==============================================================================
//
//        Filename: PartStatics.cs
//
//        Created by: CENIT AG (Jan Assmann)
//              Version: NX 8.5.2.3 MP1
//              Date: 03-12-2013  (Format: mm-dd-yyyy)
//              Time: 08:30 (Format: hh-mm)
//
//==============================================================================

using System;
using System.IO;
using NXOpen;
using NXOpen.Assemblies;

namespace Daimler.NX.BemiStructure
{
   /// <summary>Part statics class</summary>
   public static class PartStatics
   {

      private const string RniPartType = "RNI_PARTTYPE";

      private const string StartPartName = "Startpart_Name";

      private const string StartPartVersion = "Startpart_Version";

      private const string ConfigurationName = "Configuration_Name";

      private const string ConfigurationBUnit = "Business-Unit";

      private const string ConfigurationDepartment = "Department";

      private const string NoAttributeMsg = "Startpart has no Attribute: ";


      /// <summary>Get a string attribute of the part.</summary>
      /// <param name="part">The part to get the string from.</param>
      /// <param name="attributeName">The name of the attribute to get.</param>
      /// <returns>The value.</returns>
      public static string GetStringAttribute(Part part, string attributeName)
      {
         // get user attribute exists flag
         bool userAttributeExists = PartAttributeExists(part, attributeName);

         // if the user attribute does not exists
         if (!userAttributeExists)
         {
            // throw exception
            //throw ExceptionStatics.CreateException("User attribute \"" + attributeName + "\" does not exist in part \"" + ComponentPartAttributeStatics.GetArticleCode(part) + "\".");
         }

         // Get string AttributeInformation from Part for defined AttributeName
         NXObject.AttributeInformation attributeInfo = part.GetUserAttribute(attributeName, NXObject.AttributeType.String, -1);

         // Get string Value from AttributeInformation from Part for defined AttributeName
         string attributeValue = attributeInfo.StringValue;

         // Return string Value from Attribute of Part with defined Name
         return attributeValue;
      }


      /// <summary>Set a string attribute of the part.</summary>
      /// <param name="part">The part to set the string to.</param>
      /// <param name="attributeName">The name of the attribute to set.</param>
      /// <param name="attributeValue">The value of the attribute to set.</param>
      /// <param name="attributeCategory">The category of the attribute to set.</param>
      /// <param name="unlockAutomatically">If <c>true</c> the attribute is unlocked automatically to set the attribute and locked again.</param>
      public static void SetStringAttribute(Part part, string attributeName, string attributeValue, string attributeCategory = "", bool unlockAutomatically = false)
      {
         // set the unlocked attribute flag to false
         bool unlockedAttribute = false;

         // if the unlock should be performed automatically
         if (unlockAutomatically)
         {
            // if the attribute exists
            if (PartAttributeExists(part, attributeName))
            {
               // get the lock state of the attribute
               unlockedAttribute = GetAttributeLock(part, attributeName);
            }

            // if the attribute should be unlocked
            if (unlockedAttribute)
            {
               // unlock the attribute
               SetAttributeLock(part, attributeName, false, true);
            }
         }

         // build attribute
         var attribute = new NXObject.AttributeInformation
         {
            Title = attributeName,
            Type = NXObject.AttributeType.String,
            StringValue = attributeValue,
            Category = attributeCategory
         };

         // set attribute to part
         part.SetUserAttribute(attribute, Update.Option.Now);

         // if the attributes have been unlocked
         if (unlockedAttribute)
         {
            // lock the attribute again
            SetAttributeLock(part, attributeName, true, true);
         }
      }


      /// <summary>Lock or unlock the attribute of the part.</summary>
      /// <param name="part">The part of the attribute to lock or unlock.</param>
      /// <param name="attributeName">The name of the attribute to lock or unlock.</param>
      /// <param name="lockAttribute">The lock/unlock state.</param>
      /// <param name="throwExceptionIfNotAllowed">If <code>true</code> an exception is thrown, if it is not allowed to lock an attribute.</param>
      public static void SetAttributeLock(Part part, string attributeName, bool lockAttribute, bool throwExceptionIfNotAllowed)
      {
         try
         {
            part.SetUserAttributeLock(attributeName, NXObject.AttributeType.Any, lockAttribute);
         }
         catch (Exception)
         {
            if (throwExceptionIfNotAllowed)
            {
               throw new Exception("Locking of attributes is not allowed. Change in customer defaults.");
            }
         }
      }


      /// <summary>Get if an attribute of the part is locked or unlocked.</summary>
      /// <param name="part">The part of the attribute to get the locked or unlocked state.</param>
      /// <param name="attributeName">The name of the attribute to lock or unlock.</param>
      /// <returns><code>true</code>, if the attribute is locked. <code>false</code>, if the attribute is not locked.</returns>
      public static bool GetAttributeLock(Part part, string attributeName)
      {
         bool lockAttribute = part.GetUserAttributeLock(attributeName, NXObject.AttributeType.Any);

         return lockAttribute;
      }


      /// <summary>Check if an attribute exists in the part.</summary>
      /// <param name="part">The part of the attribute.</param>
      /// <param name="attributeName">The name of the attribute to check existance in the part.</param>
      /// <returns><c>true</c>, if the attribute exists in the part. <c>false</c>, if the attribute does not exist in the part.</returns>
      public static bool PartAttributeExists(Part part, string attributeName)
      {
         bool attributeExists = part.HasUserAttribute(attributeName, NXObject.AttributeType.Any, -1);

         return attributeExists;
      }


      /// <summary>Check part</summary>
      /// <param name="startPartToCheck">The part to check</param>
      /// <param name="messageIfFailed">The message if check faild</param>
      /// <returns>check ok or not</returns>
      public static bool CheckPart(Part startPartToCheck, out string messageIfFailed)
      {
         messageIfFailed = "";
   
         // check if rniPartType attribute exists and Value is "ASM"
         if (PartAttributeExists(startPartToCheck, RniPartType))
         {
            string rniType = GetStringAttribute(startPartToCheck, RniPartType);
            if ( rniType != "ASM")
            {
               messageIfFailed = "StartPart (root part) has not RNI_PARTTYPE 'ASM'. Please check StartPart!";
               return false;
            } 
         }
         else
         {
            messageIfFailed = NoAttributeMsg + RniPartType;
            return false;
         }
            
         // only check if attribute "Startpart_Name" exists
         if (!PartAttributeExists(startPartToCheck, StartPartName))
         {
            messageIfFailed = NoAttributeMsg + StartPartName;
            return false;
         }         

         // only check if attribute "Startpart_Version" exists
         if (!PartAttributeExists(startPartToCheck, StartPartVersion))
         {
            messageIfFailed = NoAttributeMsg + StartPartVersion;
            return false;
         }

         // Get rootComponent of current work part (asm start part)
         Component rootComponent = startPartToCheck.ComponentAssembly.RootComponent;
         if (null != rootComponent)
         {
            // Get children of root component 
            Component[] children = rootComponent.GetChildren();

            // Root component should have one child !!
            if (children.Length != 1)
            {
               messageIfFailed = "StartPart has more than one child. Please check StartPart!";
               return false;
            }
              
            // Get child from array
            Component child = children[0];

            // Get PRT part from child (prototype)
            Part daimlerModelPart = child.Prototype as Part;

            if (null != daimlerModelPart)
            {          
               if (PartAttributeExists(daimlerModelPart, RniPartType))
               {
                  string rniType = GetStringAttribute(daimlerModelPart, RniPartType);
                  if (rniType != "PRT")
                  {
                     messageIfFailed = "Geometry StartPart has not RNI_PARTTYPE 'PRT'. Please check StartPart!";
                     return false;
                  }
               }
               else
               {
                  messageIfFailed = NoAttributeMsg + RniPartType;
                  return false;
               }

               if (!PartAttributeExists(daimlerModelPart, ConfigurationName))
               {
                  messageIfFailed = NoAttributeMsg + ConfigurationName;
                  return false;
               }

               if (!PartAttributeExists(daimlerModelPart, ConfigurationBUnit))
               {
                  messageIfFailed = NoAttributeMsg + ConfigurationBUnit;
                  return false;
               }

               if (!PartAttributeExists(daimlerModelPart, ConfigurationDepartment))
               {
                  messageIfFailed = NoAttributeMsg + ConfigurationDepartment;
                  return false;
               }
                     
               string configName = GetStringAttribute(daimlerModelPart, ConfigurationName);
               if (String.IsNullOrEmpty(configName))
               {
                  messageIfFailed = "Geometry StartPart has not Configuration_Name. Please check StartPart!";
                  return false;
               }

               string businessUnit = GetStringAttribute(daimlerModelPart, ConfigurationBUnit);
               if (String.IsNullOrEmpty(businessUnit))
               {
                  messageIfFailed = "Geometry StartPart has no Business-Unit. Please check StartPart!";
                  return false;
               }

               string department = GetStringAttribute(daimlerModelPart, ConfigurationDepartment);
               if (String.IsNullOrEmpty(department))
               {
                  messageIfFailed = "Geometry StartPart has no Department. Please check StartPart!";
                  return false;
               }
  
               // only check if attribute "Startpart_Name" exists
               if (!PartAttributeExists(daimlerModelPart, StartPartName))
               {
                  messageIfFailed = NoAttributeMsg + StartPartName;
                  return false;
               }

               // only check if attribute "Startpart_Version" exists
               if (!PartAttributeExists(daimlerModelPart, StartPartVersion))
               {
                  messageIfFailed = NoAttributeMsg + StartPartVersion;
                  return false;
               }
            }
         }                 
         else
         {
            messageIfFailed = "StartPart has no root component. Please check !";
            return false;
         }
    
         return true;
      }


      /// <summary>Create daimlerStartPartsForEnhancement</summary>
      /// <param name="currentSession"></param>
      /// <param name="referenceGeoPartPath"></param>
      /// <param name="origModelPartPath"></param>
      /// <param name="xmlConfiguration"></param>
      /// <param name="xmlBusinessUnit"> </param>
      /// <param name="xmlDepartment"> </param>
      /// <param name="startModelPart"></param>
      /// <param name="startAsmPart"></param>
      /// <returns></returns>
      public static bool CreateDaimerStartPartsForEnhancement(Session currentSession, string referenceGeoPartPath, string origModelPartPath, string xmlConfiguration,
                                                               string xmlBusinessUnit, string xmlDepartment, out Part startModelPart, out Part startAsmPart)
      {
         startModelPart = null;
         startAsmPart = null;
    
         // Check input
         if (null == currentSession)
         {
            return false;
         }

         if (String.IsNullOrEmpty(referenceGeoPartPath))
         {
            return false;
         }

         Part workPart = currentSession.Parts.Work;
         string targetPath = Path.GetDirectoryName(workPart.FullPath);

         if (String.IsNullOrEmpty(targetPath))
         {
            return false;
         }
        
         Component tmpComp = null;
         NXObject nXObject = null;

         bool foundPartInSession = false;
 
         try
         {
            Point3d basePoint1 = new Point3d(0.0, 0.0, 0.0);
            Matrix3x3 orientation1;
            orientation1.Xx = 1.0;
            orientation1.Xy = 0.0;
            orientation1.Xz = 0.0;
            orientation1.Yx = 0.0;
            orientation1.Yy = 1.0;
            orientation1.Yz = 0.0;
            orientation1.Zx = 0.0;
            orientation1.Zy = 0.0;
            orientation1.Zz = 1.0;
            PartLoadStatus partLoadStatus1;


            // try to find part in current session 
            PartCollection partCollection =  currentSession.Parts;
            Part part;
            try
            {
               part  = partCollection.FindObject("daimler_assembly_SNR.prt") as Part;       
            }
            catch(Exception)
            {
               part = null;
            }
          
            if (null == part)
            {
               FileNew fileNew = currentSession.Parts.FileNew();

               fileNew.TemplateFileName = "daimler_startpart_assembly_mm.prt";
               fileNew.Application = FileNewApplication.Assemblies;
               fileNew.Units = Part.Units.Millimeters;
               fileNew.TemplateType = FileNewTemplateType.Item;
               fileNew.NewFileName = targetPath + "\\daimler_assembly_SNR.prt";
               fileNew.MasterFileName = "";
               fileNew.UseBlankTemplate = false;
               fileNew.MakeDisplayedPart = false;

               // Create component builder
               CreateNewComponentBuilder createNewComponentBuilder = workPart.AssemblyManager.CreateNewComponentBuilder();

               createNewComponentBuilder.ReferenceSet = CreateNewComponentBuilder.ComponentReferenceSetType.EntirePartOnly;
               createNewComponentBuilder.ReferenceSetName = "Entire Part";
               createNewComponentBuilder.NewComponentName = "ASSEMBLY_SNR1";
               createNewComponentBuilder.ReferenceSet = CreateNewComponentBuilder.ComponentReferenceSetType.Model;
               createNewComponentBuilder.NewFile = fileNew;

               // Commit
               nXObject = createNewComponentBuilder.Commit();

               // Destroy builder
               createNewComponentBuilder.Destroy();

               startAsmPart = nXObject.Prototype as Part;

               if (null != startAsmPart)
               {
                  // add component               
                  int startIndex = referenceGeoPartPath.LastIndexOf("\\", StringComparison.Ordinal);
                  string componentName = referenceGeoPartPath.Substring(startIndex + 1, referenceGeoPartPath.Length - startIndex - 1);

                  Component newComponent = startAsmPart.ComponentAssembly.AddComponent(referenceGeoPartPath,
                                                                                         "FINAL_PART",
                                                                                         componentName.ToUpper(),
                                                                                         basePoint1,
                                                                                         orientation1,
                                                                                         -1,
                                                                                         out partLoadStatus1,
                                                                                         true);
                  // Get component part
                  startModelPart = newComponent.Prototype as Part;

               }
            }
            else
            {
               startAsmPart = part;
               foundPartInSession = true;

                // Get rootComponent of current work part (asm start part)
               Component rootComponent = startAsmPart.ComponentAssembly.RootComponent;
               if (null != rootComponent)
               {
                  // Get children of root component 
                  Component[] children = rootComponent.GetChildren();

                  // Root component should have one child !!
                  if (children.Length == 1)
                  {
                     // Get child from array
                     Component child = children[0];

                     // Get PRT part from child (prototype)
                     startModelPart = child.Prototype as Part;                   
                  }
               }
            }

            if ( null != startAsmPart)
            {
               // Add tmp component to get the version to compare with
               tmpComp = startAsmPart.ComponentAssembly.AddComponent(origModelPartPath,
                                                                           "FINAL_PART",
                                                                           "MODEL_GEO1",
                                                                           basePoint1,
                                                                           orientation1,
                                                                           -1,
                                                                           out partLoadStatus1,
                                                                           true);

            }

            
         }
         catch(Exception ex)
         {
            UI.GetUI().NXMessageBox.Show("Part creation", NXMessageBox.DialogType.Error, ex.ToString());
            if ( null != nXObject)
            {
               Component newComp = nXObject as Component;
               workPart.ComponentAssembly.RemoveComponent(newComp);
               ChangeWorkPart(currentSession, startAsmPart, workPart); 
            }
            return false;        
         }
        
         if ( null == tmpComp )
         {
            Component newComp = nXObject as Component;
            workPart.ComponentAssembly.RemoveComponent(newComp);
            ChangeWorkPart(currentSession, startAsmPart, workPart);           
            return false;
         }

         // Get Part
         Part tmpPart = tmpComp.Prototype as Part;    

         // Get Version of the two parts and compare !! must be the same !
         string originalVersion;
         if (PartAttributeExists(tmpPart, StartPartVersion))
         {
            originalVersion = GetStringAttribute(tmpPart, StartPartVersion);
         }
         else
         {
            
            Component newComp = nXObject as Component;
            workPart.ComponentAssembly.RemoveComponent(newComp);
            ChangeWorkPart(currentSession, startAsmPart, workPart);
            
            string message = NoAttributeMsg + StartPartVersion;
            UI.GetUI().NXMessageBox.Show("Part creation", NXMessageBox.DialogType.Information, message);
            return false;
         }
         string version;
         if (PartAttributeExists(startModelPart, StartPartVersion))
         {
            version = GetStringAttribute(startModelPart, StartPartVersion);
         }
         else
         {
            Component newComp = nXObject as Component;
            workPart.ComponentAssembly.RemoveComponent(newComp);
            ChangeWorkPart(currentSession, startAsmPart, workPart);
            string message = NoAttributeMsg + StartPartVersion;
            UI.GetUI().NXMessageBox.Show("Part creation", NXMessageBox.DialogType.Information, message);
            return false;
         }

         // Remove temp Component
         startAsmPart.ComponentAssembly.RemoveComponent(tmpComp);

         if (PartAttributeExists(startModelPart, ConfigurationName))
         {
            string configuration = GetStringAttribute(startModelPart, ConfigurationName);
            if (configuration != xmlConfiguration)
            {
               Component newComp = nXObject as Component;
               workPart.ComponentAssembly.RemoveComponent(newComp);
               ChangeWorkPart(currentSession, startAsmPart, workPart);
               string message = "Configuration of the reference part is not the Configuration of the xml setting. \nCheck and try again.";
               UI.GetUI().NXMessageBox.Show("Part creation", NXMessageBox.DialogType.Information, message);
               return false;
            }
         }
         else
         {
            Component newComp = nXObject as Component;
            workPart.ComponentAssembly.RemoveComponent(newComp);
            ChangeWorkPart(currentSession, startAsmPart, workPart);
            string message = NoAttributeMsg + ConfigurationName;
            UI.GetUI().NXMessageBox.Show("Part creation", NXMessageBox.DialogType.Information, message);
            return false;
         }

         if (PartAttributeExists(startModelPart, ConfigurationBUnit))
         {
            string businessUnit = GetStringAttribute(startModelPart, ConfigurationBUnit);
            if (businessUnit != xmlBusinessUnit)
            {
               Component newComp = nXObject as Component;
               workPart.ComponentAssembly.RemoveComponent(newComp);
               ChangeWorkPart(currentSession, startAsmPart, workPart);
               string message = "Business-Unit of the reference part is not the Business-Unit of the xml setting. \nCheck and try again.";
               UI.GetUI().NXMessageBox.Show("Part creation", NXMessageBox.DialogType.Information, message);
               return false;
            }
         }
         else
         {
            Component newComp = nXObject as Component;
            workPart.ComponentAssembly.RemoveComponent(newComp);
            ChangeWorkPart(currentSession, startAsmPart, workPart);
            string message = NoAttributeMsg + ConfigurationBUnit;
            UI.GetUI().NXMessageBox.Show("Part creation", NXMessageBox.DialogType.Information, message);
            return false;
         }

         if (PartAttributeExists(startModelPart, ConfigurationDepartment))
         {
            string department = GetStringAttribute(startModelPart, ConfigurationDepartment);
            if (department != xmlDepartment)
            {
               Component newComp = nXObject as Component;
               workPart.ComponentAssembly.RemoveComponent(newComp);
               ChangeWorkPart(currentSession, startAsmPart, workPart);
               string message = "The Department of the reference part is not the department of the xml setting. \nCheck and try again.";
               UI.GetUI().NXMessageBox.Show("Part creation", NXMessageBox.DialogType.Information, message);
               return false;
            }
         }
         else
         {
            Component newComp = nXObject as Component;
            workPart.ComponentAssembly.RemoveComponent(newComp);
            ChangeWorkPart(currentSession, startAsmPart, workPart);
            string message = NoAttributeMsg + ConfigurationDepartment;
            UI.GetUI().NXMessageBox.Show("Part creation", NXMessageBox.DialogType.Information, message);
            return false;
         }

         // Version muss gleich sein: 
         if (version != originalVersion)
         {
            Component newComp = nXObject as Component;
            workPart.ComponentAssembly.RemoveComponent(newComp);
            ChangeWorkPart(currentSession, startAsmPart, workPart);
            string message = "Version of the reference part is not equal to Daimler part Version. \nCheck and try again.";
            UI.GetUI().NXMessageBox.Show("Part creation", NXMessageBox.DialogType.Information, message);

            return false;
         }

         // Check part 
         string messageIfFailed;
         bool checkOk = CheckPart(startAsmPart, out messageIfFailed);

         if (checkOk && (!foundPartInSession) )
         {
            // Save As... the component in the target path              
            if (null != startModelPart)
            {
               startModelPart.SaveAs(targetPath + "\\" + Path.GetFileName(startModelPart.FullPath));
            }

            // Save the part including component
            PartSaveStatus partSaveStatus = startAsmPart.Save(BasePart.SaveComponents.True, BasePart.CloseAfterSave.False);
            partSaveStatus.Dispose();

            // Remove component
            Component newComp = nXObject as Component;
            workPart.ComponentAssembly.RemoveComponent(newComp);
         }

         else
         {
            if( null != nXObject)
            {
               Component newComp = nXObject as Component;
               workPart.ComponentAssembly.RemoveComponent(newComp);
               ChangeWorkPart(currentSession, startAsmPart, workPart);
               UI.GetUI().NXMessageBox.Show("Part creation", NXMessageBox.DialogType.Information, messageIfFailed);
               return false;
            }
         }

         return true;
      }


      /// <summary>Create daimler start parts</summary>
      /// <param name="currentSession">The current session</param>
      /// <param name="targetPath">The target path.</param>
      /// <param name="referenceGeoPartPath">The reference geo part path.</param>
      /// <param name="origModelPartPath">The original model part path. </param>
      /// <param name="xmlConfiguration">The xml configuration to compare </param>
      /// <param name="xmlBusinessUnit">The xml Business unit to compare </param>
      /// <param name="xmlDepartment">The xml Department to compare </param>
      public static bool CreateDaimlerStartParts(Session currentSession,   string targetPath, 
                                                                           string referenceGeoPartPath, 
                                                                           string origModelPartPath, 
                                                                           string xmlConfiguration,
                                                                           string xmlBusinessUnit, 
                                                                           string xmlDepartment)   
      {
         
         // Check input
         if ( null == currentSession )
         {
            return false;
         }

         if (String.IsNullOrEmpty(targetPath))
         {
            return false;
         }

         if (String.IsNullOrEmpty(referenceGeoPartPath))
         {
            return false;
         }

         // initialize some variables
         Component tmpComp = null;
         Part newCurrentWorkPart = null;
         Part oldWorkPart = currentSession.Parts.Work;
         Part prtPart = null;

         bool foundPartInSession = false;

         try
         {
            Point3d basePoint1 = new Point3d(0.0, 0.0, 0.0);
            Matrix3x3 orientation1;
            orientation1.Xx = 1.0;
            orientation1.Xy = 0.0;
            orientation1.Xz = 0.0;
            orientation1.Yx = 0.0;
            orientation1.Yy = 1.0;
            orientation1.Yz = 0.0;
            orientation1.Zx = 0.0;
            orientation1.Zy = 0.0;
            orientation1.Zz = 1.0;
            PartLoadStatus partLoadStatus1;

            // try to find part in current session 
            PartCollection partCollection =  currentSession.Parts;
            Part part;
            try
            {
               part  = partCollection.FindObject("daimler_assembly_SNR.prt") as Part;       
            }
            catch(Exception)
            {
               part = null;
            }

            if (null == part)
            {

               // File new
               FileNew fileNew = currentSession.Parts.FileNew();

               fileNew.TemplateFileName = "daimler_startpart_assembly_mm.prt";
               fileNew.Application = FileNewApplication.Assemblies;
               fileNew.Units = Part.Units.Millimeters;
               fileNew.TemplateType = FileNewTemplateType.Item;
               fileNew.NewFileName = targetPath + "daimler_assembly_SNR.prt";
               fileNew.MasterFileName = "";
               fileNew.UseBlankTemplate = false;
               fileNew.MakeDisplayedPart = true;

               // Commit
               fileNew.Commit();

               // Destroy
               fileNew.Destroy();

               // Get current work part (this is the new one)
               newCurrentWorkPart = currentSession.Parts.Work;

               // add component           
               int startIndex = referenceGeoPartPath.LastIndexOf("\\", StringComparison.Ordinal);
               string componentName = referenceGeoPartPath.Substring(startIndex + 1,
                                                                     referenceGeoPartPath.Length - startIndex - 1);

               Component newComponent = newCurrentWorkPart.ComponentAssembly.AddComponent(referenceGeoPartPath,
                                                                                          "FINAL_PART",
                                                                                          componentName.ToUpper(),
                                                                                          basePoint1,
                                                                                          orientation1,
                                                                                          -1,
                                                                                          out partLoadStatus1,
                                                                                          true);
               // Get component part
               prtPart = newComponent.Prototype as Part;
              

            }
            else
            {
               newCurrentWorkPart = part;

               // Make the old work part to work part 
               PartLoadStatus partLoadStatus;
               PartCollection.SdpsStatus status = currentSession.Parts.SetDisplay(newCurrentWorkPart, true, true, out partLoadStatus);

               foundPartInSession = true;

               // Get rootComponent of current work part (asm start part)
               Component rootComponent = newCurrentWorkPart.ComponentAssembly.RootComponent;
               if (null != rootComponent)
               {
                  // Get children of root component 
                  Component[] children = rootComponent.GetChildren();

                  // Root component should have one child !!
                  if (children.Length == 1)
                  {
                     // Get child from array
                     Component child = children[0];

                     // Get PRT part from child (prototype)
                     prtPart = child.Prototype as Part;
                  }
               }
            }
     
            // Add tmp component to get the version to compare with
            tmpComp = newCurrentWorkPart.ComponentAssembly.AddComponent(origModelPartPath,
                                                                        "FINAL_PART",
                                                                        "MODEL_GEO1",
                                                                        basePoint1,
                                                                        orientation1,
                                                                        -1,
                                                                        out partLoadStatus1,
                                                                        true);

            


         }
         catch (Exception ex)
         {
            UI.GetUI().NXMessageBox.Show("Part creation", NXMessageBox.DialogType.Error, ex.ToString());

            // try to change work part (to the old one)
            ChangeWorkPart(currentSession, newCurrentWorkPart, oldWorkPart);
            return false;
         }


         // Get Part
         Part tmpPart = tmpComp.Prototype as Part;

         // Get Version of the two parts and compare !! must be the same !
         string originalVersion;
         if (PartAttributeExists(tmpPart, StartPartVersion))
         {
            originalVersion = GetStringAttribute(tmpPart, StartPartVersion);
         }
         else
         {
            ChangeWorkPart(currentSession, newCurrentWorkPart, oldWorkPart);
            string message = NoAttributeMsg + StartPartVersion;
            UI.GetUI().NXMessageBox.Show("Part creation", NXMessageBox.DialogType.Information, message);
            return false;
         }
         string version;
         if (PartAttributeExists(tmpPart, StartPartVersion))
         {
            version = GetStringAttribute(prtPart, StartPartVersion);
         }
         else
         {
            ChangeWorkPart(currentSession, newCurrentWorkPart, oldWorkPart);
            string message = NoAttributeMsg + StartPartVersion;
            UI.GetUI().NXMessageBox.Show("Part creation", NXMessageBox.DialogType.Information, message);
            return false;
         }

         // Remove temp Component
         newCurrentWorkPart.ComponentAssembly.RemoveComponent(tmpComp);

         if (PartAttributeExists(prtPart, ConfigurationName))
         {
            string configuration = GetStringAttribute(prtPart, ConfigurationName);
            if ( configuration != xmlConfiguration)
            {
               ChangeWorkPart(currentSession, newCurrentWorkPart, oldWorkPart);
               string message = "Configuration of the reference part is not the Configuration of the xml setting. \nCheck and try again.";
               UI.GetUI().NXMessageBox.Show("Part creation", NXMessageBox.DialogType.Information, message);
               return false;
            }
         }
         else
         {
            ChangeWorkPart(currentSession, newCurrentWorkPart, oldWorkPart);
            string message = NoAttributeMsg + ConfigurationName;
            UI.GetUI().NXMessageBox.Show("Part creation", NXMessageBox.DialogType.Information, message);
            return false;
         }

         if (PartAttributeExists(prtPart, ConfigurationBUnit))
         {
            string businessUnit = GetStringAttribute(prtPart, ConfigurationBUnit);
            if (businessUnit != xmlBusinessUnit)
            {
               ChangeWorkPart(currentSession, newCurrentWorkPart, oldWorkPart);
               string message = "Business-Unit of the reference part is not the Business-Unit of the xml setting. \nCheck and try again.";
               UI.GetUI().NXMessageBox.Show("Part creation", NXMessageBox.DialogType.Information, message);
               return false;
            }
         }
         else
         {
            ChangeWorkPart(currentSession, newCurrentWorkPart, oldWorkPart);
            string message = NoAttributeMsg + ConfigurationBUnit;
            UI.GetUI().NXMessageBox.Show("Part creation", NXMessageBox.DialogType.Information, message);
            return false;
         }

         if (PartAttributeExists(prtPart, ConfigurationDepartment))
         {
            string department = GetStringAttribute(prtPart, ConfigurationDepartment);
            if (department != xmlDepartment)
            {
               ChangeWorkPart(currentSession, newCurrentWorkPart, oldWorkPart);
               string message = "The Department of the reference part is not the department of the xml setting. \nCheck and try again.";
               UI.GetUI().NXMessageBox.Show("Part creation", NXMessageBox.DialogType.Information, message);
               return false;
            }
         }
         else
         {
            ChangeWorkPart(currentSession, newCurrentWorkPart, oldWorkPart);
            string message = NoAttributeMsg + ConfigurationDepartment;
            UI.GetUI().NXMessageBox.Show("Part creation", NXMessageBox.DialogType.Information, message);
            return false;
         }

         // Version muss gleich sein: 
         if (version != originalVersion)
         {
            ChangeWorkPart(currentSession, newCurrentWorkPart, oldWorkPart);
            string message = "Version of the reference part is not equal to Daimler part Version. \nCheck and try again.";
            UI.GetUI().NXMessageBox.Show("Part creation", NXMessageBox.DialogType.Information, message);

            return false;
         }

         // Check part 
         string messageIfFailed;
         bool checkOk = CheckPart(newCurrentWorkPart, out messageIfFailed);

         if (checkOk && (!foundPartInSession) )
         {
            // Save As... the component in the target path              
            if (null != prtPart)
            {
               prtPart.SaveAs(targetPath + Path.GetFileName(prtPart.FullPath));
            }

            // Save the part including component
            PartSaveStatus partSaveStatus = newCurrentWorkPart.Save(BasePart.SaveComponents.True, BasePart.CloseAfterSave.False);
            partSaveStatus.Dispose();
         }

         else
         {
            if (!foundPartInSession)
            {
               ChangeWorkPart(currentSession, newCurrentWorkPart, oldWorkPart);
               UI.GetUI().NXMessageBox.Show("Part creation", NXMessageBox.DialogType.Information, messageIfFailed);

               return false;
            }     
         }
         return true;
      }


      /// <summary>Change work part</summary>
      /// <param name="theSession">The current session</param>
      /// <param name="partToClose">The part to close</param>
      /// <param name="partToSetToWorkPart">The part to set to work part.</param>
      public static void ChangeWorkPart(Session theSession, Part partToClose, Part partToSetToWorkPart)
      {
         // check inputs
         if (null == theSession)
            return;
         if (null == partToClose)
            return;
         if (null == partToSetToWorkPart)
            return;

         // Close the current workpart
         partToClose.Close(BasePart.CloseWholeTree.False, BasePart.CloseModified.UseResponses, null);

         // Make the old work part to work part 
         PartLoadStatus partLoadStatus;
         PartCollection.SdpsStatus status = theSession.Parts.SetDisplay(partToSetToWorkPart, true, true, out partLoadStatus);

         Part newworkPart = theSession.Parts.Work;
         theSession.Parts.SetWork(newworkPart);
      }


   }
}
