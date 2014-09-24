//==============================================================================
//
//        Filename: StructureCreation.cs
//
//        Created by: CENIT AG (Jan Assmann)
//              Version: NX 8.5.2.3 MP1
//              Date: 11-11-2013  (Format: mm-dd-yyyy)
//              Time: 08:30 (Format: hh-mm)
//
//==============================================================================

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Xml;
using Microsoft.Office.Interop.Excel;
using NXOpen;
using NXOpen.Assemblies;
using NXOpen.BlockStyler;
using System.Windows.Forms;
using NXOpen.UF;
using Application = Microsoft.Office.Interop.Excel.Application;
using Assembly = System.Reflection.Assembly;


namespace Daimler.NX.BemiStructure
{
   public class StructureCreation
   {
      //------------------------------------------------------------------
      //-------------------- Class Members -------------------------------
      //------------------------------------------------------------------

      #region Class Members


      private const string RniPartType = "RNI_PARTTYPE";

      //class members
      private static Session theSession;
      private static UI theUI;
      private string theDlxFileName;

      private string ownDllLocation;

      private BlockDialog theDialog;

      private NXOpen.BlockStyler.Group groupStructureOption; // Block type: Group
      private NXOpen.BlockStyler.Group groupFile; // Block type: Group
      private NXOpen.BlockStyler.Group groupImportOptions; // Block type: Group
      private NXOpen.BlockStyler.Group groupExport; // Block type: Group
      private NXOpen.BlockStyler.Group groupEnhanceExStructure; // Block type: Group 
      private NXOpen.BlockStyler.Group groupValidation;// Block type: Group

      private NXOpen.BlockStyler.Button buttonOpenFile; // Block type: Button
      private NXOpen.BlockStyler.Button exportFilePath; // Block type: Button
      private NXOpen.BlockStyler.Button validationButton;// Block type: Button

      private NXOpen.BlockStyler.SelectObject componentSelection; // Block type: Selection

      private StringBlock fileName; // Block type: String
      private StringBlock fileExportPath; // Block type: String
      private StringBlock fileExportName; // Block type: String 

      private MultilineString multilineValResult;// Block type: Multiline String

      private Enumeration RadioBoxStructureOption; // Block type: Enumeration
      
      private Toggle toggleUseExcelSession; // Block type: Toggle 
      private Toggle toggleAddPosNb; // Block type: Toggle  
      private Toggle toggleCreateTmpFile; // Block type: Toggle

      private NXOpen.BlockStyler.Label labelSaveBeforeValidation;// Block type: Label

      private FolderSelection structurePathBrowser;// Block type: NativeFolderBrowser

      public static readonly int SnapPointTypesEnabled_UserDefined = (1 << 0);
      public static readonly int SnapPointTypesEnabled_Inferred = (1 << 1);
      public static readonly int SnapPointTypesEnabled_ScreenPosition = (1 << 2);
      public static readonly int SnapPointTypesEnabled_EndPoint = (1 << 3);
      public static readonly int SnapPointTypesEnabled_MidPoint = (1 << 4);
      public static readonly int SnapPointTypesEnabled_ControlPoint = (1 << 5);
      public static readonly int SnapPointTypesEnabled_Intersection = (1 << 6);
      public static readonly int SnapPointTypesEnabled_ArcCenter = (1 << 7);
      public static readonly int SnapPointTypesEnabled_QuadrantPoint = (1 << 8);
      public static readonly int SnapPointTypesEnabled_ExistingPoint = (1 << 9);
      public static readonly int SnapPointTypesEnabled_PointonCurve = (1 << 10);
      public static readonly int SnapPointTypesEnabled_PointonSurface = (1 << 11);
      public static readonly int SnapPointTypesEnabled_PointConstructor = (1 << 12);
      public static readonly int SnapPointTypesEnabled_TwocurveIntersection = (1 << 13);
      public static readonly int SnapPointTypesEnabled_TangentPoint = (1 << 14);
      public static readonly int SnapPointTypesEnabled_Poles = (1 << 15);
      public static readonly int SnapPointTypesEnabled_BoundedGridPoint = (1 << 16);
      public static readonly int SnapPointTypesOnByDefault_EndPoint = (1 << 3);
      public static readonly int SnapPointTypesOnByDefault_MidPoint = (1 << 4);
      public static readonly int SnapPointTypesOnByDefault_ControlPoint = (1 << 5);
      public static readonly int SnapPointTypesOnByDefault_Intersection = (1 << 6);
      public static readonly int SnapPointTypesOnByDefault_ArcCenter = (1 << 7);
      public static readonly int SnapPointTypesOnByDefault_QuadrantPoint = (1 << 8);
      public static readonly int SnapPointTypesOnByDefault_ExistingPoint = (1 << 9);
      public static readonly int SnapPointTypesOnByDefault_PointonCurve = (1 << 10);
      public static readonly int SnapPointTypesOnByDefault_PointonSurface = (1 << 11);
      public static readonly int SnapPointTypesOnByDefault_PointConstructor = (1 << 12);
      public static readonly int SnapPointTypesOnByDefault_BoundedGridPoint = (1 << 16);

      #endregion Class Members


      //------------------------------------------------------------------
      //-------------------- Dialog Handling Methods ---------------------
      //------------------------------------------------------------------

      #region Dialog Handling Methods


      //-------------------------------------------------------------------
      //Constructor for NX Styler class
      //-------------------------------------------------------------------
      public StructureCreation(Session session)
      {
         // Get session
         theSession = session;

         // Get the ui
         theUI = UI.GetUI();

         // Get executing assembly
         Assembly assembly = Assembly.GetExecutingAssembly();

         // save own dll location
         ownDllLocation = Path.GetDirectoryName(assembly.Location);

         // Get manifest resource names
         string[] names = assembly.GetManifestResourceNames();

         // Get dlx stream 
         Stream dlxStream = assembly.GetManifestResourceStream(names[0]);

         // Get temp path + structureCreation.dlx
         string fileFullPath = Path.GetTempPath() + "StructureCreation.dlx";

         // store this stream  to file --> temporary and this will immediately destroyed in dialogShown_cb
         if ( null != dlxStream )
         {
            // Create a FileStream object to write a stream to a file
            using (FileStream fileStream = File.Create(fileFullPath, (int)dlxStream.Length))
            {
               // Fill the bytes[] array with the stream data
               byte[] bytesInStream = new byte[dlxStream.Length];
               dlxStream.Read(bytesInStream, 0, (int)bytesInStream.Length);

               // Use FileStream object to write to the specified file
               fileStream.Write(bytesInStream, 0, bytesInStream.Length);
            }
         }
         // ---------------------------------------------------------------------------

         theDlxFileName = fileFullPath;   

         theDialog = theUI.CreateDialog(theDlxFileName);
         theDialog.AddOkHandler(ok_cb);
         theDialog.AddUpdateHandler(update_cb);
         theDialog.AddCancelHandler(cancel_cb);
         theDialog.AddFilterHandler(filter_cb);
         theDialog.AddInitializeHandler(initialize_cb);
         theDialog.AddEnableOKButtonHandler(enableOKButton_cb);
         theDialog.AddDialogShownHandler(dialogShown_cb);
      }

      //------------------------------------------------------------------
      //This method shows the dialog on the screen
      //------------------------------------------------------------------
      public NXOpen.UIStyler.DialogResponse Show()
      {
         try
         {
            theDialog.Show();
         }
         catch (Exception ex)
         {
            theUI.NXMessageBox.Show("Block Styler", NXMessageBox.DialogType.Error, ex.ToString());
         }
         return 0;
      }

      //------------------------------------------------------------------
      //Method Name: Dispose
      //------------------------------------------------------------------
      public void Dispose()
      {
         if (theDialog != null)
         {
            theDialog.Dispose();
            theDialog = null;
         }
      }


      #endregion Dialog Handling Methods


      //------------------------------------------------------------------
      //-------------------- Block UI Styler Callback Functions ----------
      //------------------------------------------------------------------

      #region Callback Functions


      //------------------------------------------------------------------
      //Callback Name: initialize_cb
      //------------------------------------------------------------------
      public void initialize_cb()
      {
         try
         {
            groupStructureOption = (NXOpen.BlockStyler.Group)theDialog.TopBlock.FindBlock("groupStructureOption");
            RadioBoxStructureOption = (Enumeration)theDialog.TopBlock.FindBlock("RadioBoxStructureOption");
            structurePathBrowser = (FolderSelection)theDialog.TopBlock.FindBlock("structurePathBrowser");
            groupFile = (NXOpen.BlockStyler.Group)theDialog.TopBlock.FindBlock("groupFile");
            toggleUseExcelSession = (Toggle)theDialog.TopBlock.FindBlock("toggleUseExcelSession");
            buttonOpenFile = (NXOpen.BlockStyler.Button)theDialog.TopBlock.FindBlock("buttonOpenFile");
            fileName = (StringBlock)theDialog.TopBlock.FindBlock("fileName");
            groupImportOptions = (NXOpen.BlockStyler.Group)theDialog.TopBlock.FindBlock("groupImportOptions");
            toggleAddPosNb = (Toggle)theDialog.TopBlock.FindBlock("toggleAddPosNb");
            groupValidation = (NXOpen.BlockStyler.Group)theDialog.TopBlock.FindBlock("groupValidation");
            validationButton = (NXOpen.BlockStyler.Button)theDialog.TopBlock.FindBlock("validationButton");
            multilineValResult = (MultilineString)theDialog.TopBlock.FindBlock("multilineValResult");
            groupExport = (NXOpen.BlockStyler.Group)theDialog.TopBlock.FindBlock("groupExport");
            toggleCreateTmpFile = (Toggle)theDialog.TopBlock.FindBlock("toggleCreateTmpFile");
            exportFilePath = (NXOpen.BlockStyler.Button)theDialog.TopBlock.FindBlock("exportFilePath");
            fileExportPath = (StringBlock)theDialog.TopBlock.FindBlock("fileExportPath");
            fileExportName = (StringBlock)theDialog.TopBlock.FindBlock("fileExportName");
            groupEnhanceExStructure = (NXOpen.BlockStyler.Group)theDialog.TopBlock.FindBlock("groupEnhanceExStructure");
            componentSelection = (NXOpen.BlockStyler.SelectObject)theDialog.TopBlock.FindBlock("componentSelection");
            labelSaveBeforeValidation = (NXOpen.BlockStyler.Label)theDialog.TopBlock.FindBlock("labelSaveBeforeValidation");
         }
         catch (Exception ex)
         {
            theUI.NXMessageBox.Show("Block Styler", NXMessageBox.DialogType.Error, ex.ToString());
         }
      }

      //------------------------------------------------------------------
      //Callback Name: dialogShown_cb
      //------------------------------------------------------------------
      public void dialogShown_cb()
      {
         try
         {
            // update ui blocks visibilities
            UpdateUiBlocksVisibilities();

            // set empty value
            multilineValResult.SetValue(new string[] { });    

            // set maximum scope to "Entire Assembly" for comoonent selection (export)
            componentSelection.MaximumScopeAsString = "Entire Assembly";

            // delete the temp dlx file, which is only used for dialog creation 
            File.Delete(theDlxFileName);

         }
         catch (Exception ex)
         {
            theUI.NXMessageBox.Show("Block Styler", NXMessageBox.DialogType.Error, ex.ToString());
         }
      }

      //------------------------------------------------------------------
      //Callback Name: update_cb
      //------------------------------------------------------------------
      public int update_cb(UIBlock block)
      {
         try
         {
            if (block == validationButton)
            {
               TreeStructure rootTreeNode;
               ReadEnhanceExcelfile(out rootTreeNode);

               if ( null != rootTreeNode )
               {
                  // Enhance assembly structure 
                  Dictionary<string, string> resultList = EnhanceAssemblyStructure(rootTreeNode, null, null, true);

                  // fill result list
                  FillValidationResult(resultList);
               }            
            }
            if (block == RadioBoxStructureOption)
            {
               UpdateUiBlocksVisibilities();
            }
            else if (block == buttonOpenFile)
            {
               CreateOpenFileDialog();
            }
            else if (block == toggleUseExcelSession)
            {
               if (toggleUseExcelSession.Value)
               {
                  buttonOpenFile.Enable = false;
                  labelSaveBeforeValidation.Enable = true;
               }
               else
               {
                  buttonOpenFile.Enable = true;
                  labelSaveBeforeValidation.Enable = false;
               }
            }
         }
         catch (Exception ex)
         {
            theUI.NXMessageBox.Show("Block Styler", NXMessageBox.DialogType.Error, ex.ToString());
         }
         return 0;
      }

      //------------------------------------------------------------------
      //Callback Name: ok_cb
      //------------------------------------------------------------------
      public int ok_cb()
      {
         int errorCode = 0;
         try
         {
            string structureCreationType = RadioBoxStructureOption.ValueAsString;

            if ("Load / create new structure" == structureCreationType)
            {
               // Read settings file 
               string referenceGeoPartPath, origModelPartPath, configuration, businessUnit, department;
               bool readSettingOk = ReadSettingsXmlFile(out referenceGeoPartPath, out origModelPartPath, out configuration, out businessUnit, out department);

               Part oldWorkPart = theSession.Parts.Work;
                  
               if (readSettingOk)
               {
                  // Read excel file
                  TreeStructure rootTreeNode;
                  ReadExcelFile(out rootTreeNode);

                  bool exist = CheckExistingFiles(rootTreeNode);
                  if (exist)
                  {
                     string msg = "Minimum one part exists in the current folder. \n";
                     theUI.NXMessageBox.Show("Note", NXMessageBox.DialogType.Information, msg);
                  }
                  else
                  {
                     // target path from UI
                     string targetPath = structurePathBrowser.Path;

                     // Create StartPart to copy from 
                     bool partCreationOk = PartStatics.CreateDaimlerStartParts(theSession, targetPath, referenceGeoPartPath, origModelPartPath, configuration, businessUnit, department);

                     if (partCreationOk)
                     {
                        Part currentWorkPart = theSession.Parts.Work;

                        // Create NX assembly structure
                        bool successfulCreation = CreateAssemblyStructure(rootTreeNode);

                        if ( successfulCreation )
                        {
                           // save work part and all his components
                           PartSaveStatus partSaveStatus = currentWorkPart.Save(BasePart.SaveComponents.True, BasePart.CloseAfterSave.False);
                           partSaveStatus.Dispose();

                           // show message
                           theUI.NXMessageBox.Show("Assembly creation", NXMessageBox.DialogType.Information, "Makro has completed successfully.");
                        }
                        else
                        {
                           PartStatics.ChangeWorkPart(theSession, currentWorkPart, oldWorkPart);
                           theUI.NXMessageBox.Show("Assembly creation", NXMessageBox.DialogType.Information, "Makro has not completed successfully.");
                        } 
                     }
                  }           
               }
            }
            else if ("Export nodes to enhance existing structure" == structureCreationType)
            { 
               // Export to excel
               ExportToExcel();
            }
            else
            {
               // Read settings file 
               string referenceGeoPartPath, origModelPartPath, configuration, businessUnit, department;
               bool readSettingOk = ReadSettingsXmlFile(out referenceGeoPartPath, out origModelPartPath, out configuration, out businessUnit, out department);

               if (readSettingOk)
               {
                  TreeStructure rootTreeNode;
                  ReadEnhanceExcelfile(out rootTreeNode);

                  if ( rootTreeNode.HasChildrenLevel1() )
                  {
                     Part startModelPart, startAsmPart;
                     bool startPartCreation = PartStatics.CreateDaimerStartPartsForEnhancement(theSession, referenceGeoPartPath, origModelPartPath, configuration, businessUnit, department, out startModelPart, out startAsmPart);

                     if (startPartCreation)
                     {
                        // Enhance assembly structure 
                        Dictionary<string, string> resultList = EnhanceAssemblyStructure(rootTreeNode, startAsmPart,
                                                                                         startModelPart, false);

                        // save work part and all his components
                        Part currentWorkPart = theSession.Parts.Work;
                        PartSaveStatus partSaveStatus = currentWorkPart.Save(BasePart.SaveComponents.True,
                                                                             BasePart.CloseAfterSave.False);
                        partSaveStatus.Dispose();

                        string[] listKeys = new string[resultList.Count];
                        resultList.Keys.CopyTo(listKeys, 0);

                        string result = "";
                        if (listKeys.Length > 0)
                        {
                           foreach (var key in listKeys)
                           {
                              string row = key + ": not imported !";
                              result = result + row + "\n";
                           }
                           // show result list
                           theUI.NXMessageBox.Show("Import result", NXMessageBox.DialogType.Information, result + "\n \nPlease close Excel.");
                        }
                        else
                        {
                           theUI.NXMessageBox.Show("Import result", NXMessageBox.DialogType.Information, "Import has completed successfully. \n \nPlease close Excel.");
                        }
                     }
                  }
                  else
                  {
                     theUI.NXMessageBox.Show("Import result", NXMessageBox.DialogType.Information, "There are no entries in level 1 in the excel file.");
                  }
               }
            }
         }
         catch (Exception ex)
         {
            errorCode = 1;
            theUI.NXMessageBox.Show("Block Styler", NXMessageBox.DialogType.Error, ex.ToString());
         }
         return errorCode;
      }

      //------------------------------------------------------------------
      //Callback Name: cancel_cb
      //------------------------------------------------------------------
      public int cancel_cb()
      {
         try
         {
         }
         catch (Exception ex)
         {
            theUI.NXMessageBox.Show("Block Styler", NXMessageBox.DialogType.Error, ex.ToString());
         }
         return 0;
      }

      //------------------------------------------------------------------
      //Callback Name: filter_cb
      //------------------------------------------------------------------
      public int filter_cb(UIBlock block, TaggedObject selectedObject)
      {
         // Initialize Flag for not allowed selected Object   
         int nfilterselectedObject = UFConstants.UF_UI_SEL_ACCEPT;

         if ( block == componentSelection )
         {
            Component selObject = selectedObject as Component;
            if (null == selObject)
            {
               nfilterselectedObject = UFConstants.UF_UI_SEL_REJECT;
            }
            else
            {
               Part partOfComponent = selObject.Prototype as Part;
               if ( null != partOfComponent )
               {
                  string partType = PartStatics.GetStringAttribute(partOfComponent, RniPartType);
                  if ( partType != "ASM")
                  {
                     nfilterselectedObject = UFConstants.UF_UI_SEL_REJECT;
                  }
               }
            }
         }    

         // Return Flag for not allowed selected Object
         return nfilterselectedObject;
      }

      //------------------------------------------------------------------
      //Callback Name: enableOKButton_cb
      //------------------------------------------------------------------
      public bool enableOKButton_cb()
      {
         bool enableOkButton = true;
         try
         {
            if ("Load / create new structure" == RadioBoxStructureOption.ValueAsString)
            {
               enableOkButton = false;

               string filePath = fileName.Value;
               string structureTargePath = structurePathBrowser.Path;

               if (!String.IsNullOrEmpty(filePath) && !String.IsNullOrEmpty(structureTargePath))
               {
                  enableOkButton = true;
               }
            }
            else if ("Import enhanced structure" == RadioBoxStructureOption.ValueAsString)
            {
               if (!toggleUseExcelSession.Value)
               {
                  enableOkButton = false;

                  string filePath = fileName.Value;
                  if (!String.IsNullOrEmpty(filePath))
                  {
                     enableOkButton = true;
                  }
               }

               string[] validationResult = multilineValResult.GetValue();
               if ( validationResult.Length == 0)
               {
                  enableOkButton = false;
               }
            }
         }
         catch (Exception ex)
         {
            theUI.NXMessageBox.Show("Block Styler", NXMessageBox.DialogType.Error, ex.ToString());
         }
         return enableOkButton;    
      }

      //------------------------------------------------------------------------------
      //Function Name: GetBlockProperties
      //Returns the propertylist of the specified BlockID
      //------------------------------------------------------------------------------
      public PropertyList GetBlockProperties(string blockID)
      {
         PropertyList plist = null;
         try
         {
            plist = theDialog.GetBlockProperties(blockID);
         }
         catch (Exception ex)
         {
            theUI.NXMessageBox.Show("Block Styler", NXMessageBox.DialogType.Error, ex.ToString());
         }
         return plist;
      }

      #endregion Callback Functions


      //------------------------------------------------------------------
      //-------------------- Private Dialog Methods --------------------
      //------------------------------------------------------------------

      #region Private Methods


      /// <summary>Read setting xml file.</summary>
      /// <param name="referenceGeoPartPath">Reference geo part string path.</param>
      /// <param name="origModelPartPath">Original part path.</param>
      /// <param name="configuration">Reference part configuration name.</param>
      /// <param name="businessUnit">Reference part configuration business unit </param>
      /// <param name="department">Reference part configuration department </param>
      /// <returns>If successful or not</returns>
      private bool ReadSettingsXmlFile(out string referenceGeoPartPath, out string origModelPartPath, out string configuration, out string businessUnit, out string department)
      {
         // Initialize output values:
         referenceGeoPartPath = "";
         origModelPartPath = "";
         configuration = "";
         businessUnit = "";
         department = "";


         // Check if dll location is empty (should normally not happens)
         if (String.IsNullOrEmpty(ownDllLocation))
         {
            return false;
         }
         
         // Build path of xml file path (must be same location as the own dll
         string xmlFilePath = ownDllLocation + "\\SmaragdStrukturMakroSettings.xml";

         // Check if xml setting file exists
         if (!File.Exists(xmlFilePath))
         {
            string msg = "Setting file does not exists: " + xmlFilePath + "\n Please check and try again.";
            theUI.NXMessageBox.Show("Read xml settings file", NXMessageBox.DialogType.Information, msg);
            return false;
         }

         // Create new Xml document
         XmlDocument doc = new XmlDocument();

         try
         {
            // load xml file in document
            doc.Load(xmlFilePath);
         }
         catch (Exception)
         {
            const string msg = "Problem with the setting xml file. \n Check and try again.";
            theUI.NXMessageBox.Show("Read xml settings file", NXMessageBox.DialogType.Information, msg);
            return false;
         }

         // Get settings root:
         XmlNode rootNode = doc.DocumentElement;

         // Check root node
         if ( null == rootNode)
         {
            const string msg = "Problem with the setting xml file. \n Root node is empty. \n Check and try again.";
            theUI.NXMessageBox.Show("Read xml settings file", NXMessageBox.DialogType.Information, msg);
            return false;
         }
        
         // check name of root 
         if (rootNode.Name != "GeneralSettings")
         {
            const string msg = "Problem with the setting xml file: \n Wrong root node (must be GeneralSettings) \n Check and try again.";
            theUI.NXMessageBox.Show("Read xml settings file", NXMessageBox.DialogType.Information, msg);
            return false;
         }

         // Check if root has childs
         if (!rootNode.HasChildNodes)
         {
            const string msg = "Problem with the setting xml file: \n Root node has no children. \n Check and try again.";
            theUI.NXMessageBox.Show("Read xml settings file", NXMessageBox.DialogType.Information, msg);
            return false;
         }

         // Get root child nodes
         XmlNodeList settinglist = rootNode.ChildNodes;

         Dictionary<string, string> settings = new Dictionary<string, string>();
         foreach (XmlNode setting in settinglist)
         {
            string name = setting.Name;
            if ( name == "Setting")
            {
               XmlAttributeCollection attributes = setting.Attributes;
               if (null != attributes)
               {
                  if (attributes.Count == 2 )
                  {
                     XmlAttribute key = attributes[0];
                     XmlAttribute value = attributes[1];

                     string attributeName = key.Name;
                     string attributeNameVal = value.Name;

                     if ( attributeName == "Key" && attributeNameVal == "Value")
                     {
                        string keyValue = key.Value;
                        string valValue = value.Value;

                        if (!settings.ContainsKey(keyValue))
                        {
                           settings.Add(keyValue, valValue);
                        }
                     }
                  }
               }
            }   
         }

         // Get Settings
         if (settings.ContainsKey("ReferenceModelPart"))
         {
           if (settings.TryGetValue("ReferenceModelPart", out referenceGeoPartPath))
           {
              if ( String.IsNullOrEmpty(referenceGeoPartPath))
              {
                 const string msg = "Problem with the setting xml file: \nValue of setting: 'ReferenceModelPart' is empty. \nCheck and try again.";
                 theUI.NXMessageBox.Show("Read xml settings file", NXMessageBox.DialogType.Information, msg);
                 return false;
              }
              if ( !File.Exists(referenceGeoPartPath))
              {
                 const string msg = "Problem with the setting xml file: \nReference ModelPart does not exists. \nCheck and try again.";
                 theUI.NXMessageBox.Show("Read xml settings file", NXMessageBox.DialogType.Information, msg);
                 return false;
              }
           }
         }
         else
         {
            const string msg = "Problem with the setting xml file: \nSetting: 'ReferenceModelPart' does not exist. \nCheck and try again.";
            theUI.NXMessageBox.Show("Read xml settings file", NXMessageBox.DialogType.Information, msg);
            return false;
         }


         if (settings.ContainsKey("ReferenceModelPartConfiguration_Name"))
         {
            if (settings.TryGetValue("ReferenceModelPartConfiguration_Name", out configuration))
            {
               if (String.IsNullOrEmpty(configuration))
               {
                  const string msg = "Problem with the setting xml file: \nValue of setting: 'ReferenceModelPartConfiguration_Name' is empty. \nCheck and try again.";
                  theUI.NXMessageBox.Show("Read xml settings file", NXMessageBox.DialogType.Information, msg);
                  return false;
               }
            }
         }
         else
         {
            const string msg = "Problem with the setting xml file: \nSetting: 'ReferenceModelPartConfiguration_Name' does not exist. \nCheck and try again.";
            theUI.NXMessageBox.Show("Read xml settings file", NXMessageBox.DialogType.Information, msg);
            return false;
         }

         if (settings.ContainsKey("ReferenceModelPartConfiguration_BUnit"))
         {
            if (settings.TryGetValue("ReferenceModelPartConfiguration_BUnit", out businessUnit))
            {
               if (String.IsNullOrEmpty(businessUnit))
               {
                  const string msg = "Problem with the setting xml file: \nValue of setting: 'ReferenceModelPartConfiguration_BUnit' is empty. \nCheck and try again.";
                  theUI.NXMessageBox.Show("Read xml settings file", NXMessageBox.DialogType.Information, msg);
                  return false;
               }
            }
         }
         else
         {
            const string msg = "Problem with the setting xml file: \nSetting: 'ReferenceModelPartConfiguration_BUnit' does not exist. \nCheck and try again.";
            theUI.NXMessageBox.Show("Read xml settings file", NXMessageBox.DialogType.Information, msg);
            return false;
         }

         if (settings.ContainsKey("ReferenceModelPartConfiguration_Department"))
         {
            if (settings.TryGetValue("ReferenceModelPartConfiguration_Department", out department))
            {
               if (String.IsNullOrEmpty(department))
               {
                  const string msg = "Problem with the setting xml file: \nValue of setting: 'ReferenceModelPartConfiguration_Department' is empty. \nCheck and try again.";
                  theUI.NXMessageBox.Show("Read xml settings file", NXMessageBox.DialogType.Information, msg);
                  return false;
               }
            }
         }
         else
         {
            const string msg = "Problem with the setting xml file: \nSetting: 'ReferenceModelPartConfiguration_Department' does not exist. \nCheck and try again.";
            theUI.NXMessageBox.Show("Read xml settings file", NXMessageBox.DialogType.Information, msg);
            return false;
         }

         if (settings.ContainsKey("ModelStartPart"))
         {
            if (settings.TryGetValue("ModelStartPart", out origModelPartPath))
            {
               if (String.IsNullOrEmpty(origModelPartPath))
               {
                  const string msg = "Problem with the setting xml file: \nValue of setting: 'ModelStartPart' is empty. \nCheck and try again.";
                  theUI.NXMessageBox.Show("Read xml settings file", NXMessageBox.DialogType.Information, msg);
                  return false;
               }

               if ( !File.Exists(origModelPartPath))
               {
                  const string msg = "Problem with the setting xml file: \nModel StartPart does not exists. \nCheck and try again.";
                  theUI.NXMessageBox.Show("Read xml settings file", NXMessageBox.DialogType.Information, msg);
                  return false;
               }
            }
         }
         else
         {
            const string msg = "Problem with the setting xml file: \nSetting: 'ModelStartPart' does not exist. \nCheck and try again.";
            theUI.NXMessageBox.Show("Read xml settings file", NXMessageBox.DialogType.Information, msg);
            return false;
         }

         return true;
      }


      /// <summary>Read enhance excel file </summary>
      /// <param name="rootTreeNode"></param>
      private void ReadEnhanceExcelfile(out TreeStructure rootTreeNode)
      {
         rootTreeNode = null;

         bool useCurrentExcelSession = toggleUseExcelSession.Value;

         if (!useCurrentExcelSession)
         {
            if (String.IsNullOrEmpty(fileName.Value))
            {
               const string message = "File string is empty";
               theUI.NXMessageBox.Show("Excel file failure", NXMessageBox.DialogType.Warning, message);
               return;
            }
            if (!File.Exists(fileName.Value))
            {
               const string message = "Selected file does not exists";
               theUI.NXMessageBox.Show("Excel file failure", NXMessageBox.DialogType.Warning, message);
               return;
            }
         }

         // Initialize Excel Application and Workbook
         Application excelApp = null;
         Workbook excelWorkbook = null;

         bool useOpenExcel = false;

         try
         {
            Worksheet excelworksheet;
            if (useCurrentExcelSession)
            {
               // Get current active excel application
               excelApp = (Application)Marshal.GetActiveObject("Excel.Application");

               // Get current active workbook
               excelWorkbook = excelApp.ActiveWorkbook;

               // Get current active sheet
               excelworksheet = excelWorkbook.ActiveSheet;

               useOpenExcel = true;
            }
            else
            {
               // new Excel Application
               excelApp = new Application();

               // Open workbook from file
               excelWorkbook = excelApp.Workbooks.Open(fileName.Value);

               // Get active worksheet
               excelworksheet = excelWorkbook.ActiveSheet;
            }

            if (null != excelworksheet)
            {
               //find the used range in worksheet
               Range excelRange = excelworksheet.UsedRange;

               //get an object array of all of the cells in the worksheet (their values)
               object[,] valueArray = (object[,])excelRange.Value[XlRangeValueDataType.xlRangeValueDefault];

               if (null == valueArray)
               {
                  const string message = "Excel worksheet is empty!";
                  theUI.NXMessageBox.Show("Excel file failure", NXMessageBox.DialogType.Warning, message);
                  return;
               }

               // Get root item number (and nomenclature) from excel file
               string rootItemNumber, rootNomenclature;
               ExcelStatics.GetRootValuesFromWorksheet(excelworksheet, valueArray, out rootItemNumber, out rootNomenclature);

               if ( String.IsNullOrEmpty(rootItemNumber) )
               {
                  return;
               }

               // Get root item number (and nomenclature) from current nx session root part
               string nxRootItemNumber, nxRootNomenclature;
               GetRootValuesFromNx(out nxRootItemNumber, out nxRootNomenclature);

               if (nxRootItemNumber.ToUpper() != rootItemNumber.ToUpper())
               {
                  // fehlermeldung: root hat nicht die gleiche item number ! 
                  const string message = "Root number is not the same as the Excel file root number.";
                  theUI.NXMessageBox.Show("Excel file failure", NXMessageBox.DialogType.Warning, message);
                  return;
               }

               // Build intern tree structure
               rootTreeNode = new TreeStructure(null, rootItemNumber, rootNomenclature, 8, -1, excelworksheet.UsedRange.Rows.Count);
               ExcelStatics.CreateSubTree(valueArray, excelworksheet, rootTreeNode);

            }
         }
         catch (Exception)
         {
            const string message = "Problem with excel !";
            theUI.NXMessageBox.Show("Excel file failure", NXMessageBox.DialogType.Warning, message);
         }
         finally
         {
            if (!useOpenExcel)
            {
               // Check if Excel Workbook is defined
               if (excelWorkbook != null)
               {
                  //close the Excel Workbook
                  excelWorkbook.Close(false, Type.Missing, Type.Missing);
               }

               // Check if Excel Application is defined
               if (excelApp != null)
               {
                  // Quit the Excel Application
                  excelApp.Quit();
               }
            }
            else
            {
               Marshal.ReleaseComObject(excelApp);
               Marshal.FinalReleaseComObject(excelApp);
            }
         }

      }


      /// <summary>Read excel file</summary>
      /// <param name="rootTreeNode"></param>
      /// <returns></returns>
      private void ReadExcelFile(out TreeStructure rootTreeNode)
      {
         rootTreeNode = null;

         // check input values
         if (String.IsNullOrEmpty(fileName.Value))
         {
            string outputText = "File string is empty";
            theUI.NXMessageBox.Show("Note", NXMessageBox.DialogType.Information, outputText);
            return;
         }
         if (!File.Exists(fileName.Value))
         {
            string outputText = "Selected file does not exists";
            theUI.NXMessageBox.Show("Note", NXMessageBox.DialogType.Information, outputText);
            return;
         }

         // Initialize Excel Application and Workbook
         Application excelApp = null;
         Workbook excelWorkbook = null;

         try
         {
            // Start Excel Application, add Workbook and get the Worksheets
            excelApp = new Application();
            excelWorkbook = excelApp.Workbooks.Open(fileName.Value);

            // Get worksheet (--> structure must be at the active worksheet ! )
            Worksheet worksheet = excelWorkbook.ActiveSheet;

            if (null != worksheet)
            {
               //find the used range in worksheet
               Range excelRange = worksheet.UsedRange;

               //get an object array of all of the cells in the worksheet (their values)
               object[,] valueArray = (object[,]) excelRange.Value[XlRangeValueDataType.xlRangeValueDefault];

               if (null == valueArray)
               {
                  string outputText = "Excel worksheet is empty!";
                  theUI.NXMessageBox.Show("Note", NXMessageBox.DialogType.Information, outputText);
                  return;
               }

               string rootItemNumber, rootNomenclature;
               ExcelStatics.GetRootValuesFromWorksheet(worksheet, valueArray, out rootItemNumber, out rootNomenclature);

               if (!String.IsNullOrEmpty(rootItemNumber) | !String.IsNullOrEmpty(rootItemNumber))
               {
                  // Build intern tree structure
                  rootTreeNode = new TreeStructure(null, rootItemNumber, rootNomenclature, 8, -1,
                                                   worksheet.UsedRange.Rows.Count);
                  ExcelStatics.CreateSubTree(valueArray, worksheet, rootTreeNode);

               }
            }
         }
         catch (Exception)
         {
            // show NX Message Box
            string outputText = "Problems in reading excel file.";
            theUI.NXMessageBox.Show("Note", NXMessageBox.DialogType.Information, outputText);
         }
         finally
         {
            // Check if Excel Workbook is defined
            if (excelWorkbook != null)
            {
               // close the Excel Workbook
               excelWorkbook.Close(false, Type.Missing, Type.Missing);
            }

            // Check if Excel Application is defined
            if (excelApp != null)
            {
               // Quit the Excel Application
               excelApp.Quit();
            }
         }
      }


      /// <summary>Export to excel. </summary>
      private void ExportToExcel()
      {
         // To remember first entry must be the root item 
         List<string> nomenclatures;
         List<string> partsToExport = GetPartsToExport(out nomenclatures);

         // Initialize Excel Application and Workbook
         Application excelApp = null;
         Workbook excelWorkbook = null;
         object missingValue = Missing.Value;

         try
         {
            // Start Excel Application, add Workbook and get the Worksheets
            excelApp = new Application();
            excelWorkbook = excelApp.Workbooks.Add(missingValue);

            // Get active sheet to fill 
            Worksheet worksheet = excelWorkbook.ActiveSheet;

            // Build skeletal structure of excel worksheet
            ExcelStatics.BuildExcelSkeletalStructure(worksheet);

            // Fill excel file
            ExcelStatics.FillExcelTable(worksheet, partsToExport, nomenclatures);

            //Make sure Excel is visible and give the user control of Microsoft Excel's lifetime.
            excelApp.Visible = true;
            excelApp.UserControl = true;

         }
         catch (Exception ex)
         {
            // show NX Message Box
            theUI.NXMessageBox.Show("Export fails", NXMessageBox.DialogType.Error, ex.ToString());

            // Check if Excel Workbook is defined
            if (excelWorkbook != null)
            {
               // close the Excel Workbook
               excelWorkbook.Close(false, Type.Missing, Type.Missing);
            }

            // Check if Excel Application is defined
            if (excelApp != null)
            {
               // Quit the Excel Application
               excelApp.Quit();
            }
         }
      }


      /// <summary>Fill validation result to UI</summary>
      /// <param name="resultList">the results</param>
      private void FillValidationResult(Dictionary<string, string> resultList)
      {
         if (null == resultList)
         {
            return;
         }

         string[] listKeys = new string[resultList.Count];
         resultList.Keys.CopyTo(listKeys, 0);

         if (listKeys.Length > 0)
         {
            string[] results = new string[resultList.Count];

            for (int index = 0; index < listKeys.Length; index++)
            {
               var key = listKeys[index];
               string row = key + ": ";
               string value;
               if (resultList.TryGetValue(key, out value))
               {
                  row = row + value;
               }

               results[index] = row;
            }
            multilineValResult.SetValue(results);
         }
         else
         {
            string[] results = new string[1];
            results[0] = "Validation OK";
            multilineValResult.SetValue(results);
         }
      }


      /// <summary>Get all asm parts</summary>
      /// <returns>List of paths </returns>
      private List<string> GetAllAsmParts()
      {
         List<string> listOfParts = new List<string>();

         // Get current ASM root part 
         Part asmPart = theSession.Parts.Work;

         // Get directory of part 
         string directoryOfCurrentPart = Path.GetDirectoryName(asmPart.FullPath);

         if (!String.IsNullOrEmpty(directoryOfCurrentPart))
         {     
            // Get all .prt files 
            string[] existingFilesInFolder = Directory.GetFiles(directoryOfCurrentPart, "*.prt", SearchOption.TopDirectoryOnly);
            foreach (var filePath in existingFilesInFolder)
            {
               listOfParts.Add(Path.GetFullPath(filePath));
            }
         }

         return listOfParts;
      }


      /// <summary>Get all ASM part</summary>
      /// <returns>List of all ASM parts.</returns>
      private List<Part> GetAllAsmPart(out List<string> multiplePartNames)
      {
         // initialize ASM part list
         List<Part> asmParts = new List<Part>();

         List<string> multipleParts = new List<string>();

         // Get all parts      
         Part workPart = theSession.Parts.Work;
         Component component = workPart.ComponentAssembly.RootComponent;

         // for each child component
         List<Component> components = ComponentStatics.GetChildren(component, true);
         foreach (Component currentComponent in components)
         {
            // get the part
            Part part = currentComponent.Prototype as Part;

            if (null != part)
            {
               if (PartStatics.GetStringAttribute(part, RniPartType) == "ASM")
               {
                  if (asmParts.Contains(part))
                  {
                     multipleParts.Add(part.Leaf);
                  }
                  else
                  {
                     // add part to list
                     asmParts.Add(part);
                  }
               }
            }
         }
         multiplePartNames = new List<string>(multipleParts.Distinct());

         return asmParts;
      }


      /// <summary>Check existing files.</summary>
      /// <param name="rootNode">The tree structure root node.</param>
      /// <returns>bool if min one file exists or not.</returns>
      private bool CheckExistingFiles(TreeStructure rootNode)
      {
         bool exist = false;

         if ( null == rootNode )
         {
            return false;
         }

         string targetPath = structurePathBrowser.Path;

         if (!String.IsNullOrEmpty(targetPath))
         {
            List<string> allExistingFileNames = new List<string>();

            // Get all .prt files 
            string[] existingFilesInFolder = Directory.GetFiles(targetPath, "*.prt", SearchOption.TopDirectoryOnly);
            foreach (var filePath in existingFilesInFolder)
            {
               allExistingFileNames.Add(Path.GetFileNameWithoutExtension(filePath));   
            }

            List<string> allValues = new List<string>();
            GetAllValuesFromTreeStructure(rootNode, ref allValues );

            foreach (var value in allValues)
            {
               if(allExistingFileNames.Contains(value))
               {
                  exist = true;
                  break;
               }
            }
         }
         return exist;
      }


      /// <summary>Get all values from tree structure</summary>
      /// <param name="treeStructure">Tree structure</param>
      /// <param name="values">List of values</param>
      private void GetAllValuesFromTreeStructure(TreeStructure treeStructure, ref List<string> values )
      {
         // add value and value+_1 to list
         values.Add(treeStructure.GetValue());
         values.Add(treeStructure.GetValue() + "_1");

         List<TreeStructure> children = treeStructure.GetChidren();
         foreach (var child in children)
         {
            // rekursive call of method
            GetAllValuesFromTreeStructure(child, ref values);
         }
      }
    

      /// <summary>Get parts to export</summary>
      /// <returns>The part infos in a dictionary.</returns>
      private List<string> GetPartsToExport(out List<string> nomenclatures )
      {
         List<string> partsToExport = new List<string>();
         nomenclatures = new List<string>();

         TaggedObject[] selectedObjects = componentSelection.GetSelectedObjects();

         if (selectedObjects.Length > 0 )
         {
            Component selObject = selectedObjects[0] as Component;

            if (null != selObject)
            {
               // Get Root component / part
               Component rootComponent = ComponentStatics.GetRootComponent(selObject);
               Part partOfRootComponent = rootComponent.Prototype as Part;

               if (null != partOfRootComponent)
               {
                  // Get RNI type attribute from work part
                  string rniType = PartStatics.GetStringAttribute(partOfRootComponent, RniPartType);

                  // The RNI type must be "ASM"
                  if (rniType == "ASM")
                  {
                     // Get value of part 
                     string partName = partOfRootComponent.Leaf;
                     partsToExport.Add(partName);

                     // Get Nomenclature of part
                     string nomenclature = PartStatics.GetStringAttribute(partOfRootComponent, "NOMENCLATURE");
                     nomenclatures.Add(nomenclature);
                  }
               }
            }

            foreach (var selectedObject in selectedObjects)
            {
               Component selectedComponent = selectedObject as Component;
               if (null != selectedComponent)
               {
                  Part partOfComponent = selectedComponent.Prototype as Part;
                  if (null != partOfComponent)
                  {
                     // Get RNI type attribute from work part
                     string rniType = PartStatics.GetStringAttribute(partOfComponent, RniPartType);

                     // The RNI type must be "ASM"
                     if (rniType == "ASM")
                     {
                        // Get value of part 
                        string partName = partOfComponent.Leaf;
                        if ( partName.Length == 17 )
                        {
                           string value = partName.Substring(13, 4);

                           partsToExport.Add(value);

                           // Get Nomenclature of part
                           string nomenclature = PartStatics.GetStringAttribute(partOfComponent, "NOMENCLATURE");
                           nomenclatures.Add(nomenclature);
                        }
                     }
                  }
               }
            }
         }
         return partsToExport;
      }


      /// <summary>Get root values from nx</summary>
      /// <param name="nxRootItemNumber">out: The root item number.</param>
      /// <param name="nxRootNomenclature">out: The root nomenclature.</param>
      private void GetRootValuesFromNx(out string nxRootItemNumber, out string nxRootNomenclature)
      {
         // Initialize out parameter
         nxRootItemNumber = "";
         nxRootNomenclature = "";

         Part currentWorkPart = theSession.Parts.Work;

         // Get rootComponent of currenct work part (asm start part)
         Component rootComponent = currentWorkPart.ComponentAssembly.RootComponent;

         if (null != rootComponent)
         {
            Part partOfRootComponent = rootComponent.Prototype as Part;

            if (null != partOfRootComponent)
            {
               // Get RNI type attribute from work part
               string rniType = PartStatics.GetStringAttribute(partOfRootComponent, RniPartType);

               // The RNI type must be "ASM"
               if (rniType == "ASM")
               {
                  // Get value of part 
                  nxRootItemNumber = partOfRootComponent.Leaf;

                  // Get Nomenclature of part
                  nxRootNomenclature = PartStatics.GetStringAttribute(partOfRootComponent, "NOMENCLATURE"); 
               }
            }
         }
      }

    
      /// <summary>Update visibilities of UI Blocks to the user selection.</summary>
      private void UpdateUiBlocksVisibilities()
      {
         string structureCreationType = RadioBoxStructureOption.ValueAsString;
 
         toggleUseExcelSession.Show = false;
         labelSaveBeforeValidation.Show = false;
         structurePathBrowser.Show = false;
         groupValidation.Show = false;

         if ("Load / create new structure" == structureCreationType)
         {
            // Group visibility 
            groupFile.Show = true;
            groupImportOptions.Show = true;
            groupExport.Show = false;
            groupEnhanceExStructure.Show = false; 

            buttonOpenFile.Enable = true;
            structurePathBrowser.Show = true;

         }

         else if ("Export nodes to enhance existing structure" == structureCreationType)
         {
            // Group visibility 
            groupFile.Show = false;
            groupImportOptions.Show = false;
            groupExport.Show = true;
            groupEnhanceExStructure.Show = true;
         }

         else
         {
            // Group visibility 
            groupFile.Show = true;
            groupImportOptions.Show = true;
            groupExport.Show = false;
            groupEnhanceExStructure.Show = false;
            groupValidation.Show = true;

            toggleUseExcelSession.Show = true;
            labelSaveBeforeValidation.Show = true;

            if (toggleUseExcelSession.Value)
            {
               buttonOpenFile.Enable = false;
               labelSaveBeforeValidation.Enable = true;
            }
            else
            {
               buttonOpenFile.Enable = true;
               labelSaveBeforeValidation.Enable = false;
            }
         }
      }


      /// <summary>Create open file dialog.</summary>
      private void CreateOpenFileDialog()
      {
         // Create SaveFileDialog
         const string dialogFilter = "Excel Files(.xlsx)|*.xlsx| Excel Files(.xls)|*.xls| Excel Files(*.xlsm)|*.xlsm";
         const string dialogTitle = "Browse Excel file";
         var openFileDialog = new OpenFileDialog
                                 {
                                    CheckPathExists = true,
                                    DefaultExt = "xlsx",
                                    Filter = dialogFilter,
                                    Title = dialogTitle
                                 };

         // Check if openFileDialog is closed with OK Button
         if (openFileDialog.ShowDialog() == DialogResult.OK)
         {
            fileName.Value = openFileDialog.FileName;
         }
      }


      /// <summary>Enhance NX assembly structure.</summary>
      /// <param name="rootNode">The tree structure node</param>
      /// <param name="startAsmPart">selected start asm part </param>
      /// <param name="startModelPart"> </param>
      /// <param name="onlyValidation"> </param>
      private Dictionary<string, string> EnhanceAssemblyStructure(TreeStructure rootNode, Part startAsmPart, Part startModelPart, bool onlyValidation)
      {
         Dictionary<string, string > resultList = new Dictionary<string, string>();

         // Initialize asm part list
         List<string> listOfCreatedAsmParts = GetAllAsmParts();

         // Get (Add position number to nomenclature) from ui
         bool addPosNb = toggleAddPosNb.Value;      

         // Get full path of startpart
         string fullPathOfDaimlerStartPart = "";
         string fullPathOfDaimlerModelStartPart = "";
         if ( null != startAsmPart )
         {
            fullPathOfDaimlerStartPart = Path.GetFullPath(startAsmPart.FullPath); 
         }
         if (null != startModelPart)
         {
            fullPathOfDaimlerModelStartPart = Path.GetFullPath(startModelPart.FullPath);
         }          

         // Get all ASM parts of current work part (root part)
         List<string> multiplePartnames;
         List<Part> asmParts = GetAllAsmPart(out multiplePartnames);

         // Get all root children
         List<TreeStructure> children = rootNode.GetChidren();
         foreach (var child in children)
         {
            string itemName = child.GetValue();
            string nomenclature = child.GetNomenclature();

            // found part with this itemName and nomenclature in current session part 
            Part foundedAsmPart = null;
            foreach (var part in asmParts)
            {
               string name = part.Leaf;
               string partNomenclature = PartStatics.GetStringAttribute(part, "NOMENCLATURE");

               if ( name == itemName && partNomenclature == nomenclature)
               {
                  foundedAsmPart = part;
                  if (multiplePartnames.Contains(name))
                  {
                     if (!resultList.ContainsKey(itemName) && onlyValidation)
                     {
                        resultList.Add(itemName, " is a multiple part. (only information)");
                     }
                  }
                  break;
               }

               if ( name == itemName && partNomenclature!=nomenclature)
               {
                  if (!resultList.ContainsKey(itemName))
                  {
                     resultList.Add(itemName, " Import part has different nomenclature than NX part.");
                     child.SetCreation(false);
                  }
                  break;
               }
            }

            List<TreeStructure> nodes = child.GetChidren();

            // check parts to add
            foreach (var node in nodes)
            {
               CheckPartsToAdd(node, asmParts, onlyValidation, ref resultList);
            }

            if (null == foundedAsmPart)
            {
               if (!resultList.ContainsKey(itemName))
               {
                  resultList.Add(itemName, "not found in NX component tree");
               }
            }
            else
            {
               // Get all root children
               List<TreeStructure> nodeChildren = child.GetChidren();

               if (!onlyValidation)
               {
                  // call for each child the recursive method add part copies
                  foreach (var node in nodeChildren)
                  {
                     if ( node.ShouldCreated())
                     {
                        // add part copies to NX assembly structure
                        AddDaimlerPartCopies(node, foundedAsmPart, fullPathOfDaimlerStartPart, addPosNb, ref listOfCreatedAsmParts);
                     }  
                  }
               }
            }
         }

         // Delete the temp saved start parts
         if (!String.IsNullOrEmpty(fullPathOfDaimlerModelStartPart))
         {
            File.Delete(fullPathOfDaimlerModelStartPart);
         }
         if (!String.IsNullOrEmpty(fullPathOfDaimlerStartPart))
         {
            File.Delete(fullPathOfDaimlerStartPart);
         }       

         return resultList;
      }


      /// <summary>Check parts to add</summary>
      /// <param name="node"></param>
      /// <param name="asmParts"></param>
      /// <param name="onlyValidation"> </param>
      /// <param name="resultList"></param>
      private void CheckPartsToAdd(TreeStructure node, List<Part> asmParts, bool onlyValidation, ref Dictionary<string, string> resultList)
      {
         string itemName = node.GetValue();
         string nomenclature = node.GetNomenclature();

         node.SetCreation(true);
         bool found = false;

         // found part with this itemName and nomenclature in current session part 
         foreach (var part in asmParts)
         {
            string name = part.Leaf;
            string partNomenclature = PartStatics.GetStringAttribute(part, "NOMENCLATURE");

            if (name == itemName && partNomenclature == nomenclature)
            {
               if (!resultList.ContainsKey(itemName) && onlyValidation)
               {
                  resultList.Add(itemName, " is a multiple part. (only information)");
               }
               found = true;
               break;
            }

            if (name == itemName && partNomenclature != nomenclature)
            {
               if (!resultList.ContainsKey(itemName))
               {
                  resultList.Add(itemName, " Import part has different nomenclature than NX part. (No import!)");
                  node.SetCreation(false);
                  found = true;
                  break;
               }
            }
         }
         if ( !found && (nomenclature=="NOMENCLATURE"))
         {
            if (!resultList.ContainsKey(itemName) && onlyValidation)
            {
               resultList.Add(itemName, " has empty nomenclature. (only information)");
            }         
         }

         // Get all root children
         List<TreeStructure> nodeChildren = node.GetChidren();
         foreach (var child in nodeChildren)
         {
            CheckPartsToAdd(child, asmParts, onlyValidation, ref resultList);
         }

      }


      /// <summary>Create NX assembly structure</summary>
      /// <param name="rootNode">The tree structure root node.</param>
      private bool CreateAssemblyStructure(TreeStructure rootNode)
      {
         bool creation = true;

         // Initialize asm part list
         List<string> listOfCreatedAsmParts = new List<string>();

         // Get value (Add position number to nomenclature) from ui
         bool addPosNb = toggleAddPosNb.Value;

         // Get start parts (The ASM part must be the work part and as child the PRT part
         Part asmPart, prtPart;
         GetDaimlerStartParts(out asmPart, out prtPart);

         // If start parts correct 
         if ((null != asmPart) && (null != prtPart))
         {
            // Get the full path ot the start part to make the copies from
            string fullPathOfDaimlerStartPart = Path.GetFullPath(asmPart.FullPath);
            string fullPathOfDaimlerModelStartPart = Path.GetFullPath(prtPart.FullPath);

            bool adPosNr;
            if (addPosNb)
            {
               adPosNr = !rootNode.HasChildren();
            }
            else
            {
               adPosNr = false;
            }
            // Save the parts with root name
            bool correctSave = SaveDaimlerPartsAs(asmPart, prtPart, rootNode.GetValue(), rootNode.GetNomenclature(), adPosNr, ref listOfCreatedAsmParts);

            if ( correctSave )
            {
               Component rootComponent = asmPart.ComponentAssembly.RootComponent;
               if (null != rootComponent)
               {
                  Component[] rootCompChildren = rootComponent.GetChildren();
                  Component rootCompChild = rootCompChildren[0];

                  // Add fix constraint to component
                  ComponentStatics.AddFixConstraint(asmPart, rootCompChild, theSession);
               }

               // Get all root children
               List<TreeStructure> children = rootNode.GetChidren();

               // call for each child the recursive method add part copies
               foreach (var child in children)
               {
                  // add part copies to NX assembly structure
                  AddDaimlerPartCopies(child, asmPart, fullPathOfDaimlerStartPart, addPosNb, ref listOfCreatedAsmParts);
               }
            }
            else
            {
               creation = false;          
            }

            // Delete the temp saved start parts
            if (!String.IsNullOrEmpty(fullPathOfDaimlerModelStartPart))
            {
               File.Delete(fullPathOfDaimlerModelStartPart);
            }
            if (!String.IsNullOrEmpty(fullPathOfDaimlerStartPart))
            {
               File.Delete(fullPathOfDaimlerStartPart);
            }
         }
         return creation;
      }


      /// <summary>Save ASM and PRT with new name</summary>
      /// <param name="asmPart">The ASM part to save.</param>
      /// <param name="prtPart">The PRT part to save.</param>
      /// <param name="newName">The new name.</param>
      /// <param name="nomenclature">The nomenclature to set as attribute to parts.</param>
      /// <param name="addPosNb"> </param>
      /// <param name="asmParts">List of all created ASM parts. If save as works, then part added to list.</param>
      private bool SaveDaimlerPartsAs(Part asmPart, Part prtPart, string newName, string nomenclature, bool addPosNb, ref List<string> asmParts )
      {
         // Check input asm part
         if ( null == asmPart)
         {
            return false;
         }

         // Check input prt part
         if ( null == prtPart)
         {
            return false;
         }

         // Get directory of asm Part 
         string directoryOfCurrentPart = Path.GetDirectoryName(asmPart.FullPath);

         // Create new directory path string
         string path = directoryOfCurrentPart + "\\" + newName.ToUpper();

         try
         {
            // Save prt Part as
            prtPart.SaveAs(path + "_1");

            // Save asm Part as 
            asmPart.SaveAs(path);

            // remember all path strings in asmParts list (all new created asm Parts)
            if ( !asmParts.Contains(path))
            {
               asmParts.Add(path + ".prt");
            } 
         }
         catch (Exception)
         {
            string message = "File: " + newName.ToUpper() + ".prt" + " already exists in location: \n" + directoryOfCurrentPart;
            theUI.NXMessageBox.Show("Overwrite exception", NXMessageBox.DialogType.Warning, message);
            return false;
         }

         string partNomenclature = nomenclature;
         if (addPosNb)
         {
            if(newName.Length == 17)
            {
               string posNumber = newName.Substring(13, 4);
               partNomenclature = partNomenclature + "_POS_" + posNumber;
            }
         }

         // Set nomenclature as string attribute to parts:
         PartStatics.SetStringAttribute(asmPart, "NOMENCLATURE", partNomenclature.ToUpper());
         PartStatics.SetStringAttribute(prtPart, "NOMENCLATURE", partNomenclature.ToUpper());

         return true;
      }


      /// <summary>Add part copies.</summary>
      /// <param name="treeNode">The current tree node.</param>
      /// <param name="parentPart">The parent part</param>
      /// <param name="fullPathOfDaimlerStartPart">Path of the start ASM part.</param>
      /// <param name="addPosNb"> </param>
      /// <param name="asmParts">ref list of all created ASM parts.</param>
      private void AddDaimlerPartCopies( TreeStructure treeNode, Part parentPart, string fullPathOfDaimlerStartPart, bool addPosNb, ref List<string> asmParts)
      {
         // Check parent part
         if ( null == parentPart)
         {
            return;
         }

         // Check tree node 
         if (null == treeNode)
         {
            return;
         }

         // Get directory of parent part
         string directory = Path.GetDirectoryName(parentPart.FullPath);

         // Create new part name path
         string newPartNamePath = directory + "\\" + treeNode.GetValue() + ".prt";

         Point3d basePoint1 = new Point3d(0.0, 0.0, 0.0);
         Matrix3x3 orientation1 = new Matrix3x3 { Xx = 1.0, Xy = 0.0, Xz = 0.0, Yx = 0.0, Yy = 1.0, Yz = 0.0, Zx = 0.0, Zy = 0.0, Zz = 1.0 };

         // if part still created ("Mehrfachverbauung")
         if (asmParts.Contains(newPartNamePath))
         {
            // Add component to parent part 
            Component newComponent;

            try
            {
               PartLoadStatus partLoadStatus;
               newComponent = parentPart.ComponentAssembly.AddComponent(newPartNamePath, "FINAL_PART", treeNode.GetValue(), basePoint1, orientation1, -1, out partLoadStatus, true);
            }
            catch (Exception)
            {
               const string message = "Please save first start parts and start import again!";
               theUI.NXMessageBox.Show("Start part problem", NXMessageBox.DialogType.Information, message);
               return;
            }
            

            if (null != newComponent)
            {
               // Get ASM part of new component 
               Part asmPart = newComponent.Prototype as Part;

               if (null != asmPart)
               {
                  List<TreeStructure> treeChilds = treeNode.GetChidren();
                  foreach (var child in treeChilds)
                  {
                     AddDaimlerPartCopies(child, asmPart, fullPathOfDaimlerStartPart, addPosNb, ref asmParts);
                  }
               }
            } 
         }
         else
         {
            // Add component to parent part 
            Component newComponent;

            try
            {
               PartLoadStatus partLoadStatus;
               newComponent = parentPart.ComponentAssembly.AddComponent(fullPathOfDaimlerStartPart, "FINAL_PART", "ASSEMBLY_SNR1", basePoint1, orientation1, -1, out partLoadStatus, true);
            }
            catch (Exception)
            {
               const string message = "Please save first start parts and start import again!";
               theUI.NXMessageBox.Show("Start part problem", NXMessageBox.DialogType.Information, message);
               return;
            }

            if ( null != newComponent )
            {
               // Get ASM part of new component 
               Part asmPart = newComponent.Prototype as Part;

               if (null != asmPart)
               {
                  Component[] children = newComponent.GetChildren();
                  Component child = children[0];

                  // Get PRT part from child component (prototype)
                  Part prtPart = child.Prototype as Part;
               
                  bool adPosNr;
                  if ( addPosNb )
                  {
                     adPosNr = !treeNode.HasChildren();
                  }
                  else
                  {
                     adPosNr = false;
                  }
                  // Save parts with new name
                  SaveDaimlerPartsAs(asmPart, prtPart, treeNode.GetValue(), treeNode.GetNomenclature(), adPosNr, ref asmParts);

                  // Add fix constraint to component
                  ComponentStatics.AddFixConstraint(asmPart, child, theSession);
          
                  List<TreeStructure> treeChilds = treeNode.GetChidren();
                  foreach (var node in treeChilds)
                  {
                     // Recursive call
                     AddDaimlerPartCopies(node, asmPart, fullPathOfDaimlerStartPart, addPosNb, ref asmParts);
                  }   
               }
            } 
         }
      }


      /// <summary>Get Start parts from current session.</summary>
      /// <param name="daimlerAsmPart">out ASM start part.</param>
      /// <param name="daimlerPrtPart">out PRT geometry start part.</param>
      private void GetDaimlerStartParts(out Part daimlerAsmPart, out Part daimlerPrtPart)
      {
         // Initialize parts
         daimlerAsmPart = null;
         daimlerPrtPart = null;

         // Get current work part in session
         Part currentWorkPart = theSession.Parts.Work;

         if ( null == currentWorkPart)
            return;

         // Get rootComponent of currenct work part (asm start part)
         Component rootComponent = currentWorkPart.ComponentAssembly.RootComponent;

         if (null != rootComponent)
         {
            // Get children of root component 
            Component[] children = rootComponent.GetChildren();

            if ( children.Length > 0)
            {
               // Get child from array
               Component child = children[0];

               // Get PRT part from child (prototype)
               Part daimlerModelPart = child.Prototype as Part;

               if (null != daimlerModelPart)
               {
                  // Set output parts
                  daimlerAsmPart = currentWorkPart;
                  daimlerPrtPart = daimlerModelPart;
               }    
            }
         }
      }


      #endregion Private Methods



   }

}
