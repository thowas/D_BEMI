//==============================================================================
//
//        Filename: ExcelStatics.cs
//
//        Created by: CENIT AG (Jan Assmann)
//              Version: NX 8.5.2.3 MP1
//              Date: 03-12-2013  (Format: mm-dd-yyyy)
//              Time: 08:30 (Format: hh-mm)
//
//==============================================================================

using System;
using System.Collections.Generic;
using System.Reflection;
using Microsoft.Office.Interop.Excel;
using NXOpen;


namespace  Daimler.NX.BemiStructure
{
   /// <summary>Excel static class</summary>
   public static class ExcelStatics
   {

      /// <summary>Fill excel Table </summary>
      /// <param name="worksheet">The Worksheet to fill</param>
      /// <param name="partsToExport">Parts (as strings) to export</param>
      /// <param name="nomenclatures">Nomenclatures (as strings) to export.</param>
      public static void FillExcelTable(Worksheet worksheet, List<string> partsToExport, List<string> nomenclatures)
      {
         if (null == worksheet)
            return;

         // Used range
         Range range = worksheet.UsedRange;

         // Root item number
         worksheet.Cells[3, 1] = partsToExport[0];
         worksheet.Cells[3, 2] = nomenclatures[0];

         Range cell1 = range.Cells[3, 1];
         cell1.BorderAround();
         cell1.HorizontalAlignment = Constants.xlCenter;

         Range cell2 = range.Cells[3, 2];
         cell2.BorderAround();
         cell2.HorizontalAlignment = Constants.xlCenter;

         int emptyRows = 0;
         int partSize = partsToExport.Count;
         for (int i = 1; i < partSize; i++)
         {
            string partName = partsToExport[i];
            string nomenclature = nomenclatures[i];

            int row = 8 + i + emptyRows;

            worksheet.Cells[row, 1].NumberFormat = "@";
            worksheet.Cells[row, 1] = partName;
            worksheet.Cells[row, 2] = nomenclature;

            Range range1 = range.Cells[row, 1];
            Range range2 = range.Cells[row, 2];

            range1.HorizontalAlignment = Constants.xlCenter;
            range2.HorizontalAlignment = Constants.xlCenter;

            emptyRows = emptyRows + 5;
         }

         // borders:
         int rowLength = 9 + (partSize - 1) * 6;
         string tillRange = "H" + rowLength;
         Range rg = worksheet.Range["A8", tillRange];

         rg.Borders[XlBordersIndex.xlEdgeLeft].LineStyle = XlLineStyle.xlContinuous;
         rg.Borders[XlBordersIndex.xlEdgeRight].LineStyle = XlLineStyle.xlContinuous;
         rg.Borders[XlBordersIndex.xlEdgeTop].LineStyle = XlLineStyle.xlContinuous;
         rg.Borders[XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlContinuous;
         rg.Borders.ColorIndex = 0;
      }


      /// <summary>Build skeletal structure inside excel worksheet</summary>
      /// <param name="worksheet">Excel worksheet</param>
      public static void BuildExcelSkeletalStructure(Worksheet worksheet)
      {
         // check input
         if (null == worksheet)
            return;

         object missingValue = Missing.Value;

         ((Range)worksheet.Columns[1, Type.Missing]).ColumnWidth = 17;
         ((Range)worksheet.Columns[2, Type.Missing]).ColumnWidth = 17;
         ((Range)worksheet.Columns[3, Type.Missing]).ColumnWidth = 11;
         ((Range)worksheet.Columns[4, Type.Missing]).ColumnWidth = 17;
         ((Range)worksheet.Columns[5, Type.Missing]).ColumnWidth = 11;
         ((Range)worksheet.Columns[6, Type.Missing]).ColumnWidth = 17;
         ((Range)worksheet.Columns[7, Type.Missing]).ColumnWidth = 11;
         ((Range)worksheet.Columns[8, Type.Missing]).ColumnWidth = 17;

         ((Range)worksheet.Columns[1, Type.Missing]).NumberFormat = "@";
         ((Range)worksheet.Columns[3, Type.Missing]).NumberFormat = "@";
         ((Range)worksheet.Columns[5, Type.Missing]).NumberFormat = "@";
         ((Range)worksheet.Columns[7, Type.Missing]).NumberFormat = "@";

         ((Range)worksheet.Columns[1, Type.Missing]).HorizontalAlignment = Constants.xlCenter;
         ((Range)worksheet.Columns[2, Type.Missing]).HorizontalAlignment = Constants.xlCenter;
         ((Range)worksheet.Columns[3, Type.Missing]).HorizontalAlignment = Constants.xlCenter;
         ((Range)worksheet.Columns[4, Type.Missing]).HorizontalAlignment = Constants.xlCenter;
         ((Range)worksheet.Columns[5, Type.Missing]).HorizontalAlignment = Constants.xlCenter;
         ((Range)worksheet.Columns[6, Type.Missing]).HorizontalAlignment = Constants.xlCenter;
         ((Range)worksheet.Columns[7, Type.Missing]).HorizontalAlignment = Constants.xlCenter;
         ((Range)worksheet.Columns[8, Type.Missing]).HorizontalAlignment = Constants.xlCenter;
     
         // blue columns
         ((Range)worksheet.Columns[2, Type.Missing]).Interior.Pattern = Constants.xlSolid;
         ((Range)worksheet.Columns[2, Type.Missing]).Interior.PatternColorIndex = Constants.xlAutomatic;
         ((Range)worksheet.Columns[2, Type.Missing]).Interior.ThemeColor = 4;
         ((Range)worksheet.Columns[2, Type.Missing]).Interior.TintAndShade = 0.799981688894314;
         ((Range)worksheet.Columns[2, Type.Missing]).Interior.PatternTintAndShade = 0;

         ((Range)worksheet.Columns[4, Type.Missing]).Interior.Pattern = Constants.xlSolid;
         ((Range)worksheet.Columns[4, Type.Missing]).Interior.PatternColorIndex = Constants.xlAutomatic;
         ((Range)worksheet.Columns[4, Type.Missing]).Interior.ThemeColor = 4;
         ((Range)worksheet.Columns[4, Type.Missing]).Interior.TintAndShade = 0.799981688894314;
         ((Range)worksheet.Columns[4, Type.Missing]).Interior.PatternTintAndShade = 0;

         ((Range)worksheet.Columns[6, Type.Missing]).Interior.Pattern = Constants.xlSolid;
         ((Range)worksheet.Columns[6, Type.Missing]).Interior.PatternColorIndex = Constants.xlAutomatic;
         ((Range)worksheet.Columns[6, Type.Missing]).Interior.ThemeColor = 4;
         ((Range)worksheet.Columns[6, Type.Missing]).Interior.TintAndShade = 0.799981688894314;
         ((Range)worksheet.Columns[6, Type.Missing]).Interior.PatternTintAndShade = 0;

         ((Range)worksheet.Columns[8, Type.Missing]).Interior.Pattern = Constants.xlSolid;
         ((Range)worksheet.Columns[8, Type.Missing]).Interior.PatternColorIndex = Constants.xlAutomatic;
         ((Range)worksheet.Columns[8, Type.Missing]).Interior.ThemeColor = 4;
         ((Range)worksheet.Columns[8, Type.Missing]).Interior.TintAndShade = 0.799981688894314;
         ((Range)worksheet.Columns[8, Type.Missing]).Interior.PatternTintAndShade = 0;

         worksheet.Cells[1, 1] = "Rootsachnummer" + "\n" + "Root Item Number";
         worksheet.Cells[1, 1].Characters(1, 14).Font.Bold = true;

         // Zellen verbinden
         Range rng1 = worksheet.Range["A1", "A2"];
         rng1.Merge(missingValue);
         rng1.BorderAround();
         rng1.Interior.ColorIndex = 6;
         rng1.HorizontalAlignment = Constants.xlCenter;
         rng1.VerticalAlignment = Constants.xlCenter;

         worksheet.Cells[1, 2] = "Benennung" + "\n" + "Nomenclature";
         worksheet.Cells[1, 2].Characters(1, 9).Font.Bold = true;

         Range rng2 = worksheet.Range["B1", "B2"];
         // Zellen verbinden
         rng2.Merge(missingValue);
         rng2.BorderAround();
         rng2.HorizontalAlignment = Constants.xlCenter;
         rng2.VerticalAlignment = Constants.xlCenter;

         int ebene = 0;
         for (int i = 1; i < 8; i = i + 2)
         {
            worksheet.Cells[6, i] = "Ebene_" + ebene + "\n" + "Level_" + ebene;
            worksheet.Cells[6, i].Characters(1, 7).Font.Bold = true;

            worksheet.Cells[6, i + 1] = "Benennung" + "\n" + "Nomenclature";
            worksheet.Cells[6, i + 1].Characters(1, 9).Font.Bold = true;

            Range levelRange;
            Range nomenclatureRange;
            if (ebene == 0)
            {
               levelRange = worksheet.Range["A6", "A7"];
               nomenclatureRange = worksheet.Range["B6", "B7"];
            }
            else if (ebene == 1)
            {
               levelRange = worksheet.Range["C6", "C7"];
               nomenclatureRange = worksheet.Range["D6", "D7"];
            }
            else if (ebene == 2)
            {
               levelRange = worksheet.Range["E6", "E7"];
               nomenclatureRange = worksheet.Range["F6", "F7"];
            }
            else
            {
               levelRange = worksheet.Range["G6", "G7"];
               nomenclatureRange = worksheet.Range["H6", "H7"];
            }

            levelRange.Merge(missingValue); // Zellen verbinden
            levelRange.BorderAround(); // Rahmen
            levelRange.HorizontalAlignment = Constants.xlCenter;
            levelRange.VerticalAlignment = Constants.xlCenter;
            levelRange.Interior.ColorIndex = 6; // Farbe gelb

            nomenclatureRange.Merge(missingValue); // Zellen verbinden
            nomenclatureRange.BorderAround(); // Rahmen
            nomenclatureRange.HorizontalAlignment = Constants.xlCenter;
            nomenclatureRange.VerticalAlignment = Constants.xlCenter;

            ebene++;
         }

         // Farbe rausnehmen aus paar Zellen
         Range rg1 = worksheet.Range["B4", "B5"];
         rg1.Interior.Pattern = Constants.xlNone;

         Range rg2 = worksheet.Range["D1", "D5"];
         rg2.Interior.Pattern = Constants.xlNone;

         Range rg3 = worksheet.Range["F1", "F5"];
         rg3.Interior.Pattern = Constants.xlNone;

         Range rg4 = worksheet.Range["H1", "H5"];
         rg4.Interior.Pattern = Constants.xlNone;

      }


      /// <summary>Create sub tree.</summary>
      /// <param name="valueArray">The value array of the worksheet.</param>
      /// <param name="worksheet">The worksheet.</param>
      /// <param name="treeNodeParent">The parent tree node.</param>
      public static void CreateSubTree(object[,] valueArray, Worksheet worksheet, TreeStructure treeNodeParent)
      {
         // check input 
         if (null == worksheet)
         {
            return;
         }

         int row = treeNodeParent.GetRow();
         int maxRow = treeNodeParent.GetMaxRow();
         int col = treeNodeParent.GetCol() + 2;

         for (int i = row; i <= maxRow; i++)
         {
            if (worksheet.Cells[i, col].Value != null)
            {
               string cell = (valueArray[i, col].ToString());

               // next cell must be nomenclature
               string nextCell = "nomenclature";
               if (worksheet.Cells[i, col + 1].Value != null)
               {
                  nextCell = (valueArray[i, col + 1].ToString());
               }

               // Get max range 
               int maxRange = GetRange(worksheet, col, i);

               TreeStructure child = new TreeStructure(treeNodeParent, cell, nextCell, i, col, maxRange);
               treeNodeParent.AddChild(child);

               // recursive call 
               CreateSubTree(valueArray, worksheet, child);
            }
         }
      }


      /// <summary>Get Range of a given column and row.</summary>
      /// <param name="worksheet">The current worksheet.</param>
      /// <param name="currentColumn">The current column</param>
      /// <param name="currentRow">The current row.</param>
      /// <returns>The max row</returns>
      public static int GetRange(Worksheet worksheet, int currentColumn, int currentRow)
      {
         // max row is max range of worksheet
         int maxRow = worksheet.UsedRange.Rows.Count;

         // start at next row
         currentRow++;

         for (int i = currentRow; i <= worksheet.UsedRange.Rows.Count; i++)
         {
            if (worksheet.Cells[i, currentColumn].Value != null)
            {
               maxRow = i;
               break;
            }
         }
         return maxRow;
      }


      /// <summary>Get root values from worksheet.</summary>
      /// <param name="worksheet">The worksheet to read in.</param>
      /// <param name="valueArray">The value array of the worksheet.</param>
      /// <param name="rootItemNumber">out: root item number</param>
      /// <param name="rootNomenclature">out root nomenclature</param>
      public static void GetRootValuesFromWorksheet(Worksheet worksheet, Object[,] valueArray, out string rootItemNumber, out string rootNomenclature)
      {
         // Initialize out parameter
         rootItemNumber = "";
         rootNomenclature = "";

         UI theUI = UI.GetUI();

         // try to get cell content 3,1 to get root item number
         if (worksheet.Cells[3, 1].Value != null)
         {
            rootItemNumber = valueArray[3, 1].ToString();
         }
         else
         {
            const string message = "Value of root item number (column: 1, row: 3) is empty !";
            theUI.NXMessageBox.Show("Excel file failure", NXMessageBox.DialogType.Warning, message);
            return;
         }

         if ( rootItemNumber.Length != 13 )
         {
            const string message = "Value of root item number (column: 1, row: 3) has not 13 characters !";
            theUI.NXMessageBox.Show("Excel file failure", NXMessageBox.DialogType.Warning, message);
         }

         // try to get cell content 3,2 to get root nomenclature
         if (worksheet.Cells[3, 2].Value != null)
         {
            rootNomenclature = valueArray[3, 2].ToString();
         }
         else
         {
            const string message = "Value of nomenclature (column: 2, row: 3) is empty !";
            theUI.NXMessageBox.Show("Excel file failure", NXMessageBox.DialogType.Warning, message);
         }
      }


   }



}
