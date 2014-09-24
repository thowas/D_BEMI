//==============================================================================
//
//        Filename: TreeStructure.cs
//
//        Created by: CENIT AG (Jan Assmann)
//              Version: NX 8.5.2.3 MP1
//              Date: 03-12-2013  (Format: mm-dd-yyyy)
//              Time: 08:30 (Format: hh-mm)
//
//==============================================================================

using System.Collections.Generic;


namespace Daimler.NX.BemiStructure
{

   /// <summary>Tree structure class</summary>
   public class TreeStructure
   {

      //------------------------------------------------------------------
      //-------------------- Class Members -------------------------------
      //------------------------------------------------------------------

      #region Class Members

      private TreeStructure _parent;
      private readonly List<TreeStructure> _children;
      private readonly string _value;
      private readonly string _nomenclature;
      private readonly int _max;
      private readonly int _row;
      private readonly int _col;
      private bool _creation;

      #endregion Class Members


      /// <summary>Constructor of Tree structure class</summary>
      /// <param name="parent">The parent tree node</param>
      /// <param name="value">The value</param>
      /// <param name="nomenclature">The nomenclature</param>
      /// <param name="row">The Row</param>
      /// <param name="col">The Column</param>
      /// <param name="max">The maximum</param>
      public TreeStructure(TreeStructure parent, string value, string nomenclature, int row, int col, int max)
      { 
         // write to members
         _row = row;
         _col = col;
         _max = max;
         _nomenclature = nomenclature.ToUpper();
         _parent = parent;
         _creation = true;

         _children = new List<TreeStructure>();

         string rootValue = "";
         if ( null != parent )
         {
            TreeStructure rootNode = null;
            TreeStructure treeNodeParent = parent;
            while (null != treeNodeParent)
            {
               // remember current parent as root component
               rootNode = treeNodeParent;

               // next parent is current parent
               treeNodeParent = treeNodeParent.GetParent();
            }
            rootValue = rootNode.GetValue();
         }

         _value = rootValue + value;
      }

      public void SetCreation(bool creation)
      {
         _creation = creation;
      }

      public bool ShouldCreated()
      {
         return _creation;
      }

      /// <summary>Add child to tree structure</summary>
      /// <param name="childToAdd">The child to add</param>
      public void AddChild(TreeStructure childToAdd)
      {
         _children.Add(childToAdd);
      }


      /// <summary>Set parent to tree structure</summary>
      /// <param name="parent">The parent to set</param>
      public void SetParent(TreeStructure parent)
      {
         _parent = parent;
      }


      /// <summary>Get Value of tree structure (node)</summary>
      /// <returns>The string value.</returns>
      public string GetValue()
      {
         return _value;
      }


      /// <summary>Get row of tree structure (node)</summary>
      /// <returns></returns>
      public int GetRow()
      {
         return _row;
      }


      /// <summary>Get the maximum row</summary>
      /// <returns>The maximum row.</returns>
      public int GetMaxRow()
      {
         return _max;
      }


      /// <summary>Get Column of tree structure (node)</summary>
      /// <returns>The column</returns>
      public int GetCol()
      {
         return _col;
      }


      /// <summary>Get all direct children of tree structure (node)</summary>
      /// <returns>List of all direct children.</returns>
      public List<TreeStructure> GetChidren()
      {
         return _children;
      }


      /// <summary>Has children</summary>
      /// <returns></returns>
      public bool HasChildren()
      {
         if ( _children.Count == 0 )
         {
            return false;
         }
         return true;
      }


      /// <summary>Has children in level 1</summary>
      /// <returns>true or false</returns>
      public bool HasChildrenLevel1()
      {
         bool level1 = false;

         if ( HasChildren() )
         {
            foreach (var child in _children)
            {
               if (child.GetChidren().Count > 0 )
               {
                  level1 = true;
                  break;
               }
            }        
         }
         return level1;
      }


      /// <summary>Get nomenclature of tree structure (node)</summary>
      /// <returns>The nomenclature.</returns>
      public string GetNomenclature()
      {
         return _nomenclature;
      }


      /// <summary>Get parent of tree structure (node)</summary>
      /// <returns>The parent</returns>
      public TreeStructure GetParent()
      {
         return _parent;
      }


      /// <summary>Get root value of the tree structure tree</summary>
      /// <returns>The root value.</returns>
      public string GetRootValue()
      {
         string rootValue = "";

         if (null != _parent)
         {
            TreeStructure rootNode = null;
            TreeStructure treeNodeParent = _parent;

            while (null != treeNodeParent)
            {
               // remember current parent as root component
               rootNode = treeNodeParent;

               // next parent is current parent
               treeNodeParent = treeNodeParent.GetParent();
            }
            rootValue = rootNode.GetValue();
         }

         return rootValue;
      }

   }
}
