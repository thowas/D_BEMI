//==============================================================================
//
//        Filename: ComponentStatics.cs
//
//        Created by: CENIT AG (Jan Assmann)
//              Version: NX 8.5.2.3 MP1
//              Date: 03-12-2013  (Format: mm-dd-yyyy)
//              Time: 08:30 (Format: hh-mm)
//
//==============================================================================

using System;
using System.Collections.Generic;
using NXOpen;
using NXOpen.Assemblies;
using NXOpen.Positioning;

namespace Daimler.NX.BemiStructure
{
   /// <summary>Component statics class </summary>
   public static class ComponentStatics
   {

      /// <summary>Get root component of a component.</summary>
      /// <param name="component">Component to get root component for.</param>
      /// <returns>Root component of component.</returns>
      public static Component GetRootComponent(Component component)
      {
         Component rootComponent = null;

         // for the fathers (beginning at the passed component)
         Component currentFather = component;
         while (null != currentFather)
         {
            // remember current father as root component
            rootComponent = currentFather;

            // next father is current parent
            currentFather = GetParentComponent(currentFather);
         }

         return rootComponent;
      }


      /// <summary>Get parent component of a component.</summary>
      /// <param name="component">Component to get parent component for.</param>
      /// <returns>Parent component of component.</returns>
      public static Component GetParentComponent(Component component)
      {
         // get parent of component
         Component parentComponent = component.Parent;

         return parentComponent;
      }


      /// <summary>Get all children of given component.</summary>
      /// <param name="father">The component to get the children for.</param>
      /// <param name="recursively"><c>false</c> only returns the direct children.
      ///                           <c>true</c> returns all children.</param>
      /// <returns>The children of the component.</returns>
      public static List<Component> GetChildren(Component father, bool recursively)
      {
         List<Component> children = new List<Component>();

         // get direct children of father
         List<Component> directChildren = new List<Component>(father.GetChildren());

         // if all children should be gotten recursively
         if (recursively)
         {
            // for each direct child
            foreach (Component currentComponent in directChildren)
            {
               // add current child to output list
               children.Add(currentComponent);

               // add children of current child to output list
               children.AddRange(GetChildren(currentComponent, true));
            }
         }

         // if only direct children should be returned
         else
         {
            // add all direct children to output list
            children.AddRange(directChildren);
         }

         return children;
      }



      /// <summary>Add constraint (fix)</summary>
      /// <param name="part">The parent Part inside to fix.</param>
      /// <param name="componentToFix">The component to fix.</param>
      /// <param name="theSession">The current session</param>
      public static void AddFixConstraint(Part part, Component componentToFix, Session theSession)
      {
         // Check input
         if ( null == part )
         {
            return;
         }
         if ( null == componentToFix )
         {
            return;
         }
         if ( null == theSession )
         {
            return;
         }

         // Component positioner
         ComponentPositioner componentPositioner = part.ComponentAssembly.Positioner;

         if ( null != componentPositioner )
         {
            // Begin Constraints
            componentPositioner.BeginAssemblyConstraints();

            // Create Constraint
            Constraint constraint = componentPositioner.CreateConstraint(true);

            if ( null != constraint )
            {
               ComponentConstraint componentConstraint = (ComponentConstraint)constraint;

               // Set type --> Fix
               componentConstraint.ConstraintType = Constraint.Type.Fix;

               try
               {
                  // Create constraint reference
                  ConstraintReference constraintReference = componentConstraint.CreateConstraintReference(componentToFix, componentToFix, false, false, false);

                  // Help point
                  Point3d helpPoint = new Point3d(0.0, 0.0, 0.0);
                  constraintReference.HelpPoint = helpPoint;
               }
               catch (Exception ex)
               {
                  string displayName = componentToFix.DisplayName;
                  UI.GetUI().NXMessageBox.Show("Constraint test!!", NXMessageBox.DialogType.Error, ex.ToString() + displayName);
               }

               
               // Delete non persistent constraints
               componentPositioner.DeleteNonPersistentConstraints();

               // End constraints
               componentPositioner.EndAssemblyConstraints();
            }
         }
      }


   }
}
