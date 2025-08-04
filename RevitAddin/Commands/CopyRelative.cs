//
// File: CopyRelative.cs
// Company: ReTick Solutions Pty Ltd
// Function: This script places instances of a user-selected lighting fixture in the active model.
//           The placement is based on a geometric relationship (offset and rotation)
//           established from a user-selected "template" door in a linked model and a "template"
//           light fixture in the host model. The script then iterates through all other doors
//           of the same type in the linked model and creates new light fixtures at the
//           corresponding relative locations and rotations.
//
// Change Log:
// 2025-08-01: Initial version of the script created.
// 2025-08-01: Updated namespace to RTS.Commands.
// 2025-08-02: Implemented improved error handling, code refactoring into helper methods,
//             and an optimized element collector for better performance and readability.
// 2025-08-04: Renamed file to CopyRelative.cs and class to CopyRelativeClass.
// 2025-08-04: Fixed compiler errors CS0103 and CS0122 by correcting variable scope
//             and adding a missing namespace reference.
//
using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Structure; // Added to resolve CS0122
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;

namespace RTS.Commands
{
    [Transaction(TransactionMode.Manual)]
    public class CopyRelativeClass : IExternalCommand
    {
        private UIDocument _uidoc;
        private Document _doc;

        public Result Execute(
          ExternalCommandData commandData,
          ref string message,
          ElementSet elements)
        {
            _uidoc = commandData.Application.ActiveUIDocument;
            _doc = _uidoc.Document;

            try
            {
                // Get the template elements from user selection
                FamilyInstance templateDoor;
                FamilyInstance templateLight;
                if (!GetTemplateElements(out templateDoor, out templateLight))
                {
                    // The error message is already set by the helper method
                    return Result.Failed;
                }

                // Calculate the geometric relationship
                XYZ baseRelativeVector;
                double rotationAngle;
                CalculateGeometricRelationship(templateDoor, templateLight, out baseRelativeVector, out rotationAngle);

                // Find all target doors
                List<FamilyInstance> targetDoors = FindAllTargetDoors(templateDoor);

                // Place new light fixtures
                PlaceNewLights(targetDoors, templateLight, templateDoor, baseRelativeVector, rotationAngle);

                TaskDialog.Show("Success", $"Successfully placed {targetDoors.Count - 1} new light fixtures.");
                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                message = ex.Message;
                TaskDialog.Show("Error", "An unexpected error occurred: " + ex.Message);
                return Result.Failed;
            }
        }

        /// <summary>
        /// Prompts the user to select the template door and light, and validates the selections.
        /// </summary>
        private bool GetTemplateElements(out FamilyInstance templateDoor, out FamilyInstance templateLight)
        {
            templateDoor = null;
            templateLight = null;

            // Step 1: User selects the template door in the linked model
            try
            {
                Reference doorRef = _uidoc.Selection.PickObject(
                    ObjectType.LinkedElement,
                    "Select a template door in the linked model.");

                RevitLinkInstance linkInstance = _doc.GetElement(doorRef.ElementId) as RevitLinkInstance;
                if (linkInstance == null)
                {
                    TaskDialog.Show("Selection Error", "The selected element is not in a linked model.");
                    return false;
                }

                Document linkedDoc = linkInstance.GetLinkDocument();
                if (linkedDoc == null)
                {
                    TaskDialog.Show("Selection Error", "Could not get the linked document.");
                    return false;
                }

                templateDoor = linkedDoc.GetElement(doorRef.LinkedElementId) as FamilyInstance;
                if (templateDoor == null)
                {
                    TaskDialog.Show("Selection Error", "The selected element is not a family instance (door).");
                    return false;
                }
            }
            catch (OperationCanceledException)
            {
                return false; // User cancelled
            }
            catch (Exception ex)
            {
                TaskDialog.Show("Error", "An error occurred during door selection: " + ex.Message);
                return false;
            }

            // Step 2: User selects the template light fixture in the host model
            try
            {
                Reference lightRef = _uidoc.Selection.PickObject(
                    ObjectType.Element,
                    "Select the light fixture to copy.");

                templateLight = _doc.GetElement(lightRef.ElementId) as FamilyInstance;
                if (templateLight == null)
                {
                    TaskDialog.Show("Selection Error", "The selected element is not a family instance (light fixture).");
                    return false;
                }
            }
            catch (OperationCanceledException)
            {
                return false; // User cancelled
            }
            catch (Exception ex)
            {
                TaskDialog.Show("Error", "An error occurred during light fixture selection: " + ex.Message);
                return false;
            }

            return true;
        }

        /// <summary>
        /// Calculates the relative offset and rotation angle between the template door and light.
        /// </summary>
        private void CalculateGeometricRelationship(FamilyInstance templateDoor, FamilyInstance templateLight,
            out XYZ baseRelativeVector, out double rotationAngle)
        {
            // Get locations and orientations
            XYZ templateDoorOrigin = (templateDoor.Location as LocationPoint).Point;
            XYZ templateDoorFacing = templateDoor.FacingOrientation.Normalize();
            XYZ templateLightOrigin = (templateLight.Location as LocationPoint).Point;
            XYZ templateLightFacing = templateLight.FacingOrientation.Normalize();

            // Calculate the vector from the door to the light
            XYZ relativeOffset = templateLightOrigin - templateDoorOrigin;

            // Calculate the rotation angle between the facing vectors
            rotationAngle = templateLightFacing.AngleOnPlaneTo(templateDoorFacing, XYZ.BasisZ);
            Transform rotationTransform = Transform.CreateRotation(XYZ.BasisZ, rotationAngle);

            // Rotate the relative offset vector to align it with the door's facing direction.
            // This 'baseRelativeVector' represents the offset in a normalized coordinate system
            // and will be re-rotated for each new door.
            baseRelativeVector = rotationTransform.Inverse.OfVector(relativeOffset);
        }

        /// <summary>
        /// Finds all doors of the same type as the template door in the linked model.
        /// </summary>
        private List<FamilyInstance> FindAllTargetDoors(FamilyInstance templateDoor)
        {
            // Fix for CS0103: Access the linked document directly from the template door
            Document linkedDoc = templateDoor.Document;
            ElementId doorTypeId = templateDoor.GetTypeId();

            // Use an optimized FilteredElementCollector with a native category filter
            return new FilteredElementCollector(linkedDoc)
                .OfCategory(BuiltInCategory.OST_Doors)
                .OfClass(typeof(FamilyInstance))
                .Where(fi => fi.GetTypeId() == doorTypeId)
                .Cast<FamilyInstance>()
                .ToList();
        }

        /// <summary>
        /// Places new light fixtures based on the calculated geometric relationship.
        /// </summary>
        private void PlaceNewLights(List<FamilyInstance> targetDoors, FamilyInstance templateLight, FamilyInstance templateDoor,
            XYZ baseRelativeVector, double rotationAngle)
        {
            FamilySymbol lightSymbol = _doc.GetElement(templateLight.GetTypeId()) as FamilySymbol;

            if (!lightSymbol.IsActive)
            {
                using (Transaction activateTrans = new Transaction(_doc, "Activate FamilySymbol"))
                {
                    activateTrans.Start();
                    lightSymbol.Activate();
                    activateTrans.Commit();
                }
            }

            using (Transaction trans = new Transaction(_doc, "Place Lights by Door"))
            {
                trans.Start();

                // Using a try-finally block for robust transaction handling
                try
                {
                    foreach (FamilyInstance targetDoor in targetDoors)
                    {
                        // Skip the template door to avoid creating a duplicate light at its location
                        if (targetDoor.Id == templateDoor.Id)
                        {
                            continue;
                        }

                        LocationPoint targetDoorLocation = targetDoor.Location as LocationPoint;
                        if (targetDoorLocation == null) continue;

                        XYZ targetDoorOrigin = targetDoorLocation.Point;
                        XYZ targetDoorFacing = targetDoor.FacingOrientation.Normalize();

                        // Calculate the rotation needed for the current door
                        double targetDoorRotationAngle = XYZ.BasisX.AngleOnPlaneTo(targetDoorFacing, XYZ.BasisZ);
                        Transform targetDoorRotationTransform = Transform.CreateRotation(XYZ.BasisZ, targetDoorRotationAngle);

                        // Apply the door's rotation to the base relative vector to find the new light's location
                        XYZ newLightVector = targetDoorRotationTransform.OfVector(baseRelativeVector);
                        XYZ newLightLocation = targetDoorOrigin + newLightVector;

                        // Calculate the final light rotation
                        double newLightRotation = targetDoorLocation.Rotation + rotationAngle;

                        // Create the new family instance
                        FamilyInstance newLight = _doc.Create.NewFamilyInstance(
                            newLightLocation,
                            lightSymbol,
                            StructuralType.NonStructural); // Fix for CS0122

                        // Rotate the new light instance
                        LocationPoint newLightLocationPoint = newLight.Location as LocationPoint;
                        if (newLightLocationPoint != null)
                        {
                            newLightLocationPoint.Rotate(Line.CreateBound(newLightLocation, newLightLocation + XYZ.BasisZ), newLightRotation);
                        }
                    }
                }
                finally
                {
                    if (trans.GetStatus() == TransactionStatus.Started)
                    {
                        trans.Commit();
                    }
                }
            }
        }
    }
}