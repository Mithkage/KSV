//
// File: CopyRelative.cs
// Company: ReTick Solutions Pty Ltd
// Function: This script places instances of a user-selected family instance (e.g., lighting fixture,
//           electrical outlet, etc.) in the active model. The placement is based on a geometric
//           relationship established from a user-selected "template" reference element (e.g., a door)
//           in either the host or a linked model, and a "template" element to copy in the host model.
//           The script then iterates through all other elements of the same type as the reference
//           and creates new elements at the corresponding relative locations.
//
// Change Log:
// ... (previous logs)
// 2025-08-08: (Improvement) Implemented ISelectionFilter for a better user experience, preventing invalid selections.
// 2025-08-08: (Improvement) Generalized the reference element; it can now be any FamilyInstance, not just a door.
// 2025-08-08: (Improvement) Added robustness check for available 3D views for host finding.
// 2025-08-08: (Improvement) Corrected the calculation for skipped elements in user feedback.
// 2025-08-08: (Fix) Corrected compiler error CS0117 by removing invalid 'OST_FaceBased' category from filter.
//
using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Structure;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;

namespace RTS.Commands.MEPTools.Positioning
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
                if (!GetTemplateElements(out FamilyInstance templateReferenceElement, out RevitLinkInstance linkInstance, out FamilyInstance elementToCopy))
                {
                    return Result.Cancelled;
                }

                // Calculate the geometric relationship
                if (!CalculateGeometricRelationship(templateReferenceElement, elementToCopy, linkInstance, out Transform relativeTransform))
                {
                    return Result.Failed;
                }

                // Find all target elements
                List<FamilyInstance> targetElements = FindAllTargetElements(templateReferenceElement, linkInstance);

                // Place new elements
                int placedCount = PlaceNewElements(targetElements, elementToCopy, templateReferenceElement, linkInstance, relativeTransform);
                int skippedCount = targetElements.Count - 1 - placedCount; // -1 for the template reference itself

                string categoryName = elementToCopy.Category.Name;

                // Provide detailed feedback
                string feedback = $"Successfully placed {placedCount} new {categoryName} elements.";
                if (skippedCount > 0)
                {
                    feedback += $"\nSkipped {skippedCount} locations due to invalid target geometry.";
                }
                TaskDialog.Show("Success", feedback);

                return Result.Succeeded;
            }
            catch (OperationCanceledException)
            {
                return Result.Cancelled;
            }
            catch (Exception ex)
            {
                message = ex.Message;
                TaskDialog.Show("Error", "An unexpected error occurred: " + ex.Message);
                return Result.Failed;
            }
        }

        /// <summary>
        /// Prompts the user to select template elements using selection filters for a better UX.
        /// </summary>
        private bool GetTemplateElements(out FamilyInstance templateReferenceElement, out RevitLinkInstance linkInstance, out FamilyInstance elementToCopy)
        {
            templateReferenceElement = null;
            linkInstance = null;
            elementToCopy = null;

            // Step 1: Select the template reference element (any FamilyInstance)
            var referenceFilter = new FamilyInstanceSelectionFilter();
            Reference referenceElementRef = _uidoc.Selection.PickObject(ObjectType.Element, referenceFilter, "Select a template reference element (e.g., door, window, desk).");

            Element pickedElement = _doc.GetElement(referenceElementRef.ElementId);
            if (pickedElement is RevitLinkInstance)
            {
                linkInstance = pickedElement as RevitLinkInstance;
                Document linkedDoc = linkInstance.GetLinkDocument();
                if (linkedDoc == null)
                {
                    TaskDialog.Show("Selection Error", "Could not get the linked document. Ensure the link is loaded.");
                    return false;
                }
                templateReferenceElement = linkedDoc.GetElement(referenceElementRef.LinkedElementId) as FamilyInstance;
            }
            else
            {
                templateReferenceElement = pickedElement as FamilyInstance;
            }

            if (templateReferenceElement == null) return false; // Should not happen with filter, but good practice

            // Step 2: Select the element to copy (specific categories)
            var supportedCategories = new List<BuiltInCategory>
            {
                BuiltInCategory.OST_LightingFixtures,
                BuiltInCategory.OST_ElectricalFixtures,
                BuiltInCategory.OST_ElectricalEquipment,
                BuiltInCategory.OST_LightingDevices,
                BuiltInCategory.OST_CommunicationDevices
            };
            var elementToCopyFilter = new FamilyInstanceSelectionFilter(supportedCategories);
            Reference elementToCopyRef = _uidoc.Selection.PickObject(ObjectType.Element, elementToCopyFilter, "Select the template element to copy (e.g., light, switch, outlet).");

            elementToCopy = _doc.GetElement(elementToCopyRef.ElementId) as FamilyInstance;

            return elementToCopy != null;
        }

        /// <summary>
        /// Creates a local coordinate system (Transform) from a FamilyInstance's location and orientation.
        /// </summary>
        private Transform CreateTransformFromInstance(FamilyInstance instance)
        {
            if (!(instance.Location is LocationPoint locationPoint)) return null;

            var origin = locationPoint.Point;
            var basisX = instance.HandOrientation.Normalize();
            var basisZ = instance.FacingOrientation.Normalize();
            var basisY = basisZ.CrossProduct(basisX).Normalize();

            var transform = Transform.Identity;
            transform.Origin = origin;
            transform.BasisX = basisX;
            transform.BasisY = basisY;
            transform.BasisZ = basisZ;

            return transform;
        }

        /// <summary>
        /// Calculates the relative transformation from the reference element's coordinate system to the element to copy's.
        /// </summary>
        private bool CalculateGeometricRelationship(FamilyInstance templateReference, FamilyInstance elementToCopy, RevitLinkInstance linkInstance, out Transform relativeTransform)
        {
            relativeTransform = null;
            Transform linkTransform = linkInstance?.GetTotalTransform() ?? Transform.Identity;

            Transform referenceTransform = CreateTransformFromInstance(templateReference);
            Transform elementToCopyTransform = CreateTransformFromInstance(elementToCopy);

            if (referenceTransform == null || elementToCopyTransform == null)
            {
                TaskDialog.Show("Error", "Could not determine the location of one or both template elements.");
                return false;
            }

            Transform referenceTransformInHost = linkTransform.Multiply(referenceTransform);
            relativeTransform = referenceTransformInHost.Inverse.Multiply(elementToCopyTransform);
            return true;
        }

        /// <summary>
        /// Finds all elements of the same type as the template reference element in the source document (host or link).
        /// </summary>
        private List<FamilyInstance> FindAllTargetElements(FamilyInstance templateReferenceElement, RevitLinkInstance linkInstance)
        {
            Document sourceDoc = linkInstance?.GetLinkDocument() ?? _doc;
            if (sourceDoc == null) return new List<FamilyInstance>();

            var referenceTypeId = templateReferenceElement.GetTypeId();

#if REVIT2024_OR_GREATER
            return new FilteredElementCollector(sourceDoc)
                .OfCategory((BuiltInCategory)templateReferenceElement.Category.Id.Value)
                .OfClass(typeof(FamilyInstance))
                .Where(fi => fi.GetTypeId() == referenceTypeId)
                .Cast<FamilyInstance>()
                .ToList();
#else
            return new FilteredElementCollector(sourceDoc)
                .OfCategory((BuiltInCategory)templateReferenceElement.Category.Id.IntegerValue)
                .OfClass(typeof(FamilyInstance))
                .Where(fi => fi.GetTypeId() == referenceTypeId)
                .Cast<FamilyInstance>()
                .ToList();
#endif
        }

        /// <summary>
        /// Finds a host element by casting rays from a given point.
        /// </summary>
        private Element GetHostElementAtPoint(Document doc, XYZ point)
        {
            View3D view3D = new FilteredElementCollector(doc).OfClass(typeof(View3D)).Cast<View3D>().FirstOrDefault(v => !v.IsTemplate);
            if (view3D == null)
            {
                // This is a critical failure for host finding. We can't proceed without a 3D view.
                // For this tool, we will allow unhosted placement, so we just return null.
                // A more advanced implementation could warn the user here.
                return null;
            }

            var categories = new List<BuiltInCategory>
            {
                BuiltInCategory.OST_Ceilings,
                BuiltInCategory.OST_Floors,
                BuiltInCategory.OST_Roofs,
                BuiltInCategory.OST_Walls
            };
            var categoryFilter = new ElementMulticategoryFilter(categories);

            var intersector = new ReferenceIntersector(categoryFilter, FindReferenceTarget.Face, view3D)
            {
                FindReferencesInRevitLinks = false
            };

            var refUp = intersector.FindNearest(point, XYZ.BasisZ);
            var refDown = intersector.FindNearest(point, -XYZ.BasisZ);

            if (refUp != null && refDown != null)
            {
                return refUp.Proximity < refDown.Proximity ? doc.GetElement(refUp.GetReference()) : doc.GetElement(refDown.GetReference());
            }
            return refUp != null ? doc.GetElement(refUp.GetReference()) : refDown != null ? doc.GetElement(refDown.GetReference()) : null;
        }

        /// <summary>
        /// Places new elements based on the calculated geometric relationship and the target elements.
        /// </summary>
        private int PlaceNewElements(List<FamilyInstance> targetElements, FamilyInstance elementToCopy, FamilyInstance templateReferenceElement, RevitLinkInstance linkInstance, Transform relativeTransform)
        {
            var elementSymbol = elementToCopy.Symbol;
            int placedCount = 0;

            if (!elementSymbol.IsActive)
            {
                using (var activateTrans = new Transaction(_doc, "Activate Family Symbol"))
                {
                    activateTrans.Start();
                    elementSymbol.Activate();
                    activateTrans.Commit();
                }
            }

            Transform linkTransform = linkInstance?.GetTotalTransform() ?? Transform.Identity;

            using (var trans = new Transaction(_doc, $"Place {elementToCopy.Category.Name} by Reference"))
            {
                trans.Start();

                foreach (var targetElement in targetElements)
                {
                    if (targetElement.Id == templateReferenceElement.Id) continue;

                    Transform targetTransform = CreateTransformFromInstance(targetElement);
                    if (targetTransform == null) continue;

                    Transform targetTransformInHost = linkTransform.Multiply(targetTransform);
                    Transform newElementTransform = targetTransformInHost.Multiply(relativeTransform);
                    XYZ newElementLocation = newElementTransform.Origin;

                    var hostElement = GetHostElementAtPoint(_doc, newElementLocation);

                    FamilyInstance newElement = _doc.Create.NewFamilyInstance(
                        newElementLocation,
                        elementSymbol,
                        hostElement,
                        StructuralType.NonStructural);

                    if (newElement == null) continue;

                    if (newElement.Location is LocationPoint newElementLocationPoint)
                    {
                        double angle = XYZ.BasisX.AngleOnPlaneTo(newElementTransform.BasisX, XYZ.BasisZ);

                        if (Math.Abs(angle) > 1e-9)
                        {
                            var axis = Line.CreateBound(newElementLocation, newElementLocation + XYZ.BasisZ);
                            ElementTransformUtils.RotateElement(_doc, newElement.Id, axis, angle);
                        }
                    }
                    placedCount++;
                }

                trans.Commit();
            }
            return placedCount;
        }
    }

    /// <summary>
    /// A selection filter to allow the user to select only FamilyInstance elements,
    /// optionally restricted to a list of specific categories.
    /// </summary>
    public class FamilyInstanceSelectionFilter : ISelectionFilter
    {
        private readonly List<BuiltInCategory> _allowedCategories;

        /// <summary>
        /// Constructor to allow any FamilyInstance.
        /// </summary>
        public FamilyInstanceSelectionFilter()
        {
            _allowedCategories = null; // Null means all categories are allowed
        }

        /// <summary>
        /// Constructor to restrict to specific categories.
        /// </summary>
        public FamilyInstanceSelectionFilter(List<BuiltInCategory> allowedCategories)
        {
            _allowedCategories = allowedCategories;
        }

        public bool AllowElement(Element elem)
        {
            if (!(elem is FamilyInstance)) return false;

            if (_allowedCategories == null || _allowedCategories.Count == 0)
            {
                return true; // No category restriction
            }

#if REVIT2024_OR_GREATER
            return _allowedCategories.Contains((BuiltInCategory)elem.Category.Id.Value);
#else
            return _allowedCategories.Contains((BuiltInCategory)elem.Category.Id.IntegerValue);
#endif
        }

        public bool AllowReference(Reference reference, XYZ position)
        {
            // Set to true if you need to select geometry within an element, like a face or edge.
            // For this tool, we only need to select the element itself, so false is correct.
            return false;
        }
    }
}
