//
// File: CopyRelative.cs
// Company: ReTick Solutions Pty Ltd
// Function: This script places instances of a user-selected family instance (e.g., lighting fixture,
//           electrical outlet, etc.) in the active model. The placement is based on a geometric
//           relationship established from a pre-selected "template" reference element (e.g., a door)
//           in either the host or a linked model, and a "template" element to copy in the host model.
//           The script then iterates through all other elements of the same type as the reference
//           and creates new elements at the corresponding relative locations.
//
// Change Log:
// ... (previous logs)
// 2025-08-11: (Improvement) Added logic to write reference element data to shared parameters for traceability.
// 2025-08-11: (Fix) Corrected duplicate placement logic by tracking newly placed locations within the transaction.
// 2025-08-11: (Fix) Implemented version-specific selection logic to support Revit 2022 and 2024.
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
using RTS.UI; // Added using directive for the progress bar UI
using System.Text; // For StringBuilder

namespace RTS.Commands.MEPTools.Positioning
{
    [Transaction(TransactionMode.Manual)]
    public class CopyRelativeClass : IExternalCommand
    {
        private UIDocument _uidoc;
        private Document _doc;
        // Shared Parameter GUIDs
        private readonly Guid _refTypeGuid = new Guid("4d6ce1ad-eb55-47e2-acb6-69490634990e");
        private readonly Guid _refIdGuid = new Guid("a377a731-1886-4b93-b477-a6208f672987");

        public Result Execute(
            ExternalCommandData commandData,
            ref string message,
            ElementSet elements)
        {
            _uidoc = commandData.Application.ActiveUIDocument;
            _doc = _uidoc.Document;
            var missingParamsFamilies = new HashSet<string>();

            try
            {
                // Step 1: Get the pre-selected reference element
                if (!GetPreselectedReference(out FamilyInstance templateReferenceElement, out RevitLinkInstance linkInstance))
                {
                    return Result.Failed;
                }

                // Step 2: Prompt for the elements to copy
                if (!GetElementsToCopy(out List<FamilyInstance> elementsToCopy))
                {
                    return Result.Cancelled;
                }

                // Step 3: Update parameters on the original template elements
                using (var trans = new Transaction(_doc, "Update Template Reference Parameters"))
                {
                    trans.Start();
                    foreach (var element in elementsToCopy)
                    {
                        UpdateReferenceParameters(element, templateReferenceElement, missingParamsFamilies);
                    }
                    trans.Commit();
                }

                // Step 4: Calculate the geometric relationships
                if (!CalculateGeometricRelationships(templateReferenceElement, elementsToCopy, linkInstance, out List<Transform> relativeTransforms))
                {
                    return Result.Failed;
                }

                // Step 5: Find all target elements
                List<FamilyInstance> targetElements = FindAllTargetElements(templateReferenceElement, linkInstance);

                // Step 6: Place new elements
                (int placedCount, int skippedCount) = PlaceNewElements(targetElements, elementsToCopy, templateReferenceElement, linkInstance, relativeTransforms, missingParamsFamilies);

                // Step 7: Provide final feedback
                StringBuilder feedback = new StringBuilder();
                feedback.AppendLine($"Successfully placed {placedCount} new elements.");
                if (skippedCount > 0)
                {
                    feedback.AppendLine($"Skipped {skippedCount} locations where an element of the same type already existed or the target geometry was invalid.");
                }
                if (missingParamsFamilies.Any())
                {
                    feedback.AppendLine("\nWarning: The following families are missing the required 'RTS_CopyReference' parameters and were not updated:");
                    foreach (var familyName in missingParamsFamilies)
                    {
                        feedback.AppendLine($" - {familyName}");
                    }
                }
                TaskDialog.Show("Operation Complete", feedback.ToString());

                return Result.Succeeded;
            }
            catch (OperationCanceledException)
            {
                TaskDialog.Show("Cancelled", "The operation was cancelled by the user.");
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
        /// Gets the pre-selected reference element and validates it.
        /// </summary>
        private bool GetPreselectedReference(out FamilyInstance templateReferenceElement, out RevitLinkInstance linkInstance)
        {
            templateReferenceElement = null;
            linkInstance = null;

#if REVIT2024_OR_GREATER
            ICollection<Reference> selectedReferences = null;
            try { selectedReferences = _uidoc.Selection.GetReferences(); } catch { }

            if (selectedReferences == null || selectedReferences.Count != 1)
            {
                TaskDialog.Show("Selection Error", "Please select exactly one reference element before running the command.");
                return false;
            }

            var reference = selectedReferences.First();
            var selectedElement = _doc.GetElement(reference.ElementId);

            if (selectedElement is RevitLinkInstance)
            {
                linkInstance = selectedElement as RevitLinkInstance;
                Document linkedDoc = linkInstance.GetLinkDocument();
                if (linkedDoc == null) { TaskDialog.Show("Selection Error", "Could not get the linked document."); return false; }
                templateReferenceElement = linkedDoc.GetElement(reference.LinkedElementId) as FamilyInstance;
            }
            else if (selectedElement is FamilyInstance hostInstance)
            {
                templateReferenceElement = hostInstance;
            }
#else
            try
            {
                var reference = _uidoc.Selection.PickObject(ObjectType.Element, new FamilyInstanceSelectionFilter(), "Select a reference element in the HOST model (or press Esc for linked).");
                var selectedElement = _doc.GetElement(reference.ElementId);
                if (selectedElement is FamilyInstance hostInstance) { templateReferenceElement = hostInstance; }
            }
            catch (Autodesk.Revit.Exceptions.OperationCanceledException)
            {
                try
                {
                    var reference = _uidoc.Selection.PickObject(ObjectType.LinkedElement, new FamilyInstanceSelectionFilter(), "Select a reference element in a LINKED model.");
                    var selectedElement = _doc.GetElement(reference.ElementId);
                    if (selectedElement is RevitLinkInstance li)
                    {
                        linkInstance = li;
                        Document linkedDoc = linkInstance.GetLinkDocument();
                        if (linkedDoc == null) { TaskDialog.Show("Selection Error", "Could not get the linked document."); return false; }
                        templateReferenceElement = linkedDoc.GetElement(reference.LinkedElementId) as FamilyInstance;
                    }
                }
                catch (Autodesk.Revit.Exceptions.OperationCanceledException) { return false; }
            }
#endif

            if (templateReferenceElement == null)
            {
                TaskDialog.Show("Selection Error", "The selected element is not a valid Family Instance.");
                return false;
            }
            return true;
        }

        /// <summary>
        /// Prompts the user to select multiple elements to be copied.
        /// </summary>
        private bool GetElementsToCopy(out List<FamilyInstance> elementsToCopy)
        {
            elementsToCopy = new List<FamilyInstance>();
            try
            {
                var references = _uidoc.Selection.PickObjects(ObjectType.Element, new FamilyInstanceSelectionFilter(), "Select the template elements to copy.");
                foreach (var reference in references)
                {
                    var pickedElement = _doc.GetElement(reference.ElementId);
                    if (pickedElement is FamilyInstance fi)
                    {
                        elementsToCopy.Add(fi);
                    }
                }

                if (!elementsToCopy.Any())
                {
                    TaskDialog.Show("Invalid Selection", "No valid Family Instances were selected.");
                    return false;
                }
                return true;
            }
            catch (OperationCanceledException) { return false; }
        }

        /// <summary>
        /// Creates a transform from a FamilyInstance's location and orientation.
        /// </summary>
        private Transform CreateTransformFromInstance(FamilyInstance instance, RevitLinkInstance linkInstance = null)
        {
            if (!(instance.Location is LocationPoint locationPoint))
            {
                throw new InvalidOperationException($"Element '{instance.Name}' (ID: {instance.Id}) lacks a valid insertion point.");
            }

            var origin = locationPoint.Point;
            var basisX = instance.HandOrientation.Normalize();
            var basisZ = instance.FacingOrientation.Normalize();
            var basisY = basisZ.CrossProduct(basisX).Normalize();

            if (linkInstance != null)
            {
                Transform linkTransform = linkInstance.GetTotalTransform();
                double determinant = linkTransform.BasisX.DotProduct(linkTransform.BasisY.CrossProduct(linkTransform.BasisZ));
                if (determinant < 0)
                {
                    basisY = basisX.CrossProduct(basisZ).Normalize();
                }
            }

            var transform = Transform.Identity;
            transform.Origin = origin;
            transform.BasisX = basisX;
            transform.BasisY = basisY;
            transform.BasisZ = basisZ;
            return transform;
        }

        /// <summary>
        /// Calculates the relative transformations between the reference element and each element to be copied.
        /// </summary>
        private bool CalculateGeometricRelationships(FamilyInstance templateReference, List<FamilyInstance> elementsToCopy, RevitLinkInstance linkInstance, out List<Transform> relativeTransforms)
        {
            relativeTransforms = new List<Transform>();
            Transform linkTransform = linkInstance?.GetTotalTransform() ?? Transform.Identity;
            Transform referenceTransform = CreateTransformFromInstance(templateReference, linkInstance);

            if (referenceTransform == null) return false;

            Transform referenceTransformInHost = linkTransform.Multiply(referenceTransform);

            foreach (var elementToCopy in elementsToCopy)
            {
                Transform elementToCopyTransform = CreateTransformFromInstance(elementToCopy);
                if (elementToCopyTransform == null) return false;

                relativeTransforms.Add(referenceTransformInHost.Inverse.Multiply(elementToCopyTransform));
            }
            return true;
        }

        /// <summary>
        /// Finds all elements of the same type as the reference.
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
                .Cast<FamilyInstance>().ToList();
#else
            return new FilteredElementCollector(sourceDoc)
                .OfCategory((BuiltInCategory)templateReferenceElement.Category.Id.IntegerValue)
                .OfClass(typeof(FamilyInstance))
                .Where(fi => fi.GetTypeId() == referenceTypeId)
                .Cast<FamilyInstance>().ToList();
#endif
        }

        /// <summary>
        /// Finds a host element by ray casting.
        /// </summary>
        private Element GetHostElementAtPoint(Document doc, XYZ point)
        {
            View3D view3D = new FilteredElementCollector(doc).OfClass(typeof(View3D)).Cast<View3D>().FirstOrDefault(v => !v.IsTemplate);
            if (view3D == null) return null;

            var categories = new List<BuiltInCategory> { BuiltInCategory.OST_Ceilings, BuiltInCategory.OST_Floors, BuiltInCategory.OST_Roofs, BuiltInCategory.OST_Walls };
            var categoryFilter = new ElementMulticategoryFilter(categories);
            var intersector = new ReferenceIntersector(categoryFilter, FindReferenceTarget.Face, view3D) { FindReferencesInRevitLinks = false };
            var refUp = intersector.FindNearest(point, XYZ.BasisZ);
            var refDown = intersector.FindNearest(point, -XYZ.BasisZ);

            if (refUp != null && refDown != null)
            {
                return refUp.Proximity < refDown.Proximity ? doc.GetElement(refUp.GetReference()) : doc.GetElement(refDown.GetReference());
            }
            return refUp != null ? doc.GetElement(refUp.GetReference()) : refDown != null ? doc.GetElement(refDown.GetReference()) : null;
        }

        /// <summary>
        /// Places the new elements and updates their reference parameters.
        /// </summary>
        private (int placed, int skipped) PlaceNewElements(
            List<FamilyInstance> targetElements,
            List<FamilyInstance> elementsToCopy,
            FamilyInstance templateReferenceElement,
            RevitLinkInstance linkInstance,
            List<Transform> relativeTransforms,
            HashSet<string> missingParamsFamilies)
        {
            int placedCount = 0, skippedCount = 0, processedCount = 0;
            int totalOperations = targetElements.Count * elementsToCopy.Count;

            var progressBar = new ProgressBarWindow();
            progressBar.Show();

            foreach (var element in elementsToCopy)
            {
                if (!element.Symbol.IsActive)
                {
                    using (var t = new Transaction(_doc, "Activate Symbol")) { t.Start(); element.Symbol.Activate(); t.Commit(); }
                }
            }

            Transform linkTransform = linkInstance?.GetTotalTransform() ?? Transform.Identity;

            var placedLocationsByType = new Dictionary<ElementId, HashSet<XYZ>>();
            foreach (var elementToCopy in elementsToCopy)
            {
                var existingInstances = new FilteredElementCollector(_doc)
                    .OfClass(typeof(FamilyInstance))
                    .Where(e => (e as FamilyInstance).Symbol.Id == elementToCopy.Symbol.Id)
                    .Select(e => (e.Location as LocationPoint)?.Point)
                    .Where(p => p != null);
                placedLocationsByType[elementToCopy.Symbol.Id] = new HashSet<XYZ>(existingInstances, new XyzEqualityComparer());
            }

            var hostLevelsByName = new FilteredElementCollector(_doc).OfClass(typeof(Level)).Cast<Level>().ToDictionary(l => l.Name, l => l);

            using (var trans = new Transaction(_doc, "Place Elements by Reference"))
            {
                var options = trans.GetFailureHandlingOptions();
                options.SetFailuresPreprocessor(new DuplicateInstancesPreprocessor());
                trans.SetFailureHandlingOptions(options);
                trans.Start();

                foreach (var targetElement in targetElements)
                {
                    if (targetElement.Id == templateReferenceElement.Id)
                    {
                        processedCount += elementsToCopy.Count;
                        continue;
                    }

                    Transform targetTransform = CreateTransformFromInstance(targetElement, linkInstance);
                    if (targetTransform == null) { skippedCount += elementsToCopy.Count; continue; }
                    Transform targetTransformInHost = linkTransform.Multiply(targetTransform);

                    for (int i = 0; i < elementsToCopy.Count; i++)
                    {
                        processedCount++;
                        progressBar.UpdateProgress(processedCount, totalOperations);
                        progressBar.UpdateRoomStatus($"Processing target: {targetElement.Name}", processedCount, totalOperations);

                        if (progressBar.IsCancellationPending)
                        {
                            trans.RollBack();
                            progressBar.Close();
                            throw new OperationCanceledException();
                        }

                        var elementToCopy = elementsToCopy[i];
                        var relativeTransform = relativeTransforms[i];
                        var elementSymbol = elementToCopy.Symbol;
                        bool isHosted = elementToCopy.Host != null;

                        Transform newElementTransform = targetTransformInHost.Multiply(relativeTransform);
                        XYZ newElementLocation = newElementTransform.Origin;

                        if (placedLocationsByType[elementSymbol.Id].Contains(newElementLocation)) { skippedCount++; continue; }

                        Document sourceDoc = linkInstance?.GetLinkDocument() ?? _doc;
                        Level sourceLevel = sourceDoc.GetElement(targetElement.LevelId) as Level;
                        if (sourceLevel == null || !hostLevelsByName.TryGetValue(sourceLevel.Name, out Level targetHostLevel))
                        {
                            skippedCount++;
                            continue;
                        }

                        Element hostElement = isHosted ? GetHostElementAtPoint(_doc, newElementLocation) : null;
                        FamilyInstance newElement = _doc.Create.NewFamilyInstance(newElementLocation, elementSymbol, hostElement, targetHostLevel, StructuralType.NonStructural);

                        if (newElement == null) { skippedCount++; continue; }

                        placedLocationsByType[elementSymbol.Id].Add(newElementLocation);

                        if (newElement.Location is LocationPoint)
                        {
                            double angle = XYZ.BasisX.AngleOnPlaneTo(newElementTransform.BasisX, XYZ.BasisZ);
                            if (Math.Abs(angle) > 1e-9)
                            {
                                var axis = Line.CreateBound(newElementLocation, newElementLocation + XYZ.BasisZ);
                                ElementTransformUtils.RotateElement(_doc, newElement.Id, axis, angle);
                            }
                        }

                        // Update the reference parameters for the new element
                        UpdateReferenceParameters(newElement, targetElement, missingParamsFamilies);

                        placedCount++;
                    }
                }
                trans.Commit();
            }

            progressBar.Close();
            return (placedCount, skippedCount);
        }

        /// <summary>
        /// Writes the reference element's information to the shared parameters of the target element.
        /// </summary>
        private void UpdateReferenceParameters(FamilyInstance elementToUpdate, Element referenceElement, ICollection<string> missingParamsFamilies)
        {
            var familyName = (elementToUpdate.Symbol.FamilyName);
            if (missingParamsFamilies.Contains(familyName)) return;

            var paramType = elementToUpdate.get_Parameter(_refTypeGuid);
            var paramId = elementToUpdate.get_Parameter(_refIdGuid);

            if (paramType != null && !paramType.IsReadOnly && paramId != null && !paramId.IsReadOnly)
            {
                paramType.Set(referenceElement.Name);
                paramId.Set(referenceElement.UniqueId);
            }
            else
            {
                // Add the family name to the set so we don't check it again and only warn once.
                missingParamsFamilies.Add(familyName);
            }
        }
    }

    /// <summary>
    /// A selection filter to allow the user to select only FamilyInstance elements.
    /// </summary>
    public class FamilyInstanceSelectionFilter : ISelectionFilter
    {
        public bool AllowElement(Element elem) => elem is FamilyInstance;
        public bool AllowReference(Reference reference, XYZ position) => true;
    }

    /// <summary>
    /// Custom comparer for XYZ points with a tolerance.
    /// </summary>
    public class XyzEqualityComparer : IEqualityComparer<XYZ>
    {
        private readonly double _tolerance = 1e-9;
        private readonly int _multiplier = 10000;

        public bool Equals(XYZ p1, XYZ p2)
        {
            if (ReferenceEquals(p1, p2)) return true;
            if (p1 is null || p2 is null) return false;
            return p1.IsAlmostEqualTo(p2, _tolerance);
        }

        public int GetHashCode(XYZ p)
        {
            if (p is null) return 0;
            int hashX = ((int)(p.X * _multiplier)).GetHashCode();
            int hashY = ((int)(p.Y * _multiplier)).GetHashCode();
            int hashZ = ((int)(p.Z * _multiplier)).GetHashCode();
            unchecked
            {
                int hash = 17;
                hash = hash * 23 + hashX;
                hash = hash * 23 + hashY;
                hash = hash * 23 + hashZ;
                return hash;
            }
        }
    }

    /// <summary>
    /// Preprocessor to automatically dismiss "identical instances" warnings.
    /// </summary>
    public class DuplicateInstancesPreprocessor : IFailuresPreprocessor
    {
        public FailureProcessingResult PreprocessFailures(FailuresAccessor failuresAccessor)
        {
            var failures = failuresAccessor.GetFailureMessages();
            foreach (var failure in failures)
            {
                if (failure.GetFailureDefinitionId() == BuiltInFailures.OverlapFailures.DuplicateInstances)
                {
                    failuresAccessor.DeleteWarning(failure);
                }
            }
            return FailureProcessingResult.Continue;
        }
    }
}
