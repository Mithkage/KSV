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
// 2025-08-12: (Fix) Corrected shared parameter lookup to check category bindings, not the instance.
// 2025-08-12: (Improvement) Enhanced warning message to specify which parameter is missing.
// 2025-08-12: (Refactor) Updated to use the new SetParameterValue methods from the SharedParameters utility.
// 2025-08-11: (Fix) Reverted to a pre-selection workflow for both Revit 2022 and 2024.
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
using RTS.Utilities; // For the SharedParameters utility
using System.Globalization; // For number formatting

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
            var missingParamsMessages = new HashSet<string>();
            var overwrittenElements = new List<string>();

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

                // Step 3: Ensure shared parameters exist and are bound to the correct categories
                SharedParameters.AddMyParametersToProject(_doc);

                // Step 4: Calculate the geometric relationships
                if (!CalculateGeometricRelationships(templateReferenceElement, elementsToCopy, linkInstance, out List<Transform> relativeTransforms))
                {
                    return Result.Failed;
                }

                // Step 5: Update parameters on the original template elements
                using (var trans = new Transaction(_doc, "Update Template Reference Parameters"))
                {
                    trans.Start();
                    for (int i = 0; i < elementsToCopy.Count; i++)
                    {
                        UpdateReferenceParameters(elementsToCopy[i], templateReferenceElement, relativeTransforms[i], missingParamsMessages, overwrittenElements);
                    }
                    trans.Commit();
                }

                // Step 6: Find all target elements
                List<FamilyInstance> targetElements = FindAllTargetElements(templateReferenceElement, linkInstance);

                // Step 7: Place new elements
                (int placedCount, int skippedCount) = PlaceNewElements(targetElements, elementsToCopy, templateReferenceElement, linkInstance, relativeTransforms, missingParamsMessages, overwrittenElements);

                // Step 8: Provide final feedback
                StringBuilder feedback = new StringBuilder();
                feedback.AppendLine($"Successfully placed {placedCount} new elements.");
                if (skippedCount > 0)
                {
                    feedback.AppendLine($"Skipped {skippedCount} locations where an element of the same type already existed or the target geometry was invalid.");
                }
                if (overwrittenElements.Any())
                {
                    feedback.AppendLine($"\nWarning: Overwrote the 'RTS_CopyRelative_Position' parameter for {overwrittenElements.Count} element(s), including the original templates.");
                }
                if (missingParamsMessages.Any())
                {
                    feedback.AppendLine("\nWarning: Could not update traceability parameters for the following families:");
                    foreach (var msg in missingParamsMessages)
                    {
                        feedback.AppendLine(msg);
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
        /// Gets the pre-selected reference element and validates it. Handles both host and linked element selections.
        /// Uses different logic for Revit 2022 due to API limitations.
        /// </summary>
        private bool GetPreselectedReference(out FamilyInstance templateReferenceElement, out RevitLinkInstance linkInstance)
        {
            templateReferenceElement = null;
            linkInstance = null;

#if REVIT2024_OR_GREATER
            // For Revit 2024+, we can reliably get references from a pre-selection.
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
            // For Revit 2022, we must handle host and linked selections differently.
            var selectedIds = _uidoc.Selection.GetElementIds();
            if (selectedIds.Count == 1)
            {
                var selectedElement = _doc.GetElement(selectedIds.First());
                if (selectedElement is FamilyInstance hostInstance)
                {
                    // Pre-selection of a host element is successful.
                    templateReferenceElement = hostInstance;
                    return true;
                }
            }
            
            // If we're here, either nothing was selected, or a linked element was selected.
            // We must prompt the user to get a valid reference for a linked element.
            try
            {
                var referenceFilter = new FamilyInstanceSelectionFilter();
                Reference reference = _uidoc.Selection.PickObject(ObjectType.LinkedElement, referenceFilter, "Please select a reference element in a LINKED model.");
                
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
#endif

            if (templateReferenceElement == null)
            {
                TaskDialog.Show("Selection Error", "The selected element is not a valid Family Instance.");
                return false;
            }
            return true;
        }

        /// <summary>
        /// Updates the traceability parameters on an element.
        /// Uses the SetParameterValue method from the SharedParameters utility.
        /// </summary>
        private void UpdateReferenceParameters(FamilyInstance elementToUpdate, Element referenceElement, Transform relativeTransform, ICollection<string> missingParamsMessages, ICollection<string> overwrittenElements)
        {
            var familyName = elementToUpdate.Symbol.FamilyName;

            // Set the reference type and ID using the helper method.
            if (!SharedParameters.SetParameterValue(elementToUpdate, SharedParameters.General.RTS_CopyReference_Type, referenceElement.Name))
            {
                missingParamsMessages.Add($" - {familyName}: Missing or read-only 'RTS_CopyReference_Type' parameter.");
                return; // Exit if a critical parameter is missing.
            }

            if (!SharedParameters.SetParameterValue(elementToUpdate, SharedParameters.General.RTS_CopyReference_ID, referenceElement.UniqueId))
            {
                missingParamsMessages.Add($" - {familyName}: Missing or read-only 'RTS_CopyReference_ID' parameter.");
                return;
            }

            // Calculate the position string.
            var origin = relativeTransform.Origin;
            double angle = Math.Atan2(relativeTransform.BasisY.X, relativeTransform.BasisX.X) * (180.0 / Math.PI);
            string positionString = $"{origin.X.ToString("F6", CultureInfo.InvariantCulture)},{origin.Y.ToString("F6", CultureInfo.InvariantCulture)},{origin.Z.ToString("F6", CultureInfo.InvariantCulture)},{angle.ToString("F6", CultureInfo.InvariantCulture)}";

            // Check if the position parameter already has a value before overwriting.
            var paramPos = elementToUpdate.get_Parameter(SharedParameters.General.RTS_CopyRelative_Position);
            if (paramPos != null && paramPos.HasValue && !string.IsNullOrEmpty(paramPos.AsString()))
            {
                overwrittenElements.Add(elementToUpdate.UniqueId);
            }

            // Set the position parameter.
            if (!SharedParameters.SetParameterValue(elementToUpdate, SharedParameters.General.RTS_CopyRelative_Position, positionString))
            {
                missingParamsMessages.Add($" - {familyName}: Missing or read-only 'RTS_CopyRelative_Position' parameter.");
            }
        }

        #region Unchanged Methods
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

        private (int placed, int skipped) PlaceNewElements(
            List<FamilyInstance> targetElements,
            List<FamilyInstance> elementsToCopy,
            FamilyInstance templateReferenceElement,
            RevitLinkInstance linkInstance,
            List<Transform> relativeTransforms,
            HashSet<string> missingParamsMessages,
            List<string> overwrittenElements)
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

                        UpdateReferenceParameters(newElement, targetElement, relativeTransform, missingParamsMessages, overwrittenElements);

                        placedCount++;
                    }
                }
                trans.Commit();
            }

            progressBar.Close();
            return (placedCount, skippedCount);
        }
        #endregion
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
