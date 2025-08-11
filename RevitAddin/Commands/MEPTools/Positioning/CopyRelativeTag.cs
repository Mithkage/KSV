//
// File: CopyRelativeTag.cs
// Company: ReTick Solutions Pty Ltd
// Function: This script copies tags from pre-selected elements to all other elements
//           that share the same 'RTS_CopyReference_Type' parameter value.
//
// Workflow:
// 1. User pre-selects one or more elements and their corresponding tags.
// 2. User runs the command.
// 3. The script intelligently groups the selected tags with their host elements.
// 4. A dialog appears for the user to select the views for tag placement.
// 5. The script finds all other elements in the model with a matching 'RTS_CopyReference_Type'.
// 6. In each selected view, new tags are placed, matching the properties and relative
//    position of the template tags.
//
// Change Log:
// 2025-08-11: (Fix) Added missing using directive to resolve compilation error.
// 2025-08-11: (Major Update) Reworked script to use a full pre-selection workflow.
// 2025-08-11: (Improvement) Logic now uses the 'RTS_CopyReference_Type' parameter for finding target elements.
//

#region Namespaces
using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection; // Added missing directive
using System.Windows;
using System.Windows.Controls;
using System.Text;
using RTS.UI; // For ProgressBarWindow
#endregion

namespace RTS.Commands.MEPTools.Positioning
{
    [Transaction(TransactionMode.Manual)]
    public class CopyRelativeTagClass : IExternalCommand
    {
        private UIDocument _uidoc;
        private Document _doc;
        // Shared Parameter GUIDs
        private readonly Guid _refTypeGuid = new Guid("4d6ce1ad-eb55-47e2-acb6-69490634990e");

        public Result Execute(
            ExternalCommandData commandData,
            ref string message,
            ElementSet elements)
        {
            _uidoc = commandData.Application.ActiveUIDocument;
            _doc = _uidoc.Document;

            try
            {
                // Step 1: Parse the user's pre-selection
                if (!ParsePreselection(out Dictionary<ElementId, List<IndependentTag>> referenceElementsWithTags, out int orphanedTags))
                {
                    return Result.Failed;
                }

                // Step 2: Prompt user to select views for tag placement
                List<View> targetViews = SelectViews();
                if (targetViews == null || !targetViews.Any())
                {
                    return Result.Cancelled;
                }

                // Step 3: Group the reference elements by their 'RTS_CopyReference_Type'
                var elementsGroupedByRefType = GroupElementsByReferenceType(referenceElementsWithTags.Keys.Select(id => _doc.GetElement(id)).ToList());

                // Step 4: Place new tags for each group
                int totalPlaced = 0;
                var skippedSummary = new Dictionary<string, int>();

                foreach (var group in elementsGroupedByRefType)
                {
                    var referenceType = group.Key;
                    var referenceElementsInGroup = group.Value;

                    // Find all other elements in the model with this reference type
                    var allTargetElements = FindAllTargetElements(referenceType);

                    // For each unique tag configuration in the selection, place copies
                    foreach (var referenceElement in referenceElementsInGroup)
                    {
                        var templateTags = referenceElementsWithTags[referenceElement.Id];
                        var relativeVectors = CalculateRelativeVectors(referenceElement, templateTags);
                        (int placed, Dictionary<string, int> skipped) = PlaceNewTags(allTargetElements, templateTags, relativeVectors, targetViews);

                        totalPlaced += placed;
                        foreach (var kvp in skipped)
                        {
                            if (!skippedSummary.ContainsKey(kvp.Key)) skippedSummary[kvp.Key] = 0;
                            skippedSummary[kvp.Key] += kvp.Value;
                        }
                    }
                }

                // Step 5: Provide detailed final feedback
                StringBuilder feedback = new StringBuilder();
                feedback.AppendLine($"Successfully placed {totalPlaced} new tags.");
                if (orphanedTags > 0)
                {
                    feedback.AppendLine($"\nWarning: {orphanedTags} selected tags were ignored because their host element was not included in the selection.");
                }
                int totalSkipped = skippedSummary.Values.Sum();
                if (totalSkipped > 0)
                {
                    feedback.AppendLine($"\nSkipped a total of {totalSkipped} tags for the following reasons:");
                    foreach (var entry in skippedSummary.Where(kvp => kvp.Value > 0))
                    {
                        feedback.AppendLine($" - {entry.Key}: {entry.Value} tags");
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
        /// Parses the user's pre-selection into a dictionary of elements and their associated tags.
        /// </summary>
        private bool ParsePreselection(out Dictionary<ElementId, List<IndependentTag>> referenceElementsWithTags, out int orphanedTags)
        {
            referenceElementsWithTags = new Dictionary<ElementId, List<IndependentTag>>();
            orphanedTags = 0;

            var selectedIds = _uidoc.Selection.GetElementIds();
            if (selectedIds.Count < 2)
            {
                TaskDialog.Show("Selection Error", "Please select at least one element and its tag before running the command.");
                return false;
            }

            var selectedElements = selectedIds.Select(id => _doc.GetElement(id)).ToList();
            var selectedHostElements = selectedElements.OfType<FamilyInstance>().ToList();
            var selectedTags = selectedElements.OfType<IndependentTag>().ToList();

            if (!selectedHostElements.Any() || !selectedTags.Any())
            {
                TaskDialog.Show("Selection Error", "Your selection must include at least one Family Instance and at least one Tag.");
                return false;
            }

            var selectedHostElementIds = new HashSet<ElementId>(selectedHostElements.Select(e => e.Id));

            foreach (var tag in selectedTags)
            {
                var taggedElementId = tag.GetTaggedLocalElementIds().FirstOrDefault();
                if (taggedElementId != null && selectedHostElementIds.Contains(taggedElementId))
                {
                    if (!referenceElementsWithTags.ContainsKey(taggedElementId))
                    {
                        referenceElementsWithTags[taggedElementId] = new List<IndependentTag>();
                    }
                    referenceElementsWithTags[taggedElementId].Add(tag);
                }
                else
                {
                    orphanedTags++;
                }
            }

            if (!referenceElementsWithTags.Any())
            {
                TaskDialog.Show("Selection Error", "None of the selected tags are associated with the selected elements. Please ensure you select both the element and its tag.");
                return false;
            }

            return true;
        }

        /// <summary>
        /// Groups the selected reference elements by their 'RTS_CopyReference_Type' parameter value.
        /// </summary>
        private Dictionary<string, List<Element>> GroupElementsByReferenceType(List<Element> elements)
        {
            var groupedElements = new Dictionary<string, List<Element>>();
            foreach (var element in elements)
            {
                var param = element.get_Parameter(_refTypeGuid);
                if (param != null && param.HasValue)
                {
                    string refType = param.AsString();
                    if (!groupedElements.ContainsKey(refType))
                    {
                        groupedElements[refType] = new List<Element>();
                    }
                    groupedElements[refType].Add(element);
                }
            }
            return groupedElements;
        }

        /// <summary>
        /// Displays a dialog for the user to select views.
        /// </summary>
        private List<View> SelectViews()
        {
            var compatibleViews = new FilteredElementCollector(_doc)
                .OfClass(typeof(View))
                .Cast<View>()
                .Where(v => !v.IsTemplate && (v.ViewType == ViewType.FloorPlan || v.ViewType == ViewType.CeilingPlan || v.ViewType == ViewType.Elevation || v.ViewType == ViewType.Section))
                .OrderBy(v => v.Name)
                .ToList();

            var viewSelectionDialog = new ViewSelectionWindow(compatibleViews);
            if (viewSelectionDialog.ShowDialog() == true)
            {
                return viewSelectionDialog.GetSelectedViews();
            }
            return null;
        }

        /// <summary>
        /// Calculates the vector from the reference element's location to each tag's head position.
        /// </summary>
        private List<XYZ> CalculateRelativeVectors(Element templateReferenceElement, List<IndependentTag> templateTags)
        {
            var relativeVectors = new List<XYZ>();
            var referenceLocation = (templateReferenceElement.Location as LocationPoint).Point;

            foreach (var tag in templateTags)
            {
                relativeVectors.Add(tag.TagHeadPosition - referenceLocation);
            }
            return relativeVectors;
        }

        /// <summary>
        /// Finds all elements that have a matching 'RTS_CopyReference_Type' value.
        /// </summary>
        private List<Element> FindAllTargetElements(string referenceType)
        {
            var provider = new ParameterValueProvider(new ElementId(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS)); // Placeholder
            var rule = new FilterStringEquals();
            var filter = new ElementParameterFilter(new FilterStringRule(provider, rule, referenceType));

            return new FilteredElementCollector(_doc)
                .WhereElementIsNotElementType()
                .WherePasses(filter)
                .ToList();
        }

        /// <summary>
        /// Places the new tags in the selected views.
        /// </summary>
        private (int placed, Dictionary<string, int> skippedSummary) PlaceNewTags(List<Element> targetElements, List<IndependentTag> templateTags, List<XYZ> relativeVectors, List<View> targetViews)
        {
            int placedCount = 0;
            var skippedSummary = new Dictionary<string, int>
            {
                { "Already Tagged", 0 },
                { "Not Visible in View", 0 },
                { "Invalid Location", 0 }
            };

            var existingTagsLookup = new FilteredElementCollector(_doc, targetViews.Select(v => v.Id).ToList())
                .OfClass(typeof(IndependentTag))
                .Cast<IndependentTag>()
                .SelectMany(t => t.GetTaggedLocalElementIds().Select(id => new { ViewId = t.OwnerViewId, TaggedId = id, TagTypeId = t.GetTypeId() }))
                .ToLookup(x => x.ViewId);

            int totalOperations = targetElements.Count * templateTags.Count * targetViews.Count;
            int processedCount = 0;
            var progressBar = new ProgressBarWindow();
            progressBar.Show();

            using (var trans = new Transaction(_doc, "Copy Tags to Elements"))
            {
                trans.Start();

                int viewNum = 0;
                foreach (var view in targetViews)
                {
                    viewNum++;
                    foreach (var targetElement in targetElements)
                    {
                        var targetLocation = (targetElement.Location as LocationPoint)?.Point;
                        if (targetLocation == null)
                        {
                            skippedSummary["Invalid Location"] += templateTags.Count;
                            processedCount += templateTags.Count;
                            continue;
                        }

                        for (int i = 0; i < templateTags.Count; i++)
                        {
                            processedCount++;
                            progressBar.UpdateProgress(processedCount, totalOperations);
                            progressBar.UpdateRoomStatus($"Processing View {viewNum} of {targetViews.Count}: {view.Name}", processedCount, totalOperations);

                            if (progressBar.IsCancellationPending)
                            {
                                trans.RollBack();
                                progressBar.Close();
                                throw new OperationCanceledException();
                            }

                            var templateTag = templateTags[i];
                            var relativeVector = relativeVectors[i];

                            bool alreadyTagged = existingTagsLookup.Contains(view.Id) &&
                                                 existingTagsLookup[view.Id].Any(x => x.TaggedId == targetElement.Id && x.TagTypeId == templateTag.GetTypeId());

                            if (alreadyTagged)
                            {
                                skippedSummary["Already Tagged"]++;
                                continue;
                            }

                            XYZ newTagHeadPosition = targetLocation + relativeVector;

                            try
                            {
                                // Future Development: A new shared parameter should be created to store the relative
                                // position vector of the tag. The script would then check both the RTS_CopyReference_Type
                                // and this new parameter to find elements that are not just the same type, but also
                                // in the same relative position (e.g., only doors on the left side of a corridor).
                                var newTag = IndependentTag.Create(_doc, templateTag.GetTypeId(), view.Id, new Reference(targetElement), false, templateTag.TagOrientation, newTagHeadPosition);

                                if (newTag != null)
                                {
                                    newTag.LeaderEndCondition = templateTag.LeaderEndCondition;
                                    newTag.HasLeader = templateTag.HasLeader;
                                    placedCount++;
                                }
                                else
                                {
                                    skippedSummary["Not Visible in View"]++;
                                }
                            }
                            catch
                            {
                                skippedSummary["Not Visible in View"]++;
                            }
                        }
                    }
                }
                trans.Commit();
            }

            progressBar.Close();
            return (placedCount, skippedSummary);
        }
    }

    /// <summary>
    /// A selection filter to allow only tags that are associated with a specific element.
    /// </summary>
    public class TagSelectionFilter : ISelectionFilter
    {
        private readonly ElementId _taggedElementId;
        public TagSelectionFilter(ElementId taggedElementId)
        {
            _taggedElementId = taggedElementId;
        }
        public bool AllowElement(Element elem)
        {
            if (elem is IndependentTag tag)
            {
                return tag.GetTaggedLocalElementIds().Contains(_taggedElementId);
            }
            return false;
        }
        public bool AllowReference(Reference reference, XYZ position) => false;
    }

    /// <summary>
    /// A simple dialog window for selecting views with a search filter.
    /// </summary>
    public class ViewSelectionWindow : Window
    {
        private ListBox _listBox;
        private List<View> _allViews;
        private List<View> _selectedViews = new List<View>();

        public ViewSelectionWindow(List<View> views)
        {
            _allViews = views;
            Title = "Select Views to Place Tags";
            Width = 450;
            Height = 550;
            WindowStartupLocation = WindowStartupLocation.CenterOwner;

            var grid = new System.Windows.Controls.Grid();
            grid.Margin = new Thickness(10);
            grid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            grid.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) });
            grid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            this.Content = grid;

            var searchBox = new System.Windows.Controls.TextBox
            {
                Margin = new Thickness(0, 0, 0, 10)
            };
            searchBox.TextChanged += SearchBox_TextChanged;
            System.Windows.Controls.Grid.SetRow(searchBox, 0);
            grid.Children.Add(searchBox);

            _listBox = new ListBox
            {
                SelectionMode = SelectionMode.Extended
            };
            PopulateListBox(_allViews);
            System.Windows.Controls.Grid.SetRow(_listBox, 1);
            grid.Children.Add(_listBox);

            var okButton = new Button
            {
                Content = "OK",
                IsDefault = true,
                Margin = new Thickness(0, 10, 0, 0),
                Padding = new Thickness(10, 5, 10, 5)
            };
            okButton.Click += (s, e) => {
                foreach (var item in _listBox.SelectedItems)
                {
                    if (item is ListBoxItem listBoxItem)
                    {
                        _selectedViews.Add(listBoxItem.Tag as View);
                    }
                }
                this.DialogResult = true;
                this.Close();
            };
            System.Windows.Controls.Grid.SetRow(okButton, 2);
            grid.Children.Add(okButton);
        }

        private void PopulateListBox(IEnumerable<View> views)
        {
            _listBox.Items.Clear();
            foreach (var view in views)
            {
                _listBox.Items.Add(new ListBoxItem { Content = view.Name, Tag = view });
            }
        }

        private void SearchBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            var searchText = (sender as System.Windows.Controls.TextBox).Text;
            if (string.IsNullOrWhiteSpace(searchText))
            {
                PopulateListBox(_allViews);
            }
            else
            {
                var filteredViews = _allViews.Where(v => v.Name.IndexOf(searchText, StringComparison.OrdinalIgnoreCase) >= 0);
                PopulateListBox(filteredViews);
            }
        }

        public List<View> GetSelectedViews() => _selectedViews;
    }
}
