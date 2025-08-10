// File: RT_TrayConduits.cs

// Application: Revit 2022 External Command

// Description: This script identifies cable tray elements that contain cable data

//              (i.e., have values in RTS_Cable_XX shared parameters).

//              For each non-empty RTS_Cable_XX value, it models a 25mm diameter conduit

//              along the cable tray, evenly spaced from the right edge.

//              All created conduits are placed in the currently active workset.

//              The associated cable value is written to the 'RTS_ID' parameter of the conduit.

//              The script removes existing conduits identified by the 'RTS_ID' parameter

//              at the beginning of its execution.

//              A report listing successfully created conduits and their RTS_ID is provided.



using System;

using System.Collections.Generic;

using System.Linq;

using Autodesk.Revit.Attributes;

using Autodesk.Revit.DB;

using Autodesk.Revit.DB.Electrical; // Required for CableTray and Conduit classes

using Autodesk.Revit.UI;



namespace RTS.Commands.MEPTools.CableTray

{

    [Transaction(TransactionMode.Manual)]

    [Regeneration(RegenerationOption.Manual)]

    public class RT_TrayConduitsClass : IExternalCommand

    {

        // Define GUIDs for the shared parameters

        // These GUIDs must match those defined in your Revit shared parameters file

        private static readonly Guid RTS_ID_GUID = new Guid("3175a27e-d386-4567-bf10-2da1a9cbb73b");



        private static readonly List<Guid> RTS_Cable_GUIDs = new List<Guid>

    {

      new Guid("cf0d478e-1e98-4e83-ab80-6ee867f61798"), // RTS_Cable_01

            new Guid("2551d308-44ed-405c-8aad-fb78624d086e"), // RTS_Cable_02

            new Guid("c1dfc402-2101-4e53-8f52-f6af64584a9f"), // RTS_Cable_03

            new Guid("f297daa6-a9e0-4dd5-bda3-c628db7c28bd"), // RTS_Cable_04

            new Guid("b0ef396d-6ec0-4ab7-b7cc-9318e9e9b3ab"), // RTS_Cable_05

            new Guid("7c08095a-a3b2-4b78-ba15-dde09a7bc3a9"), // RTS_Cable_06

            new Guid("9bc78bce-0d39-4538-b507-7b98e8a13404"), // RTS_Cable_07

            new Guid("e9d50153-a0e9-4685-bc92-d89f244f7e8e"), // RTS_Cable_08

            new Guid("5713d65a-91df-4d2e-97bf-1c3a10ea5225"), // RTS_Cable_09

            new Guid("64af3105-b2fd-44bc-9ad3-17264049ff62"), // RTS_Cable_10

            new Guid("f3626002-0e62-4b75-93cc-35d0b11dfd67"), // RTS_Cable_11

            new Guid("63dc0a2e-0770-4002-a859-a9d40a2ce023"), // RTS_Cable_12

            new Guid("eb7c4b98-d676-4e2b-a408-e3578b2c0ef2"), // RTS_Cable_13

            new Guid("0e0572e5-c568-42b7-8730-a97433bd9b54"), // RTS_Cable_14

            new Guid("bf9cd3e8-e38f-4250-9daa-c0fc67eca10f"), // RTS_Cable_15

            new Guid("f6d2af67-027e-4b9c-9def-336ebaa87336"), // RTS_Cable_16

            new Guid("f6a4459d-46a1-44c0-8545-ee44e4778854"), // RTS_Cable_17

            new Guid("0d66d2fa-f261-4daa-8041-9eadeefac49a"), // RTS_Cable_18

            new Guid("af483914-c8d2-4ce6-be6e-ab81661e5bf1"), // RTS_Cable_19

            new Guid("c8d2d2fc-c248-483f-8d52-e630eb730cd7"), // RTS_Cable_20

            new Guid("aa41bc4a-e3e7-45b0-81fa-74d3e71ca506"), // RTS_Cable_21

            new Guid("6cffdb25-8270-4b34-8bb4-cf5d0a224dc2"), // RTS_Cable_22

            new Guid("7fdaad3a-454e-47f3-8189-7eda9cb9f6a2"), // RTS_Cable_23

            new Guid("7f745b2b-a537-42d9-8838-7a5521cc7d0c"), // RTS_Cable_24

            new Guid("9a76c2dc-1022-4a54-ab66-5ca625b50365"), // RTS_Cable_25

            new Guid("658e39c4-bbac-4e2e-b649-2f2f5dd05b5e"), // RTS_Cable_26

            new Guid("8ad24640-036b-44d2-af9c-b891f6e64271"), // RTS_Cable_27

            new Guid("c046c4d7-e1fd-4cf7-a99f-14ae96b722be"), // RTS_Cable_28

            new Guid("cdf00587-7e11-4af4-8e54-48586481cf22"), // RTS_Cable_29

            new Guid("a92bb0f9-2781-4971-a3b1-9c47d62b947b")  // RTS_Cable_30

        };



        // Inner class to hold conduit creation data for reporting

        private class CreatedConduitInfo

        {

            public ElementId ConduitId { get; set; }

            public string CableValue { get; set; }

            public string TrayId { get; set; }

            public string TrayLocation { get; set; }

        }



        public Result Execute(

          ExternalCommandData commandData,

          ref string message,

          ElementSet elements)

        {

            UIDocument uidoc = commandData.Application.ActiveUIDocument;

            Document doc = uidoc.Document;



            List<CreatedConduitInfo> successfulConduits = new List<CreatedConduitInfo>();

            string reportMessage = "";



            // Desired conduit diameter in millimeters

            const double desiredConduitDiameterMM = 25.0;

            // Additional spacing from the tray edge in millimeters

            const double additionalEdgeOffsetMM = 10.0;



            // Convert to Revit's internal units (feet) manually

            // 1 foot = 304.8 millimeters

            double desiredConduitDiameterFeet = desiredConduitDiameterMM / 304.8;

            double conduitRadiusFeet = desiredConduitDiameterFeet / 2.0;

            double additionalEdgeOffsetFeet = additionalEdgeOffsetMM / 304.8;



            // --- 1. Find closest conduit type ---

            ConduitType conduitType = null;

            try

            {

                conduitType = new FilteredElementCollector(doc)

                  .OfClass(typeof(ConduitType))

                  .Cast<ConduitType>()

        // Accessing NominalDiameter via parameter BuiltInParameter.RBS_CONDUIT_DIAMETER_PARAM

                            .OrderBy(ct => Math.Abs((ct.get_Parameter(BuiltInParameter.RBS_CONDUIT_DIAMETER_PARAM)?.AsDouble() ?? 0.0) - desiredConduitDiameterFeet))

                  .FirstOrDefault();



                if (conduitType == null)

                {

                    message = "No Conduit Types found in the project. Please load one and try again.";

                    TaskDialog.Show("RT_TrayConduits Error", message);

                    return Result.Failed;

                }

            }

            catch (Exception ex)

            {

                message = $"Error finding conduit type: {ex.Message}";

                TaskDialog.Show("RT_TrayConduits Error", message);

                return Result.Failed;

            }



            // --- 2. Remove existing conduits created by this script ---

            using (Transaction t_delete = new Transaction(doc, "Delete Existing Conduits (RTS_ID)"))

            {

                try

                {

                    t_delete.Start();

                    // Identify conduits previously created by this script by checking for RTS_ID_GUID parameter

                    var conduitsToDelete = new FilteredElementCollector(doc)

            .OfCategory(BuiltInCategory.OST_Conduit)

            .WhereElementIsNotElementType()

            .Where(e => e.get_Parameter(RTS_ID_GUID) != null && !string.IsNullOrEmpty(e.get_Parameter(RTS_ID_GUID).AsString()))

            .ToList();



                    if (conduitsToDelete.Any())

                    {

                        doc.Delete(conduitsToDelete.Select(e => e.Id).ToList());

                        System.Diagnostics.Debug.WriteLine($"Deleted {conduitsToDelete.Count} existing conduits identified by 'RTS_ID' parameter.");

                    }

                    t_delete.Commit();

                }

                catch (Exception ex)

                {

                    t_delete.RollBack();

                    message = $"Error deleting existing conduits: {ex.Message}";

                    TaskDialog.Show("RT_TrayConduits Error", message);

                    return Result.Failed;

                }

            }



            // --- 3. Process Cable Trays and Create Conduits ---

            using (TransactionGroup tg = new TransactionGroup(doc, "Create Conduits from Cable Trays"))

            {

                tg.Start(); // Start the transaction group



                try

                {

                    FilteredElementCollector trayCollector = new FilteredElementCollector(doc);

                    List<Element> cableTraysWithCableData = trayCollector

                      .OfCategory(BuiltInCategory.OST_CableTray)

                      .WhereElementIsNotElementType()

                      .Where(elem => RTS_Cable_GUIDs.Any(guid =>

                      {

                          Parameter p = elem.get_Parameter(guid);

                          return p != null && !string.IsNullOrEmpty(p.AsString());

                      }))

                      .ToList();



                    if (!cableTraysWithCableData.Any())

                    {

                        TaskDialog.Show("RT_TrayConduits", "No cable trays with RTS_Cable_XX data found.");

                        tg.RollBack();

                        return Result.Succeeded;

                    }



                    foreach (Element elem in cableTraysWithCableData)

                    {

                        LocationCurve locationCurve = elem.Location as LocationCurve;

                        if (locationCurve == null)

                        {

                            System.Diagnostics.Debug.WriteLine($"Warning: Cable tray {elem.Id} has no valid LocationCurve. Skipping.");

                            continue;

                        }



                        Curve trayCurve = locationCurve.Curve;

                        XYZ startPoint = trayCurve.GetEndPoint(0);

                        XYZ endPoint = trayCurve.GetEndPoint(1);



                        // Get the width of the cable tray

                        Parameter widthParam = elem.get_Parameter(BuiltInParameter.RBS_CABLETRAY_WIDTH_PARAM);

                        double trayWidthFeet = widthParam != null && widthParam.HasValue ? widthParam.AsDouble() : 0.0;



                        if (trayWidthFeet <= 0)

                        {

                            System.Diagnostics.Debug.WriteLine($"Warning: Cable tray {elem.Id} has invalid or zero width. Skipping conduit creation for this tray.");

                            continue;

                        }



                        // Determine the tray's normal vector in the XY plane for offsetting

                        XYZ trayDirection = (endPoint - startPoint).Normalize();

                        // Fully qualify XYZ.BasisZ

                        XYZ trayNormal = XYZ.BasisZ.CrossProduct(trayDirection).Normalize();



                        // Iterate through RTS_Cable_GUIDs to create a conduit for each non-empty value

                        int conduitCountOnTray = 0;

                        foreach (Guid cableParamGuid in RTS_Cable_GUIDs)

                        {

                            Parameter cableParam = elem.get_Parameter(cableParamGuid);

                            if (cableParam != null && !string.IsNullOrEmpty(cableParam.AsString()))

                            {

                                string cableValue = cableParam.AsString();



                                using (Transaction t_create_conduit = new Transaction(doc, $"Create Conduit for {elem.Id} - {cableValue}"))

                                {

                                    try

                                    {

                                        t_create_conduit.Start();



                                        // Calculate current offset from the tray's centerline

                                        // The first conduit's center is (trayWidth / 2) - conduitRadius - additionalEdgeOffset from the center

                                        // Subsequent conduits move inwards by 2 * conduitDiameter

                                        double currentOffsetFromCenterline = trayWidthFeet / 2.0 - conduitRadiusFeet - additionalEdgeOffsetFeet - conduitCountOnTray * 2 * desiredConduitDiameterFeet;



                                        // Ensure we don't place conduits off the left edge.

                                        // A simple check: if the conduit's leftmost point would be past the tray's leftmost point.

                                        // Tray's left edge from centerline: -(trayWidthFeet / 2.0)

                                        // Conduit's leftmost point from centerline: currentOffsetFromCenterline - conduitRadiusFeet

                                        if (currentOffsetFromCenterline - conduitRadiusFeet < -(trayWidthFeet / 2.0))

                                        {

                                            System.Diagnostics.Debug.WriteLine($"Warning: Would place conduit for tray {elem.Id} ({cableValue}) outside tray boundaries on the left. Skipping.");

                                            t_create_conduit.RollBack(); // Rollback individual conduit creation

                                            continue;

                                        }



                                        // Calculate the start and end points of the conduit line

                                        // Offset direction is -trayNormal (to move towards the right side of the tray)

                                        XYZ offsetVector = trayNormal.Negate().Multiply(currentOffsetFromCenterline);



                                        XYZ conduitStart = startPoint + offsetVector;

                                        XYZ conduitEnd = endPoint + offsetVector;



                                        // Create the conduit. It will automatically be placed in the active workset.

                                        Conduit newConduit = Conduit.Create(doc, conduitType.Id, conduitStart, conduitEnd, elem.LevelId);



                                        if (newConduit != null)

                                        {

                                            // Write associated cable value to RTS_ID parameter of the conduit

                                            Parameter rtsIdConduitParam = newConduit.get_Parameter(RTS_ID_GUID);

                                            if (rtsIdConduitParam != null && rtsIdConduitParam.IsReadOnly == false)

                                            {

                                                rtsIdConduitParam.Set(cableValue);

                                            }

                                            else

                                            {

                                                System.Diagnostics.Debug.WriteLine($"Warning: Conduit {newConduit.Id} does not have the 'RTS_ID' shared parameter or it's read-only.");

                                            }



                                            successfulConduits.Add(new CreatedConduitInfo

                                            {

                                                ConduitId = newConduit.Id,

                                                CableValue = cableValue,

                                                TrayId = elem.Id.ToString(),

                                                TrayLocation = $"Start: {startPoint.X:F2},{startPoint.Y:F2},{startPoint.Z:F2} - End: {endPoint.X:F2},{endPoint.Y:F2},{endPoint.Z:F2}"

                                            });



                                            conduitCountOnTray++;

                                        }

                                        else

                                        {

                                            System.Diagnostics.Debug.WriteLine($"Error: Failed to create conduit for cable tray {elem.Id} with cable value {cableValue}.");

                                            t_create_conduit.RollBack(); // Rollback individual conduit creation

                                            continue;

                                        }



                                        t_create_conduit.Commit(); // Commit individual conduit creation

                                    }

                                    catch (Exception ex)

                                    {

                                        t_create_conduit.RollBack();

                                        System.Diagnostics.Debug.WriteLine($"Error creating conduit for cable tray {elem.Id}, cable value '{cableValue}': {ex.Message}");

                                    }

                                }

                            }

                        }

                    }



                    tg.Assimilate(); // Commit the entire transaction group

                }

                catch (Exception ex)

                {

                    tg.RollBack(); // Rollback the entire transaction group on any error

                    message = $"Error creating conduits: {ex.Message}";

                    TaskDialog.Show("RT_TrayConduits Error", message);

                    return Result.Failed;

                }

            }



            // --- 4. Report Results ---

            if (successfulConduits.Any())

            {

                System.Text.StringBuilder sb = new System.Text.StringBuilder();

                sb.AppendLine("Successfully Created Conduits:");

                sb.AppendLine("--------------------------------");

                foreach (var info in successfulConduits)

                {

                    sb.AppendLine($"Conduit ID: {info.ConduitId}, RTS_ID: '{info.CableValue}'");

                    sb.AppendLine($"  (Associated with Tray ID: {info.TrayId}, Location: {info.TrayLocation})");

                }

                reportMessage = sb.ToString();

                TaskDialog.Show("RT_TrayConduits Report", reportMessage);

            }

            else

            {

                TaskDialog.Show("RT_TrayConduits", "No conduits were created. Check warnings for details.");

            }



            return Result.Succeeded;

        }

    }

}