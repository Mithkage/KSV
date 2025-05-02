using System;
using System.Collections.Generic;
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;

namespace RTS_Project_Initiate
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class RTS_InitiateProject : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            Application app = uiapp.Application;
            Document doc = uiapp.ActiveUIDocument.Document;

            // 1. Define Shared Parameter Information
            Guid sharedParameterGuid = new Guid("c7e96b3f-aee1-4ecb-b287-6c3a57548acb"); // Your GUID
            string parameterName = "RTS_Type"; // Your Parameter Name
            string parameterDescription = "Cleaned Type name across all elements (unique code for key value)";
            
            // 2. Get the Detail Items Category
            Category detailItemCategory = doc.Settings.Categories.get_Item(BuiltInCategory.OST_DetailComponents);
            if (detailItemCategory == null)
            {
                message = "Failed to retrieve the Detail Items category.";
                return Result.Failed;
            }
            
            // 3. Create a CategorySet and add the Detail Items category to it.
            CategorySet categorySet = app.Create.NewCategorySet();
            categorySet.Insert(detailItemCategory);
            
            // 4. Get the Shared Parameter Element
             DefinitionFile sharedParameterFile = app.OpenSharedParameterFile();
            if (sharedParameterFile == null)
            {
                message = "Could not open shared parameter file.  Ensure a shared parameter file is set in Revit.";
                return Result.Failed;
            }

            DefinitionGroup targetGroup = null;
            foreach (DefinitionGroup group in sharedParameterFile.Groups)
            {
                if (group.Name == "Your Parameter Group Name") // <<== CHANGE THIS TO YOUR GROUP NAME
                {
                    targetGroup = group;
                    break;
                }
            }
            if (targetGroup == null)
            {
                 message = "Could not find the specified parameter group in the shared parameter file.  Ensure the group exists.";
                return Result.Failed;
            }

            Definition sharedParameterDefinition = targetGroup.Definitions.get_Item(parameterName);
            if (sharedParameterDefinition == null)
            {
                message = "Could not find the specified shared parameter definition.";
                return Result.Failed;
            }
            
            // 5. Add the Project Parameter
            using (Transaction t = new Transaction(doc, "Add Project Parameter"))
            {
                t.Start();
                try //fix this!
                {
                    //App.NewProjectParameter takes the catSet
                    //doc.ParameterBindings.Insert(sharedParameterDefinition, categorySet, ParameterBindingOptions.Instance); // Or .Type/Instance, depending on your needs 
                    //t.Commit();
                }
                catch (Exception ex)
                {
                    t.RollBack();
                    message = "Error adding project parameter: " + ex.Message;
                    return Result.Failed;
                }
            }

            return Result.Succeeded;
        }
    }
}
