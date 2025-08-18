using Autodesk.Revit.DB;
using System.Collections.Generic;
using System.Linq;

namespace RTS.Reports.Generators.Routing
{
    /// <summary>
    /// A helper class to collect specific categories of elements from the Revit document.
    /// </summary>
    public class ContainmentCollector
    {
        private readonly Document _doc;
        public ContainmentCollector(Document doc) { _doc = doc; }

        /// <summary>
        /// Gets all cable tray, conduit, and their corresponding fittings from the project.
        /// </summary>
        public List<Element> GetAllContainmentElements()
        {
            var categories = new List<BuiltInCategory>
            {
                BuiltInCategory.OST_Conduit, BuiltInCategory.OST_CableTray,
                BuiltInCategory.OST_ConduitFitting, BuiltInCategory.OST_CableTrayFitting
            };
            return new FilteredElementCollector(_doc)
                .WherePasses(new ElementMulticategoryFilter(categories))
                .WhereElementIsNotElementType().ToList();
        }

        /// <summary>
        /// Gets all electrical equipment and fixtures from the project.
        /// </summary>
        public List<Element> GetAllElectricalEquipment()
        {
            var equip = new FilteredElementCollector(_doc).OfCategory(BuiltInCategory.OST_ElectricalEquipment).WhereElementIsNotElementType().ToList();
            var fixtures = new FilteredElementCollector(_doc).OfCategory(BuiltInCategory.OST_ElectricalFixtures).WhereElementIsNotElementType().ToList();
            return equip.Concat(fixtures).ToList();
        }
    }
}