using Autodesk.Revit.DB;
using System.Collections.Generic;

namespace RTS.Reports.Generators.Routing
{
    public enum ContainmentType { Tray, Conduit, Mixed }

    /// <summary>
    /// Represents a single, continuously connected group of containment elements.
    /// </summary>
    public class ContainmentIsland
    {
        public int Id { get; }
        public HashSet<ElementId> ElementIds { get; } = new HashSet<ElementId>();
        public BoundingBoxXYZ BoundingBox { get; set; }
        public ContainmentType Type { get; set; } = ContainmentType.Mixed;
        public ContainmentIsland(int id) { Id = id; }
    }
}