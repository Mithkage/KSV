using Autodesk.Revit.DB;
using System.Collections.Generic;

namespace RTS.Reports.Generators.Routing
{
    public class VirtualPathResult
    {
        public string RoutingSequence { get; set; }
        public string BranchSequence { get; set; }
        public List<ElementId> StitchedPath { get; set; }
        public double VirtualLength { get; set; }
        public int IslandCount { get; set; }
        public string StatusMessage { get; set; }
    }
}