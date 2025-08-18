//-----------------------------------------------------------------------------
// <copyright file="ContainmentIsland.cs" company="ReTick Solutions Pty Ltd">
//     Copyright (c) ReTick Solutions Pty Ltd. All rights reserved.
// </copyright>
// <summary>
//   Represents a single, continuously connected group of containment elements.
// </summary>
//-----------------------------------------------------------------------------

using Autodesk.Revit.DB;
using System;
using System.Collections.Generic;

namespace RTS.Reports.Generators.Routing
{
    /// <summary>
    /// Represents a single, continuously connected group of containment elements.
    /// </summary>
    public class ContainmentIsland
    {
        /// <summary>
        /// Gets or sets the list of elements that make up this island.
        /// </summary>
        public List<Element> Elements { get; set; } = new List<Element>();

        /// <summary>
        /// Gets or sets the potential entry points to this island from the start equipment.
        /// Each tuple contains the ElementId of the entry element and the distance to it.
        /// </summary>
        public List<Tuple<ElementId, double>> EntryPoints { get; set; } = new List<Tuple<ElementId, double>>();

        /// <summary>
        /// Gets or sets the potential exit points from this island to the end equipment.
        /// Each tuple contains the ElementId of the exit element and the distance from it.
        /// </summary>
        public List<Tuple<ElementId, double>> ExitPoints { get; set; } = new List<Tuple<ElementId, double>>();
    }
}