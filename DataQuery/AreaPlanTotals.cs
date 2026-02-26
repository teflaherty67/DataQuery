using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DataQuery.Common;

namespace DataQuery
{
    internal static class AreaPlanTotals
    {
        private const string AreaCategoryParamName = "Area Category";
        private const string TotalCoveredValue = "Total Covered";
        private const string LivingName = "Living";

        public static int GetLivingArea(Document doc)
        {
            ElementId schemeId = FindBestSchemeByLiving(doc);
            if (schemeId == ElementId.InvalidElementId) return 0;

            double living = CollectAreas(doc, schemeId)
                .Where(a => string.Equals((a.Name ?? "").Trim(), LivingName, StringComparison.OrdinalIgnoreCase))
                .Sum(a => a.Area);

            return (int)Math.Round(living, MidpointRounding.AwayFromZero);
        }

        public static int GetTotalCoveredArea(Document doc)
        {
            ElementId schemeId = FindBestSchemeByLiving(doc);
            if (schemeId == ElementId.InvalidElementId) return 0;

            double totalCovered = CollectAreas(doc, schemeId)
                .Where(a => IsAreaCategory(a, TotalCoveredValue))
                .Sum(a => a.Area);

            return (int)Math.Round(totalCovered, MidpointRounding.AwayFromZero);
        }

        /// <summary>
        /// Finds the AreaSchemeId for the elevation that actually has a non-zero Living area.
        /// If multiple schemes have Living > 0, picks the one with the largest Living sum;
        /// tie-breaker is largest Total Covered sum.
        /// </summary>
        public static ElementId FindBestSchemeByLiving(Document doc)
        {
            // Schemes that are actually used by at least one Area Plan view in this model
            var schemeIdsInUse = new FilteredElementCollector(doc)
                .OfClass(typeof(ViewPlan))
                .Cast<ViewPlan>()
                .Where(v => !v.IsTemplate && v.ViewType == ViewType.AreaPlan)
                .Select(v => v.AreaSchemeId)
                .Where(id => id != ElementId.InvalidElementId)
                .Distinct()
                .ToList();

            if (schemeIdsInUse.Count == 0)
                return ElementId.InvalidElementId;

            // Pull all areas once for performance
            var allAreas = new FilteredElementCollector(doc)
                .OfCategory(BuiltInCategory.OST_Areas)
                .WhereElementIsNotElementType()
                .Cast<Area>()
                .Where(a => a.AreaSchemeId != ElementId.InvalidElementId && a.Area > 0)
                .ToList();

            ElementId bestScheme = ElementId.InvalidElementId;
            double bestLiving = 0.0;
            double bestTotalCovered = 0.0;

            foreach (var schemeId in schemeIdsInUse)
            {
                var schemeAreas = allAreas.Where(a => a.AreaSchemeId == schemeId);

                double living = schemeAreas
                    .Where(a => string.Equals((a.Name ?? "").Trim(), LivingName, StringComparison.OrdinalIgnoreCase))
                    .Sum(a => a.Area);

                if (living <= 0)
                    continue;

                double totalCovered = schemeAreas
                    .Where(a => IsAreaCategory(a, TotalCoveredValue))
                    .Sum(a => a.Area);

                bool better =
                    living > bestLiving ||
                    (Math.Abs(living - bestLiving) < 1e-9 && totalCovered > bestTotalCovered);

                if (better)
                {
                    bestLiving = living;
                    bestTotalCovered = totalCovered;
                    bestScheme = schemeId;
                }
            }

            return bestScheme;
        }

        private static IEnumerable<Area> CollectAreas(Document doc, ElementId schemeId)
        {
            return new FilteredElementCollector(doc)
                .OfCategory(BuiltInCategory.OST_Areas)
                .WhereElementIsNotElementType()
                .Cast<Area>()
                .Where(a => a.AreaSchemeId == schemeId && a.Area > 0);
        }

        private static bool IsAreaCategory(Area area, string expected)
        {
            Parameter p = area?.LookupParameter(AreaCategoryParamName);
            if (p == null) return false;

            string val = (p.StorageType == StorageType.String ? p.AsString() : p.AsValueString()) ?? "";
            return string.Equals(val.Trim(), expected, StringComparison.OrdinalIgnoreCase);
        }
    }
}
