using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DataQuery
{
    /// <summary>
    /// House plan data record class to store and manage information
    /// about individual house plans, including dimensions, area,
    /// bedroom/bathroom counts, and client/division details.
    /// </summary>
    internal class clsPlanData
    {
        public string PlanName { get; set; }
        public string SpecLevel { get; set; }
        public string Client { get; set; }
        public string Division { get; set; }
        public string Subdivision { get; set; }
        public string OverallWidth { get; set; }
        public string OverallDepth { get; set; }
        public int Stories { get; set; }
        public int Bedrooms { get; set; }
        public decimal Bathrooms { get; set; }
        public int GarageBays { get; set; }
        public string GarageLoading { get; set; }
        public int LivingArea { get; set; }
        public int TotalArea { get; set; }


        public override string ToString()
        {
            return $"{PlanName} - {SpecLevel} | {LivingArea} SF | {Bedrooms}BR/{Bathrooms}BA | {Stories} Story";
        }

        /// <summary>
        /// Get a detailed multi-line description
        /// </summary>
        public string ToDetailedString()
        {
            return $@"
                Plan: {PlanName} ({SpecLevel})
                Client: {Client ?? "N/A"}
                Division: {Division ?? "N/A"}
                Subdivision: {Subdivision ?? "N/A"}
                Dimensions: {OverallWidth} W x {OverallDepth} D
                Stories: {Stories}
                Bedrooms: {Bedrooms} | Bathrooms: {Bathrooms}
                Garage: {GarageBays} bay(s)
                Garage Loading: {GarageLoading} Load
                Living Area: {LivingArea:N0} SF
                Total Area: {TotalArea:N0} SF";
        }
    }
}