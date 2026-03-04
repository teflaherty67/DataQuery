using DataQuery.Common;
using System.Windows.Input;

namespace DataQuery
{
    /// <summary>
    /// One-time migration command that swaps the Sq Ft sheet parameter to Living Area.
    /// Reads the value of Sq Ft from each sheet, adds the Living Area shared parameter,
    /// assigns the stored values, then removes the Sq Ft binding.
    /// Called automatically by cmdDataQuery if Living Area is not found on sheets.
    /// </summary>
    [Transaction(TransactionMode.Manual)]
    public class cmdLivingParam : IExternalCommand
    {
        // define constants for the shared parameter file, group, and parameter names
        private const string SharedParamFile = @"S:\Shared Folders\Lifestyle USA Design\Library 2026\LD_Shared-Parameters_Master.txt";
        private const string SharedParamGroup = "Title Block";
        private const string NewParamName = "Living Area";
        private const string OldParamName = "Sq Ft";

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            // Revit application and document variables
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document curDoc = uidoc.Document;

            // wrap the entire code in a try-catch to handle any exceptions gracefully
            try
            {
                // step 1: verify the shared parameter file exists
                if (!File.Exists(SharedParamFile))
                {
                    Utils.TaskDialogError("Data Export", "Error",
                        $"Shared parameter file not found:\n{SharedParamFile}");
                    return Result.Failed;
                }

                // step 2: collect all sheets in the project
                List<ViewSheet> sheets = new FilteredElementCollector(curDoc)
                    .OfClass(typeof(ViewSheet))
                    .Cast<ViewSheet>()
                    .Where(s => !s.IsTemplate)
                    .ToList();

                // step 3: read the Sq Ft values from each sheet and store them in a dictionary
                Dictionary<ElementId, double> dicSqFtValues = new Dictionary<ElementId, double>();

                // step 4: loop through the sheets, get & store the Sq Ft value in the dictionary
                foreach (ViewSheet sheet in sheets)
                {
                    Parameter sqFt = Utils.GetParameterByName(sheet, OldParamName);
                    if (sqFt == null) continue;

                    dicSqFtValues[sheet.Id] = sqFt.AsDouble();
                }

                // notify the user if no sheets with Sq Ft values were found and exit
                if (dicSqFtValues.Count == 0)
                {
                    Utils.TaskDialogWarning("Data Export", "No Values Found",
                        $"No sheets were found with a '{OldParamName}' parameter value.");
                    return Result.Failed;
                }

                // step 5: open the shared parameter file and get the definition for Living Area
                string originalParamFile = uiapp.Application.SharedParametersFilename;
                uiapp.Application.SharedParametersFilename = SharedParamFile;

                // open the shared parameter file
                DefinitionFile defFile = uiapp.Application.OpenSharedParameterFile();

                // null check the definition file
                if (defFile == null)
                {
                    // notify the user if the shared parameter file could not be opened
                    Utils.TaskDialogError("Living Param", "Error",
                        "Could not open the shared parameter file.");

                    // restore the original shared parameter file path and exit
                    uiapp.Application.SharedParametersFilename = originalParamFile;
                    return Result.Failed;
                }

                // step 6: get the Living Area definition from the shared parameter file
                Definition newDef = Utils.GetParameterDefinitionFromFile(defFile, SharedParamGroup, NewParamName);

                // null check the new definition
                if (newDef == null)
                {
                    // notify the user if the Living Area definition could not be found in the shared parameter file
                    Utils.TaskDialogError("Data Export", "Error",
                        $"Could not find '{NewParamName}' in shared parameter file group '{SharedParamGroup}'.");

                    // restore the original shared parameter file before exiting
                    uiapp.Application.SharedParametersFilename = originalParamFile;

                    return Result.Failed;
                }

                // step 7: get Sq Ft definition for removal later
                DefinitionBindingMapIterator oldIter = curDoc.ParameterBindings.ForwardIterator();
                Definition oldDef = null;
                while (oldIter.MoveNext())
                {
                    if (oldIter.Key.Name.Equals(OldParamName, StringComparison.OrdinalIgnoreCase))
                    {
                        oldDef = oldIter.Key;
                        break;
                    }
                }

                // step 8: add Living Area parameter to the project
                CategorySet catSet = new CategorySet();
                catSet.Insert(curDoc.Settings.Categories.get_Item(BuiltInCategory.OST_Sheets));
                InstanceBinding instBinding = uiapp.Application.Create.NewInstanceBinding(catSet);

                // create a transaction to add Living Area, assign values, and remove Sq Ft
                using (Transaction t = new Transaction(curDoc, "Migrate Parameter Values"))
                {
                    // start the transaction
                    t.Start();

                    // add Living Area parameter to Sheets
#if REVIT2024
curDoc.ParameterBindings.Insert(newDef, instBinding, BuiltInParameterGroup.PG_IDENTITY_DATA);
#else
                    curDoc.ParameterBindings.Insert(newDef, instBinding, GroupTypeId.IdentityData);
#endif
                    // loop through the sheets and set Living Area value to the stored Sq Ft value
                    foreach (ViewSheet sheet in sheets)
                    {
                        if (!dicSqFtValues.TryGetValue(sheet.Id, out double storedValue)) continue;
                        if (storedValue <= 0) continue;

                        Parameter livingArea = Utils.GetParameterByNameAndWritable(sheet, NewParamName);
                        if (livingArea == null) continue;

                        livingArea.Set(storedValue);
                    }

                    // remove the Sq Ft parameter binding from the project
                    if (oldDef != null)
                        curDoc.ParameterBindings.Remove(oldDef);

                    // commit the transaction
                    t.Commit();
                }

                // notify the user of success
                Utils.TaskDialogInformation("Data Export", "Migration Complete",
                   $"'{OldParamName}' values have been migrated to '{NewParamName}' on {dicSqFtValues.Count} sheet(s).\n" +
                   $"'{OldParamName}' parameter has been removed.");

                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                message = ex.Message;
                Utils.TaskDialogError("Living Param", "Error", $"An error occurred:\n{ex.Message}");
                return Result.Failed;
            }
        }

        internal static PushButtonData GetButtonData()
        {
            // use this method to define the properties for this command in the Revit ribbon
            string buttonInternalName = "btnCommand2";
            string buttonTitle = "Button 2";

            Common.ButtonDataClass myButtonData = new Common.ButtonDataClass(
                buttonInternalName,
                buttonTitle,
                MethodBase.GetCurrentMethod().DeclaringType?.FullName,
                Properties.Resources.Blue_32,
                Properties.Resources.Blue_16,
                "This is a tooltip for Button 2");

            return myButtonData.Data;
        }
    }
}
