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



            }
            catch (Exception)
            {

                throw;
            }


            return Result.Succeeded;
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
