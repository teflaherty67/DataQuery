using DataQuery.Common;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Text.Json;
using System.Text.Json.Nodes;

namespace DataQuery
{
    /// <summary>
    /// Revit command to:
    ///   1. Add any missing shared parameters to the project
    ///   2. Launch frmProjInfo for user input
    ///   3. Write the values to Project Information parameters
    ///   4. Extract plan data from the active model
    ///   5. Write the data to Airtable
    /// Requires: clsPlanData.cs, frmProjInfo.xaml
    /// </summary>

    [Transaction(TransactionMode.Manual)]
    public class cmdDataExport : IExternalCommand
    {
        // declare variable for Airtable API client
        private const string AirtableApiKey = apiSecrets.AirtableApiKey;
        private const string AirtableBaseId = "appwAYciO1uHJiC7u";
        private const string AirtableTable = "tblCGlniNbnq76ifv";

        // Shared HTTP client instance for all Airtable API requests
        private static readonly HttpClient _http = new HttpClient();

        // declare variable for Shared Parameter file path
        private const string SharedParamFile = @"S:\Shared Folders\Lifestyle USA Design\Library 2026\LD_Shared-Parameters_Master.txt";

        // Maps each required parameter name to its shared parameter group and binding category.
        // Code Masonry lives in the "Title Block" group and is bound to sheets;
        // all other parameters live in "Project Information" and are bound to ProjectInformation.
        private static readonly Dictionary<string, (string Group, BuiltInCategory Category)> RequiredParams = new()
        {
            { "Spec Level",         ("Project Information", BuiltInCategory.OST_ProjectInformation) },
            { "Client Division",    ("Project Information", BuiltInCategory.OST_ProjectInformation) },
            { "Client Subdivision", ("Project Information", BuiltInCategory.OST_ProjectInformation) },
            { "Garage Loading",     ("Project Information", BuiltInCategory.OST_ProjectInformation) },
            { "Code Masonry",       ("Title Block",         BuiltInCategory.OST_Sheets)             }
        };

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            // Revit application and document variables
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document curDoc = uidoc.Document;

            // wrap the main logic in a try-catch block to handle any unexpected errors gracefully
            try
            {
                // 1. Add any missing shared parameters to the project
                Result paramResult = AddMissingParameters(uiapp, curDoc);
                if (paramResult != Result.Succeeded)
                    return paramResult;

                // 2. Launch frmProjInfo for user input
                frmProjInfo curForm = new frmProjInfo(curDoc);
                bool? dialogResult = curForm.ShowDialog();
                if (dialogResult != true)
                    return Result.Cancelled;

                // 3. Write form values to Project Information (and Code Masonry to Cover sheets if applicable)
                WriteValuesToProjectInfo(curDoc, curForm);

                // 4. Extract plan data
                clsPlanData planData = ExtractPlanData(curDoc);

                if (planData == null)
                {
                    Utils.TaskDialogError("Data Query", "Error", "Unable to extract plan data from the model.");
                    return Result.Succeeded; // ← not Failed, to preserve local parameter writes
                }

                // 5. Show confirmation
                if (!ShowConfirmationDialog(planData))
                    return Result.Succeeded; // ← not Cancelled, to preserve local parameter writes

                // 6. Write to Airtable - isolated try-catch so Revit doesn't roll back on failure
                try
                {
                    string existingRecordId = FindExistingRecord(planData.PlanName, planData.SpecLevel, planData.Subdivision);

                    if (existingRecordId != null)
                    {
                        if (!Utils.TaskDialogAccept("Data Query", "Plan Exists",
                            $"Plan '{planData.PlanName}' already exists. Update it?"))
                            return Result.Succeeded;

                        UpdateRecord(existingRecordId, planData);
                        Utils.TaskDialogInformation("Data Query", "Success", $"Updated plan '{planData.PlanName}' in Airtable.");
                    }
                    else
                    {
                        InsertRecord(planData);
                        Utils.TaskDialogInformation("Data Query", "Success", $"Added plan '{planData.PlanName}' to Airtable.");
                    }
                }
                catch (Exception ex)
                {
                    Utils.TaskDialogError("Data Query", "Airtable Error",
                        $"Project Information was saved, but Airtable write failed:\n{ex.Message}");
                }

                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                message = ex.Message;
                Utils.TaskDialogError("Data Query", "Error", $"An error occurred:\n{ex.Message}");
                return Result.Failed;
            }
        }

        #region Add Missing Parameters

        private Result AddMissingParameters(UIApplication uiapp, Document curDoc)
        {
            if (!File.Exists(SharedParamFile))
            {
                Utils.TaskDialogError("Data Query", "Error",
                    $"Shared parameter file not found:\n{SharedParamFile}");
                return Result.Failed;
            }

            List<string> addedParams = new List<string>();
            List<string> existingParams = new List<string>();

            foreach (string paramName in RequiredParams.Keys)
            {
                if (Utils.DoesProjectParamExist(curDoc, paramName))
                    existingParams.Add(paramName);
                else
                    addedParams.Add(paramName);
            }

            if (addedParams.Count == 0)
            {
                Utils.TaskDialogInformation("Data Query", "Parameters Verified",
                    "All required parameters already exist in the project.");
                return Result.Succeeded;
            }

            uiapp.Application.SharedParametersFilename = SharedParamFile;
            DefinitionFile curDefFile = uiapp.Application.OpenSharedParameterFile();

            if (curDefFile == null)
            {
                Utils.TaskDialogError("Data Query", "Error",
                    $"Could not open shared parameter file:\n{SharedParamFile}");
                return Result.Failed;
            }

            using (Transaction t = new Transaction(curDoc, "Add Shared Parameters"))
            {
                t.Start();

                // Group params to add by binding category so each category gets one InstanceBinding
                foreach (var catGroup in addedParams.GroupBy(p => RequiredParams[p].Category))
                {
                    CategorySet catSet = new CategorySet();
                    catSet.Insert(curDoc.Settings.Categories.get_Item(catGroup.Key));
                    InstanceBinding instBinding = uiapp.Application.Create.NewInstanceBinding(catSet);

                    foreach (string paramName in catGroup)
                    {
                        Definition curDef = Utils.GetParameterDefinitionFromFile(
                            curDefFile, RequiredParams[paramName].Group, paramName);

                        if (curDef == null)
                        {
                            Utils.TaskDialogError("Data Query", "Error",
                                $"Could not find definition for '{paramName}' in shared parameter file:\n{SharedParamFile}");
                            continue;
                        }

                        curDoc.ParameterBindings.Insert(curDef, instBinding);
                    }
                }

                t.Commit();
            }

            // Build and display result message
            string resultMessage = $"Added {addedParams.Count} parameter(s):\n";
            foreach (string name in addedParams)
                resultMessage += $"  - {name}\n";

            if (existingParams.Count > 0)
            {
                resultMessage += $"\n{existingParams.Count} parameter(s) already existed:\n";
                foreach (string name in existingParams)
                    resultMessage += $"  - {name}\n";
            }

            Utils.TaskDialogInformation("Data Query", "Parameters Added", resultMessage);

            return Result.Succeeded;
        }

        private void WriteValuesToProjectInfo(Document curDoc, frmProjInfo curForm)
        {
            using (Transaction t = new Transaction(curDoc, "Save Project Information"))
            {
                t.Start();

                ProjectInfo projInfo = curDoc.ProjectInformation;

                Common.Utils.SetParameterByName(projInfo, "Project Name", curForm.PlanName);
                Common.Utils.SetParameterByName(projInfo, "Spec Level", curForm.SpecLevel);
                Common.Utils.SetParameterByName(projInfo, "Client Name", curForm.ClientName);
                Common.Utils.SetParameterByName(projInfo, "Client Division", curForm.ClientDivision);
                Common.Utils.SetParameterByName(projInfo, "Client Subdivision", curForm.ClientSubdivision);
                Common.Utils.SetParameterByName(projInfo, "Garage Loading", curForm.GarageLoading);

                // If Code Masonry was newly added and the user entered a value, write it to all Cover sheets
                if (!string.IsNullOrEmpty(curForm.CodeMasonry) &&
                    int.TryParse(curForm.CodeMasonry, out int masonryValue))
                {
                    IEnumerable<ViewSheet> coverSheets = new FilteredElementCollector(curDoc)
                        .OfClass(typeof(ViewSheet))
                        .Cast<ViewSheet>()
                        .Where(s => s.Name.IndexOf("Cover", StringComparison.OrdinalIgnoreCase) >= 0
                                 && !s.IsTemplate);

                    foreach (ViewSheet sheet in coverSheets)
                        Common.Utils.SetParameterByName(sheet, "Code Masonry", masonryValue);
                }

                t.Commit();
            }
        }

        #endregion

        #region Plan Data Extraction

        private clsPlanData ExtractPlanData(Document curDoc)
        {
            var planData = new clsPlanData();
            ProjectInfo curProjInfo = curDoc.ProjectInformation;

            planData.PlanName = Utils.GetParameterValueByName(curProjInfo, "Project Name");
            planData.SpecLevel = Utils.GetParameterValueByName(curProjInfo, "Spec Level");
            planData.Client = Utils.GetParameterValueByName(curProjInfo, "Client Name");
            planData.Division = Utils.GetParameterValueByName(curProjInfo, "Client Division");
            planData.Subdivision = Utils.GetParameterValueByName(curProjInfo, "Client Subdivision");
            planData.GarageLoading = Utils.GetParameterValueByName(curProjInfo, "Garage Loading");

            GetBuildingDimensions(curDoc, out string width, out string depth);
            planData.OverallWidth = width;
            planData.OverallDepth = depth;

            planData.Stories = CountStories(curDoc);

            GetRoomCounts(curDoc, out int bedrooms, out decimal bathrooms);
            planData.Bedrooms = bedrooms;
            planData.Bathrooms = bathrooms;

            planData.MasterBedLoc = GetMasterBedLoc(curDoc);
            planData.MasonryPercentage = GetMasonryPercentage(curDoc);
            planData.GarageBays = CountGarageBays(curDoc);
            planData.LivingArea = GetLivingArea(curDoc);
            planData.TotalArea = GetTotalArea(curDoc);

            return planData;
        }

        private void GetBuildingDimensions(Document curDoc, out string width, out string depth)
        {
            // assign default values for the out parameters
            width = "0'-0\"";
            depth = "0'-0\"";

            // find the Form/Foundation Plan view to use as the source for dimension extraction
            View foundationView = new FilteredElementCollector(curDoc)
                .OfClass(typeof(View))
                .Cast<View>()
                .FirstOrDefault(v => v.Name.IndexOf("Form/Foundation Plan", StringComparison.OrdinalIgnoreCase) >= 0
                      && !v.IsTemplate);

            // null check - if the view isn't found, show an error and return default dimensions
            if (foundationView == null)
            {
                Utils.TaskDialogWarning("Data Query", "Error",
                    "No Form/Foundation Plan view found. Cannot extract building dimensions.");
                return;
            }

            // get the view's RightDirection and UpDirection to determine width and depth orientation
            XYZ up = foundationView.UpDirection.Normalize();
            XYZ right = foundationView.RightDirection.Normalize();

            // if the view is a dependent view, collect dimensions from the parent view instead
            ElementId parentId = foundationView.GetPrimaryViewId();
            ElementId collectFromId = (parentId != ElementId.InvalidElementId)
                ? parentId
                : foundationView.Id;

            // collect all single-segment dimensions in the view
            List<Dimension> listDims = new FilteredElementCollector(curDoc, collectFromId)
                .OfClass(typeof(Dimension))
                .Cast<Dimension>()
                .Where(d => d.Segments.Size == 0 && d.Value.HasValue)
                .ToList();

            // null check - if no dimensions are found, show an error and return default dimensions
            if (!listDims.Any())
            {
                Utils.TaskDialogWarning("Data Query", "Error",
                    "No single-segment dimensions found in the Form/Foundation Plan view. Cannot extract building dimensions.");
                return;
            }

            Dimension widthDim = null;
            Dimension depthDim = null;
            double maxWidthValue = double.MinValue;
            double maxDepthValue = double.MinValue;

            foreach (Dimension curDim in listDims)
            {
                Line dimLine = curDim.Curve as Line;
                if (dimLine == null) continue;

                XYZ dimDir = dimLine.Direction.Normalize();

                bool isWidth = Math.Abs(dimDir.DotProduct(right)) > 0.99;
                bool isDepth = Math.Abs(dimDir.DotProduct(up)) > 0.99;

                double value = curDim.Value.Value;

                if (isWidth && value > maxWidthValue)
                {
                    maxWidthValue = value;
                    widthDim = curDim;
                }
                else if (isDepth && value > maxDepthValue)
                {
                    maxDepthValue = value;
                    depthDim = curDim;
                }
            }

            if (widthDim != null) width = widthDim.ValueString;
            if (depthDim != null) depth = depthDim.ValueString;
        }

        private int CountStories(Document curDoc)
        {
            string[] storyLevels = { "first floor", "second floor", "main level", "upper level" };

            return new FilteredElementCollector(curDoc)
                .OfClass(typeof(Level))
                .Cast<Level>()
                .Where(l => storyLevels.Contains(l.Name.ToLower()))
                .Count();
        }

        private void GetRoomCounts(Document curDoc, out int bedrooms, out decimal bathrooms)
        {
            bedrooms = 0;
            decimal fullBaths = 0;
            decimal halfBaths = 0;

            foreach (Room room in new FilteredElementCollector(curDoc)
                .OfClass(typeof(SpatialElement))
                .OfType<Room>())
            {
                if (room.Area <= 0) continue;

                string name = room.Name.ToLower();

                if (name.Contains("bedroom") || name.Contains("bed"))
                    bedrooms++;

                if (name.Contains("bath"))
                    fullBaths++;

                if (name.Contains("powder") || name.Contains("pwdr"))
                    halfBaths++;
            }

            bathrooms = fullBaths + (halfBaths * 0.5m);
        }

        private string GetMasterBedLoc(Document curDoc)
        {
            foreach (Room room in new FilteredElementCollector(curDoc)
                .OfClass(typeof(SpatialElement))
                .OfType<Room>())
            {
                if (room.Area <= 0) continue;
                if (!room.Name.ToLower().Contains("master")) continue;

                string levelName = room.Level?.Name ?? string.Empty;

                if (levelName.ToLower().Contains("first") || levelName.ToLower().Contains("main"))
                    return "Down";
                else
                    return "Up";
            }

            return null;
        }

        private string GetMasonryPercentage(Document curDoc)
        {
            int maxMasonry = 0;

            // find all Cover sheets and return the largest Code Masonry value across them
            IEnumerable<ViewSheet> coverSheets = new FilteredElementCollector(curDoc)
                .OfClass(typeof(ViewSheet))
                .Cast<ViewSheet>()
                .Where(s => s.Name.IndexOf("Cover", StringComparison.OrdinalIgnoreCase) >= 0
                         && !s.IsTemplate);

            foreach (ViewSheet sheet in coverSheets)
            {
                Parameter codeMasonry = Utils.GetParameterByName(sheet, "Code Masonry");
                if (codeMasonry == null) continue;

                int value = (int)codeMasonry.AsDouble();
                if (value > maxMasonry)
                    maxMasonry = value;
            }

            return maxMasonry == 0 ? "N/A" : $"{maxMasonry}%";
        }

        private int CountGarageBays(Document curDoc)
        {
            int bays = 0;

            foreach (Room room in new FilteredElementCollector(curDoc)
                .OfClass(typeof(SpatialElement))
                .OfType<Room>())
            {
                if (room.Area <= 0) continue;

                string name = room.Name.ToLower();
                if (!name.Contains("garage")) continue;

                if (name.Contains("three")) bays += 3;
                else if (name.Contains("two")) bays += 2;
                else if (name.Contains("one")) bays += 1;
            }

            return bays;
        }

        private int GetLivingArea(Document curDoc)
        {
            ViewSchedule schedule = Utils.GetFloorAreaSchedule(curDoc);
            if (schedule == null) return 0;

            TableSectionData body = schedule.GetTableData().GetSectionData(SectionType.Body);
            int rowCount = body.NumberOfRows;
            int areaCol = body.NumberOfColumns - 1;

            for (int row = 0; row < rowCount; row++)
            {
                if (!body.GetCellText(row, 0).Trim().Equals("Living", StringComparison.OrdinalIgnoreCase))
                    continue;

                string areaText = body.GetCellText(row, areaCol).Trim();
                if (!string.IsNullOrEmpty(areaText)) return ParseAreaValue(areaText);

                for (int sub = row + 1; sub < rowCount; sub++)
                {
                    string subName = body.GetCellText(sub, 0).Trim();
                    string subArea = body.GetCellText(sub, areaCol).Trim();

                    if (!string.IsNullOrEmpty(subName) && !subName.Contains("Floor")) break;
                    if (string.IsNullOrEmpty(subName) && !string.IsNullOrEmpty(subArea))
                        return ParseAreaValue(subArea);
                }
            }

            return 0;
        }

        private int GetTotalArea(Document curDoc)
        {
            ViewSchedule schedule = Utils.GetFloorAreaSchedule(curDoc);
            if (schedule == null) return 0;

            TableSectionData body = schedule.GetTableData().GetSectionData(SectionType.Body);
            int rowCount = body.NumberOfRows;
            int areaCol = body.NumberOfColumns - 1;

            for (int row = 0; row < rowCount; row++)
            {
                if (body.GetCellText(row, 0).Trim().Equals("Total Covered", StringComparison.OrdinalIgnoreCase))
                    return ParseAreaValue(body.GetCellText(row, areaCol).Trim());
            }

            return 0;
        }

        private int ParseAreaValue(string areaText)
        {
            string cleaned = areaText.Replace("SF", "").Replace(",", "").Trim();
            return int.TryParse(cleaned, out int result) ? result : 0;
        }

        #endregion

        #region Airtable Operations

        private string FindExistingRecord(string planName, string specLevel, string subdivision)
        {
            string formula = $"AND({{Plan Name}}=\"{planName}\",{{Spec Level}}=\"{specLevel}\",{{Client Subdivision}}=\"{subdivision}\")";
            string url = $"https://api.airtable.com/v0/{AirtableBaseId}/{AirtableTable}" +
                             $"?filterByFormula={Uri.EscapeDataString(formula)}&maxRecords=1";

            HttpRequestMessage request = BuildRequest(HttpMethod.Get, url);
            HttpResponseMessage response = _http.SendAsync(request).GetAwaiter().GetResult();
            response.EnsureSuccessStatusCode();

            string json = response.Content.ReadAsStringAsync().GetAwaiter().GetResult();
            JsonNode root = JsonNode.Parse(json);
            JsonArray records = root["records"]?.AsArray();

            return (records != null && records.Count > 0)
                ? records[0]["id"]?.GetValue<string>()
                : null;
        }

        private void InsertRecord(clsPlanData plan)
        {
            string url = $"https://api.airtable.com/v0/{AirtableBaseId}/{AirtableTable}";
            string body = BuildRecordJson(plan);

            HttpRequestMessage request = BuildRequest(HttpMethod.Post, url, body);
            HttpResponseMessage response = _http.SendAsync(request).GetAwaiter().GetResult();

            if (!response.IsSuccessStatusCode)
            {
                string errorBody = response.Content.ReadAsStringAsync().GetAwaiter().GetResult();
                throw new Exception($"HTTP {(int)response.StatusCode}: {errorBody}");
            }
        }

        private void UpdateRecord(string recordId, clsPlanData plan)
        {
            string url = $"https://api.airtable.com/v0/{AirtableBaseId}/{AirtableTable}/{recordId}";
            string body = BuildRecordJson(plan);

            HttpRequestMessage request = BuildRequest(HttpMethod.Patch, url, body);
            HttpResponseMessage response = _http.SendAsync(request).GetAwaiter().GetResult();
            response.EnsureSuccessStatusCode();
        }

        private string BuildRecordJson(clsPlanData plan)
        {
            var fields = new Dictionary<string, object>
            {
                { "Plan Name",          plan.PlanName          },
                { "Spec Level",         plan.SpecLevel         },
                { "Client Name",        plan.Client            },
                { "Client Division",    plan.Division          },
                { "Client Subdivision", plan.Subdivision       },
                { "Overall Width",      plan.OverallWidth      },
                { "Overall Depth",      plan.OverallDepth      },
                { "Stories",            plan.Stories           },
                { "Bedrooms",           plan.Bedrooms          },
                { "Bathrooms",          (double)plan.Bathrooms },
                { "Garage Bays",        plan.GarageBays        },
                { "Garage Loading",     plan.GarageLoading     },
                { "Master Bedroom",     plan.MasterBedLoc      },
                { "Masonry",            plan.MasonryPercentage },
                { "Living Area",        plan.LivingArea        },
                { "Total Area",         plan.TotalArea         }
            };

            return JsonSerializer.Serialize(new { fields });
        }

        private HttpRequestMessage BuildRequest(HttpMethod method, string url, string jsonBody = null)
        {
            var request = new HttpRequestMessage(method, url);
            request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", AirtableApiKey);
            request.Headers.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));

            if (jsonBody != null)
                request.Content = new StringContent(jsonBody, Encoding.UTF8, "application/json");

            return request;
        }

        #endregion

        #region UI

        private bool ShowConfirmationDialog(clsPlanData planData)
        {
            string message = $@"Ready to save this plan to Airtable:

                Plan Name:      {planData.PlanName}
                Spec Level:     {planData.SpecLevel}
                Client:         {planData.Client}
                Division:       {planData.Division}
                Subdivision:    {planData.Subdivision}

                Dimensions:     {planData.OverallWidth} W x {planData.OverallDepth} D
                Total Area:     {planData.TotalArea:N0} SF
                Living Area:    {planData.LivingArea:N0} SF
                Bedrooms:       {planData.Bedrooms}
                Bathrooms:      {planData.Bathrooms}
                Stories:        {planData.Stories}
                Master Bedroom: {planData.MasterBedLoc ?? "Not Found"}
                Masonry:        {planData.MasonryPercentage ?? "N/A"}
                Garage Bays:    {planData.GarageBays}
                Garage Loading: {planData.GarageLoading}

                Do you want to proceed?";

            return Utils.TaskDialogAccept("Data Query", "Confirm Plan Data", message);
        }

        internal static PushButtonData GetButtonData()
        {
            string buttonInternalName = "btnCommand1";
            string buttonTitle = "Button 1";

            Common.ButtonDataClass myButtonData = new Common.ButtonDataClass(
                buttonInternalName,
                buttonTitle,
                MethodBase.GetCurrentMethod().DeclaringType?.FullName,
                Properties.Resources.Blue_32,
                Properties.Resources.Blue_16,
                "This is a tooltip for Button 1");

            return myButtonData.Data;
        }

        #endregion
    }
}