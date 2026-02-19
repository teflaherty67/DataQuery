namespace DataQuery.Common
{
    internal static class Utils
    {
        #region Parameters

        public struct ParameterData
        {
            public Definition def;
            public ElementBinding binding;
            public string name;
            public bool IsSharedStatusKnown;
            public bool IsShared;
            public string GUID;
            public ElementId id;
        }

        public static List<ParameterData> GetAllProjectParameters(Document curDoc)
        {
            if (curDoc.IsFamilyDocument)
            {
                TaskDialog.Show("Error", "Cannot be a family curDocument.");
                return null;
            }

            List<ParameterData> paraList = new List<ParameterData>();

            BindingMap map = curDoc.ParameterBindings;
            DefinitionBindingMapIterator iter = map.ForwardIterator();
            iter.Reset();
            while (iter.MoveNext())
            {
                ParameterData pd = new ParameterData();
                pd.def = iter.Key;
                pd.name = iter.Key.Name;
                pd.binding = iter.Current as ElementBinding;
                paraList.Add(pd);
            }

            return paraList;
        }

        public static bool DoesProjectParamExist(Document curDoc, string pName)
        {
            List<ParameterData> pdList = GetAllProjectParameters(curDoc);
            foreach (ParameterData pd in pdList)
            {
                if (pd.name == pName)
                {
                    return true;
                }
            }
            return false;
        }

        public static void CreateSharedParam(Document curDoc, string groupName, string paramName, BuiltInCategory cat)
        {
            Definition curDef = null;

            //check if current file has shared param file - if not then exit
            DefinitionFile defFile = curDoc.Application.OpenSharedParameterFile();

            //check if file has shared parameter file
            if (defFile == null)
            {
                TaskDialog.Show("Error", "No shared parameter file.");
                //Throw New Exception("No Shared Parameter File!")
            }

            //check if shared parameter exists in shared param file - if not then create
            if (ParamExists(defFile.Groups, groupName, paramName) == false)
            {
                //create param
                curDef = AddParamToFile(defFile, groupName, paramName);
            }
            else
            {
                curDef = GetParameterDefinitionFromFile(defFile, groupName, paramName);
            }

            //check if param is added to views - if not then add
            if (ParamAddedToFile(curDoc, paramName) == false)
            {
                //add parameter to current Revitfile
                AddParamToDocument(curDoc, curDef, cat);
            }
        }

        internal static string GetParameterValueByName(Element element, string paramName)
        {
            IList<Parameter> paramList = element.GetParameters(paramName);

            if (paramList != null)
                try
                {
                    Parameter param = paramList[0];
                    string paramValue = param.AsValueString();
                    return paramValue;
                }
                catch (System.ArgumentOutOfRangeException)
                {
                    return null;
                }

            return "";
        }

        internal static Parameter GetParameterByName(Element curElem, string paramName)
        {
            foreach (Parameter curParam in curElem.Parameters)
            {
                if (curParam.Definition.Name.ToString() == paramName)
                    return curParam;
            }

            return null;
        }

        internal static Parameter GetParameterByNameAndWritable(Element curElem, string paramName)
        {
            foreach (Parameter curParam in curElem.Parameters)
            {
                if (curParam.Definition.Name.ToString() == paramName && curParam.IsReadOnly == false)
                    return curParam;
            }

            return null;
        }

        internal static ElementId GetProjectParameterId(Document curDoc, string name)
        {
            ParameterElement pElem = new FilteredElementCollector(curDoc)
                .OfClass(typeof(ParameterElement))
                .Cast<ParameterElement>()
                .Where(e => e.Name.Equals(name))
                .FirstOrDefault();

            return pElem?.Id;
        }

        internal static ElementId GetBuiltInParameterId(Document curDoc, BuiltInCategory cat, BuiltInParameter bip)
        {
            FilteredElementCollector collector = new FilteredElementCollector(curDoc);
            collector.OfCategory(cat);

            Parameter curParam = collector.FirstElement().get_Parameter(bip);

            return curParam?.Id;
        }

        internal static string SetParameterByNameAndWritable(Element curElem, string paramName, string value)
        {
            Parameter curParam = GetParameterByNameAndWritable(curElem, paramName);

            curParam.Set(value);
            return curParam.ToString();
        }

        internal static void SetParameterByName(Element element, string paramName, string value)
        {
            IList<Parameter> paramList = element.GetParameters(paramName);

            if (paramList != null)
            {
                Parameter param = paramList[0];

                param.Set(value);
            }
        }

        internal static void SetParameterByName(Element element, string paramName, int value)
        {
            IList<Parameter> paramList = element.GetParameters(paramName);

            if (paramList != null)
            {
                Parameter param = paramList[0];

                param.Set(value);
            }
        }

        internal static bool SetParameterValue(Element curElem, string paramName, string value)
        {
            Parameter curParam = GetParameterByName(curElem, paramName);

            if (curParam != null)
            {
                curParam.Set(value);
                return true;
            }

            return false;
        }

        internal static Definition GetParameterDefinitionFromFile(DefinitionFile defFile, string groupName, string paramName)
        {
            // iterate the Definition groups of this file
            foreach (DefinitionGroup group in defFile.Groups)
            {
                if (group.Name == groupName)
                {
                    // iterate the difinitions
                    foreach (Definition definition in group.Definitions)
                    {
                        if (definition.Name == paramName)
                            return definition;
                    }
                }
            }
            return null;
        }

        //check if specified parameter is already added to Revit file
        public static bool ParamAddedToFile(Document curDoc, string paramName)
        {
            foreach (Parameter curParam in curDoc.ProjectInformation.Parameters)
            {
                if (curParam.Definition.Name.Equals(paramName))
                {
                    return true;
                }
            }

            return false;
        }

        public static bool AddParamToDocument(Document curDoc, Definition curDef, BuiltInCategory cat)
        {
            bool paramAdded = false;

            //define category for shared param
            Category myCat = curDoc.Settings.Categories.get_Item(cat);
            CategorySet myCatSet = curDoc.Application.Create.NewCategorySet();
            myCatSet.Insert(myCat);

            //create binding
            ElementBinding curBinding = curDoc.Application.Create.NewInstanceBinding(myCatSet);

            //do something
            //paramAdded = curDoc.ParameterBindings.Insert(curDef, curBinding, BuiltInParameterGroup.PG_IDENTITY_DATA);

            return paramAdded;
        }


        //check if specified parameter exists in shared parameter file
        public static bool ParamExists(DefinitionGroups groupList, string groupName, string paramName)
        {
            //loop through groups and look for match
            foreach (DefinitionGroup curGroup in groupList)
            {
                if (curGroup.Name.Equals(groupName) == true)
                {
                    //check if param exists
                    foreach (Definition curDef in curGroup.Definitions)
                    {
                        if (curDef.Name.Equals(paramName))
                        {
                            return true;
                        }
                    }
                }
            }

            return false;
        }

        //add parameter to specified shared parameter file
        public static Definition AddParamToFile(DefinitionFile defFile, string groupName, string paramName)
        {
            //create new shared parameter in specified file
            DefinitionGroup defGroup = GetDefinitionGroup(defFile, groupName);

            //check if group exists - if not then create
            if (defGroup == null)
            {
                //create group
                defGroup = defFile.Groups.Create(groupName);
            }

            //create parameter in group
            ExternalDefinitionCreationOptions curOptions = new ExternalDefinitionCreationOptions(paramName, SpecTypeId.String.Text);
            curOptions.Visible = true;

            Definition newParam = defGroup.Definitions.Create(curOptions);

            return newParam;
        }

        public static DefinitionGroup GetDefinitionGroup(DefinitionFile defFile, string groupName)
        {
            //loop through groups and look for match
            foreach (DefinitionGroup curGroup in defFile.Groups)
            {
                if (curGroup.Name.Equals(groupName))
                {
                    return curGroup;
                }
            }

            return null;
        }

        #endregion

        #region Ribbon
        internal static RibbonPanel CreateRibbonPanel(UIControlledApplication app, string tabName, string panelName)
        {
            RibbonPanel curPanel;

            if (GetRibbonPanelByName(app, tabName, panelName) == null)
                curPanel = app.CreateRibbonPanel(tabName, panelName);

            else
                curPanel = GetRibbonPanelByName(app, tabName, panelName);

            return curPanel;
        }

        internal static RibbonPanel GetRibbonPanelByName(UIControlledApplication app, string tabName, string panelName)
        {
            foreach (RibbonPanel tmpPanel in app.GetRibbonPanels(tabName))
            {
                if (tmpPanel.Name == panelName)
                    return tmpPanel;
            }

            return null;
        }

        #endregion

        #region Schedules

        /// <summary>
        /// Find the Floor Areas schedule that contains non-zero area values
        /// </summary>
        internal static ViewSchedule GetFloorAreaSchedule(Document curDoc)
        {
            FilteredElementCollector collector = new FilteredElementCollector(curDoc)
                .OfClass(typeof(ViewSchedule));

            foreach (ViewSchedule vs in collector.Cast<ViewSchedule>())
            {
                if (!vs.Name.StartsWith("Floor Areas", StringComparison.OrdinalIgnoreCase))
                    continue;

                TableData tableData = vs.GetTableData();
                TableSectionData bodyData = tableData.GetSectionData(SectionType.Body);

                int rowCount = bodyData.NumberOfRows;
                int areaCol = bodyData.NumberOfColumns - 1;

                for (int row = 0; row < rowCount; row++)
                {
                    string areaText = bodyData.GetCellText(row, areaCol).Trim();
                    string cleaned = areaText.Replace("SF", "").Trim();

                    if (int.TryParse(cleaned, out int areaValue) && areaValue > 0)
                        return vs;
                }
            }

            return null;
        }

        #endregion

        #region Task Dialog

        /// <summary>
        /// Displays an accept/decline dialog to the user with Yes/No buttons
        /// </summary>
        /// <param name="tdName">The internal name of the TaskDialog</param>
        /// <param name="tdTitle">The title displayed in the dialog header</param>
        /// <param name="textMessage">The main message content to display to the user</param>
        /// <returns>True if user clicked Yes, false if user clicked No</returns>
        internal static bool TaskDialogAccept(string tdName, string tdTitle, string textMessage)
        {
            // Create a new TaskDialog with the specified name
            TaskDialog m_Dialog = new TaskDialog(tdName);

            // Set the warning icon to indicate this is a warning message
            m_Dialog.MainIcon = Icon.TaskDialogIconInformation;

            // Set the custom title for the dialog
            m_Dialog.Title = tdTitle;

            // Disable automatic title prefixing to use our custom title exactly as specified
            m_Dialog.TitleAutoPrefix = false;

            // Set the main message content that will be displayed to the user
            m_Dialog.MainContent = textMessage;

            // Add Yes and No buttons for the user to accept or decline
            m_Dialog.CommonButtons = TaskDialogCommonButtons.Yes | TaskDialogCommonButtons.No;

            // Display the dialog and capture the result
            TaskDialogResult m_DialogResult = m_Dialog.Show();

            // Return true if Yes was clicked, false otherwise
            return m_DialogResult == TaskDialogResult.Yes;
        }

        /// <summary>
        /// Displays a warning dialog to the user with custom title and message
        /// </summary>
        /// <param name="tdName">The internal name of the TaskDialog</param>
        /// <param name="tdTitle">The title displayed in the dialog header</param>
        /// <param name="textMessage">The main message content to display to the user</param>
        internal static void TaskDialogWarning(string tdName, string tdTitle, string textMessage)
        {
            // Create a new TaskDialog with the specified name
            TaskDialog m_Dialog = new TaskDialog(tdName);

            // Set the warning icon to indicate this is a warning message
            m_Dialog.MainIcon = Icon.TaskDialogIconWarning;

            // Set the custom title for the dialog
            m_Dialog.Title = tdTitle;

            // Disable automatic title prefixing to use our custom title exactly as specified
            m_Dialog.TitleAutoPrefix = false;

            // Set the main message content that will be displayed to the user
            m_Dialog.MainContent = textMessage;

            // Add a Close button for the user to dismiss the dialog
            m_Dialog.CommonButtons = TaskDialogCommonButtons.Close;

            // Display the dialog and capture the result (though we don't use it for warnings)
            TaskDialogResult m_DialogResult = m_Dialog.Show();
        }

        /// <summary>
        /// Displays an information dialog to the user with custom title and message
        /// </summary>
        /// <param name="tdName">The internal name of the TaskDialog</param>
        /// <param name="tdTitle">The title displayed in the dialog header</param>
        /// <param name="textMessage">The main message content to display to the user</param>
        internal static void TaskDialogInformation(string tdName, string tdTitle, string textMessage)
        {
            // Create a new TaskDialog with the specified name
            TaskDialog m_Dialog = new TaskDialog(tdName);

            // Set the warning icon to indicate this is a warning message
            m_Dialog.MainIcon = Icon.TaskDialogIconInformation;

            // Set the custom title for the dialog
            m_Dialog.Title = tdTitle;

            // Disable automatic title prefixing to use our custom title exactly as specified
            m_Dialog.TitleAutoPrefix = false;

            // Set the main message content that will be displayed to the user
            m_Dialog.MainContent = textMessage;

            // Add a Close button for the user to dismiss the dialog
            m_Dialog.CommonButtons = TaskDialogCommonButtons.Close;

            // Display the dialog and capture the result (though we don't use it for warnings)
            TaskDialogResult m_DialogResult = m_Dialog.Show();
        }

        /// <summary>
        /// Displays an error dialog to the user with custom title and message
        /// </summary>
        /// <param name="tdName">The internal name of the TaskDialog</param>
        /// <param name="tdTitle">The title displayed in the dialog header</param>
        /// <param name="textMessage">The main message content to display to the user</param>
        internal static void TaskDialogError(string tdName, string tdTitle, string textMessage)
        {
            // Create a new TaskDialog with the specified name
            TaskDialog m_Dialog = new TaskDialog(tdName);

            // Set the warning icon to indicate this is a warning message
            m_Dialog.MainIcon = Icon.TaskDialogIconError;

            // Set the custom title for the dialog
            m_Dialog.Title = tdTitle;

            // Disable automatic title prefixing to use our custom title exactly as specified
            m_Dialog.TitleAutoPrefix = false;

            // Set the main message content that will be displayed to the user
            m_Dialog.MainContent = textMessage;

            // Add a Close button for the user to dismiss the dialog
            m_Dialog.CommonButtons = TaskDialogCommonButtons.Close;

            // Display the dialog and capture the result (though we don't use it for warnings)
            TaskDialogResult m_DialogResult = m_Dialog.Show();
        }

        #endregion
    }
}
