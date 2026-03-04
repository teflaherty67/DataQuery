using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;


namespace DataQuery
{
    /// <summary>
    /// Interaction logic for frmProjInfo.xaml
    /// </summary>
    public partial class frmProjInfo : Window
    {
        #region Fields

        private readonly Document CurDoc;

        private static readonly List<string> SpecLevels = new()
        {
            "Complete Home",
            "Complete Home Plus",
            "Terrata",
            "N/A"
        };

        private static readonly List<string> ClientNames = new()
        {
            "DRB Group",
            "Lennar Homes",
            "LGI Homes"
        };

        private static readonly List<string> ClientDivisions = new()
        {
            "Central Texas",
            "Dallas-Fort Worth",
            "Florida",
            "Houston",
            "Maryland",
            "Minnesota",
            "Pensylvania",
            "Oklahoma",
            "Southeast",
            "Virginia",
            "West Virginia",
            "Taylor"
        };

        private static readonly List<string> GarageLoadings = new()
        {
            "Front",
            "Side",
            "Rear"
        };

        #endregion

        #region Properties

        public string PlanName => tbxPlanName.Text.Trim();
        public string SpecLevel => cbxSpecLevel.Text.Trim();
        public string ClientName => cbxClientName.Text.Trim();
        public string ClientDivision => cbxClientDivision.Text.Trim();
        public string ClientSubdivision => tbxClientSubdivision.Text.Trim();
        public string GarageLoading => cbxGarageLoading.Text.Trim();

        /// <summary>
        /// Returns the Code Masonry value if the field was shown (parameter was newly added),
        /// or null if the field was hidden (parameter already had values on Cover sheets).
        /// </summary>
        public string CodeMasonry => tbxCodeMasonry.Visibility == Visibility.Visible
            ? tbxCodeMasonry.Text.Trim()
            : null;

        #endregion

        #region Constructor

        public frmProjInfo(Document curDoc)
        {
            InitializeComponent();
            CurDoc = curDoc;
            PopulateDropdowns();
            LoadExistingValues();

            // Register paste handler for numeric-only validation on Code Masonry TextBox
            DataObject.AddPastingHandler(tbxCodeMasonry, OnCodeMasonryPaste);

            // Hide the Code Masonry row if Cover sheets already have values for the parameter
            if (CodeMasonryAlreadyHasValues())
                HideCodeMasonryRow();
        }

        #endregion

        #region Initialization

        private void PopulateDropdowns()
        {
            cbxSpecLevel.ItemsSource = SpecLevels;
            cbxClientName.ItemsSource = ClientNames;
            cbxClientDivision.ItemsSource = ClientDivisions;
            cbxGarageLoading.ItemsSource = GarageLoadings;
        }

        private void LoadExistingValues()
        {
            ProjectInfo projInfo = CurDoc.ProjectInformation;

            tbxPlanName.Text = Common.Utils.GetParameterValueByName(projInfo, "Project Name") ?? string.Empty;
            tbxClientSubdivision.Text = Common.Utils.GetParameterValueByName(projInfo, "Client Subdivision") ?? string.Empty;

            SetComboValue(cbxSpecLevel, Common.Utils.GetParameterValueByName(projInfo, "Spec Level"));
            SetComboValue(cbxClientDivision, Common.Utils.GetParameterValueByName(projInfo, "Client Division"));
            SetComboValue(cbxGarageLoading, Common.Utils.GetParameterValueByName(projInfo, "Garage Loading"));

            // Client Name is a built-in Revit parameter — use AsString() to avoid
            // AsValueString() returning the parameter name when the value is empty
            Parameter clientNameParam = Common.Utils.GetParameterByName(projInfo, "Client Name");
            SetComboValue(cbxClientName, clientNameParam?.AsString());
        }

        private static void SetComboValue(ComboBox combo, string value)
        {
            if (string.IsNullOrEmpty(value)) return;

            int index = combo.Items.IndexOf(value);
            if (index >= 0)
                combo.SelectedIndex = index;
            else
                combo.Text = value;
        }

        /// <summary>
        /// Returns true if any Cover sheet already has a non-zero Code Masonry value,
        /// indicating the parameter was populated before this command ran.
        /// Returns false if the parameter was just added (default value of 0) or doesn't exist.
        /// </summary>
        private bool CodeMasonryAlreadyHasValues()
        {
            return new FilteredElementCollector(CurDoc)
                .OfClass(typeof(ViewSheet))
                .Cast<ViewSheet>()
                .Where(s => s.Name.IndexOf("Cover", StringComparison.OrdinalIgnoreCase) >= 0
                         && !s.IsTemplate)
                .Any(s =>
                {
                    Parameter p = Common.Utils.GetParameterByName(s, "Code Masonry");
                    return p != null && p.AsDouble() > 0;
                });
        }

        /// <summary>
        /// Collapses the Code Masonry label, TextBox, and its spacer row.
        /// </summary>
        private void HideCodeMasonryRow()
        {
            gridContent.RowDefinitions[11].Height = new GridLength(0); // spacer before Code Masonry
            gridContent.RowDefinitions[12].Height = new GridLength(0); // Code Masonry row
            lblCodeMasonry.Visibility = Visibility.Collapsed;
            tbxCodeMasonry.Visibility = Visibility.Collapsed;
        }

        #endregion

        #region Validation

        private bool ValidateInputs(out string errorMessage)
        {
            errorMessage = string.Empty;
            var missing = new List<string>();

            if (string.IsNullOrWhiteSpace(tbxPlanName.Text)) missing.Add("Plan Name");
            if (string.IsNullOrWhiteSpace(cbxSpecLevel.Text)) missing.Add("Spec Level");
            if (string.IsNullOrWhiteSpace(cbxClientName.Text)) missing.Add("Client Name");
            if (string.IsNullOrWhiteSpace(cbxClientDivision.Text)) missing.Add("Client Division");
            if (string.IsNullOrWhiteSpace(tbxClientSubdivision.Text)) missing.Add("Client Subdivision");
            if (string.IsNullOrWhiteSpace(cbxGarageLoading.Text)) missing.Add("Garage Loading");

            if (missing.Count > 0)
            {
                errorMessage = "The following fields are required:\n\n" +
                               string.Join("\n", missing.Select(f => $"  \u2022 {f}"));
                return false;
            }

            return true;
        }

        #endregion

        #region Event Handlers

        private void btnOK_Click(object sender, RoutedEventArgs e)
        {
            if (!ValidateInputs(out string errorMessage))
            {
                Common.Utils.TaskDialogWarning("frmProjInfo", "Missing Information", errorMessage);
                return;
            }

            DialogResult = true;
            Close();
        }

        private void btnCancel_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }

        private void btnHelp_Click(object sender, RoutedEventArgs e)
        {
            // Add help content here
        }

        private void TitleBar_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            DragMove();
        }

        /// <summary>
        /// Restricts Code Masonry TextBox to numeric input only.
        /// </summary>
        private void tbxCodeMasonry_PreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            e.Handled = !e.Text.All(char.IsDigit);
        }

        /// <summary>
        /// Restricts Code Masonry TextBox to numeric paste only.
        /// </summary>
        private void OnCodeMasonryPaste(object sender, DataObjectPastingEventArgs e)
        {
            if (e.DataObject.GetDataPresent(typeof(string)))
            {
                string text = (string)e.DataObject.GetData(typeof(string));
                if (!text.All(char.IsDigit))
                    e.CancelCommand();
            }
            else
            {
                e.CancelCommand();
            }
        }

        #endregion
    }
}