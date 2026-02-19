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
            // Add your client names here
            "DRB Group",
            "Lennar Homes",
            "LGI Homes"
        };

        private static readonly List<string> ClientDivisions = new()
        {
            // Add your division names here
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

        #endregion

        #region Constructor

        public frmProjInfo(Document curDoc)
        {
            InitializeComponent();
            CurDoc = curDoc;
            PopulateDropdowns();
            LoadExistingValues();
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


        private static void SetComboValue(System.Windows.Controls.ComboBox combo, string value)
        {
            if (string.IsNullOrEmpty(value)) return;

            int index = combo.Items.IndexOf(value);
            if (index >= 0)
                combo.SelectedIndex = index;
            else
                combo.Text = value;
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

        #endregion
    }
}
