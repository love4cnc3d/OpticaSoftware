namespace OpticalShop
{
    using ADOX;
    using System;
    using System.ComponentModel;
    using System.Data;
    using System.Data.OleDb;
    using System.Drawing;
    using System.IO;
    using System.Runtime.InteropServices;
    using System.Windows.Forms;

    public class MainForm : Form
    {
        private OleDbConnection databaseConnection;
        private string databaseConnectionString = 
            "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:/OpticalShop/Database/OpticalDb.mdb;Jet OLEDB:Engine Type=5";

        private IContainer components;
        private TabControl mainTabControl;
        private TabPage tabPageCustomer;
        private TabPage tabPagePrescription;

        private TextBox textBoxCustomerName;
        private TextBox textBoxCustomerContact;
        private TextBox textBoxCustomerAge;
        private TextBox textBoxCustomerAddress;

        private Button buttonSaveCustomer;
        private Button buttonClearCustomer;
        private Button buttonDeleteCustomer;

        private TextBox textBoxLensType;
        private TextBox textBoxLensPrice;
        private Button buttonSavePrescription;
        private Button buttonUpdatePrescription;

        private DataGridView dataGridViewCustomerHistory;
        private DataGridView dataGridViewPrescriptionHistory;

        public MainForm()
        {
            InitializeComponent();
            InitializeFormComponents();
            InitializePrescriptionTab();
        }

        private void InitializeFormComponents()
        {
            // Inicialización de controles para clientes
            textBoxCustomerName = new TextBox
            {
                PlaceholderText = "Enter customer name",
                Location = new Point(20, 30),
                Size = new Size(200, 25)
            };

            textBoxCustomerContact = new TextBox
            {
                PlaceholderText = "Enter contact number",
                Location = new Point(20, 70),
                Size = new Size(200, 25)
            };

            textBoxCustomerAge = new TextBox
            {
                PlaceholderText = "Enter age",
                Location = new Point(20, 110),
                Size = new Size(200, 25)
            };

            textBoxCustomerAddress = new TextBox
            {
                PlaceholderText = "Enter address",
                Location = new Point(20, 150),
                Size = new Size(200, 50),
                Multiline = true
            };

            buttonSaveCustomer = new Button
            {
                Text = "Save Customer",
                Location = new Point(20, 220),
                Size = new Size(100, 30)
            };
            buttonSaveCustomer.Click += SaveCustomer_Click;

            buttonClearCustomer = new Button
            {
                Text = "Clear",
                Location = new Point(130, 220),
                Size = new Size(100, 30)
            };
            buttonClearCustomer.Click += ClearCustomerForm;

            buttonDeleteCustomer = new Button
            {
                Text = "Delete Customer",
                Location = new Point(240, 220),
                Size = new Size(120, 30)
            };
            buttonDeleteCustomer.Click += DeleteCustomer_Click;

            dataGridViewCustomerHistory = new DataGridView
            {
                Location = new Point(20, 270),
                Size = new Size(600, 200),
                AllowUserToAddRows = false,
                ReadOnly = true
            };

            // Tab para clientes
            tabPageCustomer = new TabPage
            {
                Text = "Customer Details"
            };
            tabPageCustomer.Controls.AddRange(new Control[]
            {
                textBoxCustomerName, textBoxCustomerContact, textBoxCustomerAge, textBoxCustomerAddress,
                buttonSaveCustomer, buttonClearCustomer, buttonDeleteCustomer, dataGridViewCustomerHistory
            });

            // Configuración del TabControl
            mainTabControl = new TabControl
            {
                Dock = DockStyle.Fill,
                TabPages = { tabPageCustomer }
            };

            this.Controls.Add(mainTabControl);
        }

        private void InitializePrescriptionTab()
        {
            // Inicialización de controles para recetas
            textBoxLensType = new TextBox
            {
                PlaceholderText = "Enter lens type",
                Location = new Point(20, 30),
                Size = new Size(200, 25)
            };

            textBoxLensPrice = new TextBox
            {
                PlaceholderText = "Enter lens price",
                Location = new Point(20, 70),
                Size = new Size(200, 25)
            };

            buttonSavePrescription = new Button
            {
                Text = "Save Prescription",
                Location = new Point(20, 110),
                Size = new Size(120, 30)
            };
            buttonSavePrescription.Click += SavePrescription_Click;

            buttonUpdatePrescription = new Button
            {
                Text = "Update Prescription",
                Location = new Point(150, 110),
                Size = new Size(120, 30)
            };
            buttonUpdatePrescription.Click += UpdatePrescription_Click;

            dataGridViewPrescriptionHistory = new DataGridView
            {
                Location = new Point(20, 160),
                Size = new Size(600, 200),
                AllowUserToAddRows = false,
                ReadOnly = true
            };

            // Configuración de la pestaña de recetas
            tabPagePrescription = new TabPage
            {
                Text = "Prescriptions"
            };
            tabPagePrescription.Controls.AddRange(new Control[]
            {
                textBoxLensType, textBoxLensPrice, buttonSavePrescription, buttonUpdatePrescription, dataGridViewPrescriptionHistory
            });

            // Agregar la pestaña al TabControl principal
            mainTabControl.TabPages.Add(tabPagePrescription);
        }

        private void SaveCustomer_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(textBoxCustomerName.Text))
            {
                MessageBox.Show("Please enter a valid customer name.", "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            try
            {
                using (var connection = new OleDbConnection(databaseConnectionString))
                {
                    connection.Open();
                    var command = new OleDbCommand(
                        "INSERT INTO Customer (Name, Contact, Age, Address) VALUES (?, ?, ?, ?)",
                        connection);

                    command.Parameters.AddWithValue("Name", textBoxCustomerName.Text);
                    command.Parameters.AddWithValue("Contact", textBoxCustomerContact.Text);
                    command.Parameters.AddWithValue("Age", textBoxCustomerAge.Text);
                    command.Parameters.AddWithValue("Address", textBoxCustomerAddress.Text);

                    command.ExecuteNonQuery();

                    MessageBox.Show("Customer saved successfully.", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error saving customer: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void ClearCustomerForm(object sender, EventArgs e)
        {
            textBoxCustomerName.Clear();
            textBoxCustomerContact.Clear();
            textBoxCustomerAge.Clear();
            textBoxCustomerAddress.Clear();
        }

        private void DeleteCustomer_Click(object sender, EventArgs e)
        {
            MessageBox.Show("Feature to delete customer is not implemented yet.", "Info", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void SavePrescription_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(textBoxLensType.Text) || string.IsNullOrWhiteSpace(textBoxLensPrice.Text))
            {
                MessageBox.Show("Please fill out all fields.", "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            try
            {
                using (var connection = new OleDbConnection(databaseConnectionString))
                {
                    connection.Open();
                    var command = new OleDbCommand(
                        "INSERT INTO Prescriptions (LensType, LensPrice) VALUES (?, ?)",
                        connection);

                    command.Parameters.AddWithValue("LensType", textBoxLensType.Text);
                    command.Parameters.AddWithValue("LensPrice", textBoxLensPrice.Text);

                    command.ExecuteNonQuery();

                    MessageBox.Show("Prescription saved successfully.", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    RefreshPrescriptionHistory();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error saving prescription: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void UpdatePrescription_Click(object sender, EventArgs e)
        {
            MessageBox.Show("Feature to update prescription is not implemented yet.", "Info", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void RefreshPrescriptionHistory()
        {
            try
            {
                using (var connection = new OleDbConnection(databaseConnectionString))
                {
                    connection.Open();
                    var adapter = new OleDbDataAdapter("SELECT * FROM Prescriptions", connection);
                    var dataTable = new DataTable();
                    adapter.Fill(dataTable);

                    dataGridViewPrescriptionHistory.DataSource = dataTable;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error refreshing prescription history: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}
