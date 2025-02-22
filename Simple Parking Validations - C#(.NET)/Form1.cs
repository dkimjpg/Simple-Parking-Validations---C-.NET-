using System;
using System.Data;
using System.Drawing;
using System.IO;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace ParkingValidation
{
    public partial class MainForm : Form
    {
        private Panel currentPanel;
        private readonly string employeeLogsPath;
        private readonly string parkingCodesPath;

        public MainForm()
        {
            InitializeComponent();
            CenterToScreen();

            // Initialize paths
            string baseDir = AppDomain.CurrentDomain.BaseDirectory;
            employeeLogsPath = Path.Combine(baseDir, "employee_logs");
            parkingCodesPath = Path.Combine(baseDir, "parking_codes");

            // Ensure directories exist
            Directory.CreateDirectory(employeeLogsPath);
            Directory.CreateDirectory(parkingCodesPath);

            ShowMainPage();
        }

        private void InitializeComponent()
        {
            this.Size = new Size(1000, 720);
            this.Text = "Parking Validations";
            this.StartPosition = FormStartPosition.CenterScreen;
        }

        private void ShowMainPage()
        {
            if (currentPanel != null)
                this.Controls.Remove(currentPanel);

            var panel = new Panel
            {
                Dock = DockStyle.Fill,
                Padding = new Padding(40)
            };

            var welcomeLabel = new Label
            {
                Text = "Welcome!",
                Font = new Font("Segoe UI", 28, FontStyle.Bold),
                TextAlign = ContentAlignment.MiddleCenter,
                Dock = DockStyle.Top,
                Height = 60
            };

            var optionLabel = new Label
            {
                Text = "Please select an option below.",
                Font = new Font("Segoe UI", 12, FontStyle.Bold),
                TextAlign = ContentAlignment.MiddleCenter,
                Dock = DockStyle.Top,
                Height = 40
            };

            var buttonPanel = new FlowLayoutPanel
            {
                FlowDirection = FlowDirection.LeftToRight,
                Dock = DockStyle.Top,
                Height = 40,
                Padding = new Padding(10)
            };

            var parkingValButton = new Button
            {
                Text = "Parking Validation",
                Padding = new Padding(10, 5, 10, 5),
                Width = 150
            };
            parkingValButton.Click += (s, e) => ShowParkingValidation();

            var tempBadgeButton = new Button
            {
                Text = "Temporary Badge",
                Padding = new Padding(10, 5, 10, 5),
                Width = 150
            };
            tempBadgeButton.Click += (s, e) => ShowTemporaryBadgeMessage();

            buttonPanel.Controls.AddRange(new Control[] { parkingValButton, tempBadgeButton });
            panel.Controls.AddRange(new Control[] { welcomeLabel, optionLabel, buttonPanel });

            currentPanel = panel;
            this.Controls.Add(panel);
        }

        private void ShowParkingValidation()
        {
            if (currentPanel != null)
                this.Controls.Remove(currentPanel);

            var panel = new Panel
            {
                Dock = DockStyle.Fill,
                Padding = new Padding(40)
            };

            var entryLabel = new Label
            {
                Text = "Please enter your Name and ID below.",
                Dock = DockStyle.Top,
                Height = 30
            };

            var nameLabel = new Label { Text = "Name:", Width = 100 };
            var nameTextBox = new TextBox { Width = 200 };

            var idLabel = new Label { Text = "ID:", Width = 100 };
            var idTextBox = new TextBox { Width = 200 };

            var buttonPanel = new FlowLayoutPanel
            {
                FlowDirection = FlowDirection.LeftToRight,
                Dock = DockStyle.Top,
                Height = 40
            };

            var cancelButton = new Button
            {
                Text = "Cancel",
                Padding = new Padding(10, 5, 10, 5)
            };
            cancelButton.Click += (s, e) => ShowMainPage();

            var submitButton = new Button
            {
                Text = "Submit",
                Padding = new Padding(10, 5, 10, 5),
                Enabled = false
            };
            submitButton.Click += (s, e) => SubmitParkingValidation(nameTextBox.Text, idTextBox.Text);

            var prepaidButton = new Button
            {
                Text = "I have a prepaid code",
                Padding = new Padding(10, 5, 10, 5)
            };
            prepaidButton.Click += (s, e) => ShowPrepaidCode();

            EventHandler textChanged = (s, e) =>
            {
                submitButton.Enabled = !string.IsNullOrWhiteSpace(nameTextBox.Text) &&
                                     !string.IsNullOrWhiteSpace(idTextBox.Text);
            };
            nameTextBox.TextChanged += textChanged;
            idTextBox.TextChanged += textChanged;

            var namePanel = new FlowLayoutPanel { FlowDirection = FlowDirection.LeftToRight, Dock = DockStyle.Top, Height = 30 };
            namePanel.Controls.AddRange(new Control[] { nameLabel, nameTextBox });

            var idPanel = new FlowLayoutPanel { FlowDirection = FlowDirection.LeftToRight, Dock = DockStyle.Top, Height = 30 };
            idPanel.Controls.AddRange(new Control[] { idLabel, idTextBox });

            buttonPanel.Controls.AddRange(new Control[] { cancelButton, submitButton });

            panel.Controls.AddRange(new Control[] { entryLabel, namePanel, idPanel, buttonPanel, prepaidButton });

            currentPanel = panel;
            this.Controls.Add(panel);
        }

        private void ShowPrepaidCode()
        {
            if (currentPanel != null)
                this.Controls.Remove(currentPanel);

            var panel = new Panel
            {
                Dock = DockStyle.Fill,
                Padding = new Padding(40)
            };

            var entryLabel = new Label
            {
                Text = "Please enter your Name, Provider, and Prepaid Code below.",
                Dock = DockStyle.Top,
                Height = 30
            };

            var guestLabel = new Label { Text = "Name:", Width = 100 };
            var guestTextBox = new TextBox { Width = 200 };

            var providerLabel = new Label { Text = "Provider:", Width = 100 };
            var providerTextBox = new TextBox { Width = 200 };

            var codeLabel = new Label { Text = "Prepaid Code:", Width = 100 };
            var codeTextBox = new TextBox { Width = 200 };

            var buttonPanel = new FlowLayoutPanel
            {
                FlowDirection = FlowDirection.LeftToRight,
                Dock = DockStyle.Top,
                Height = 40
            };

            var cancelButton = new Button
            {
                Text = "Cancel",
                Padding = new Padding(10, 5, 10, 5)
            };
            cancelButton.Click += (s, e) => ShowMainPage();

            var submitButton = new Button
            {
                Text = "Submit",
                Padding = new Padding(10, 5, 10, 5),
                Enabled = false
            };
            submitButton.Click += (s, e) => SubmitPrepaidCode(guestTextBox.Text, providerTextBox.Text, codeTextBox.Text);

            var regularButton = new Button
            {
                Text = "I don't have a prepaid code",
                Padding = new Padding(10, 5, 10, 5)
            };
            regularButton.Click += (s, e) => ShowParkingValidation();

            EventHandler textChanged = (s, e) =>
            {
                submitButton.Enabled = !string.IsNullOrWhiteSpace(guestTextBox.Text) &&
                                     !string.IsNullOrWhiteSpace(providerTextBox.Text) &&
                                     !string.IsNullOrWhiteSpace(codeTextBox.Text);
            };

            guestTextBox.TextChanged += textChanged;
            providerTextBox.TextChanged += textChanged;
            codeTextBox.TextChanged += textChanged;

            var guestPanel = new FlowLayoutPanel { FlowDirection = FlowDirection.LeftToRight, Dock = DockStyle.Top, Height = 30 };
            guestPanel.Controls.AddRange(new Control[] { guestLabel, guestTextBox });

            var providerPanel = new FlowLayoutPanel { FlowDirection = FlowDirection.LeftToRight, Dock = DockStyle.Top, Height = 30 };
            providerPanel.Controls.AddRange(new Control[] { providerLabel, providerTextBox });

            var codePanel = new FlowLayoutPanel { FlowDirection = FlowDirection.LeftToRight, Dock = DockStyle.Top, Height = 30 };
            codePanel.Controls.AddRange(new Control[] { codeLabel, codeTextBox });

            buttonPanel.Controls.AddRange(new Control[] { cancelButton, submitButton });

            panel.Controls.AddRange(new Control[] {
                entryLabel,
                guestPanel,
                providerPanel,
                codePanel,
                buttonPanel,
                regularButton
            });

            currentPanel = panel;
            this.Controls.Add(panel);
        }

        private void ShowValidationCode()
        {
            if (currentPanel != null)
                this.Controls.Remove(currentPanel);

            var panel = new Panel
            {
                Dock = DockStyle.Fill,
                Padding = new Padding(40)
            };

            string code;
            try
            {
                code = GetNextParkingCode();
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    "Error generating parking code. Please try again or contact support.",
                    "Error",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error
                );
                ShowMainPage();
                return;
            }

            var successLabel = new Label
            {
                Text = "Your code!",
                Dock = DockStyle.Top,
                Height = 30,
                TextAlign = ContentAlignment.MiddleCenter
            };

            var codeLabel = new Label
            {
                Text = code,
                Dock = DockStyle.Top,
                Height = 40,
                Font = new Font("Segoe UI", 16, FontStyle.Bold),
                TextAlign = ContentAlignment.MiddleCenter
            };

            var buttonPanel = new FlowLayoutPanel
            {
                FlowDirection = FlowDirection.TopDown,
                Dock = DockStyle.Top,
                Height = 120,
                Padding = new Padding(10)
            };

            var tutorialButton = new Button
            {
                Text = "Learn more about how to use a Parking Validation Code",
                Width = 300,
                Padding = new Padding(10, 5, 10, 5)
            };
            tutorialButton.Click += (s, e) => ShowTutorial();

            var homeButton = new Button
            {
                Text = "Return to Start",
                Width = 300,
                Padding = new Padding(10, 5, 10, 5)
            };
            homeButton.Click += (s, e) => ShowMainPage();

            var newCodeButton = new Button
            {
                Text = "Code not working? Get a new one here.",
                Width = 300,
                Padding = new Padding(10, 5, 10, 5)
            };
            newCodeButton.Click += (s, e) => ShowParkingValidation();

            buttonPanel.Controls.AddRange(new Control[] { tutorialButton, homeButton, newCodeButton });
            panel.Controls.AddRange(new Control[] { successLabel, codeLabel, buttonPanel });

            currentPanel = panel;
            this.Controls.Add(panel);
        }

        private void ShowTemporaryBadgeMessage()
        {
            MessageBox.Show(
                "Ask the receptionist about getting a temporary badge, or call the service center if there is no receptionist present.",
                "Temporary Badge Info",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information
            );
        }

        private void SubmitParkingValidation(string name, string id)
        {
            SaveParkingData(name, id);
            ShowValidationCode();
        }

        private void SubmitPrepaidCode(string guest, string provider, string code)
        {
            string filePath = Path.Combine(employeeLogsPath, "parkingDataTestIOCC.xlsx");

            // Create file if it doesn't exist
            /*
            if (!File.Exists(filePath))
            {
                Excel.Application excel = new Excel.Application();
                Excel.Workbook workbook = null;
                try
                {
                    workbook = excel.Workbooks.Add();
                    Excel.Worksheet worksheet = (Excel.Worksheet)workbook.Sheets[1];

                    // Add headers
                    worksheet.Cells[1, 1].Value = "Date";
                    worksheet.Cells[1, 2].Value = "Code Provider";
                    worksheet.Cells[1, 3].Value = "Guest";
                    worksheet.Cells[1, 4].Value = "Code";

                    workbook.SaveAs(filePath);
                }
                finally
                {
                    workbook?.Close();
                    excel.Quit();
                    if (workbook != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook);
                    if (excel != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(excel);
                }
            }
            */

            Excel.Application excelApp = new Excel.Application();
            Excel.Workbook wb = null;
            Excel.Worksheet ws = null;

            try
            {
                wb = excelApp.Workbooks.Open(filePath);
                ws = (Excel.Worksheet)wb.Sheets[1];

                int lastRow = ws.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row + 1;

                ws.Cells[lastRow, 1].Value = DateTime.Now.ToString("MM/dd/yyyy");
                ws.Cells[lastRow, 2].Value = provider;
                ws.Cells[lastRow, 3].Value = guest;
                ws.Cells[lastRow, 4].Value = code;

                wb.Save();
                ShowValidationCode();
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    "Error saving prepaid code data: " + ex.Message,
                    "Error",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error
                );
            }
            finally
            {
                wb?.Close();
                excelApp.Quit();

                if (ws != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(ws);
                if (wb != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(wb);
                if (excelApp != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);
            }
        }

        private void ShowTutorial()
        {
            MessageBox.Show(
                "To use your parking validation code:\n\n" +
                "1. Park in the designated visitor parking area\n" +
                "2. Take a parking ticket when entering\n" +
                "3. Before leaving, go to the parking payment kiosk\n" +
                "4. Select 'Use Validation Code'\n" +
                "5. Enter the code shown above\n" +
                "6. The parking fee will be automatically adjusted\n\n" +
                "If you experience any issues, please contact the front desk.",
                "How to Use Parking Validation",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information
            );
        }

        private string GetNextParkingCode()
        {
            string codePath = Path.Combine(parkingCodesPath, "parkingCodeTest1.xlsx");

            // Create file if it doesn't exist
            if (!File.Exists(codePath))
            {
                CreateNewParkingCodeFile(codePath);
            }

            Excel.Application excel = new Excel.Application();
            Excel.Workbook workbook = null;
            Excel.Worksheet worksheet = null;

            try
            {
                workbook = excel.Workbooks.Open(codePath);
                worksheet = (Excel.Worksheet)workbook.Sheets[1];

                // Get current row and code
                int currentRow = (int)worksheet.Cells[1, 3].Value;
                string code = worksheet.Cells[currentRow, 1].Value?.ToString();

                // Update current row
                worksheet.Cells[1, 3].Value = currentRow + 1;
                workbook.Save();

                return code ?? throw new Exception("Invalid parking code retrieved");
            }
            finally
            {
                workbook?.Close();
                excel.Quit();

                if (worksheet != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(worksheet);
                if (workbook != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook);
                if (excel != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(excel);
            }
        }

        private void CreateNewParkingCodeFile(string path)
        {
            Excel.Application excel = new Excel.Application();
            Excel.Workbook workbook = null;

            try
            {
                workbook = excel.Workbooks.Add();
                Excel.Worksheet worksheet = (Excel.Worksheet)workbook.Sheets[1];

                // Initialize with some codes and current row pointer
                worksheet.Cells[1, 1].Value = "CODE1";
                worksheet.Cells[2, 1].Value = "CODE2";
                worksheet.Cells[3, 1].Value = "CODE3";
                worksheet.Cells[1, 3].Value = 1; // Current row pointer

                workbook.SaveAs(path);
            }
            finally
            {
                workbook?.Close();
                excel.Quit();

                if (workbook != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook);
                if (excel != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(excel);
            }
        }

        private void SaveParkingData(string name, string id)
        {
            string filePath = Path.Combine(employeeLogsPath, "parkingDataTest1.xlsx");

            // Create file if it doesn't exist
            if (!File.Exists(filePath))
            {
                Excel.Application excel = new Excel.Application();
                Excel.Workbook workbook = null;
                try
                {
                    workbook = excel.Workbooks.Add();
                    Excel.Worksheet worksheet = (Excel.Worksheet)workbook.Sheets[1];

                    // Add headers
                    worksheet.Cells[1, 1].Value = "Date";
                    worksheet.Cells[1, 2].Value = "ID";
                    worksheet.Cells[1, 3].Value = "Name";

                    workbook.SaveAs(filePath);
                }
                finally
                {
                    workbook?.Close();
                    excel.Quit();
                    if (workbook != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook);
                    if (excel != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(excel);
                }
            }

            Excel.Application excelApp = new Excel.Application();
            Excel.Workbook wb = null;
            Excel.Worksheet ws = null;

            try
            {
                wb = excelApp.Workbooks.Open(filePath);
                ws = (Excel.Worksheet)wb.Sheets[1];

                // Find last row
                int lastRow = ws.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row + 1;

                // Add new record
                ws.Cells[lastRow, 1].Value = DateTime.Now.ToString("MM/dd/yyyy");
                ws.Cells[lastRow, 2].Value = id;
                ws.Cells[lastRow, 3].Value = name;

                wb.Save();
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    "Error saving parking data: " + ex.Message,
                    "Error",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error
                );
            }
            finally
            {
                wb?.Close();
                excelApp.Quit();

                if (ws != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(ws);
                if (wb != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(wb);
                if (excelApp != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);
            }
        }

    }
}

        