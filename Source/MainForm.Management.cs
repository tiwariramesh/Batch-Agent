using System;
using System.Drawing;
using System.IO;
using System.Threading;
using System.Windows.Forms;
using System.IO.Compression;

namespace DocumentConverter
{
    public partial class MainForm
    {
        private string manageInputDir = "";
        private string manageOutputDir = "";
        private bool isExtracting = false;
        
        private ListBox manageInputList;
        private Label manageInputStatus;
        private Label manageOutputLabel;
        private AnimatedFlatButton manageStartBtn;
        private AnimatedFlatButton manageInputBtn;
        private AnimatedFlatButton manageOutputBtn;
        private ProgressBar manageProgressBar;
        private Label manageStatusLabel;

        private void SetupManagementViewLayout()
        {
            Panel headerPanel = new Panel();
            headerPanel.Dock = DockStyle.Top;
            headerPanel.Height = 60;
            Label lblHeader = new Label();
            lblHeader.Text = "Batch File Management";
            lblHeader.Font = new Font("Segoe UI", 16F, FontStyle.Bold);
            lblHeader.AutoSize = true;
            lblHeader.Location = new Point(10, 20);
            headerPanel.Controls.Add(lblHeader);
            
            fileManagementView.Controls.Add(headerPanel);

            TableLayoutPanel tlp = new TableLayoutPanel();
            tlp.Dock = DockStyle.Fill;
            tlp.ColumnCount = 2;
            tlp.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50F));
            tlp.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50F));
            tlp.RowCount = 1;
            tlp.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));
            tlp.Padding = new Padding(10, 10, 10, 20);
            
            fileManagementView.Controls.Add(tlp);
            tlp.BringToFront(); 

            Panel leftPanel = new Panel();
            leftPanel.Dock = DockStyle.Fill;
            leftPanel.BackColor = panelBg;
            leftPanel.Margin = new Padding(10);
            tlp.Controls.Add(leftPanel, 0, 0);

            Panel rightPanel = new Panel();
            rightPanel.Dock = DockStyle.Fill;
            rightPanel.BackColor = panelBg;
            rightPanel.Margin = new Padding(10);
            tlp.Controls.Add(rightPanel, 1, 0);

            Label lblInputSection = new Label { Text = "1. Source Selection (.zip, .rar)", Font = new Font("Segoe UI", 12F, FontStyle.Bold), Location = new Point(20, 20), AutoSize = true };
            leftPanel.Controls.Add(lblInputSection);

            manageInputBtn = new AnimatedFlatButton();
            manageInputBtn.Text = "Browse Archives Folder";
            manageInputBtn.Size = new Size(200, 35);
            manageInputBtn.Location = new Point(20, 60);
            manageInputBtn.Click += delegate(object s, EventArgs e) { ManageBrowseFolder(true); };
            leftPanel.Controls.Add(manageInputBtn);

            manageInputList = new ListBox();
            manageInputList.BackColor = Color.FromArgb(35, 35, 35);
            manageInputList.ForeColor = Color.White;
            manageInputList.BorderStyle = BorderStyle.None;
            manageInputList.Location = new Point(20, 105);
            manageInputList.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
            manageInputList.SelectionMode = SelectionMode.None;
            leftPanel.Controls.Add(manageInputList);
            leftPanel.Resize += delegate(object s, EventArgs e) { manageInputList.Height = leftPanel.Height - 165; };

            Panel statusBox = new Panel();
            statusBox.BackColor = Color.FromArgb(50, 50, 50);
            statusBox.Height = 40;
            statusBox.Location = new Point(20, 200);
            statusBox.Anchor = AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
            leftPanel.Controls.Add(statusBox);
            leftPanel.Resize += delegate(object s, EventArgs e) { statusBox.Location = new Point(20, leftPanel.Height - 50); };

            manageInputStatus = new Label();
            manageInputStatus.Text = "Selected: 0 files ready to process.";
            manageInputStatus.AutoSize = true;
            manageInputStatus.Location = new Point(10, 10);
            manageInputStatus.ForeColor = Color.LightGray;
            statusBox.Controls.Add(manageInputStatus);

            Label lblOutputSection = new Label { Text = "2. Extraction Settings", Font = new Font("Segoe UI", 12F, FontStyle.Bold), Location = new Point(20, 20), AutoSize = true };
            rightPanel.Controls.Add(lblOutputSection);
            
            Label lblOutputFolder = new Label { Text = "Destination Folder:", Location = new Point(20, 60), AutoSize = true, ForeColor = Color.LightGray };
            rightPanel.Controls.Add(lblOutputFolder);

            manageOutputBtn = new AnimatedFlatButton();
            manageOutputBtn.Text = "Browse Extract To";
            manageOutputBtn.Size = new Size(150, 35);
            manageOutputBtn.Location = new Point(20, 85);
            manageOutputBtn.Click += delegate(object s, EventArgs e) { ManageBrowseFolder(false); };
            rightPanel.Controls.Add(manageOutputBtn);

            manageOutputLabel = new Label();
            manageOutputLabel.Text = "Same as input folder (Default)";
            manageOutputLabel.ForeColor = Color.Gray;
            manageOutputLabel.AutoSize = false;
            manageOutputLabel.AutoEllipsis = true;
            manageOutputLabel.Location = new Point(180, 92);
            manageOutputLabel.Size = new Size(200, 25);
            manageOutputLabel.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
            rightPanel.Controls.Add(manageOutputLabel);

            Panel advancedGroup = new Panel();
            advancedGroup.BackColor = Color.FromArgb(35, 35, 35);
            advancedGroup.Location = new Point(20, 140);
            advancedGroup.Size = new Size(250, 100);
            advancedGroup.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
            rightPanel.Controls.Add(advancedGroup);
            
            Label lblAdv = new Label { Text = "Advanced File Operations", Font = new Font("Segoe UI", 9F, FontStyle.Bold), Location = new Point(10, 10), AutoSize = true, ForeColor = Color.LightGray };
            advancedGroup.Controls.Add(lblAdv);
            Label lblAdvDesc = new Label { Text = "RAR extraction requires external tools.\nZIP extraction is natively supported.", Location = new Point(10, 35), AutoSize = true, ForeColor = Color.Gray };
            advancedGroup.Controls.Add(lblAdvDesc);

            manageStartBtn = new AnimatedFlatButton();
            manageStartBtn.IsPrimary = true;
            manageStartBtn.Text = "EXECUTE EXTRACTION";
            manageStartBtn.Font = new Font("Segoe UI", 12F, FontStyle.Bold);
            manageStartBtn.Click += new EventHandler(ManageStartBtn_Click);
            rightPanel.Controls.Add(manageStartBtn);

            manageStatusLabel = new Label();
            manageStatusLabel.Text = "Ready for execution";
            manageStatusLabel.AutoSize = true;
            rightPanel.Controls.Add(manageStatusLabel);

            manageProgressBar = new ProgressBar();
            rightPanel.Controls.Add(manageProgressBar);

            rightPanel.Resize += delegate(object s, EventArgs e) 
            {
                manageStartBtn.Size = new Size(rightPanel.Width - 40, 50);
                manageStartBtn.Location = new Point(20, rightPanel.Height - 140);
                
                manageStatusLabel.Location = new Point(20, rightPanel.Height - 75);
                
                manageProgressBar.Size = new Size(rightPanel.Width - 40, 15);
                manageProgressBar.Location = new Point(20, rightPanel.Height - 50);
            };
        }

        private void ManageBrowseFolder(bool isInput)
        {
            using (FolderBrowserDialog fbd = new FolderBrowserDialog())
            {
                fbd.Description = isInput ? "Select Archive Folder (.zip, .rar)" : "Select Extraction Folder";
                if (fbd.ShowDialog() == DialogResult.OK)
                {
                    if (isInput)
                    {
                        manageInputDir = fbd.SelectedPath;
                        if (string.IsNullOrEmpty(manageOutputDir)) 
                        {
                            manageOutputDir = manageInputDir;
                            manageOutputLabel.Text = manageOutputDir + " (Default)";
                            manageOutputLabel.ForeColor = textColor;
                        }
                        
                        manageInputList.Items.Clear();
                        try
                        {
                            var files = Directory.GetFiles(manageInputDir, "*.*", SearchOption.TopDirectoryOnly);
                            int count = 0;
                            foreach(var f in files)
                            {
                                if(f.EndsWith(".zip", StringComparison.OrdinalIgnoreCase) || f.EndsWith(".rar", StringComparison.OrdinalIgnoreCase))
                                {
                                    manageInputList.Items.Add(Path.GetFileName(f));
                                    count++;
                                }
                            }
                            manageInputStatus.Text = string.Format("Selected: {0} archives ready to process.", count);
                        }
                        catch
                        {
                            manageInputStatus.Text = "Error reading folder.";
                        }
                    }
                    else
                    {
                        manageOutputDir = fbd.SelectedPath;
                        manageOutputLabel.Text = manageOutputDir;
                        manageOutputLabel.ForeColor = textColor;
                    }
                }
            }
        }

        private void ManageUpdateProgress(int current, int total, string text)
        {
            if (this.InvokeRequired)
            {
                this.Invoke(new Action(() => ManageUpdateProgress(current, total, text)));
                return;
            }

            manageStatusLabel.Text = text;
            if (total > 0)
            {
                manageProgressBar.Maximum = total;
                manageProgressBar.Value = Math.Min(current, total);
            }
            else
            {
                manageProgressBar.Value = 0;
            }
        }

        private void ManageStartBtn_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(manageInputDir) || string.IsNullOrWhiteSpace(manageOutputDir))
            {
                MessageBox.Show("Please specify both source and destination folders.", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (manageInputList.Items.Count == 0)
            {
                MessageBox.Show("No .zip or .rar files found in the source directory.", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (!Directory.Exists(manageOutputDir))
            {
                try { Directory.CreateDirectory(manageOutputDir); }
                catch (Exception ex)
                {
                    MessageBox.Show(string.Format("Failed to create destination folder: {0}", ex.Message), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
            }

            SetUiState(false);
            isExtracting = true;
            ManageUpdateProgress(0, 1, "Initializing extraction...");

            Thread workerThread = new Thread(() => RunExtractionProcess(manageInputDir, manageOutputDir));
            workerThread.IsBackground = true;
            workerThread.Start();
        }

        private void RunExtractionProcess(string inDir, string outDir)
        {
            try
            {
                var files = Directory.GetFiles(inDir, "*.*", SearchOption.TopDirectoryOnly);
                var validFiles = new System.Collections.Generic.List<string>();
                
                foreach(var f in files)
                {
                    if (f.EndsWith(".zip", StringComparison.OrdinalIgnoreCase) || f.EndsWith(".rar", StringComparison.OrdinalIgnoreCase))
                    {
                        validFiles.Add(f);
                    }
                }

                int total = validFiles.Count;
                int warnings = 0;

                for (int i = 0; i < total; i++)
                {
                    if (!isExtracting) break;

                    string file = validFiles[i];
                    string fileName = Path.GetFileName(file);
                    ManageUpdateProgress(i, total, string.Format("Extracting ({0}/{1}): {2}", i + 1, total, fileName));

                    if (file.EndsWith(".zip", StringComparison.OrdinalIgnoreCase))
                    {
                        string extractPath = Path.Combine(outDir, Path.GetFileNameWithoutExtension(fileName));
                        if (!Directory.Exists(extractPath))
                        {
                            Directory.CreateDirectory(extractPath);
                        }
                        
                        try
                        {
                            // Requires System.IO.Compression.FileSystem.dll
                            ZipFile.ExtractToDirectory(file, extractPath);
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(string.Format("Failed ZIP extraction {0}: {1}", file, ex.Message));
                            warnings++;
                        }
                    }
                    else if (file.EndsWith(".rar", StringComparison.OrdinalIgnoreCase))
                    {
                        Console.WriteLine(string.Format("Skipping RAR {0}: Native extraction not supported.", file));
                        warnings++;
                    }
                }

                if (isExtracting)
                {
                    ManageUpdateProgress(total, total, warnings > 0 ? string.Format("Completed with {0} warnings (RAR unsupported).", warnings) : "All extractions completed securely.");
                    this.Invoke(new Action(() => MessageBox.Show(warnings > 0 ? "Batch extraction completed, but some files (like RAR) could not be processed natively." : "Batch extraction process executed successfully!", "Process Complete", MessageBoxButtons.OK, MessageBoxIcon.Information)));
                }
            }
            catch (Exception ex)
            {
                ManageUpdateProgress(0, 1, string.Format("Error: {0}", ex.Message));
                this.Invoke(new Action(() => MessageBox.Show(ex.Message, "Execution Fault", MessageBoxButtons.OK, MessageBoxIcon.Error)));
            }
            finally
            {
                isExtracting = false;
                SetUiState(true);
            }
        }
    }
}
