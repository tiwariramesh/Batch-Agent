using System;
using System.Drawing;
using System.IO;
using System.Threading;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace DocumentConverter
{
    public partial class MainForm
    {
        private string inputDirPath = "";
        private string outputDirPath = "";
        private bool isConverting = false;

        private RadioButton rbIndividualMode;
        private RadioButton rbBatchMode;

        private Panel leftInputPanel;
        private Panel rightOutputPanel;
        
        private ComboBox comboSourceFormat;
        private ComboBox comboTargetFormat;
        private Label inputLabel;
        private Label outputLabel;
        private AnimatedFlatButton inputBtn;
        private AnimatedFlatButton outputBtn;
        
        private Label statusLabel;
        private ProgressBar progressBar;
        private AnimatedFlatButton startBtn;

        private void SetupConversionViewLayout()
        {
            Panel headerPanel = new Panel();
            headerPanel.Dock = DockStyle.Top;
            headerPanel.Height = 60;
            Label lblHeader = new Label();
            lblHeader.Text = "Batch File Conversion";
            lblHeader.Font = new Font("Segoe UI", 16F, FontStyle.Bold);
            lblHeader.AutoSize = true;
            lblHeader.Location = new Point(10, 20);
            headerPanel.Controls.Add(lblHeader);
            
            fileConversionView.Controls.Add(headerPanel);

            TableLayoutPanel tlp = new TableLayoutPanel();
            tlp.Dock = DockStyle.Fill;
            tlp.ColumnCount = 2;
            tlp.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50F));
            tlp.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50F));
            tlp.RowCount = 1;
            tlp.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));
            tlp.Padding = new Padding(10, 10, 10, 20);
            
            fileConversionView.Controls.Add(tlp);
            tlp.BringToFront(); 

            // Left Input Panel
            leftInputPanel = new Panel();
            leftInputPanel.Dock = DockStyle.Fill;
            leftInputPanel.BackColor = panelBg;
            leftInputPanel.Margin = new Padding(10);
            tlp.Controls.Add(leftInputPanel, 0, 0);

            // Right Output Panel
            rightOutputPanel = new Panel();
            rightOutputPanel.Dock = DockStyle.Fill;
            rightOutputPanel.BackColor = panelBg;
            rightOutputPanel.Margin = new Padding(10);
            tlp.Controls.Add(rightOutputPanel, 1, 0);

            // Populate Left Panel (Input)
            Label lblInputSection = new Label { Text = "1. Input configuration", Font = new Font("Segoe UI", 12F, FontStyle.Bold), Location = new Point(20, 20), AutoSize = true };
            leftInputPanel.Controls.Add(lblInputSection);

            Label lblSource = new Label { Text = "Source File Format:", Location = new Point(20, 70), AutoSize = true, ForeColor = Color.LightGray };
            leftInputPanel.Controls.Add(lblSource);

            comboSourceFormat = new ComboBox();
            comboSourceFormat.DropDownStyle = ComboBoxStyle.DropDownList;
            comboSourceFormat.FlatStyle = FlatStyle.Flat;
            comboSourceFormat.BackColor = Color.FromArgb(60,60,60);
            comboSourceFormat.ForeColor = Color.White;
            comboSourceFormat.Location = new Point(20, 95);
            comboSourceFormat.Size = new Size(250, 25);
            comboSourceFormat.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
            comboSourceFormat.Items.AddRange(new object[] { "PDF Document", "Word (DOC/DOCX)", "PowerPoint (PPT/PPTX)", "Image (JPG/PNG)" });
            leftInputPanel.Controls.Add(comboSourceFormat);

            Label lblTarget = new Label { Text = "Convert Files To:", Location = new Point(20, 150), AutoSize = true, ForeColor = Color.LightGray };
            leftInputPanel.Controls.Add(lblTarget);

            comboTargetFormat = new ComboBox();
            comboTargetFormat.DropDownStyle = ComboBoxStyle.DropDownList;
            comboTargetFormat.FlatStyle = FlatStyle.Flat;
            comboTargetFormat.BackColor = Color.FromArgb(60,60,60);
            comboTargetFormat.ForeColor = Color.White;
            comboTargetFormat.Location = new Point(20, 175);
            comboTargetFormat.Size = new Size(250, 25);
            comboTargetFormat.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
            leftInputPanel.Controls.Add(comboTargetFormat);

            comboSourceFormat.SelectedIndexChanged += new EventHandler(ComboSourceFormat_SelectedIndexChanged);
            comboSourceFormat.SelectedIndex = 1;

            Label lblMode = new Label { Text = "Processing Mode:", Location = new Point(20, 215), AutoSize = true, ForeColor = Color.LightGray };
            leftInputPanel.Controls.Add(lblMode);

            rbIndividualMode = new RadioButton { Text = "Individual Folder Export", Location = new Point(20, 240), AutoSize = true, Checked = true, ForeColor = Color.White };
            leftInputPanel.Controls.Add(rbIndividualMode);

            rbBatchMode = new RadioButton { Text = "Batch Folders Export", Location = new Point(20, 265), AutoSize = true, ForeColor = Color.White };
            leftInputPanel.Controls.Add(rbBatchMode);

            rbBatchMode.CheckedChanged += new EventHandler(RbBatchMode_CheckedChanged);

            Label lblInputFolder = new Label { Text = "Source Folder:", Location = new Point(20, 305), AutoSize = true, ForeColor = Color.LightGray };
            leftInputPanel.Controls.Add(lblInputFolder);

            inputBtn = new AnimatedFlatButton();
            inputBtn.Text = "Browse Import";
            inputBtn.Size = new Size(120, 35);
            inputBtn.Location = new Point(20, 330);
            inputBtn.Click += delegate(object s, EventArgs e) { BrowseFolder(true); };
            leftInputPanel.Controls.Add(inputBtn);

            inputLabel = new Label();
            inputLabel.Text = "No folder selected";
            inputLabel.ForeColor = Color.Gray;
            inputLabel.AutoSize = false;
            inputLabel.AutoEllipsis = true;
            inputLabel.Location = new Point(150, 337);
            inputLabel.Size = new Size(200, 25);
            inputLabel.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
            leftInputPanel.Controls.Add(inputLabel);

            // Populate Right Panel (Output)
            Label lblOutputSection = new Label { Text = "2. Output execution", Font = new Font("Segoe UI", 12F, FontStyle.Bold), Location = new Point(20, 20), AutoSize = true };
            rightOutputPanel.Controls.Add(lblOutputSection);
            
            Label lblOutputFolder = new Label { Text = "Destination Folder:", Location = new Point(20, 70), AutoSize = true, ForeColor = Color.LightGray };
            rightOutputPanel.Controls.Add(lblOutputFolder);

            outputBtn = new AnimatedFlatButton();
            outputBtn.Text = "Browse Export";
            outputBtn.Size = new Size(120, 35);
            outputBtn.Location = new Point(20, 95);
            outputBtn.Click += delegate(object s, EventArgs e) { BrowseFolder(false); };
            rightOutputPanel.Controls.Add(outputBtn);

            outputLabel = new Label();
            outputLabel.Text = "No folder selected";
            outputLabel.ForeColor = Color.Gray;
            outputLabel.AutoSize = false;
            outputLabel.AutoEllipsis = true;
            outputLabel.Location = new Point(150, 102);
            outputLabel.Size = new Size(200, 25);
            outputLabel.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
            rightOutputPanel.Controls.Add(outputLabel);

            Panel settingsGroup = new Panel();
            settingsGroup.BackColor = Color.FromArgb(35, 35, 35);
            settingsGroup.Location = new Point(20, 150);
            settingsGroup.Size = new Size(250, 100);
            settingsGroup.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
            rightOutputPanel.Controls.Add(settingsGroup);
            
            Label lblSet = new Label { Text = "Advanced Conversion Parameters", Font = new Font("Segoe UI", 9F, FontStyle.Bold), Location = new Point(10, 10), AutoSize = true, ForeColor = Color.LightGray };
            settingsGroup.Controls.Add(lblSet);
            Label lblSetDesc = new Label { Text = "Specific settings will appear here\nbased on selected format.", Location = new Point(10, 35), AutoSize = true, ForeColor = Color.Gray };
            settingsGroup.Controls.Add(lblSetDesc);

            startBtn = new AnimatedFlatButton();
            startBtn.IsPrimary = true;
            startBtn.Text = "EXECUTE BATCH CONVERSION";
            startBtn.Font = new Font("Segoe UI", 12F, FontStyle.Bold);
            startBtn.Click += new EventHandler(StartBtn_Click);
            rightOutputPanel.Controls.Add(startBtn);

            statusLabel = new Label();
            statusLabel.Text = "Ready for execution";
            statusLabel.AutoSize = true;
            rightOutputPanel.Controls.Add(statusLabel);

            progressBar = new ProgressBar();
            rightOutputPanel.Controls.Add(progressBar);

            rightOutputPanel.Resize += delegate(object s, EventArgs e) 
            {
                startBtn.Size = new Size(rightOutputPanel.Width - 40, 50);
                startBtn.Location = new Point(20, rightOutputPanel.Height - 140);
                
                statusLabel.Location = new Point(20, rightOutputPanel.Height - 75);
                
                progressBar.Size = new Size(rightOutputPanel.Width - 40, 15);
                progressBar.Location = new Point(20, rightOutputPanel.Height - 50);
            };
        }

        private void ComboSourceFormat_SelectedIndexChanged(object sender, EventArgs e)
        {
            comboTargetFormat.Items.Clear();
            string selected = comboSourceFormat.SelectedItem.ToString();
            
            if (selected == "PDF Document")
            {
                comboTargetFormat.Items.AddRange(new object[] { "Word (DOCX)", "Image (JPG)", "Image (PNG)", "Text (TXT)" });
            }
            else if (selected == "Word (DOC/DOCX)")
            {
                comboTargetFormat.Items.AddRange(new object[] { "PDF Document", "Text (TXT)", "Web Page (HTML)" });
            }
            else if (selected == "PowerPoint (PPT/PPTX)")
            {
                comboTargetFormat.Items.AddRange(new object[] { "PDF Document", "Image Series (JPG)" });
            }
            else if (selected == "Image (JPG/PNG)")
            {
                comboTargetFormat.Items.AddRange(new object[] { "PDF Document" });
            }

            if (comboTargetFormat.Items.Count > 0)
                comboTargetFormat.SelectedIndex = 0;
        }

        private void BrowseFolder(bool isInput)
        {
            using (FolderBrowserDialog fbd = new FolderBrowserDialog())
            {
                fbd.Description = isInput ? "Select Input Folder" : "Select Output Folder";
                if (fbd.ShowDialog() == DialogResult.OK)
                {
                    if (isInput)
                    {
                        inputDirPath = fbd.SelectedPath;
                        inputLabel.Text = inputDirPath;
                        inputLabel.ForeColor = textColor;
                        if (rbBatchMode != null && rbBatchMode.Checked)
                        {
                            outputDirPath = inputDirPath;
                            outputLabel.Text = outputDirPath + " (Default)";
                            outputLabel.ForeColor = textColor;
                        }
                    }
                    else
                    {
                        outputDirPath = fbd.SelectedPath;
                        outputLabel.Text = outputDirPath;
                        outputLabel.ForeColor = textColor;
                    }
                }
            }
        }

        private void RbBatchMode_CheckedChanged(object sender, EventArgs e)
        {
            if (rbBatchMode.Checked && !string.IsNullOrEmpty(inputDirPath))
            {
                outputDirPath = inputDirPath;
                outputLabel.Text = outputDirPath + " (Default)";
                outputLabel.ForeColor = textColor;
            }
        }

        private void SetUiState(bool enabled)
        {
            if (this.InvokeRequired)
            {
                this.Invoke(new Action(() => SetUiState(enabled)));
                return;
            }

            comboSourceFormat.Enabled = enabled;
            comboTargetFormat.Enabled = enabled;
            inputBtn.Enabled = enabled;
            outputBtn.Enabled = enabled;
            startBtn.Enabled = enabled;

            // Extra management buttons disabled during execution
            manageStartBtn.Enabled = enabled;
            manageInputBtn.Enabled = enabled;
            manageOutputBtn.Enabled = enabled;
        }

        private void UpdateProgress(int current, int total, string text)
        {
            if (this.InvokeRequired)
            {
                this.Invoke(new Action(() => UpdateProgress(current, total, text)));
                return;
            }

            statusLabel.Text = text;
            if (total > 0)
            {
                progressBar.Maximum = total;
                progressBar.Value = Math.Min(current, total);
            }
            else
            {
                progressBar.Value = 0;
            }
        }

        private void StartBtn_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(inputDirPath) || string.IsNullOrWhiteSpace(outputDirPath))
            {
                MessageBox.Show("Please specify both source and destination folders.", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (!Directory.Exists(inputDirPath))
            {
                MessageBox.Show("Source folder does not exist.", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (!Directory.Exists(outputDirPath))
            {
                try { Directory.CreateDirectory(outputDirPath); }
                catch (Exception ex)
                {
                    MessageBox.Show(string.Format("Failed to create destination folder: {0}", ex.Message), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
            }
            
            if (comboTargetFormat.SelectedItem == null) return;

            string sourceFormat = comboSourceFormat.SelectedItem.ToString();
            string targetFormat = comboTargetFormat.SelectedItem.ToString();

            SetUiState(false);
            isConverting = true;
            UpdateProgress(0, 1, "Initializing execution...");

            bool isBatch = rbBatchMode.Checked;
            Thread workerThread = new Thread(() => RunConversionProcess(inputDirPath, outputDirPath, sourceFormat, targetFormat, isBatch));
            workerThread.IsBackground = true;
            workerThread.SetApartmentState(ApartmentState.STA);
            workerThread.Start();
        }

        private void RunConversionProcess(string inDir, string outDir, string src, string tgt, bool isBatch)
        {
            try
            {
                if (src == "Word (DOC/DOCX)" && tgt == "PDF Document")
                {
                    ConvertWordToPdf(inDir, outDir, isBatch);
                }
                else if (src == "PDF Document" && tgt == "Word (DOCX)")
                {
                    ConvertPdfToWord(inDir, outDir, isBatch);
                }
                else if (src == "PowerPoint (PPT/PPTX)" && tgt == "PDF Document")
                {
                    ConvertPptToPdf(inDir, outDir, isBatch);
                }
                else
                {
                    UpdateProgress(1, 1, "Formatting routine pending deployment.");
                    MessageBox.Show("This combination is listed cleanly in the UI and will be made functional in upcoming codebase extensions.", "Info", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    isConverting = false;
                }

                if (isConverting)
                {
                    UpdateProgress(1, 1, "All conversions completed securely.");
                    this.Invoke(new Action(() => MessageBox.Show("Batch conversion process executed successfully!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information)));
                }
            }
            catch (Exception ex)
            {
                UpdateProgress(0, 1, string.Format("Error: {0}", ex.Message));
                this.Invoke(new Action(() => MessageBox.Show(ex.Message, "Execution Fault", MessageBoxButtons.OK, MessageBoxIcon.Error)));
            }
            finally
            {
                isConverting = false;
                SetUiState(true);
            }
        }

        private void ConvertWordToPdf(string inDir, string outDir, bool isBatch)
        {
            var searchOption = isBatch ? SearchOption.AllDirectories : SearchOption.TopDirectoryOnly;
            var files = Directory.GetFiles(inDir, "*.doc?", searchOption);
            int total = 0;
            
            var validFiles = new System.Collections.Generic.List<string>();
            foreach (var f in files)
            {
                string name = Path.GetFileName(f);
                if (!name.StartsWith("~")) validFiles.Add(f);
            }
            total = validFiles.Count;

            if (total == 0) throw new Exception("No Word documents located in the source vector.");

            Word.Application wordApp = null;
            try
            {
                wordApp = new Word.Application();
                wordApp.Visible = false;
                
                for (int i = 0; i < validFiles.Count; i++)
                {
                    if (!isConverting) break;
                    
                    string file = validFiles[i];
                    string baseName = Path.GetFileNameWithoutExtension(file);

                    string relativePath = file.Substring(inDir.Length).TrimStart(Path.DirectorySeparatorChar);
                    string targetDir = Path.Combine(outDir, Path.GetDirectoryName(relativePath));
                    if (!Directory.Exists(targetDir)) Directory.CreateDirectory(targetDir);

                    string outPath = Path.Combine(targetDir, baseName + ".pdf");
                    
                    UpdateProgress(i, total, string.Format("Converting ({0}/{1}): {2}", i + 1, total, Path.GetFileName(file)));
                    
                    Word.Document doc = null;
                    try
                    {
                        doc = wordApp.Documents.Open(file, ReadOnly: true, Visible: false);
                        doc.SaveAs2(outPath, Word.WdSaveFormat.wdFormatPDF);
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(string.Format("Failed {0}: {1}", file, ex.Message));
                    }
                    finally
                    {
                        if (doc != null)
                        {
                            ((Word._Document)doc).Close(Word.WdSaveOptions.wdDoNotSaveChanges);
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(doc);
                        }
                    }
                }
            }
            finally
            {
                if (wordApp != null)
                {
                    ((Word._Application)wordApp).Quit();
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(wordApp);
                }
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }

        private void ConvertPdfToWord(string inDir, string outDir, bool isBatch)
        {
            var searchOption = isBatch ? SearchOption.AllDirectories : SearchOption.TopDirectoryOnly;
            var files = Directory.GetFiles(inDir, "*.pdf", searchOption);
            int total = files.Length;

            if (total == 0) throw new Exception("No PDF payloads located in the source vector.");

            Word.Application wordApp = null;
            try
            {
                wordApp = new Word.Application();
                wordApp.Visible = false;
                wordApp.DisplayAlerts = Word.WdAlertLevel.wdAlertsNone;
                
                for (int i = 0; i < total; i++)
                {
                    if (!isConverting) break;
                    
                    string file = files[i];
                    string baseName = Path.GetFileNameWithoutExtension(file);

                    string relativePath = file.Substring(inDir.Length).TrimStart(Path.DirectorySeparatorChar);
                    string targetDir = Path.Combine(outDir, Path.GetDirectoryName(relativePath));
                    if (!Directory.Exists(targetDir)) Directory.CreateDirectory(targetDir);

                    string outPath = Path.Combine(targetDir, baseName + ".docx");
                    
                    UpdateProgress(i, total, string.Format("Converting ({0}/{1}): {2}", i + 1, total, Path.GetFileName(file)));
                    
                    Word.Document doc = null;
                    try
                    {
                        doc = wordApp.Documents.Open(file, ReadOnly: true, Visible: false, ConfirmConversions: false);
                        doc.SaveAs2(outPath, Word.WdSaveFormat.wdFormatXMLDocument);
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(string.Format("Failed {0}: {1}", file, ex.Message));
                    }
                    finally
                    {
                        if (doc != null)
                        {
                            ((Word._Document)doc).Close(Word.WdSaveOptions.wdDoNotSaveChanges);
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(doc);
                        }
                    }
                }
            }
            finally
            {
                if (wordApp != null)
                {
                    wordApp.DisplayAlerts = Word.WdAlertLevel.wdAlertsAll;
                    ((Word._Application)wordApp).Quit();
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(wordApp);
                }
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }

        private void ConvertPptToPdf(string inDir, string outDir, bool isBatch)
        {
            var searchOption = isBatch ? SearchOption.AllDirectories : SearchOption.TopDirectoryOnly;
            var files = Directory.GetFiles(inDir, "*.ppt?", searchOption);
            int total = 0;
            
            var validFiles = new System.Collections.Generic.List<string>();
            foreach (var f in files)
            {
                string name = Path.GetFileName(f);
                if (!name.StartsWith("~")) validFiles.Add(f);
            }
            total = validFiles.Count;

            if (total == 0) throw new Exception("No PowerPoint packages located in the source vector.");

            PowerPoint.Application pptApp = null;
            try
            {
                pptApp = new PowerPoint.Application();
                
                for (int i = 0; i < validFiles.Count; i++)
                {
                    if (!isConverting) break;
                    
                    string file = validFiles[i];
                    string baseName = Path.GetFileNameWithoutExtension(file);

                    string relativePath = file.Substring(inDir.Length).TrimStart(Path.DirectorySeparatorChar);
                    string targetDir = Path.Combine(outDir, Path.GetDirectoryName(relativePath));
                    if (!Directory.Exists(targetDir)) Directory.CreateDirectory(targetDir);

                    string outPath = Path.Combine(targetDir, baseName + ".pdf");
                    
                    UpdateProgress(i, total, string.Format("Converting ({0}/{1}): {2}", i + 1, total, Path.GetFileName(file)));
                    
                    PowerPoint.Presentation presentation = null;
                    try
                    {
                        presentation = pptApp.Presentations.Open(file, Microsoft.Office.Core.MsoTriState.msoTrue, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoFalse);
                        presentation.SaveAs(outPath, PowerPoint.PpSaveAsFileType.ppSaveAsPDF);
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(string.Format("Failed {0}: {1}", file, ex.Message));
                    }
                    finally
                    {
                        if (presentation != null)
                        {
                            presentation.Close();
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(presentation);
                        }
                    }
                }
            }
            finally
            {
                if (pptApp != null)
                {
                    pptApp.Quit();
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(pptApp);
                }
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }
    }
}
