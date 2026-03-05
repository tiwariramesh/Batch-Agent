using System;
using System.Drawing;
using System.Windows.Forms;

[assembly: System.Reflection.AssemblyTitle("Batch File Agent")]
[assembly: System.Reflection.AssemblyProduct("Batch File Agent")]
[assembly: System.Reflection.AssemblyVersion("3.0.0.0")]
[assembly: System.Runtime.InteropServices.ComVisible(false)]

namespace DocumentConverter
{
    public partial class MainForm : Form
    {
        private Color darkBg = Color.FromArgb(28, 28, 28);
        private Color sidebarBg = Color.FromArgb(20, 20, 20);
        private Color panelBg = Color.FromArgb(43, 43, 43);
        private Color accentColor = Color.FromArgb(52, 116, 212);
        private Color textColor = Color.White;

        private Panel sidebar;
        private Panel mainContentPanel;
        private Panel fileConversionView;
        private Panel fileManagementView;
        private AnimatedFlatButton menuToggle;
        private PictureBox logoBox;
        private Label titleLabel;
        
        private System.Windows.Forms.Timer _sidebarTimer;
        private bool _isSidebarExpanded;

        public MainForm()
        {
            _isSidebarExpanded = true;
            InitializeComponent();
            this.FormClosing += new FormClosingEventHandler(MainForm_FormClosing);
        }

        private void InitializeComponent()
        {
            this.Text = "Batch File Agent - Professional File Manager & Converter";
            this.Size = new Size(1100, 650);
            this.MinimumSize = new Size(800, 500);
            this.BackColor = darkBg;
            this.ForeColor = textColor;
            this.Font = new Font("Segoe UI", 10F, FontStyle.Regular, GraphicsUnit.Point);
            this.StartPosition = FormStartPosition.CenterScreen;

            try
            {
                string iconPath = System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Assets", "app_icon.ico");
                if (System.IO.File.Exists(iconPath))
                {
                    this.Icon = new Icon(iconPath);
                }
            }
            catch { /* Fallback to default if icon fails to load */ }

            // --- Sidebar Components ---
            sidebar = new Panel();
            sidebar.BackColor = sidebarBg;
            sidebar.Dock = DockStyle.Left;
            sidebar.Width = 220;
            this.Controls.Add(sidebar);

            _sidebarTimer = new System.Windows.Forms.Timer();
            _sidebarTimer.Interval = 20;
            _sidebarTimer.Tick += new EventHandler(SidebarTimer_Tick);

            menuToggle = new AnimatedFlatButton();
            menuToggle.Text = "≡";
            menuToggle.Font = new Font("Segoe UI", 16F, FontStyle.Bold);
            menuToggle.Size = new Size(40, 40);
            menuToggle.Location = new Point(10, 10);
            menuToggle.BackColor = sidebarBg;
            menuToggle.ForeColor = textColor;
            menuToggle.Click += new EventHandler(MenuToggle_Click);
            sidebar.Controls.Add(menuToggle);

            logoBox = new PictureBox();
            logoBox.Size = new Size(32, 32);
            logoBox.Location = new Point(14, 65);
            logoBox.SizeMode = PictureBoxSizeMode.Zoom;
            logoBox.BackColor = Color.Transparent;
            try
            {
                string iconPath = System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Assets", "app_icon.ico");
                if (System.IO.File.Exists(iconPath))
                {
                    logoBox.Image = new Icon(iconPath, 32, 32).ToBitmap();
                }
            }
            catch { }
            sidebar.Controls.Add(logoBox);

            titleLabel = new Label();
            titleLabel.Text = "Batch File Agent";
            titleLabel.Font = new Font("Segoe UI", 12F, FontStyle.Bold);
            titleLabel.ForeColor = textColor;
            titleLabel.Location = new Point(55, 70);
            titleLabel.Width = 150;
            titleLabel.AutoSize = false;
            sidebar.Controls.Add(titleLabel);

            // Shift tabs down to accommodate the logo
            int tabStartY = 120;
            TabButton tabConvert = new TabButton();
            tabConvert.Text = "File Conversion";
            tabConvert.IconText = "🔄";
            tabConvert.Size = new Size(220, 50);
            tabConvert.Location = new Point(0, tabStartY);
            tabConvert.BackColor = sidebarBg;
            tabConvert.ForeColor = textColor;
            tabConvert.IsSelected = true;
            sidebar.Controls.Add(tabConvert);

            TabButton tabManage = new TabButton();
            tabManage.Text = "File Management";
            tabManage.IconText = "🗂";
            tabManage.Size = new Size(220, 50);
            tabManage.Location = new Point(0, tabStartY + 50);
            tabManage.BackColor = sidebarBg;
            tabManage.ForeColor = Color.Gray; 
            sidebar.Controls.Add(tabManage);

            tabConvert.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
            tabManage.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;

            TabButton tabAbout = new TabButton();
            tabAbout.Text = "About App";
            tabAbout.IconText = "ℹ️";
            tabAbout.Size = new Size(220, 50);
            tabAbout.Location = new Point(0, this.ClientSize.Height - 50);
            tabAbout.BackColor = sidebarBg;
            tabAbout.ForeColor = Color.Gray; 
            tabAbout.Anchor = AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
            tabAbout.Click += delegate(object sender, EventArgs e) 
            { 
                MessageBox.Show("Batch File Agent\nVersion: 3.0.0.0\nAuthor: Ramesh Tiwari\n\nPurpose: This application was created to make file conversion and management accessible for general-purpose use to all. While thousands of applications online and in stores charge payments for batch processing, these native Windows operations should not be payable. They should instead be convenient and genuinely free to all users.", "About Batch File Agent", MessageBoxButtons.OK, MessageBoxIcon.Information);
            };
            sidebar.Controls.Add(tabAbout);

            // Sidebar Tabs Action
            EventHandler tabConvertClick = delegate(object sender, EventArgs e) 
            { 
                tabConvert.IsSelected = true; 
                tabManage.IsSelected = false; 
                tabConvert.Invalidate(); 
                tabManage.Invalidate(); 
                ShowView(fileConversionView); 
            };
            EventHandler tabManageClick = delegate(object sender, EventArgs e) 
            { 
                tabManage.IsSelected = true; 
                tabConvert.IsSelected = false; 
                tabConvert.Invalidate(); 
                tabManage.Invalidate(); 
                ShowView(fileManagementView); 
            };
            tabConvert.Click += tabConvertClick;
            tabManage.Click += tabManageClick;

            // --- Main Content Area ---
            mainContentPanel = new Panel();
            mainContentPanel.BackColor = darkBg;
            mainContentPanel.Dock = DockStyle.Fill;
            this.Controls.Add(mainContentPanel);

            // --- File Management View ---
            fileManagementView = new Panel();
            fileManagementView.Dock = DockStyle.Fill;
            fileManagementView.Visible = false;
            mainContentPanel.Controls.Add(fileManagementView);
            SetupManagementViewLayout();

            // --- File Conversion View ---
            fileConversionView = new Panel();
            fileConversionView.Dock = DockStyle.Fill;
            mainContentPanel.Controls.Add(fileConversionView);
            SetupConversionViewLayout();

            fileConversionView.BringToFront();
            mainContentPanel.BringToFront();
        }

        private void MenuToggle_Click(object sender, EventArgs e)
        {
            _isSidebarExpanded = !_isSidebarExpanded;
            _sidebarTimer.Start();
        }

        private void SidebarTimer_Tick(object sender, EventArgs e)
        {
            if (_isSidebarExpanded)
            {
                sidebar.Width += 20;
                if (sidebar.Width >= 220)
                {
                    sidebar.Width = 220;
                    _sidebarTimer.Stop();
                    ToggleSidebarText(true);
                }
            }
            else
            {
                sidebar.Width -= 20;
                if (sidebar.Width <= 60)
                {
                    sidebar.Width = 60;
                    _sidebarTimer.Stop();
                }
                ToggleSidebarText(false);
            }
        }

        private void ToggleSidebarText(bool show)
        {
            foreach(Control c in sidebar.Controls)
            {
                if(c is TabButton) ((TabButton)c).ShowText = show;
            }
            titleLabel.Visible = show;
        }

        private void ShowView(Panel view)
        {
            fileConversionView.Visible = false;
            fileManagementView.Visible = false;
            view.Visible = true;
        }

        private void MainForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            isConverting = false;
            isExtracting = false;
        }
    }
}
