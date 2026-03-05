using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Windows.Forms;

namespace DocumentConverter
{
    public static class GraphicsExtension
    {
        public static void FillRoundedRectangle(this Graphics g, Brush brush, Rectangle bounds, int cornerRadius)
        {
            if (cornerRadius < 1) { g.FillRectangle(brush, bounds); return; }
            using (GraphicsPath path = new GraphicsPath())
            {
                path.AddArc(bounds.X, bounds.Y, cornerRadius, cornerRadius, 180, 90);
                path.AddArc(bounds.X + bounds.Width - cornerRadius, bounds.Y, cornerRadius, cornerRadius, 270, 90);
                path.AddArc(bounds.X + bounds.Width - cornerRadius, bounds.Y + bounds.Height - cornerRadius, cornerRadius, cornerRadius, 0, 90);
                path.AddArc(bounds.X, bounds.Y + bounds.Height - cornerRadius, cornerRadius, cornerRadius, 90, 90);
                path.CloseAllFigures();
                g.FillPath(brush, path);
            }
        }
    }

    public class AnimatedFlatButton : Button
    {
        private Color _normalColor;
        private Color _hoverColor;
        private Color _accentColor;
        private Color _currentColor;
        private System.Windows.Forms.Timer _animTimer;
        private bool _isHovering;
        private int _step;
        private bool _isPrimary;
        
        public bool IsPrimary 
        { 
            get { return _isPrimary; } 
            set { _isPrimary = value; } 
        }

        public AnimatedFlatButton()
        {
            DoubleBuffered = true;
            _normalColor = Color.FromArgb(80, 80, 80);
            _hoverColor = Color.FromArgb(100, 100, 100);
            _accentColor = Color.FromArgb(52, 116, 212);
            _currentColor = _normalColor;
            _isHovering = false;
            _step = 10;
            _isPrimary = false;

            _animTimer = new System.Windows.Forms.Timer();
            _animTimer.Interval = 15;
            _animTimer.Tick += new EventHandler(AnimTimer_Tick);
        }

        private void AnimTimer_Tick(object sender, EventArgs e)
        {
            Color target = _isHovering ? (_isPrimary ? Color.FromArgb(70, 130, 230) : _hoverColor) : (_isPrimary ? _accentColor : _normalColor);
            if (!Enabled) target = Color.FromArgb(45, 45, 45);

            int r = StepColor(_currentColor.R, target.R);
            int g = StepColor(_currentColor.G, target.G);
            int b = StepColor(_currentColor.B, target.B);
            
            _currentColor = Color.FromArgb(r, g, b);
            Invalidate();
            
            if (r == target.R && g == target.G && b == target.B)
                _animTimer.Stop();
        }

        private int StepColor(int current, int target)
        {
            if (current < target) return Math.Min(current + _step, target);
            if (current > target) return Math.Max(current - _step, target);
            return current;
        }

        public override void NotifyDefault(bool value) {}
        
        protected override void OnPaint(PaintEventArgs pevent)
        {
            Graphics g = pevent.Graphics;
            g.SmoothingMode = SmoothingMode.AntiAlias;
            
            if (_currentColor.A == 0) _currentColor = _isPrimary ? _accentColor : _normalColor;
            if (!Enabled) _currentColor = Color.FromArgb(45, 45, 45);

            using (SolidBrush brush = new SolidBrush(_currentColor))
            {
                g.FillRoundedRectangle(brush, this.ClientRectangle, 5);
            }

            TextRenderer.DrawText(g, Text, Font, this.ClientRectangle, Enabled ? ForeColor : Color.Gray, TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter);
        }
        
        protected override void OnMouseEnter(EventArgs e) { _isHovering = true; _animTimer.Start(); base.OnMouseEnter(e); }
        protected override void OnMouseLeave(EventArgs e) { _isHovering = false; _animTimer.Start(); base.OnMouseLeave(e); }
    }

    public class TabButton : Button
    {
        private bool _isSelected;
        public bool IsSelected 
        { 
            get { return _isSelected; } 
            set { _isSelected = value; Invalidate(); } 
        }
        
        private bool _showText;
        public bool ShowText 
        { 
            get { return _showText; } 
            set { _showText = value; Invalidate(); } 
        }
        
        private string _iconText;
        public string IconText 
        { 
            get { return _iconText; } 
            set { _iconText = value; Invalidate(); } 
        }

        public TabButton()
        {
            _isSelected = false;
            _showText = true;
            _iconText = "";
            DoubleBuffered = true;
        }

        public override void NotifyDefault(bool value) {}
        
        protected override void OnPaint(PaintEventArgs pevent)
        {
            Graphics g = pevent.Graphics;
            g.SmoothingMode = SmoothingMode.AntiAlias;
            
            Color bgColor = _isSelected ? Color.FromArgb(45, 45, 50) : (this.ClientRectangle.Contains(this.PointToClient(Cursor.Position)) ? Color.FromArgb(35, 35, 35) : BackColor);
            
            using (SolidBrush brush = new SolidBrush(bgColor))
            {
                g.FillRectangle(brush, this.ClientRectangle);
            }
            
            if (_isSelected)
            {
                using (SolidBrush indicatorBrush = new SolidBrush(Color.FromArgb(52, 116, 212)))
                {
                    g.FillRectangle(indicatorBrush, new Rectangle(0, 0, 4, this.Height));
                }
            }

            if (_showText)
            {
                TextRenderer.DrawText(g, _iconText + "  " + Text, Font, new Rectangle(15, 0, this.Width - 15, this.Height), ForeColor, TextFormatFlags.Left | TextFormatFlags.VerticalCenter);
            }
            else
            {
                TextRenderer.DrawText(g, _iconText, new Font(Font.FontFamily, 14F, FontStyle.Bold), this.ClientRectangle, ForeColor, TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter);
            }
        }
        
        protected override void OnMouseEnter(EventArgs e) { Invalidate(); base.OnMouseEnter(e); }
        protected override void OnMouseLeave(EventArgs e) { Invalidate(); base.OnMouseLeave(e); }
    }
}
