namespace Tilda {
    partial class TildaRibbon : Microsoft.Office.Tools.Ribbon.RibbonBase {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public TildaRibbon()
            : base(Globals.Factory.GetRibbonFactory()) {
            InitializeComponent();
        }

        /// <summary> 
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing) {
            if (disposing && (components != null)) {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Component Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent() {
            this.tab1 = this.Factory.CreateRibbonTab();
            this.exportMenu = this.Factory.CreateRibbonGroup();
            this.exportTildaShape = this.Factory.CreateRibbonButton();
            this.exportTildaSlide = this.Factory.CreateRibbonButton();
            this.exportTildaPresentation = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.exportMenu.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Groups.Add(this.exportMenu);
            this.tab1.Label = "TabAddIns";
            this.tab1.Name = "tab1";
            // 
            // exportMenu
            // 
            this.exportMenu.Items.Add(this.exportTildaShape);
            this.exportMenu.Items.Add(this.exportTildaSlide);
            this.exportMenu.Items.Add(this.exportTildaPresentation);
            this.exportMenu.Label = "Tilda Export";
            this.exportMenu.Name = "exportMenu";
            // 
            // exportTildaShape
            // 
            this.exportTildaShape.Image = global::Tilda.Properties.Resources._74_location;
            this.exportTildaShape.Label = "Selected";
            this.exportTildaShape.Name = "exportTildaShape";
            this.exportTildaShape.ShowImage = true;
            // 
            // exportTildaSlide
            // 
            this.exportTildaSlide.Image = global::Tilda.Properties.Resources._41_picture_frame;
            this.exportTildaSlide.Label = "Active Slide";
            this.exportTildaSlide.Name = "exportTildaSlide";
            this.exportTildaSlide.ShowImage = true;
            this.exportTildaSlide.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.exportTildaSlide_Click);
            // 
            // exportTildaPresentation
            // 
            this.exportTildaPresentation.Image = global::Tilda.Properties.Resources._137_presentation;
            this.exportTildaPresentation.Label = "Presentation";
            this.exportTildaPresentation.Name = "exportTildaPresentation";
            this.exportTildaPresentation.ShowImage = true;
            // 
            // TildaRibbon
            // 
            this.Name = "TildaRibbon";
            this.RibbonType = "Microsoft.PowerPoint.Presentation";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.TildaRibbon_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.exportMenu.ResumeLayout(false);
            this.exportMenu.PerformLayout();

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup exportMenu;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton exportTildaSlide;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton exportTildaPresentation;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton exportTildaShape;
    }

    partial class ThisRibbonCollection {
        internal TildaRibbon TildaRibbon {
            get { return this.GetRibbon<TildaRibbon>(); }
        }
    }
}
