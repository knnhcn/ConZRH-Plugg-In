namespace ConZRHAddIn
{
    partial class ControlsZrh : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public ControlsZrh()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

        /// <summary> 
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Component Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.Controls = this.Factory.CreateRibbonTab();
            this.richControl = this.Factory.CreateRibbonGroup();
            this.label1 = this.Factory.CreateRibbonLabel();
            this.separator1 = this.Factory.CreateRibbonSeparator();
            this.addLockButton = this.Factory.CreateRibbonButton();
            this.removeLockButton = this.Factory.CreateRibbonButton();
            this.creatLockButton = this.Factory.CreateRibbonButton();
            this.button3 = this.Factory.CreateRibbonButton();
            this.button2 = this.Factory.CreateRibbonButton();
            this.button1 = this.Factory.CreateRibbonButton();
            this.separator2 = this.Factory.CreateRibbonSeparator();
            this.separator3 = this.Factory.CreateRibbonSeparator();
            this.Controls.SuspendLayout();
            this.richControl.SuspendLayout();
            this.SuspendLayout();
            // 
            // Controls
            // 
            this.Controls.Groups.Add(this.richControl);
            this.Controls.Label = "ControlsZRH";
            this.Controls.Name = "Controls";
            // 
            // richControl
            // 
            this.richControl.Items.Add(this.creatLockButton);
            this.richControl.Items.Add(this.separator2);
            this.richControl.Items.Add(this.removeLockButton);
            this.richControl.Items.Add(this.addLockButton);
            this.richControl.Items.Add(this.separator3);
            this.richControl.Items.Add(this.button3);
            this.richControl.Items.Add(this.label1);
            this.richControl.Items.Add(this.separator1);
            this.richControl.Items.Add(this.button2);
            this.richControl.Items.Add(this.button1);
            this.richControl.Label = "Locked Text";
            this.richControl.Name = "richControl";
            // 
            // label1
            // 
            this.label1.Enabled = false;
            this.label1.Label = "Selected Item";
            this.label1.Name = "label1";
            // 
            // separator1
            // 
            this.separator1.Name = "separator1";
            // 
            // addLockButton
            // 
            this.addLockButton.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.addLockButton.Label = "Lock RichText";
            this.addLockButton.Name = "addLockButton";
            this.addLockButton.OfficeImageId = "Lock";
            this.addLockButton.ShowImage = true;
            this.addLockButton.SuperTip = "This button adds a lock to the RichTextContentControl.";
            this.addLockButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.addLockButton_Click);
            // 
            // removeLockButton
            // 
            this.removeLockButton.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.removeLockButton.Label = "Unlock RichText";
            this.removeLockButton.Name = "removeLockButton";
            this.removeLockButton.OfficeImageId = "ClearRow";
            this.removeLockButton.ShowImage = true;
            this.removeLockButton.SuperTip = "This button removes the lock from the RichTextContentControl.";
            this.removeLockButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.removeLockButton_Click);
            // 
            // creatLockButton
            // 
            this.creatLockButton.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.creatLockButton.Label = "Lock selected Text";
            this.creatLockButton.Name = "creatLockButton";
            this.creatLockButton.OfficeImageId = "GroupProtect";
            this.creatLockButton.ShowImage = true;
            this.creatLockButton.SuperTip = "Select a text and turn it into a RichTextContentControl.";
            this.creatLockButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.creatLockButton_Click);
            // 
            // button3
            // 
            this.button3.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button3.Label = "Remove RichText element";
            this.button3.Name = "button3";
            this.button3.OfficeImageId = "DeclineInvitation";
            this.button3.ShowImage = true;
            this.button3.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button3_Click);
            // 
            // button2
            // 
            this.button2.Label = "Reload AddIn";
            this.button2.Name = "button2";
            this.button2.OfficeImageId = "Repeat";
            this.button2.ShowImage = true;
            this.button2.SuperTip = "Reload the AddIn if you added sub documents.";
            this.button2.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button2_Click);
            // 
            // button1
            // 
            this.button1.Label = "About";
            this.button1.Name = "button1";
            this.button1.OfficeImageId = "Info";
            this.button1.ShowImage = true;
            this.button1.SuperTip = "About";
            this.button1.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button1_Click);
            // 
            // separator2
            // 
            this.separator2.Name = "separator2";
            // 
            // separator3
            // 
            this.separator3.Name = "separator3";
            // 
            // ControlsZrh
            // 
            this.Name = "ControlsZrh";
            this.RibbonType = "Microsoft.Word.Document";
            this.Tabs.Add(this.Controls);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.ControlsZrh_Load);
            this.Controls.ResumeLayout(false);
            this.Controls.PerformLayout();
            this.richControl.ResumeLayout(false);
            this.richControl.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup richControl;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton addLockButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton removeLockButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton creatLockButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonLabel label1;
        internal Microsoft.Office.Tools.Ribbon.RibbonTab Controls;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button2;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button1;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button3;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator2;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator3;
    }

    partial class ThisRibbonCollection
    {
        internal ControlsZrh ControlsZrh
        {
            get { return this.GetRibbon<ControlsZrh>(); }
        }
    }
}
