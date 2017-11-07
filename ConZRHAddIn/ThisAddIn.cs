using System;
using Word = Microsoft.Office.Interop.Word;
using Microsoft.Office.Tools.Word;
using System.Windows.Forms;
using System.Runtime.InteropServices;


namespace ConZRHAddIn
{

    public partial class ThisAddIn
    {
        private RichTextContentControl richTextControl = null;
        private Microsoft.Office.Interop.Word.ContentControl rt = null;
        private int index = 0;
        private Document vstoDocument;


        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {

            vstoDocument = Globals.Factory.GetVstoObject(this.Application.ActiveDocument);

            if (vstoDocument.Subdocuments.Count > 0)
            {
                MessageBox.Show("Dieses Dokument enthält "
                    + vstoDocument.Subdocuments.Count
                    + " Filialdokumente.\n\n"
                    + "Sie können das AddIn unter 'Developer->COM Add-Ins' aktivieren/deaktivieren.",
                    "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }


            vstoDocument.SelectionChange += VstoDocument_SelectionChange;
            vstoDocument.BeforeSave += VstoDocument_BeforeSave;
            vstoDocument.ContentControlOnEnter += VstoDocument_ContentControlOnEnter;
            vstoDocument.ContentControlOnExit += VstoDocument_ContentControlOnExit;

            this.Application.ActiveDocument.Subdocuments.Expanded = true;
            this.Application.ActiveDocument.ActiveWindow.View.Type = Word.WdViewType.wdPrintView;

            reloadIndexing();

        }

        internal void delRichText()
        {
            if (this.rt != null)
            {

                DialogResult result = MessageBox.Show("Do you really want to delete?", 
                    "Confirm deletion", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
                if (result == DialogResult.Yes)
                {
                    this.rt.LockContents = false;
                    this.rt.Delete(false);
                    this.rt = null;
                }
                else
                {
                    return;
                }


            }
            else
            {
                MessageBox.Show("Cursor not placed in Rich Text element!",
                    "Cursor not placed in Rich Text!",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
            }

        }

        internal void about()
        {
            MessageBox.Show("Version 3.0\n\n" + 
                "10. August 2017\n" + 
                "KH / SB\n" + 
                "MAN Controls ZH\n" + 
                "Tel. intern: 3267", "About",

                MessageBoxButtons.OK, MessageBoxIcon.Information);
        }


        private void VstoDocument_ContentControlOnEnter(Word.ContentControl cc)
        {
            cc.Type = Word.WdContentControlType.wdContentControlRichText;

            this.rt = cc;

            if (!rt.LockContents)
            {
                Globals.Ribbons.ControlsZrh.label1.Label = "NO LOCK";
            }

            if (rt.LockContents)
            {
                Globals.Ribbons.ControlsZrh.label1.Label = "LOCKED";
            }
        }

        private void VstoDocument_ContentControlOnExit(Word.ContentControl ContentControl, ref bool Cancel)
        {
            this.rt = null;
            Globals.Ribbons.ControlsZrh.label1.Label = "NO LOCK";
        }


        private void VstoDocument_SelectionChange(object sender, SelectionEventArgs e)
        {
            Word.Selection selection = this.Application.Selection;

            foreach (Word.ContentControl c in selection.ContentControls)
            {
                if (c != null)
                {
                    selection.SetRange(0, 0);
                    MessageBox.Show("Content Controls not selectable!", "Warning",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Error);
                }
            }


        }


        internal void reload()
        {

            ThisAddIn_Startup(null, null);

            MessageBox.Show("AddIn reloaded successfully!",
                    "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);

        }


        private void VstoDocument_BeforeSave(object sender, SaveEventArgs e)
        {
            foreach (Word.ContentControl cc in vstoDocument.ContentControls)
            {
                if (cc.Type == Word.WdContentControlType.wdContentControlRichText)
                {
                    cc.LockContentControl = false;
                    cc.LockContents = true;
                    cc.Title = "LOCKED";

                }
            }

        }



        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }


        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>

        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);

        }


        #endregion

        private void reloadIndexing()
        {
            vstoDocument = Globals.Factory.GetVstoObject(this.Application.ActiveDocument);

            if (vstoDocument.ContentControls.Count != 0)
            {
                foreach (Word.ContentControl cc in vstoDocument.ContentControls)
                {
                    if (cc.Type == Word.WdContentControlType.wdContentControlRichText)
                    {
                        try
                        {
                            if (int.Parse(cc.Tag.ToString()) > index)
                            {
                                index = int.Parse(cc.Tag.ToString());
                            }

                        }
                        catch (FormatException)
                        {
                            continue;
                        }
                    }
                }
            }
        }



        internal void SetRichTextControlOnDocument()
        {

            vstoDocument = Globals.Factory.GetVstoObject(this.Application.ActiveDocument);
            index++;

            string name = "CC_" + System.Convert.ToString(index);
            Word.Selection selection = this.Application.Selection;

            if (selection != null)
            {
                try
                {
                    if (selection.ParentContentControl == null)
                    {
                        if (selection != null && selection.Range != null)
                        {
                            try
                            {
                                this.richTextControl = vstoDocument.Controls.AddRichTextContentControl(
                                    selection.Range, name);
                                this.richTextControl.Tag = index.ToString();

                                //this.richTextControl.LockContentControl = true;
                                this.richTextControl.LockContents = true;

                                this.richTextControl.Title = "LOCKED";

                                Globals.Ribbons.ControlsZrh.label1.Label = "LOCKED";

                            }
                            catch (COMException)
                            {
                                MessageBox.Show("Nested RichTextContentControls not allowed", 
                                    "COMException", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            }
                        }
                    }
                    else
                    {
                        MessageBox.Show("Nested RichTextContentControls not allowed", 
                            "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }

                }
                catch (NullReferenceException)
                {
                    MessageBox.Show("Try again!", "NullReferenceException", 
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
        }

        internal void addLock()
        {

            if (this.rt != null)
            {
                if (!rt.LockContents)
                {
                    //this.rt.LockContentControl = true;
                    this.rt.LockContents = true;
                    this.rt.Title = "LOCKED";

                    Globals.Ribbons.ControlsZrh.label1.Label = "LOCKED";

                    MessageBox.Show("Lock has been set", "Lock set!", 
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            else
            {
                MessageBox.Show("Place cursor into RichtextField and try again!", 
                    "NullReferenceException", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }


        }


        internal void remLock()
        {

            if (this.rt != null)
            {
                if (rt.LockContents)
                {
                    DialogResult dialogResult = MessageBox.Show("Are you sure you want to make " + 
                        "changes to mandatory requirements???\n" + 
                        "You will need to consider other preventive measures to fulfill the risk assessment!!", 
                        "Change MUCK Requirements", MessageBoxButtons.YesNo);

                    if (dialogResult == DialogResult.Yes)
                    {
                        this.rt.LockContentControl = false;
                        this.rt.LockContents = false;
                        this.rt.Title = "UNLOCKED";

                        Globals.Ribbons.ControlsZrh.label1.Label = "NO LOCK";

                        MessageBox.Show("Lock has been removed", "Lock removed!", 
                            MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    else if (dialogResult == DialogResult.No)
                    {
                        //no action
                    }


                }
            }
            else
            {
                MessageBox.Show("Place cursor into RichtextField and try again!", 
                    "NullReferenceException", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }


        }


    }
}
