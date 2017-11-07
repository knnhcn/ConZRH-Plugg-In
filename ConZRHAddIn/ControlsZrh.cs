using Microsoft.Office.Tools.Ribbon;

namespace ConZRHAddIn
{
    public partial class ControlsZrh
    {
        private void ControlsZrh_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void addLockButton_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.addLock();
        }

        private void removeLockButton_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.remLock();
        }

        private void creatLockButton_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.SetRichTextControlOnDocument();
        }


        private void button2_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.reload();
        }

        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.about();
        }

        private void button3_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.delRichText();
        }
    }
}
