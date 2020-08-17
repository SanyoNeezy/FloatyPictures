using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Microsoft.Office.Interop.Word;
using Microsoft.Office.Tools.Ribbon;
using Word = Microsoft.Office.Interop.Word;

namespace FloatyButtons
{
    public partial class MainRibbon
    {
        

        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {
         
        }

        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            DocumentManipulator.activeSelection = ThisAddIn.WordApplication.Selection;
            DocumentManipulator.activeDocument = ThisAddIn.WordApplication.ActiveDocument;
            
            DocumentManipulator.FloatImages();
        }
    }
}
