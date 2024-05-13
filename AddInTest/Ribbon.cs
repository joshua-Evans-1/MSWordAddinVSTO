using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;
using Office = Microsoft.Office.Core;
using Word = Microsoft.Office.Interop.Word;

namespace AddInTest
{
    [ComVisible(true)]
    public class Ribbon : Office.IRibbonExtensibility
    {
        private Office.IRibbonUI ribbon;
        // when first button is clicked-
        public void OnFirstButton(Office.IRibbonControl control)
        {
            // iterates through all text in document adding highlight at each instance of the word "of"
            Word.Find find = Globals.ThisAddIn.Application.ActiveDocument.Content.Find;
            find.Replacement.Font.ColorIndexBi = Word.WdColorIndex.wdYellow;
            find.Execute(FindText: "of", MatchCase: false, MatchWholeWord: true, Replace: Word.WdReplace.wdReplaceAll);
            // displays the occurrences of the word of in a dialog box
            Word.Range range = Globals.ThisAddIn.Application.ActiveDocument.Range();
            MessageBox.Show( "Occurrences of the word \"of\" - " + range.Text.Split(' ').Count( word => word.Equals("of") ) );
        }
        // ribbon constructor
        public Ribbon()
        {
        }

        #region IRibbonExtensibility Members

        public string GetCustomUI(string ribbonID)
        {
            return GetResourceText("AddInTest.Ribbon.xml");
        }

        #endregion

        #region Ribbon Callbacks
        //Create callback methods here. For more information about adding callback methods, visit https://go.microsoft.com/fwlink/?LinkID=271226

        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        {
            this.ribbon = ribbonUI;
        }

        #endregion

        #region Helpers

        private static string GetResourceText(string resourceName)
        {
            Assembly asm = Assembly.GetExecutingAssembly();
            string[] resourceNames = asm.GetManifestResourceNames();
            for (int i = 0; i < resourceNames.Length; ++i)
            {
                if (string.Compare(resourceName, resourceNames[i], StringComparison.OrdinalIgnoreCase) == 0)
                {
                    using (StreamReader resourceReader = new StreamReader(asm.GetManifestResourceStream(resourceNames[i])))
                    {
                        if (resourceReader != null)
                        {
                            return resourceReader.ReadToEnd();
                        }
                    }
                }
            }
            return null;
        }

        #endregion
    }
}
