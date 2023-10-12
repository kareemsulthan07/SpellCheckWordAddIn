using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using Office = Microsoft.Office.Core;
using Forms = System.Windows.Forms;
using Microsoft.Office.Interop.Word;
using System.Net.NetworkInformation;

// TODO:  Follow these steps to enable the Ribbon (XML) item:

// 1: Copy the following code block into the ThisAddin, ThisWorkbook, or ThisDocument class.

//  protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
//  {
//      return new MySpellCheckRibbon();
//  }

// 2. Create callback methods in the "Ribbon Callbacks" region of this class to handle user
//    actions, such as clicking a button. Note: if you have exported this Ribbon from the Ribbon designer,
//    move your code from the event handlers to the callback methods and modify the code to work with the
//    Ribbon extensibility (RibbonX) programming model.

// 3. Assign attributes to the control tags in the Ribbon XML file to identify the appropriate callback methods in your code.  

// For more information, see the Ribbon XML documentation in the Visual Studio Tools for Office Help.


namespace SpellCheckWordAddIn
{
    [ComVisible(true)]
    public class MySpellCheckRibbon : Office.IRibbonExtensibility
    {
        private Office.IRibbonUI ribbon;

        public MySpellCheckRibbon()
        {
        }

        public void openSpellCheckButton_onAction(Office.IRibbonControl control)
        {
            try
            {
                SpellCheckUserControl userControl = new SpellCheckUserControl();

                Forms.Integration.ElementHost elementHost = new Forms.Integration.ElementHost()
                {
                    Dock = Forms.DockStyle.Fill,
                    Child = userControl,
                };

                Forms.UserControl wfUserControl = new Forms.UserControl()
                {
                    Font = new System.Drawing.Font("Calibri", 10f, System.Drawing.FontStyle.Regular),
                };
                wfUserControl.Controls.Add(elementHost);

                var taskPane = Globals.ThisAddIn.CustomTaskPanes.Add(wfUserControl, "My Spell Check");
                taskPane.Width = 450;
                taskPane.DockPositionRestrict = Office.MsoCTPDockPositionRestrict.msoCTPDockPositionRestrictNoHorizontal;
                taskPane.DockPosition = Office.MsoCTPDockPosition.msoCTPDockPositionRight;
                taskPane.Visible = true;
            }
            catch (Exception)
            {

                throw;
            }
        }

        #region IRibbonExtensibility Members

        public string GetCustomUI(string ribbonID)
        {
            return GetResourceText("SpellCheckWordAddIn.MySpellCheckRibbon.xml");
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
