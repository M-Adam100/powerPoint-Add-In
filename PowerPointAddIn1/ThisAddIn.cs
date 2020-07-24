using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;
using System.Windows.Forms;
using System.Net;

namespace PowerPointAddIn1
{
    public partial class ThisAddIn
    {
        private CustomePane pane;
        private Microsoft.Office.Tools.CustomTaskPane myCustomTaskPane;
        void Application_PresentationNewSlide(PowerPoint.Slide Sld)
        {
           
            PowerPoint.Shape textBox = Sld.Shapes.AddTextbox(
                Office.MsoTextOrientation.msoTextOrientationHorizontal, 100, 0, 500, 50);
            textBox.TextFrame.TextRange.InsertAfter("Hello World");
        }
       
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            pane = new CustomePane();
            myCustomTaskPane = this.CustomTaskPanes.Add(pane, "My Task Pane");
            myCustomTaskPane.Visible = true;

            //this.Application.PresentationNewSlide +=
            //new PowerPoint.EApplication_PresentationNewSlideEventHandler(
            //Application_PresentationNewSlide);
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
    }
}
