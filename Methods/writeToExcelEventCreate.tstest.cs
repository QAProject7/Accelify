using Telerik.TestingFramework.Controls.KendoUI;
using Telerik.WebAii.Controls.Html;
using Telerik.WebAii.Controls.Xaml;
using System;
using System.Collections.Generic;
using System.Text;
using System.Linq;

using ArtOfTest.Common.UnitTesting;
using ArtOfTest.WebAii.Core;
using ArtOfTest.WebAii.Controls.HtmlControls;
using ArtOfTest.WebAii.Controls.HtmlControls.HtmlAsserts;
using ArtOfTest.WebAii.Design;
using ArtOfTest.WebAii.Design.Execution;
using ArtOfTest.WebAii.ObjectModel;
using ArtOfTest.WebAii.Silverlight;
using ArtOfTest.WebAii.Silverlight.UI;

namespace Accelify
{

    public class writeToExcel : BaseWebAiiTest
    {
        #region [ Dynamic Pages Reference ]

        private Pages _pages;

        /// <summary>
        /// Gets the Pages object that has references
        /// to all the elements, frames or regions
        /// in this project.
        /// </summary>
        public Pages Pages
        {
            get
            {
                if (_pages == null)
                {
                    _pages = new Pages(Manager.Current);
                }
                return _pages;
            }
        }

        #endregion
        
        // Add your test methods here...
    
        [CodedStep(@"WriteToExcel")]
        public void writeToExcel_run()
        {
            writeToExcel_CodedStep();
        }
         [CodedStep(@"WriteToExcel2")]
       public void writeToExcel_CodedStep()
        {
        string dataSourcePath = this.ExecutionContext.DeploymentDirectory + @"\Data\PerformanceTestDataPSSP.xls";
string myPath = "C:\\report\\domainResults.xls";

if (!System.IO.File.Exists(myPath))
{
    System.IO.File.Copy(dataSourcePath, myPath);
}

Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
Microsoft.Office.Interop.Excel.Workbook workbook = excelApp.Workbooks.Open(myPath);

System.Threading.Thread.Sleep(1000);
//ActiveBrowser.RefreshDomTree();
var eventgroup = Data["EventName"].ToString();
if (Utility.eventCreated)
{
    
    excelApp.Cells[Data.IterationIndex + 2 , 3] = "Event "+eventgroup+ " is created";
}
else
{
    excelApp.Cells[Data.IterationIndex + 2 , 3] = "Event "+eventgroup+ " is NOT created";
}

excelApp.Visible = true;
excelApp.ActiveWorkbook.Save();

workbook.Close(false, Type.Missing, Type.Missing);
excelApp.Workbooks.Close();
System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook);

excelApp.Quit();
GC.Collect();
System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);
 Utility.eventCreated = false;
    }
    }
}
