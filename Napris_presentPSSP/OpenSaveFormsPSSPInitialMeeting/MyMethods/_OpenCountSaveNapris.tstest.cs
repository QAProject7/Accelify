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

    public class _OpenCountSaveNapris : BaseWebAiiTest
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
    
        [CodedStep(@"Open Save Form And Count")]
        public void OpenSaveForm_Count()
        {
            Console.Out.WriteLine("Open Current form: " + Data["FormName"].ToString());
            Log.WriteLine("Open Current form: " + Data["FormName"].ToString());
               var watch = System.Diagnostics.Stopwatch.StartNew();
               this.ExecuteTest("Napris_presentPSSP\\OpenSaveFormsPSSPInitialMeeting\\MyMethods\\_OpenFormNapris.tstest");
               watch.Stop();
            myUtility.title = GetExtractedValue("FormTitleName");
            Log.WriteLine("Time elapsed: "+watch.ElapsedMilliseconds+" ms");
            myUtility.opentime = watch.ElapsedMilliseconds;
            
            Console.Out.WriteLine("Saving Current form: " + Data["FormName"].ToString());
            Log.WriteLine("Saving Current form: " + Data["FormName"].ToString());
               watch = System.Diagnostics.Stopwatch.StartNew();
               this.ExecuteTest("Napris_presentPSSP\\OpenSaveFormsPSSPInitialMeeting\\MyMethods\\_SaveFormNapris.tstest");
               watch.Stop();
            Log.WriteLine("Time elapsed: "+watch.ElapsedMilliseconds+" ms");
            myUtility.savetime = watch.ElapsedMilliseconds;
            
            this.ExecuteTest("Napris_presentPSSP\\OpenSaveFormsPSSPInitialMeeting\\MyMethods\\writeToExcelFormsData.tstest");
            myUtility.row = Data.IterationIndex+2;
        }
    }
}
