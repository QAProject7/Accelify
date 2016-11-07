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

    public class Count_Time_for_Save_Present_Levels : BaseWebAiiTest
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
    
        [CodedStep(@"Save Present Levels and Count")]
       public void Count_Time_for_Save_Present_Levels_CodedStep()
        {
             Console.Out.WriteLine("Current form: "+Data["formname"].ToString());
             var watch = System.Diagnostics.Stopwatch.StartNew();
this.ExecuteTest("ElenaKochegina_presentPSS\\my_methods\\Save_Pressent_Levels.tstest");
watch.Stop();
            Log.WriteLine("Time elapsed: "+watch.ElapsedMilliseconds);
Utility.opentime = watch.ElapsedMilliseconds;
            
        }
    }
}