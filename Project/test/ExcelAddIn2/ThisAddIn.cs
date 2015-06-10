using Microsoft.Office.Tools.Excel;
using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;

namespace ExcelAddIn2
{
    public partial class ThisAddIn
    {
        public static Excel._Application App;
 
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            App = this.Application;
            
 
            
            
 //Ribbon1
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }


        protected override IRibbonExtension[] CreateRibbonObjects()
        {
            IRibbonExtension[] RibbonExtension = new IRibbonExtension[1];

            RibbonExtension[0] = new Ribbon1();
            return RibbonExtension;
        }

        #region VSTO에서 생성한 코드

        /// <summary>
        /// 디자이너 지원에 필요한 메서드입니다.
        /// 이 메서드의 내용을 코드 편집기로 수정하지 마십시오.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
    
        }
        
        #endregion
    }
}
