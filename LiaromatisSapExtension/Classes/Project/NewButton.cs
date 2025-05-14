// ===============================================================================================
// 1.0.1
// ===============================================================================================
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using SAPbouiCOM.Framework;

namespace LiaromatisSapExtension.Classes.Project
{
    class NewButton
    {
        public static void ImportAttachments(string _FormUID)
        {
            try
            {
                SAPbouiCOM.Form oForm;
                oForm = Application.SBO_Application.Forms.Item(_FormUID);
                SAPbouiCOM.Item oItem = null;
                SAPbouiCOM.Item oItemRef = null;
                SAPbouiCOM.Item oItemRef2 = null;
                SAPbouiCOM.Button oButton = null;

                // Add a new Buttons in Goods Receipt PO
                oItemRef = oForm.Items.Item("1");
                oItemRef2 = oForm.Items.Item("234000049");
                oItem = oForm.Items.Add("btExpt", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
                oItem.Left = oItemRef2.Left;
                oItem.Width = oItemRef2.Width;
                oItem.Top = oItemRef.Top;
                oItem.Height = oItemRef.Height;
                oItem.Enabled = true;
                oButton = (SAPbouiCOM.Button)oItem.Specific;
                oButton.Caption = "Export Certificates";
            }
            catch (Exception)
            { }
        }

        public static void Report_ListOutCert(string _FormUID)
        {
            try
            {
                SAPbouiCOM.Form oForm;
                oForm = Application.SBO_Application.Forms.Item(_FormUID);
                SAPbouiCOM.Item oItem = null;
                SAPbouiCOM.Item oItemRef = null;
                SAPbouiCOM.Item oItemRef2 = null;
                SAPbouiCOM.Button oButton = null;

                // Add a new Buttons in Goods Receipt PO
                oItemRef = oForm.Items.Item("1");
                oItemRef2 = oForm.Items.Item("234000028");
                oItem = oForm.Items.Add("btRrt1", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
                oItem.Left = oItemRef2.Left - oItemRef.Width - 10;
                oItem.Width = oItemRef.Width;
                oItem.Top = oItemRef2.Top;
                oItem.Height = oItemRef.Height;
                oItem.Enabled = true;
                oButton = (SAPbouiCOM.Button)oItem.Specific;
                oButton.Caption = "Report";
            }
            catch (Exception)
            { }
        }
    }
}
