// ===============================================================================================
// 1.0.1
// ===============================================================================================
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using SAPbouiCOM.Framework;

namespace LiaromatisSapExtension.Classes.GoodsReceiptPO
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
                SAPbouiCOM.Button oButton = null;

                // 1.1.1.1 Add a new Button "TAXISnet" in BP
                oItemRef = oForm.Items.Item("10000329");
                oItem = oForm.Items.Add("btAtch", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
                oItem.Left = oItemRef.Left; //oItemRef.Left - oItemRef.Width - 10;
                oItem.Width = oItemRef.Width;
                oItem.Top = oItemRef.Top - oItemRef.Height - 10;
                oItem.Height = oItemRef.Height;
                oItem.Enabled = true;
                oButton = (SAPbouiCOM.Button)oItem.Specific;
                oButton.Caption = "Attachments";
            }
            catch (Exception)
            { }
        }
    }
}
