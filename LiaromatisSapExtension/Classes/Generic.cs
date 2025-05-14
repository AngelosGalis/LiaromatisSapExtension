using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using SAPbouiCOM.Framework;

namespace LiaromatisSapExtension.Classes
{
    class Generic
    {
        public static void SetGLAccount(string _FormUID, SAPbobsCOM.Company _Company)
        {
            SAPbouiCOM.Form oForm = null;
            SAPbouiCOM.EditText oEditText = null;
            SAPbouiCOM.ComboBox oComboBox = null;
            SAPbouiCOM.Item oItemMatrix = null;
            SAPbouiCOM.Matrix oMatrix = null;
            string DescFPA;
            string ValFPA;
            string Account = null;

            try
            {
                oForm = Application.SBO_Application.Forms.Item(_FormUID);
                oForm.Freeze(true);

                oItemMatrix = oForm.Items.Item("38");
                oMatrix = (SAPbouiCOM.Matrix)oItemMatrix.Specific;

                for (int i = 1; i < oMatrix.VisualRowCount; i++)
                {
                    try
                    {
                        oComboBox = (SAPbouiCOM.ComboBox)oMatrix.Columns.Item("18").Cells.Item(i).Specific;
                        DescFPA = oComboBox.Selected.Description;
                        ValFPA = oComboBox.Selected.Value;

                        string qGLAccount = $@"SELECT Q1.VatGroup, Q1.Account 
                                                FROM (  SELECT VatGroup, ECIncome AS Account FROM OGAR
		                                                WHERE ISNULL(ECIncome, '') != ''
		                                                UNION
		                                                SELECT VatGroup, DfltIncom AS Account FROM OGAR
		                                                WHERE ISNULL(DfltIncom, '') != ''
		                                                UNION
		                                                SELECT VatGroup, ForgnIncm AS Account FROM OGAR
		                                                WHERE ISNULL(ForgnIncm, '') != ''
	                                                ) Q1
                                                WHERE Q1.VatGroup = N'{ValFPA}'";

                        SAPbobsCOM.Recordset rsGLAccount;
                        rsGLAccount = (SAPbobsCOM.Recordset)_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                        rsGLAccount.DoQuery(qGLAccount);
                        rsGLAccount.MoveFirst();
                        Account = rsGLAccount.Fields.Item("Account").Value.ToString();
                        ComObjectDisposer.ReleaseComObject(rsGLAccount, null, null, null, null);

                        oEditText = (SAPbouiCOM.EditText)oMatrix.Columns.Item("29").Cells.Item(i).Specific;
                        oEditText.Value = Account;
                    }
                    catch (Exception ex) 
                    {
                        Application.SBO_Application.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                    }
                }
                oForm.Freeze(false);
            }
            catch (Exception ex)
            {
                oForm.Freeze(false);
            }
        }
    }
}
