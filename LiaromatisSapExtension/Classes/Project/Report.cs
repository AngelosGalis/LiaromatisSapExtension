using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using SAPbouiCOM.Framework;

namespace LiaromatisSapExtension.Classes.Project
{
    class Report
    {
        public static void ImportAttachments(SAPbobsCOM.Company _Company, string _FormUID, Models.ReportMenuUidModel _ReportMenuUidModel, string _UserName)
        {
            SAPbouiCOM.Form oForm = null;
            SAPbouiCOM.Form oFormReport = null;
            SAPbouiCOM.Item oItem = null;
            SAPbouiCOM.EditText oEditText = null;
            SAPbouiCOM.Item oItemMatrix = null;
            SAPbouiCOM.Matrix oMatrix = null;
            SAPbouiCOM.ComboBox oComboBox = null;

            try
            {
                oForm = Application.SBO_Application.Forms.Item(_FormUID);
                oForm.Freeze(true);
                               
                string Project = null;
                string AllUniqueID = null;
                Boolean flagExist = false;

                oItem = oForm.Items.Item("234000049");
                oEditText = (SAPbouiCOM.EditText)oItem.Specific;
                Project = oEditText.Value.ToString();

                if (!String.IsNullOrEmpty(Project))
                {
                    try //Delete old Data from the Table
                    {
                        string Delete = $@" DELETE FROM [dbo].[DDS_T_ListOfOutstandingCertificates]
                                            WHERE UserName = N'{_UserName}' ";

                        SAPbobsCOM.Recordset rs;
                        rs = (SAPbobsCOM.Recordset)_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                        try
                        {
                            rs.DoQuery(Delete);
                        }
                        catch (Exception)
                        {
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(rs);
                        }

                        System.Runtime.InteropServices.Marshal.ReleaseComObject(rs);
                    }
                    catch (Exception)
                    { }

                    oItemMatrix = oForm.Items.Item("234000062");
                    oMatrix = (SAPbouiCOM.Matrix)oItemMatrix.Specific;

                    for (int i = 1; i < oMatrix.VisualRowCount + 1; i++)
                    {
                        oComboBox = (SAPbouiCOM.ComboBox)oMatrix.Columns.Item("U_ExStage").Cells.Item(i).Specific;
                        oComboBox.Value.ToString();

                        if (oComboBox.Value.ToString() == "Yes")
                        {
                            oEditText = (SAPbouiCOM.EditText)oMatrix.Columns.Item("234000033").Cells.Item(i).Specific;

                            if (AllUniqueID == null && !String.IsNullOrEmpty(oEditText.Value.ToString())) 
                                AllUniqueID = oEditText.Value.ToString();
                            else if (!String.IsNullOrEmpty(oEditText.Value.ToString())) 
                                AllUniqueID = AllUniqueID + "," + oEditText.Value.ToString();

                            //Insert selected stages to the Table
                            try
                            {
                                string Insert = $@" INSERT INTO [dbo].[DDS_T_ListOfOutstandingCertificates]
                                                           ([Project]
                                                           ,[UniqueID]
                                                           ,[UserName])
                                                    VALUES
                                                           (N'{Project}'
                                                           ,N'{oEditText.Value.ToString()}'
                                                           ,N'{_UserName}') ";

                                SAPbobsCOM.Recordset rs;
                                rs = (SAPbobsCOM.Recordset)_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                                try
                                {
                                    rs.DoQuery(Insert);
                                }
                                catch (Exception)
                                {
                                    System.Runtime.InteropServices.Marshal.ReleaseComObject(rs);
                                }
                                
                                System.Runtime.InteropServices.Marshal.ReleaseComObject(rs);
                            }
                            catch (Exception)
                            { }
                            flagExist = true;
                        }
                    }

                    //If no flag is yes get all stages
                    if (flagExist == false)
                    {
                        for (int i = 1; i < oMatrix.VisualRowCount + 1; i++)
                        {
                            oEditText = (SAPbouiCOM.EditText)oMatrix.Columns.Item("234000033").Cells.Item(i).Specific;

                            if (AllUniqueID == null && !String.IsNullOrEmpty(oEditText.Value.ToString()))
                                AllUniqueID = oEditText.Value.ToString();
                            else if (!String.IsNullOrEmpty(oEditText.Value.ToString()))
                                AllUniqueID = AllUniqueID + "," + oEditText.Value.ToString();

                            //Insert all stages to the Table
                            try
                            {
                                string Insert = $@" INSERT INTO [dbo].[DDS_T_ListOfOutstandingCertificates]
                                                           ([Project]
                                                           ,[UniqueID]
                                                           ,[UserName])
                                                    VALUES
                                                           (N'{Project}'
                                                           ,N'{oEditText.Value.ToString()}'
                                                           ,N'{_UserName}') ";

                                SAPbobsCOM.Recordset rs;
                                rs = (SAPbobsCOM.Recordset)_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                                try
                                {
                                    rs.DoQuery(Insert);
                                }
                                catch (Exception)
                                {
                                    System.Runtime.InteropServices.Marshal.ReleaseComObject(rs);
                                }

                                System.Runtime.InteropServices.Marshal.ReleaseComObject(rs);
                            }
                            catch (Exception)
                            { }
                        }
                    }

                    try //Update field AllUniqueID with all stages in the Table
                    {
                        string Update = $@" UPDATE [dbo].[DDS_T_ListOfOutstandingCertificates]
                                            SET AllUniqueID = N'{AllUniqueID}'
                                            WHERE UserName = N'{_UserName}' ";

                        SAPbobsCOM.Recordset rs;
                        rs = (SAPbobsCOM.Recordset)_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                        try
                        {
                            rs.DoQuery(Update);
                        }
                        catch (Exception ex)
                        {
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(rs);
                        }

                        System.Runtime.InteropServices.Marshal.ReleaseComObject(rs);
                    }
                    catch (Exception)
                    { }
                }
                oForm.Freeze(false);
                
                if (oForm.Mode.ToString() != "fm_ADD_MODE" && oForm.Mode.ToString() != "fm_FIND_MODE")
                {
                    // Triger Report and set Parameters to it
                    Application.SBO_Application.ActivateMenuItem(_ReportMenuUidModel.MenuUid_ListOfOutstandingCertificates);
                   
                    oFormReport = Application.SBO_Application.Forms.ActiveForm;
                    //oFormReport.Visible = false;
                    oItem = oFormReport.Items.Item("1000003");
                    oEditText = (SAPbouiCOM.EditText)oItem.Specific;
                    oEditText.Value = _UserName;

                    oItem = oFormReport.Items.Item("1");
                    oItem.Click();
                    //oFormReport.Visible = true;
                }
            }
            catch (Exception ex)
            {
                oForm.Freeze(false);
                Application.SBO_Application.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
        }
    }
}
