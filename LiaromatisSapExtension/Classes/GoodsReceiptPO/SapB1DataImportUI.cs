// ===============================================================================================
// 1.0.1
// ===============================================================================================
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using SAPbouiCOM.Framework;
using System.Text.RegularExpressions;
using System.IO;

namespace LiaromatisSapExtension.Classes.GoodsReceiptPO
{
    class SapB1DataImportUI
    {
        public static void ImportUDO(string _FormUID, SAPbobsCOM.Company _Company)
        {
            SAPbouiCOM.Form oForm = null;
            //SAPbouiCOM.Item oItem = null;
            SAPbouiCOM.Item oItemMatrix = null;
            SAPbouiCOM.EditText oEditText = null;
            SAPbouiCOM.EditText oEditTextPath = null;
            //SAPbouiCOM.ComboBox oComboBox = null;

            try
            {
                oForm = Application.SBO_Application.Forms.Item(_FormUID);
                oForm.Freeze(true);

                if (oForm.Mode.ToString() == "fm_ADD_MODE" && oForm.Items.Item("14").Enabled == true) //only in add mode. I use this (Item("14").Enabled) to avoid the creation of udo in Cancellation mode
                {
                    oItemMatrix = oForm.Items.Item("38");
                    SAPbouiCOM.Matrix oMatrix = ((SAPbouiCOM.Matrix)(oItemMatrix.Specific));

                    //string DocNum = null;
                    //oItem = oForm.Items.Item("8");
                    //oEditText = (SAPbouiCOM.EditText)oItem.Specific;
                    //DocNum =  oEditText.Value.ToString();

                    for (int i = 1; i < oMatrix.VisualRowCount + 1; i++)
                    {
                        //oComboBox = (SAPbouiCOM.ComboBox)oMatrix.Columns.Item("U_Cert").Cells.Item(i).Specific;
                        //oComboBox.Value.ToString();

                        oEditText = (SAPbouiCOM.EditText)oMatrix.Columns.Item("U_Cert").Cells.Item(i).Specific;
                        oEditTextPath = (SAPbouiCOM.EditText)oMatrix.Columns.Item("U_Path").Cells.Item(i).Specific;

                        if ((oEditText.Value.ToString() == "Y" || oEditText.Value.ToString() == "Yes") && String.IsNullOrEmpty(oEditTextPath.Value.ToString()))
                        {
                            string DocEntry = null;

                            //Create a New UDO
                            DocEntry = AddUDO(_Company);
                            
                            //Add Udo DocEntry in Lines
                            oEditTextPath.Value = DocEntry;
                        }
                    }
                }
                oForm.Freeze(false);
            }
            catch (Exception ex)
            {
                oForm.Freeze(false);
            }
        }

        public static void ImportPath(SAPbobsCOM.Company _Company)
        {
            string VisLineNum = null;
            string DocNum = null;
            string AtchPath = null;
            string SAPAttachPath = null;
            string SourceClientPath = null;
            string TargetClientPath = null;

            try
            {
                AtchPath = $@" SELECT AttachPath FROM OADP ";
                SAPbobsCOM.Recordset rsAttachPath;
                rsAttachPath = (SAPbobsCOM.Recordset)_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                rsAttachPath.DoQuery(AtchPath);
                rsAttachPath.MoveFirst();

                SAPAttachPath = rsAttachPath.Fields.Item("AttachPath").Value.ToString();

                System.Runtime.InteropServices.Marshal.ReleaseComObject(rsAttachPath);
                AtchPath = null;

                AtchPath = $@" SELECT Code, Name FROM [@DDS_ATTACHPATH] ";
                SAPbobsCOM.Recordset rsAttachPath2;
                rsAttachPath2 = (SAPbobsCOM.Recordset)_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                rsAttachPath2.DoQuery(AtchPath);

                rsAttachPath2.MoveFirst();
                while (rsAttachPath2.EoF == false)
                {
                    if (rsAttachPath2.Fields.Item("Code").Value.ToString() == "SourcePath")
                        SourceClientPath = rsAttachPath2.Fields.Item("Name").Value.ToString();
                    else if (rsAttachPath2.Fields.Item("Code").Value.ToString() == "TargetPath")
                        TargetClientPath = rsAttachPath2.Fields.Item("Name").Value.ToString();

                    rsAttachPath2.MoveNext();
                }
                System.Runtime.InteropServices.Marshal.ReleaseComObject(rsAttachPath2);                

                if (!String.IsNullOrEmpty(SAPAttachPath) && !String.IsNullOrWhiteSpace(SAPAttachPath) && 
                    !String.IsNullOrEmpty(SourceClientPath) && !String.IsNullOrWhiteSpace(SourceClientPath) && 
                    !String.IsNullOrEmpty(TargetClientPath) && !String.IsNullOrWhiteSpace(TargetClientPath))
                {
                    DirectoryInfo d = new DirectoryInfo(SourceClientPath);//Assuming Test is your Folder
                    FileInfo[] Files = d.GetFiles("*.pdf"); //Getting Text files
                    foreach (FileInfo file in Files)
                    {
                        try
                        {
                            string fileName = file.Name;
                            int charLocation = file.Name.IndexOf("_", StringComparison.Ordinal);
                            int charLocation2 = file.Name.IndexOf("@", StringComparison.Ordinal);
                            int charLocation3 = file.Name.Length;
                            string UdoEntry = null;

                            if (charLocation > 0)
                            {
                                DocNum = file.Name.Substring(0, charLocation);
                                if (charLocation2 > 0)
                                    VisLineNum = file.Name.Substring(charLocation + 1, charLocation2 - charLocation - 1);
                                else
                                    VisLineNum = file.Name.Substring(charLocation + 1, charLocation3 - charLocation - 5);

                                string UpdateUdo = null;
                                UpdateUdo = $@" SELECT Q1.U_Path, Q1.LineNum , Q1.VisOrder, Q1.VisOrder + SUM(Q1.AddLine) + 1 AS [VisLineNum]  FROM
		                                                (SELECT T1.U_Path, T1.LineNum,t1.VisOrder,T2.AftLineNum, CASE WHEN (VisOrder > AftLineNum ) THEN 1 ELSE 0 END AS AddLine FROM OPDN T0
		                                                INNER JOIN PDN1 T1 ON T0.DocEntry = T1.DocEntry
		                                                LEFT JOIN PDN10 T2 ON T1.DocEntry = T2.DocEntry
		                                                WHERE T0.DocNum = '{DocNum}'
		                                                AND T0.CANCELED = 'N') Q1
                                                WHERE ISNULL(Q1.U_Path,'') != ''
                                                GROUP BY Q1.U_Path, Q1.LineNum , Q1.VisOrder
                                                HAVING (Q1.VisOrder + SUM(Q1.AddLine) +1) = '{VisLineNum}' ";

                                SAPbobsCOM.Recordset rsUpdateUdo;
                                rsUpdateUdo = (SAPbobsCOM.Recordset)_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                                rsUpdateUdo.DoQuery(UpdateUdo);

                                rsUpdateUdo.MoveFirst();
                                UdoEntry = rsUpdateUdo.Fields.Item("U_Path").Value.ToString();
                                if (rsUpdateUdo.RecordCount == 1)
                                {
                                    string sourceFile = System.IO.Path.Combine(SourceClientPath, fileName);
                                    string SAPdestFile = System.IO.Path.Combine(SAPAttachPath, fileName);
                                    string destinationFile = System.IO.Path.Combine(TargetClientPath, fileName);
                                    Boolean flagUpdateUDOPath = false;
                                    try
                                    {
                                        //Copy or overwrite the destination file if it already exists.
                                        System.IO.File.Copy(sourceFile, SAPdestFile, true);
                                        flagUpdateUDOPath = UpdateUDO(_Company, UdoEntry, SAPdestFile);
                                    }
                                    catch (Exception ex)
                                    {
                                        Application.SBO_Application.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                                    }
                                    try
                                    {
                                        if (flagUpdateUDOPath == true)
                                        {
                                            System.IO.File.Copy(sourceFile, destinationFile, true);
                                            System.IO.File.Delete(sourceFile);
                                        }
                                    }
                                    catch (Exception ex)
                                    {
                                        Application.SBO_Application.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                                    }
                                }
                                else if (rsUpdateUdo.RecordCount == 0)
                                {
                                    Application.SBO_Application.StatusBar.SetText($@"In Goods Receipt PO either the combination of Line No.: {VisLineNum} and Document No.: {DocNum} or the field Path does not exist.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                                }
                            }
                            else
                            {
                                Application.SBO_Application.StatusBar.SetText($@"Wrong pdf filename. Please try this form: DocNum_LineNo@notes.pdf", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                            }
                        }
                        catch (Exception ex)
                        {
                            Application.SBO_Application.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                        }
                    }
                }
                else
                {
                    Application.SBO_Application.StatusBar.SetText($@"Attachments folders not defined, or Attachments folders has been changed or removed", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);                    
                }
            }
            catch (Exception ex)
            {}
        }

        public static Boolean UpdateUDO(SAPbobsCOM.Company _Company, string _DocEntry, string _SAPdestFile)
        {
            try
            {
                Boolean flagExistPath = false;
                string ExistPath = null;
                ExistPath = $@" SELECT U_Path FROM [@DDS_PATHL]
                                WHERE DocEntry = {_DocEntry} ";

                SAPbobsCOM.Recordset rsExistPath;
                rsExistPath = (SAPbobsCOM.Recordset)_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                rsExistPath.DoQuery(ExistPath);

                rsExistPath.MoveFirst();
                while (rsExistPath.EoF == false)
                {
                    if (rsExistPath.Fields.Item("U_Path").Value.ToString() == _SAPdestFile)
                        flagExistPath = true;

                    rsExistPath.MoveNext();
                }
                System.Runtime.InteropServices.Marshal.ReleaseComObject(rsExistPath);

                if (flagExistPath == false)
                {
                    SAPbobsCOM.GeneralService oGeneralService = null;
                    SAPbobsCOM.GeneralData oGeneralData = null;
                    SAPbobsCOM.GeneralDataParams oGeneralParams = null;
                    SAPbobsCOM.GeneralDataCollection oSons = null;
                    SAPbobsCOM.GeneralData oSon = null;
                    SAPbobsCOM.CompanyService sCmp = null;

                    sCmp = (SAPbobsCOM.CompanyService)_Company.GetCompanyService();
                    oGeneralService = sCmp.GetGeneralService("PATH");
                    oGeneralParams = (SAPbobsCOM.GeneralDataParams) oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams);
                    oGeneralParams.SetProperty("DocEntry", _DocEntry);
                    oGeneralData = (SAPbobsCOM.GeneralData)oGeneralService.GetByParams(oGeneralParams);

                    //Update UDO Header
                    //oData.SetProperty("U_NumAtCard", _NumAtCard);

                    //Update UDO Child
                    oSons = oGeneralData.Child("DDS_PATHL");
                    //oSons.Item(1).SetProperty("","");

                    oSon = oSons.Add();
                    oSon.SetProperty("U_Path", _SAPdestFile);

                    oGeneralService.Update(oGeneralData);
                }
                return true;
            }
            catch (Exception ex)
            {
                return false;
            }
        }

        public static string AddUDO(SAPbobsCOM.Company _Company)
        {
            string DocEntry = null;
            try
            {
                SAPbobsCOM.GeneralService oGeneralService = null;
                SAPbobsCOM.GeneralData oGeneralData = null;
                SAPbobsCOM.GeneralDataParams oGeneralParams = null;
                SAPbobsCOM.CompanyService sCmp = null;

                sCmp = (SAPbobsCOM.CompanyService)_Company.GetCompanyService();
                oGeneralService = sCmp.GetGeneralService("PATH");
                oGeneralData = (SAPbobsCOM.GeneralData)oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData);

                //oGeneralData.SetProperty("U_NumAtCard", _NumAtCard);

                //Add Udo PATH
                oGeneralParams = oGeneralService.Add(oGeneralData);
                //Get the DocEntry of the Added Udo
                DocEntry = oGeneralParams.GetProperty("DocEntry").ToString();
            }
            catch (Exception ex)
            { }
            return DocEntry;
        }

        public static void CancelUDO(string _FormUID, SAPbobsCOM.Company _Company)
        {
            string DocSeries = null;
            SAPbouiCOM.Form oForm = null;
            SAPbouiCOM.EditText oEditText = null;
            SAPbouiCOM.ComboBox oComboBox = null;
            SAPbouiCOM.Item oItemMatrix = null;

            try
            {
                oForm = Application.SBO_Application.Forms.Item(_FormUID);
                oForm.Freeze(true);

                // Get Cancellation Series
                DocSeries = GetSeries(_Company);

                string DescSeries;
                string ValSeries;
                try
                {
                    oComboBox = (SAPbouiCOM.ComboBox)oForm.Items.Item("88").Specific;
                    DescSeries = oComboBox.Selected.Description;
                    ValSeries = oComboBox.Selected.Value;

                    //    oItem = oForm.Items.Item("2");
                    //    oEditText = (SAPbouiCOM.EditText)oItem.Specific;
                    //    x = oEditText.Value.ToString();

                    if (ValSeries == DocSeries)
                    {
                        oItemMatrix = oForm.Items.Item("38");
                        SAPbouiCOM.Matrix oMatrix = ((SAPbouiCOM.Matrix)(oItemMatrix.Specific));
                        for (int i = 1; i < oMatrix.VisualRowCount + 1; i++)
                        {
                            string UdoDocEntry = null;
                            oEditText = (SAPbouiCOM.EditText)oMatrix.Columns.Item("U_Path").Cells.Item(i).Specific;
                            UdoDocEntry = oEditText.Value.ToString();

                            if (!String.IsNullOrEmpty(UdoDocEntry) && int.Parse(UdoDocEntry) >= 0)
                            {
                                //Cancel UDO
                                SAPbobsCOM.GeneralService oGeneralService = null;
                                SAPbobsCOM.GeneralDataParams oGeneralParams = null;
                                SAPbobsCOM.CompanyService sCmp = null;

                                sCmp = (SAPbobsCOM.CompanyService)_Company.GetCompanyService();
                                oGeneralService = sCmp.GetGeneralService("PATH");
                                oGeneralParams = (SAPbobsCOM.GeneralDataParams)oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams);
                                oGeneralParams.SetProperty("DocEntry", UdoDocEntry);
                                
                                oGeneralService.Cancel(oGeneralParams);
                            }
                        }
                    }
                }
                catch (Exception ex){}

                oForm.Freeze(false);
            }
            catch (Exception ex)
            {
                oForm.Freeze(false);
            }
        }

        public static string GetSeries(SAPbobsCOM.Company _Company)
        {
            string Series = null;
            try
            {
                
                string strSeries = $@" SELECT Series 
                                       FROM NNM1
                                       WHERE ObjectCode = '20'
                                       AND IsForCncl = 'Y' ";

                SAPbobsCOM.Recordset rs;
                rs = (SAPbobsCOM.Recordset)_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                try
                {
                    rs.DoQuery(strSeries);
                }
                catch (Exception)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(rs);
                }

                rs.MoveFirst();
                while (rs.EoF == false)
                {
                    Series = rs.Fields.Item("Series").Value.ToString();
                    rs.MoveNext();
                }
                System.Runtime.InteropServices.Marshal.ReleaseComObject(rs);
            }
            catch (Exception ex)
            { }
            return Series;
        }
    }
}
