// Editor: Aggelos  Date: 08/01/2020
// 1.4.0
// 1.4.0.1 (System Form)
// 1.4.0.1.1 Set GLAccount from FPA in Marketing Documents
// *************************************************************************************
// Editor: Aggelos  Date: 17/12/2019
// 1.3.0
// 1.3.0.1 Item Master Data (System Form)
// 1.3.0.1.1 When an ItemMD is Added: Create two empty UDOs MasterData and linked them to ItemMD
// *************************************************************************************
// Editor: Aggelos  Date: 21/10/2019
// 1.2.0
// 1.2.0.1 Delivery (System Form)
// 1.2.0.1.1 When a Delivery is Added another Draft Delivery with mapping SalesOrder added too and existing Delivery Updater too (U_DrfDeliv, U_Delivery).
// 1.2.1.1.1 PopUp an existing Draft Delivery.
// 1.2.1.1.2 When the user click Button "Add" in a Draft Delivery, Update filed stage
// 1.2.1.1.3 When the Draft Delivery Added: Update the fields (U_Delivery, U_DrfDeliv) in Source delivery and Update Packages in UDO PCK with Y
// 1.2.2.1.1 When a Delivery is Cancelled: another Delivery with mapping SalesOrder is Cancelled too and Update Packages in UDO PCK with N OR Remove a Draft Delivery
// 1.2.3.1.1 When a Delivery is Added: Insert Data in Tables DDS_StrumisPck, DDS_StrumisCnt, DDS_StrumisSit
// *************************************************************************************
// Editor: Aggelos  Date: 11/10/2019
// 1.1.0
// 1.1.0.1 Goods Receipt (System Form)
// 1.1.0.1.1 When a Goods Receipt is Added a Goods Issue is created
// 1.1.1 Small changes in (Series, Remarks) // 1.1.0.1.1 
// 1.1.2 Small changes in (Remarks) // 1.1.0.1.1 
// *************************************************************************************
// Editor: Aggelos  Date: 12/09/2019
// 1.0.9
// 1.0.9.1 Project (System Form)
// 1.0.9.1.1 Add Report Parametrically
// 1.0.9.1.2 Add a new Buttons in Project
// 1.0.9.1.3 When the user click Button "btRrt1" PopUp Report
// *************************************************************************************
// Editor: Aggelos  Date: 11/09/2019
// 1.0.8
// 1.0.8.1 Changes in Exportation of csv (1.0.1.2.2). Export additional info for ReceiptDate.
// 1.0.8.2 Comment "Sort" functionality in Project Stage (1.0.0.1.1). When SAP Fixed the bug with line ID and POS in stages in Project we should uncomment it.
// *************************************************************************************
// Editor: Aggelos  Date: 09/09/2019
// 1.0.7 
// 1.0.7.1 Goods Receipt PO (System Form)
// 1.0.7.1.1 When a Goods Receipt PO is Canceled all UDO_Path that are linked with it, are Cancelled too.
// *************************************************************************************
// Editor: Aggelos  Date: 06/09/2019
// 1.0.6 
// 1.0.6.1 Changes in Exportation of pdfs (1.0.1.2.2). Additional distinct with DocNum of GoodsReceipt PO.
// *************************************************************************************
// Editor: Aggelos  Date: 04/09/2019
// 1.0.5 
// 1.0.5.1 Changes in Exportation of pdfs and csv (1.0.1.2.2). Distinct with less columns than them we need to display.
// 1.0.5.2 Changes in (1.0.2.1.1). Set PopUp Report Parametrically
// *************************************************************************************
// Editor: Aggelos  Date: 02/08/2019
// 1.0.4 Bug in Querry csv 
// *************************************************************************************
// Editor: Aggelos  Date: 02/08/2019
// 1.0.3 Changes in Certifications (1.0.1) 
// *************************************************************************************
// Editor: Aggelos  Date: 29/07/2019
// 1.0.2 
// 1.0.2.1 BP (System Form)
// 1.0.2.1.1 PopUp Report BPCategory when user click TAB in BP's field U_Category
// *************************************************************************************
// Editor: Aggelos  Date: 18/06/2019
// 1.0.1 
// 1.0.1.1 Goods Receipt PO (System Form)
// 1.0.1.1.1 Add a new Buttons in Goods Receipt PO
// 1.0.1.1.2 When the user click Button "Add" in Goods Receipt PO, Add a udo "PATH" and set it in U_path field in lines
// 1.0.1.1.3 When the user click Button "btAtch" in Goods Receipt PO, Update the all appropriate UDOs "PATH" with the appropriate pdf files in UDO lines if does not exist

// 1.0.1.2 Project (System Form)
// 1.0.1.2.1 Add a new Buttons in Project
// 1.0.1.2.2 When the user click Button "btExpt" in Project, Export all appropriate pdf files and create a list of them in a csv file
// =====================================================================================
// 1.0.0 
// 1.0.0.1 Project (System Form)
// 1.0.0.1.1 Sort the matrix's columns in Project Stage
// =====================================================================================
// Editor: Aggelos 
// Date: 13/06/2019
// 1.0.0 version (Major_release_number(significant new features that needs parameterization and changes to run correctly).Minor_release_number(include fixes and new features but nothing ground breaking).Maintenance_release_number (bug fixes, small updates))
// =====================================================================================
// Imports: SAPBusinessOneSDK
// =====================================================================================

using System;
using System.Text;
using System.Collections.Generic;
using SAPbouiCOM.Framework;
using Microsoft.VisualBasic;
using System.Text.RegularExpressions;


namespace LiaromatisSapExtension
{
    class Program
    {
        public static SAPbouiCOM.EventFilters oFilters;
        public static SAPbouiCOM.EventFilter oFilter;
        public static SAPbobsCOM.Company oCompany;
        private const string DevConnString = @"0030002C0030002C00530041005000420044005F00440061007400650076002C0050004C006F006D0056004900490056";
        public static Models.ReportMenuUidModel _ReportMenuUidModel;
        public static string DocEntry = null;
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main(string[] args)
        {
            try
            {
                Application oApp = new Application(GetConnectionString());      //SAP Business One Application to connect to
                oCompany = (SAPbobsCOM.Company)Application.SBO_Application.Company.GetDICompany();

                SetFilters();
                Initialize_Event();

                oApp.Run();
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message);
            }
        }

        static void Initialize_Event()
        {
            Application.SBO_Application.AppEvent += new SAPbouiCOM._IApplicationEvents_AppEventEventHandler(SBO_Application_AppEvent);
            //Application.SBO_Application.MenuEvent += new SAPbouiCOM._IApplicationEvents_MenuEventEventHandler(SBO_Application_MenuEvent);
            Application.SBO_Application.ItemEvent += new SAPbouiCOM._IApplicationEvents_ItemEventEventHandler(SBO_Application_ItemEvent);
            Application.SBO_Application.FormDataEvent += new SAPbouiCOM._IApplicationEvents_FormDataEventEventHandler(SBO_Application_FormDataEvent);
            //Application.SBO_Application.ProgressBarEvent += new SAPbouiCOM._IApplicationEvents_ProgressBarEventEventHandler(SBO_Application_ProgressBarEvent);
            //Application.SBO_Application.StatusBarEvent += new SAPbouiCOM._IApplicationEvents_StatusBarEventEventHandler(SBO_Application_StatusBarEvent);
            //Application.SBO_Application.RightClickEvent += new SAPbouiCOM._IApplicationEvents_RightClickEventEventHandler(SBO_Application_RightClickEvent);
        }

        static void SBO_Application_AppEvent(SAPbouiCOM.BoAppEventTypes EventType)
        {
            switch (EventType)
            {
                case SAPbouiCOM.BoAppEventTypes.aet_ShutDown:
                    System.Environment.Exit(0);
                    //System.Windows.Forms.Application.Exit();
                    break;
                case SAPbouiCOM.BoAppEventTypes.aet_CompanyChanged:
                    System.Environment.Exit(0);
                    //System.Windows.Forms.Application.Exit();
                    break;
                case SAPbouiCOM.BoAppEventTypes.aet_FontChanged:
                    break;
                case SAPbouiCOM.BoAppEventTypes.aet_LanguageChanged:
                    break;
                case SAPbouiCOM.BoAppEventTypes.aet_ServerTerminition:
                    System.Environment.Exit(0);
                    //System.Windows.Forms.Application.Exit();
                    break;
                default:
                    break;
            }
        }

        static void SBO_Application_ItemEvent(string FormUID, ref SAPbouiCOM.ItemEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;

            // 1.0.0.1 Project (System Form)
            if (pVal.FormType == 234000045)
            {
                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK && pVal.Before_Action == true && pVal.ItemUID == "234000062" && pVal.Row == 0 && "fm_FIND_MODE" != Application.SBO_Application.Forms.ActiveForm.Mode.ToString())
                {
                    try
                    {
                        //Classes.Project.Stages.SortColumns(pVal.FormUID, pVal.ItemUID, pVal.ColUID); // 1.0.0.1.1 Sort the matrix's columns in Project Stage                        
                    }
                    catch (Exception) { }
                }
                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_LOAD && pVal.Before_Action == true)
                {
                    try
                    {
                        Classes.Project.NewButton.ImportAttachments(pVal.FormUID); // 1.0.1.2.1 Add a new Buttons in Project
                    }
                    catch (Exception) { }
                    try
                    {
                        Classes.Project.NewButton.Report_ListOutCert(pVal.FormUID); // 1.0.9.1.2 Add a new Buttons in Project
                    }
                    catch (Exception) { }
                }
                if (pVal.ItemUID == "btExpt" && pVal.EventType == SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED && pVal.Before_Action == false)
                {
                    try
                    {
                        Classes.Project.Stages.ExportPdfCSV(oCompany, pVal.FormUID); // 1.0.1.2.2 When the user click Button "btExpt" in Project, Export all appropriate pdf files and create a list of them in a csv file
                    }
                    catch (Exception) { }
                }
                if (pVal.ItemUID == "btRrt1" && pVal.EventType == SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED && pVal.Before_Action == false)
                {
                    try // 1.0.9.1.3 When the user click Button "btRrt1" PopUp Report
                    {

                        EnableUiFunctionality();
                        Classes.Project.Report.ImportAttachments(oCompany, pVal.FormUID, _ReportMenuUidModel, Application.SBO_Application.Company.UserName);
                    }
                    catch (Exception) { }
                }
            }
            // 1.0.1.1 Goods Receipt PO (System Form)
            if (pVal.FormType == 143) //Goods Receipt PO
            {
                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_LOAD && pVal.Before_Action == true)
                {
                    try
                    {
                        Classes.GoodsReceiptPO.NewButton.ImportAttachments(pVal.FormUID); // 1.0.1.1.1 Add a new Buttons in Goods Receipt PO
                    }
                    catch (Exception) { }
                }
                if (pVal.ItemUID == "btAtch" && pVal.EventType == SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED && pVal.Before_Action == false)
                {
                    try
                    {
                        Classes.GoodsReceiptPO.SapB1DataImportUI.ImportPath(oCompany); // 1.0.1.1.3 When the user click Button "btAtch" in Goods Receipt PO, Update the all appropriate UDOs "PATH" with the appropriate pdf files in UDO lines if does not exist
                    }
                    catch (Exception) { }
                }
                if (pVal.ItemUID == "1" && pVal.EventType == SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED && pVal.Before_Action == true)
                {
                    try
                    {
                        Classes.GoodsReceiptPO.SapB1DataImportUI.ImportUDO(FormUID, oCompany); // 1.0.1.1.2 When the user click Button "Add" in Goods Receipt PO, Add a udo "PATH" and set it in U_path field in lines
                    }
                    catch (Exception) { }
                }
            }
            // 1.0.2.1 BP (System Form)
            if (pVal.FormTypeEx == "-134")
            {
                if (pVal.ItemUID == "U_Category" && pVal.EventType == SAPbouiCOM.BoEventTypes.et_KEY_DOWN && pVal.CharPressed == 9 && pVal.Before_Action == false && "fm_FIND_MODE" == Application.SBO_Application.Forms.ActiveForm.Mode.ToString())
                {
                    try
                    {
                        EnableUiFunctionality();
                        // 1.0.2.1.1 PopUp Report BPCategory when user click TAB in BP's field U_Category
                        Application.SBO_Application.ActivateMenuItem(_ReportMenuUidModel.MenuUid_BPCategory);
                    }
                    catch (Exception) { }
                }
            }
            // 1.2.0.1 Delivery (System Form)
            if (pVal.FormType == 140) //Delivery
            {
                if (pVal.ItemUID == "1" && pVal.EventType == SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED)
                {
                    if (pVal.Before_Action == true)
                    {
                        try
                        {
                            Classes.Delivery.SapB1DataImport.AddDraftDelivery(pVal.FormUID, oCompany); // 1.2.1.1.2 When the user click Button "Add" in a Draft Delivery, Update filed stage
                        }
                        catch (Exception) { }                        
                        try
                        {
                            Classes.Generic.SetGLAccount(pVal.FormUID, oCompany); // 1.4.0.1.1 Set GLAccount from FPA in Marketing Documents
                        }
                        catch (Exception)
                        { }
                    }
                    if (pVal.Before_Action == false)
                    {
                        try
                        {
                            if (!String.IsNullOrEmpty(DocEntry))
                            {
                                Application.SBO_Application.OpenForm((SAPbouiCOM.BoFormObjectEnum)112, "", DocEntry); // 1.2.1.1.1 PopUp an existing Draft Delivery.
                                DocEntry = null;
                            }
                        }
                        catch (Exception)
                        {
                            DocEntry = null;
                        }

                    }
                }
            }
            // 1.3.0.1 Item Master Data (System Form)
            if (pVal.FormType == 150) //Item Master Data
            {
                if (pVal.ItemUID == "1" && pVal.EventType == SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED && pVal.Before_Action == true)
                {
                    try
                    {
                        Classes.ItemMasterData.SapB1DataImportUI.ImportUDO(FormUID, oCompany); // 1.3.0.1.1 When an ItemMD is Added: Create two empty UDOs MasterData and linked them to ItemMD
                    }
                    catch (Exception) { }
                }
            }
        }

        static void SBO_Application_FormDataEvent(ref SAPbouiCOM.BusinessObjectInfo pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            // 1.0.7.1 Goods Receipt PO (System Form)
            if (pVal.FormTypeEx == "143") //Goods Receipt PO
            {
                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD && pVal.ActionSuccess == true)
                {
                    try
                    {
                        // 1.0.7.1.1 When a Goods Receipt PO is Canceled all UDO_Path that are linked with it, are Cancelled too.
                        Classes.GoodsReceiptPO.SapB1DataImportUI.CancelUDO(pVal.FormUID, oCompany);
                    }
                    catch (Exception ex) { }
                }
            }

            // 1.1.0.1 Goods Receipt (System Form)
            if (pVal.FormTypeEx == "721") //Goods Receipt
            {
                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD && pVal.ActionSuccess == true)
                {
                    try
                    {
                        // 1.1.0.1.1 When a Goods Receipt is Added a Goods Issue is created
                        Classes.GoodsReceipt.SapB1DataImport.AddGoodsIssue(pVal.FormUID, oCompany);
                    }
                    catch (Exception ex) { }
                }
            }

            // 1.2.0.1 Delivery (System Form)
            if (pVal.FormTypeEx == "140") //Delivery
            {
                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD && pVal.ActionSuccess == true)
                {
                    try
                    {
                        // 1.2.0.1.1 When a Delivery is Added another Delivery with mapping SalesOrder added too. ("да")
                        // 1.2.1.1.3 When the Draft Delivery Added: Update the fields (U_Delivery, U_DrfDeliv) in Source delivery and Update Packages in UDO PCK with Y ("амакысг")
                        // 1.2.2.1.1 When a Delivery is Cancelled: another Delivery with mapping SalesOrder is Cancelled too and Update Packages in UDO PCK with N OR Remove a Draft Delivery ("айда")
                        // 1.2.3.1.1 When a Delivery is Added: Insert Data in Tables DDS_StrumisPck, DDS_StrumisCnt, DDS_StrumisSit
                        DocEntry = Classes.Delivery.SapB1DataImport.AddDelivery(pVal.FormUID, oCompany);
                    }
                    catch (Exception ex) { }
                }
            }
        }

        static void SetFilters()
        {
            oFilters = new SAPbouiCOM.EventFilters();

            //oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_MENU_CLICK);
            //oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_CLICK);
            //oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_RIGHT_CLICK);
            oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK);
            oFilter.Add(234000045);

            oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD);
            oFilter.Add(143);
            oFilter.Add(721);
            oFilter.Add(140);

            oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED);
            oFilter.Add(143);
            oFilter.Add(140);
            oFilter.Add(150);
            oFilter.Add(234000045);

            oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_FORM_LOAD);
            oFilter.Add(143);
            oFilter.Add(234000045);

            oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_KEY_DOWN);
            oFilter.Add(-134);

            Application.SBO_Application.SetFilter(oFilters);
        }

        private static string GetConnectionString()
        {
            var args = Environment.GetCommandLineArgs();

            if (args.Length < 2 || !args[0].Contains("SAP Business One"))
            {
                return DevConnString;
            }
            return args[1];
        }

        static void EnableUiFunctionality() // 1.0.5.2 Changes in (1.0.2.1.1). Set PopUp Report Parametrically.  
        {
            try
            {
                _ReportMenuUidModel = new Models.ReportMenuUidModel();
                string ReportMenuUid = $@" SELECT Code, Name FROM [@DDS_REPORT_UI] ";

                SAPbobsCOM.Recordset rs;
                rs = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                try
                {
                    rs.DoQuery(ReportMenuUid);
                }
                catch (Exception)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(rs);
                }
                rs.MoveFirst();
                while (rs.EoF == false)
                {
                    if (rs.Fields.Item("Name").Value.ToString() == "BPCategory")
                    {
                        _ReportMenuUidModel.MenuUid_BPCategory = rs.Fields.Item("Code").Value.ToString();
                    }

                    if (rs.Fields.Item("Name").Value.ToString() == "ListOfOutstandingCertificates") //1.0.9.1.1
                    {
                        _ReportMenuUidModel.MenuUid_ListOfOutstandingCertificates = rs.Fields.Item("Code").Value.ToString();
                    }
                    rs.MoveNext();
                }
                System.Runtime.InteropServices.Marshal.ReleaseComObject(rs);
            }
            catch (Exception)
            { }
        }

    }
}
