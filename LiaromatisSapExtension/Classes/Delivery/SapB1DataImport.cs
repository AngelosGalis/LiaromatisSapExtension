using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using SAPbouiCOM.Framework;
using System.Text.RegularExpressions;
using System.IO;
using LiaromatisSapExtension.Models;

namespace LiaromatisSapExtension.Classes.Delivery
{
    class SapB1DataImport
    {
        public static string AddDelivery(string _FormUID, SAPbobsCOM.Company _Company)
        {
            int ResCode = 0;
            int ErrCode = 0;
            string ErrMsg = null;
            string DocEntryTarget = null;
            int DocNum = 0;
            string Draft = null;
            SAPbouiCOM.Form oForm = null;
            SAPbouiCOM.Item oItem = null;
            SAPbouiCOM.EditText oEditText = null;
            SAPbouiCOM.ComboBox oComboBox = null;
            SAPbouiCOM.Item oItemMatrix = null;
            SAPbouiCOM.Matrix oMatrix = null;
            string DescSeries;
            //string ValSeries;

            SAPbouiCOM.Form oFormUDF = null;
            string U_Delivery = null;
            string U_DrfDeliv = null;

            try
            {
                oForm = Application.SBO_Application.Forms.Item(_FormUID);
                oForm.Freeze(true);

                oComboBox = (SAPbouiCOM.ComboBox)oForm.Items.Item("88").Specific;
                DescSeries = oComboBox.Selected.Description;
                //ValSeries = oComboBox.Selected.Value;

                //Add a Draft Delivery (mapped with SO), Update the field (U_DrfDeliv) in Source delivery to be Mapped with Draft Delivery
                //Insert Data in Tables DDS_StrumisPck, DDS_StrumisCnt, DDS_StrumisSit
                if (DescSeries == "ΔΑ" || DescSeries == "AP" || DescSeries == "EN")
                {
                    oComboBox = (SAPbouiCOM.ComboBox)oForm.Items.Item("81").Specific;
                    Draft = oComboBox.Selected.Description;
                    if (Draft == $@"Draft" || Draft == $@"Σχέδιο")
                    {
                        string qDraft = $@" SELECT DocNum
                                            FROM ODLN
                                            WHERE DocEntry = (SELECT max(DocEntry) FROM ODLN) ";

                        SAPbobsCOM.Recordset rsDraft;
                        rsDraft = (SAPbobsCOM.Recordset)_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                        rsDraft.DoQuery(qDraft);
                        rsDraft.MoveFirst();

                        DocNum = int.Parse(rsDraft.Fields.Item("DocNum").Value.ToString());
                    }
                    else
                    {
                        oItem = oForm.Items.Item("8");
                        oEditText = (SAPbouiCOM.EditText)oItem.Specific;
                        DocNum = int.Parse(oEditText.Value.ToString());
                    }
                    //Add a Draft Delivery (mapped with SO), Update the field (U_DrfDeliv) in Source delivery to be Mapped with Draft Delivery
                    try
                    {
                        //Delivery & SO source Data 
                        List<DeliverySAPPackageModel> oDeliverySAPPackageModel = new List<DeliverySAPPackageModel>();
                        int DocEntry = 0;
                        string qSourceData = $@" 	SELECT Q1.DocEntryDev, Q1.OpenQty, Q1.Quantity, Q1.U_AltUoM, Q1.U_Quantity, SUM(ISNULL(Q1.NoOfPackages,0)) AS SUMpckQty, Q1.DocEntrySO, Q1.LineNumSO
		                                        FROM (
				                                        SELECT T0.DocEntry AS DocEntryDev,
				                                        T2.OpenQty, T2.Quantity, ISNULL(T2.U_AltUoM,'') AS U_AltUoM, ISNULL(T2.U_Quantity,0) AS U_Quantity,
				                                        T2.U_SAPPackage AS SAPPackage, T2.U_NoOfPackages1 AS NoOfPackages, 
				                                        T2.DocEntry AS DocEntrySO, T2.LineNum AS LineNumSO,
				                                        T0.DocNum, T3.CANCELED, T2.LineStatus
				                                        FROM ODLN T0
				                                        INNER JOIN DLN1 T1 ON T0.DocEntry = T1.DocEntry AND T1.U_PackageType = 'SAP'
				                                        INNER JOIN RDR1 T2 ON T1.U_PckCnt = T2.U_SAPPackage
				                                        INNER JOIN ORDR T3 ON T2.DocEntry = T3.DocEntry AND T0.CardCode = T3.CardCode
														WHERE T2.U_SAPPackage NOT IN (SELECT DISTINCT T1.U_PckCnt 
																				      FROM ODLN T0
																				      INNER JOIN DLN1 T1 ON T0.DocEntry = T1.DocEntry
																				      WHERE T0.CANCELED = 'N'
																				      AND T1.U_PackageType = 'SAP'
                                                                                      AND ISNULL(T1.U_PckCnt,'') != ''
																				      AND T0.DocNum != {DocNum})
				                                        GROUP BY T0.DocEntry, T2.OpenQty, T2.Quantity, ISNULL(T2.U_AltUoM,''), ISNULL(T2.U_Quantity,0), T2.U_SAPPackage, T2.U_NoOfPackages1, T2.DocEntry, T2.LineNum, T0.DocNum, T3.CANCELED, T2.LineStatus
				                                        UNION ALL 
				                                        SELECT T0.DocEntry, 
				                                        T2.OpenQty, T2.Quantity, ISNULL(T2.U_AltUoM,''), ISNULL(T2.U_Quantity,0),
				                                        T2.U_SAPPackage2, T2.U_NoOfPackages2, 
				                                        T2.DocEntry, T2.LineNum,
				                                        T0.DocNum, T3.CANCELED, T2.LineStatus
				                                        FROM ODLN T0
				                                        INNER JOIN DLN1 T1 ON T0.DocEntry = T1.DocEntry AND T1.U_PackageType = 'SAP'
				                                        INNER JOIN RDR1 T2 ON T1.U_PckCnt = T2.U_SAPPackage2
				                                        INNER JOIN ORDR T3 ON T2.DocEntry = T3.DocEntry AND T0.CardCode = T3.CardCode
														WHERE T2.U_SAPPackage2 NOT IN (SELECT DISTINCT T1.U_PckCnt 
																				      FROM ODLN T0
																				      INNER JOIN DLN1 T1 ON T0.DocEntry = T1.DocEntry
																				      WHERE T0.CANCELED = 'N'
																				      AND T1.U_PackageType = 'SAP'
                                                                                      AND ISNULL(T1.U_PckCnt,'') != ''
																				      AND T0.DocNum != {DocNum})
				                                        GROUP BY T0.DocEntry, T2.OpenQty, T2.Quantity, ISNULL(T2.U_AltUoM,''), ISNULL(T2.U_Quantity,0), T2.U_SAPPackage2, T2.U_NoOfPackages2, T2.DocEntry, T2.LineNum, T0.DocNum, T3.CANCELED, T2.LineStatus
				                                        UNION ALL 
				                                        SELECT T0.DocEntry, 
				                                        T2.OpenQty, T2.Quantity, ISNULL(T2.U_AltUoM,''), ISNULL(T2.U_Quantity,0),
				                                        T2.U_SAPPackage3, T2.U_NoOfPackages3, 
				                                        T2.DocEntry, T2.LineNum,
				                                        T0.DocNum, T3.CANCELED, T2.LineStatus
				                                        FROM ODLN T0
				                                        INNER JOIN DLN1 T1 ON T0.DocEntry = T1.DocEntry AND T1.U_PackageType = 'SAP'
				                                        INNER JOIN RDR1 T2 ON T1.U_PckCnt = T2.U_SAPPackage3
				                                        INNER JOIN ORDR T3 ON T2.DocEntry = T3.DocEntry AND T0.CardCode = T3.CardCode
														WHERE T2.U_SAPPackage3 NOT IN (SELECT DISTINCT T1.U_PckCnt 
																				      FROM ODLN T0
																				      INNER JOIN DLN1 T1 ON T0.DocEntry = T1.DocEntry
																				      WHERE T0.CANCELED = 'N'
																				      AND T1.U_PackageType = 'SAP'
                                                                                      AND ISNULL(T1.U_PckCnt,'') != ''
																				      AND T0.DocNum != {DocNum})
				                                        GROUP BY T0.DocEntry, T2.OpenQty, T2.Quantity, ISNULL(T2.U_AltUoM,''), ISNULL(T2.U_Quantity,0), T2.U_SAPPackage3, T2.U_NoOfPackages3, T2.DocEntry, T2.LineNum, T0.DocNum, T3.CANCELED, T2.LineStatus
				                                        UNION ALL 
				                                        SELECT T0.DocEntry,
				                                        T2.OpenQty, T2.Quantity, ISNULL(T2.U_AltUoM,''), ISNULL(T2.U_Quantity,0),
				                                        T2.U_SAPPackage4, T2.U_NoOfPackages4, 
				                                        T2.DocEntry, T2.LineNum,
				                                        T0.DocNum, T3.CANCELED, T2.LineStatus
				                                        FROM ODLN T0
				                                        INNER JOIN DLN1 T1 ON T0.DocEntry = T1.DocEntry AND T1.U_PackageType = 'SAP'
				                                        INNER JOIN RDR1 T2 ON T1.U_PckCnt = T2.U_SAPPackage4
				                                        INNER JOIN ORDR T3 ON T2.DocEntry = T3.DocEntry AND T0.CardCode = T3.CardCode
														WHERE T2.U_SAPPackage4 NOT IN (SELECT DISTINCT T1.U_PckCnt 
																				      FROM ODLN T0
																				      INNER JOIN DLN1 T1 ON T0.DocEntry = T1.DocEntry
																				      WHERE T0.CANCELED = 'N'
																				      AND T1.U_PackageType = 'SAP'
                                                                                      AND ISNULL(T1.U_PckCnt,'') != ''
																				      AND T0.DocNum != {DocNum})
				                                        GROUP BY T0.DocEntry, T2.OpenQty, T2.Quantity, ISNULL(T2.U_AltUoM,''), ISNULL(T2.U_Quantity,0), T2.U_SAPPackage4, T2.U_NoOfPackages4, T2.DocEntry, T2.LineNum, T0.DocNum, T3.CANCELED, T2.LineStatus
				                                        UNION ALL
				                                        SELECT T0.DocEntry,
				                                        T2.OpenQty, T2.Quantity, ISNULL(T2.U_AltUoM,''), ISNULL(T2.U_Quantity,0),
				                                        T2.U_SAPPackage5, T2.U_NoOfPackages5, 
				                                        T2.DocEntry, T2.LineNum,
				                                        T0.DocNum, T3.CANCELED, T2.LineStatus
				                                        FROM ODLN T0
				                                        INNER JOIN DLN1 T1 ON T0.DocEntry = T1.DocEntry AND T1.U_PackageType = 'SAP'
				                                        INNER JOIN RDR1 T2 ON T1.U_PckCnt = T2.U_SAPPackage5
				                                        INNER JOIN ORDR T3 ON T2.DocEntry = T3.DocEntry AND T0.CardCode = T3.CardCode
														WHERE T2.U_SAPPackage5 NOT IN (SELECT DISTINCT T1.U_PckCnt 
																				      FROM ODLN T0
																				      INNER JOIN DLN1 T1 ON T0.DocEntry = T1.DocEntry
																				      WHERE T0.CANCELED = 'N'
																				      AND T1.U_PackageType = 'SAP'
                                                                                      AND ISNULL(T1.U_PckCnt,'') != ''
																				      AND T0.DocNum != {DocNum})
				                                        GROUP BY T0.DocEntry, T2.OpenQty, T2.Quantity, ISNULL(T2.U_AltUoM,''), ISNULL(T2.U_Quantity,0), T2.U_SAPPackage5, T2.U_NoOfPackages5, T2.DocEntry, T2.LineNum, T0.DocNum, T3.CANCELED, T2.LineStatus
			                                        ) Q1
		                                        WHERE Q1.DocNum = {DocNum}
		                                        AND Q1.CANCELED = 'N'
												AND Q1.LineStatus = 'O'
		                                        GROUP BY Q1.DocEntryDev, Q1.OpenQty, Q1.Quantity, Q1.U_AltUoM, Q1.U_Quantity, Q1.DocEntrySO, Q1.LineNumSO
		                                        ORDER BY Q1.DocEntrySO, Q1.LineNumSO  ";

                        SAPbobsCOM.Recordset rsSourceData;
                        rsSourceData = (SAPbobsCOM.Recordset)_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                        rsSourceData.DoQuery(qSourceData);

                        rsSourceData.MoveFirst();
                        while (rsSourceData.EoF == false)
                        {
                            oDeliverySAPPackageModel.Add(new Models.DeliverySAPPackageModel()
                            {
                                DocEntryDev = int.Parse(rsSourceData.Fields.Item("DocEntryDev").Value.ToString()),
                                OpenQty = double.Parse(rsSourceData.Fields.Item("OpenQty").Value.ToString()),
                                Quantity = double.Parse(rsSourceData.Fields.Item("Quantity").Value.ToString()),
                                U_AltUoM = rsSourceData.Fields.Item("U_AltUoM").Value.ToString(),
                                U_Quantity = double.Parse(rsSourceData.Fields.Item("U_Quantity").Value.ToString()),
                                SUMpckQty = double.Parse(rsSourceData.Fields.Item("SUMpckQty").Value.ToString()),
                                DocEntrySO = int.Parse(rsSourceData.Fields.Item("DocEntrySO").Value.ToString()),
                                LineNumSO = int.Parse(rsSourceData.Fields.Item("LineNumSO").Value.ToString())
                            });

                            DocEntry = int.Parse(rsSourceData.Fields.Item("DocEntryDev").Value.ToString());
                            rsSourceData.MoveNext();
                        }
                        rsSourceData = null;

                        //Delivery target Series
                        int Series = 0;
                        string qSeries = $@" SELECT Series FROM NNM1
                                     WHERE ObjectCode = '15'
                                     AND SeriesName = N'ΑΝΑΛΩΣΗ'  ";

                        SAPbobsCOM.Recordset rsSeries;
                        rsSeries = (SAPbobsCOM.Recordset)_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                        rsSeries.DoQuery(qSeries);
                        Series = int.Parse(rsSeries.Fields.Item("Series").Value.ToString());
                        rsSeries = null;

                        //Update Delivery Source
                        SAPbobsCOM.Documents oDeliverySource;
                        oDeliverySource = (SAPbobsCOM.Documents)_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oDeliveryNotes);
                        if (oDeliverySource.GetByKey(DocEntry) == true)
                        {
                            //Add Draft Delivery Target
                            SAPbobsCOM.Documents oDeliveryTarget;
                            oDeliveryTarget = (SAPbobsCOM.Documents)_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oDrafts);
                            oDeliveryTarget.DocObjectCode = SAPbobsCOM.BoObjectTypes.oDeliveryNotes;


                            //Delivery Target Header
                            oDeliveryTarget.Series = Series;
                            oDeliveryTarget.CardCode = oDeliverySource.CardCode;
                            oDeliveryTarget.DocDate = oDeliverySource.DocDate;
                            oDeliveryTarget.TaxDate = oDeliverySource.TaxDate;
                            oDeliveryTarget.DocDueDate = oDeliverySource.DocDueDate;
                            oDeliveryTarget.UserFields.Fields.Item("U_Delivery").Value = DocEntry;

                            foreach (var i in oDeliverySAPPackageModel.Select(x => new { x.DocEntrySO, x.LineNumSO, x.OpenQty, x.Quantity, x.U_AltUoM, x.U_Quantity, x.SUMpckQty }))
                            {
                                //Get Stage
                                string Stage = null;
                                string qStage = $@" SELECT DISTINCT T1.DocEntry , T1.LineNum, T1.Project, T4.UniqueID FROM ORDR T0
                                                INNER JOIN RDR1 T1 ON T0.DocEntry = T1.DocEntry
                                                INNER JOIN OPMG T3 ON T3.FIPROJECT = T1.Project
                                                INNER JOIN PMG1 T4 ON T4.AbsEntry = T3.AbsEntry
                                                INNER JOIN PMG4 T5 ON T5.AbsEntry = T4.AbsEntry AND T4.LineID = T5.StageID AND T1.DocEntry = T5.DocEntry AND T1.LineNum = T5.LineNum
                                                WHERE T1.DocEntry = {i.DocEntrySO}
                                                AND T1.LineNum = {i.LineNumSO}
                                                AND T5.TYP = 17  ";

                                SAPbobsCOM.Recordset rsStage;
                                rsStage = (SAPbobsCOM.Recordset)_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                                rsStage.DoQuery(qStage);
                                Stage = rsStage.Fields.Item("UniqueID").Value.ToString();
                                rsStage = null;

                                //Delivery Target  Lines
                                oDeliveryTarget.Lines.UserFields.Fields.Item("U_UniqueID").Value = Stage;
                                oDeliveryTarget.Lines.UserFields.Fields.Item("U_Quantity").Value = i.SUMpckQty;
                                if (String.IsNullOrEmpty(i.U_AltUoM))
                                    oDeliveryTarget.Lines.Quantity = i.SUMpckQty;
                                else
                                {
                                    double Quantity = (i.Quantity / i.U_Quantity) * i.SUMpckQty;
                                    if (i.OpenQty > Quantity + 0.00001)
                                        oDeliveryTarget.Lines.Quantity = Quantity;
                                    else
                                        oDeliveryTarget.Lines.Quantity = i.OpenQty;
                                }
                                oDeliveryTarget.Lines.BaseType = 17;
                                oDeliveryTarget.Lines.BaseEntry = i.DocEntrySO;
                                oDeliveryTarget.Lines.BaseLine = i.LineNumSO;
                                oDeliveryTarget.Lines.Add();
                            }

                            ResCode = oDeliveryTarget.Add();
                            if (ResCode != 0)
                            {
                                _Company.GetLastError(out ErrCode, out ErrMsg);
                                Application.SBO_Application.StatusBar.SetText($@"-0001 Error {ErrCode} {ErrMsg}", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                            }
                            else
                            {
                                DocEntryTarget = _Company.GetNewObjectKey();
                                oDeliverySource.UserFields.Fields.Item("U_DrfDeliv").Value = DocEntryTarget;
                                oDeliverySource.UserFields.Fields.Item("U_Delivery").SetNullValue();
                                ResCode = oDeliverySource.Update();
                                if (ResCode != 0)
                                {
                                    _Company.GetLastError(out ErrCode, out ErrMsg);
                                    Application.SBO_Application.StatusBar.SetText($@"-0002 Error {ErrCode} {ErrMsg}, DocEntryTarget:{DocEntryTarget}, DocEntrySource:{DocEntry}", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                                }
                                else
                                {
                                    Application.SBO_Application.StatusBar.SetText($@"Success, DocEntryTarget:{DocEntryTarget}, DocEntrySource:{DocEntry}", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
                                }
                            }
                        }
                    }
                    catch (Exception ex) { }
                                                                                                                       
                    //Insert Data in Table DDS_StrumisPck
                    try
                    {
                        int DocEntry = 0;
                        int LineNum = 0;
                        string PckCnt = null;
                        string qPckName = $@"   SELECT T1.DocEntry, T1.LineNum, T1.U_PckCnt FROM ODLN T0
                                                INNER JOIN DLN1 T1 ON T0.DocEntry = T1.DocEntry
                                                WHERE T0.DocNum = {DocNum}
                                                AND ISNULL(T1.U_PckCnt,'') != ''
                                                AND T1.U_PackageType = 'StrumisPck'";

                        SAPbobsCOM.Recordset rsPckName;
                        rsPckName = (SAPbobsCOM.Recordset)_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                        rsPckName.DoQuery(qPckName);
                        rsPckName.MoveFirst();
                        while (rsPckName.EoF == false)
                        {
                            DocEntry= int.Parse(rsPckName.Fields.Item("DocEntry").Value.ToString());
                            LineNum = int.Parse(rsPckName.Fields.Item("LineNum").Value.ToString());
                            PckCnt = rsPckName.Fields.Item("U_PckCnt").Value.ToString();
                            try
                            {
                                string qInsertPck = $@" INSERT INTO DDS_StrumisPck
                                                               ([ContractMark],[Description],[MainMember],[PaintFinish],[Length],[Width],[UnitWeight]
                                                               ,[Package],[Drawing],[SITEDELIVERYNOTE],[Contract],[Qty],[DocEntry],[LineNum],[Exception])
                                                        (SELECT *,'{DocEntry}','{LineNum}','' FROM DDS_F_StrumisPck (N'{PckCnt}')) ";
                                SAPbobsCOM.Recordset rsInsertPck;
                                rsInsertPck = (SAPbobsCOM.Recordset)_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                                rsInsertPck.DoQuery(qInsertPck);
                                rsInsertPck = null;
                            }
                            catch (Exception ex)
                            {
                                string qInsertPck = $@" INSERT INTO DDS_StrumisPck
                                                            ([Package], [DocEntry], [LineNum], [Exception])
                                                        VALUES
                                                            (N'{PckCnt}', '{DocEntry}', '{LineNum}', '{ex.Message.Replace("'", "")}') ";
                                SAPbobsCOM.Recordset rsInsertPck;
                                rsInsertPck = (SAPbobsCOM.Recordset)_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                                rsInsertPck.DoQuery(qInsertPck);
                                rsInsertPck = null;
                            }
                            rsPckName.MoveNext();
                        }
                        rsPckName = null;
                    }
                    catch (Exception)
                    { }

                    //Insert Data in Table DDS_StrumisCnt
                    try
                    {
                        int DocEntry = 0;
                        int LineNum = 0;
                        string PckCnt = null;
                        string qCntName = $@"   SELECT T1.DocEntry, T1.LineNum, T1.U_PckCnt FROM ODLN T0
                                                INNER JOIN DLN1 T1 ON T0.DocEntry = T1.DocEntry
                                                WHERE T0.DocNum = {DocNum}
                                                AND ISNULL(T1.U_PckCnt,'') != ''
                                                AND T1.U_PackageType = 'StrumisCnt'";

                        SAPbobsCOM.Recordset rsCntName;
                        rsCntName = (SAPbobsCOM.Recordset)_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                        rsCntName.DoQuery(qCntName);
                        rsCntName.MoveFirst();
                        while (rsCntName.EoF == false)
                        {
                            DocEntry = int.Parse(rsCntName.Fields.Item("DocEntry").Value.ToString());
                            LineNum = int.Parse(rsCntName.Fields.Item("LineNum").Value.ToString());
                            PckCnt = rsCntName.Fields.Item("U_PckCnt").Value.ToString();
                            try
                            {
                                string qInsertCnt = $@" INSERT INTO DDS_StrumisCnt
                                                               ([query],[ContainerName],[ContractMark],[Description],[MainMember],[PaintFinish],[Length],[UnitWeight],[UnitArea]
                                                               ,[Package],[SITEDELIVERYNOTE],[Drawing],[Contract],[Qty],[DocEntry],[LineNum],[Exception])
                                                        (SELECT *,'{DocEntry}','{LineNum}','' FROM DDS_F_StrumisCnt (N'{PckCnt}')) ";
                                SAPbobsCOM.Recordset rsInsertCnt;
                                rsInsertCnt = (SAPbobsCOM.Recordset)_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                                rsInsertCnt.DoQuery(qInsertCnt);
                                rsInsertCnt = null;
                            }
                            catch (Exception ex)
                            {
                                string qInsertCnt = $@" INSERT INTO DDS_StrumisCnt
                                                            ([ContainerName], [DocEntry], [LineNum], [Exception])
                                                        VALUES
                                                            (N'{PckCnt}', '{DocEntry}', '{LineNum}', '{ex.Message.Replace("'", "")}') ";
                                SAPbobsCOM.Recordset rsInsertCnt;
                                rsInsertCnt = (SAPbobsCOM.Recordset)_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                                rsInsertCnt.DoQuery(qInsertCnt);
                                rsInsertCnt = null;
                            }
                            rsCntName.MoveNext();
                        }
                        rsCntName = null;
                    }
                    catch (Exception)
                    { }

                    //Insert Data in Table DDS_StrumisSit
                    try
                    {
                        int DocEntry = 0;
                        int LineNum = 0;
                        string PckCnt = null;
                        string qSitName = $@"   SELECT T1.DocEntry, T1.LineNum, T1.U_PckCnt FROM ODLN T0
                                                INNER JOIN DLN1 T1 ON T0.DocEntry = T1.DocEntry
                                                WHERE T0.DocNum = {DocNum}
                                                AND ISNULL(T1.U_PckCnt,'') != ''
                                                AND T1.U_PackageType = 'StrumisSit'";

                        SAPbobsCOM.Recordset rsSitName;
                        rsSitName = (SAPbobsCOM.Recordset)_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                        rsSitName.DoQuery(qSitName);
                        rsSitName.MoveFirst();
                        while (rsSitName.EoF == false)
                        {
                            DocEntry = int.Parse(rsSitName.Fields.Item("DocEntry").Value.ToString());
                            LineNum = int.Parse(rsSitName.Fields.Item("LineNum").Value.ToString());
                            PckCnt = rsSitName.Fields.Item("U_PckCnt").Value.ToString();
                            try
                            {
                                string qInsertSit = $@" INSERT INTO DDS_StrumisSit
                                                                ([ContractMark],[Description],[MainMember],[PaintFinish],[Length],[Width],[UnitWeight]
                                                               ,[Package],[Drawing],[SITEDELIVERYNOTE],[Contract],[Qty],[DocEntry],[LineNum],[Exception])
                                                        (SELECT *,'{DocEntry}','{LineNum}','' FROM DDS_F_StrumisSit (N'{PckCnt}')) ";
                                SAPbobsCOM.Recordset rsInsertSit;
                                rsInsertSit = (SAPbobsCOM.Recordset)_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                                rsInsertSit.DoQuery(qInsertSit);
                                rsInsertSit = null;
                            }
                            catch (Exception ex)
                            {
                                string qInsertSit = $@" INSERT INTO DDS_StrumisSit
                                                            ([SITEDELIVERYNOTE], [DocEntry], [LineNum], [Exception])
                                                        VALUES
                                                            (N'{PckCnt}', '{DocEntry}', '{LineNum}', '{ex.Message.Replace("'", "")}') ";
                                SAPbobsCOM.Recordset rsInsertSit;
                                rsInsertSit = (SAPbobsCOM.Recordset)_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                                rsInsertSit.DoQuery(qInsertSit);
                                rsInsertSit = null;
                            }
                            rsSitName.MoveNext();
                        }
                        rsSitName = null;
                    }
                    catch (Exception)
                    { }
                }
                //when Draft Delivery is becoming Delivery Update the fields (U_Delivery, U_DrfDeliv) in Source delivery
                //Update Packages in UDO PCK
                else if (DescSeries == "ΑΝΑΛΩΣΗ")
                {
                    oFormUDF = Application.SBO_Application.Forms.Item(oForm.UDFFormUID); //Get UDF Form
                    oItem = oFormUDF.Items.Item("U_Delivery");
                    oEditText = (SAPbouiCOM.EditText)oItem.Specific;
                    U_Delivery = oEditText.Value.ToString();

                    oItem = oFormUDF.Items.Item("U_DrfDeliv");
                    oEditText = (SAPbouiCOM.EditText)oItem.Specific;
                    U_DrfDeliv = oEditText.Value.ToString();

                    oItem = oForm.Items.Item("8");
                    oEditText = (SAPbouiCOM.EditText)oItem.Specific;
                    DocNum = int.Parse(oEditText.Value.ToString());

                    if (String.IsNullOrEmpty(U_DrfDeliv) && !String.IsNullOrEmpty(U_Delivery))
                    {
                        try
                        {
                            SAPbobsCOM.Documents oDelivery;
                            oDelivery = (SAPbobsCOM.Documents)_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oDeliveryNotes);
                            if (oDelivery.GetByKey(int.Parse(U_Delivery)) == true)
                            {
                                int DocEntry = 0;
                                string qDocEntry = $@"  SELECT TOP 1 DocEntry FROM ODLN
                                                    WHERE U_Delivery = {U_Delivery}
                                                    AND ISNULL(U_DrfDeliv,'0') = '0'
                                                    AND CANCELED = 'N'
                                                    ORDER BY DocEntry DESC";

                                SAPbobsCOM.Recordset rsDocEntry;
                                rsDocEntry = (SAPbobsCOM.Recordset)_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                                rsDocEntry.DoQuery(qDocEntry);
                                rsDocEntry.MoveFirst();
                                if (rsDocEntry.RecordCount == 1)
                                {
                                    DocEntry = int.Parse(rsDocEntry.Fields.Item("DocEntry").Value.ToString());
                                    oDelivery.UserFields.Fields.Item("U_Delivery").Value = DocEntry;
                                    oDelivery.UserFields.Fields.Item("U_DrfDeliv").SetNullValue();                                    
                                    ResCode = oDelivery.Update();
                                    if (ResCode != 0)
                                    {
                                        _Company.GetLastError(out ErrCode, out ErrMsg);
                                        Application.SBO_Application.StatusBar.SetText($@"-0003 Error {ErrCode} {ErrMsg}, DocEntryTarget:{DocEntry}, DocEntrySource:{U_Delivery}", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                                    }
                                    else
                                    {
                                        Application.SBO_Application.StatusBar.SetText($@"Success, DocEntryTarget:{DocEntry}, DocEntrySource:{U_Delivery}", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
                                    }
                                }
                                rsDocEntry = null;
                            }
                        }
                        catch (Exception)
                        { }

                        //Update UDO PCK
                        try
                        {
                            string qPckName = $@"  SELECT U_PckCnt FROM DLN1
                                                   WHERE DocEntry = {U_Delivery}
                                                   AND ISNULL(U_PckCnt,'') != ''
                                                   AND U_PackageType = 'SAP'";

                            SAPbobsCOM.Recordset rsPckName;
                            rsPckName = (SAPbobsCOM.Recordset)_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                            rsPckName.DoQuery(qPckName);
                            rsPckName.MoveFirst();
                            while (rsPckName.EoF == false)
                            {
                                SAPbobsCOM.GeneralService oGeneralService = null;
                                SAPbobsCOM.GeneralData oGeneralData = null;
                                SAPbobsCOM.GeneralDataParams oGeneralParams = null;
                                SAPbobsCOM.CompanyService sCmp = null;
                                sCmp = (SAPbobsCOM.CompanyService)_Company.GetCompanyService();

                                oGeneralService = sCmp.GetGeneralService("PCK");
                                oGeneralParams = (SAPbobsCOM.GeneralDataParams)oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams);

                                oGeneralParams.SetProperty("Code", rsPckName.Fields.Item("U_PckCnt").Value.ToString());
                                oGeneralData = oGeneralService.GetByParams(oGeneralParams);
                                oGeneralData.SetProperty("U_PckUse", "Y");
                                oGeneralService.Update(oGeneralData);

                                rsPckName.MoveNext();
                            }
                            rsPckName = null;
                        }
                        catch (Exception)
                        { }
                    }
                }
                // When a Delivery is Cancelled: another Delivery with mapping SalesOrder is Cancelled too 
                // Update Packages in UDO PCK with N 
                // OR Remove only a Draft Delivery if Delivery does not exist yet
                else if (DescSeries == "ΑΚΔΑ")
                {
                    Boolean SuccessCancel = false;
                    oFormUDF = Application.SBO_Application.Forms.Item(oForm.UDFFormUID); //Get UDF Form
                    oItem = oFormUDF.Items.Item("U_Delivery");
                    oEditText = (SAPbouiCOM.EditText)oItem.Specific;
                    U_Delivery = oEditText.Value.ToString();

                    oItem = oFormUDF.Items.Item("U_DrfDeliv");
                    oEditText = (SAPbouiCOM.EditText)oItem.Specific;
                    U_DrfDeliv = oEditText.Value.ToString();

                    //Cancel Source Delivery 
                    //Update UDO PCK
                    if (String.IsNullOrEmpty(U_DrfDeliv) && !String.IsNullOrEmpty(U_Delivery))
                    {
                        try
                        {
                            oItemMatrix = oForm.Items.Item("38");
                            oMatrix = (SAPbouiCOM.Matrix)oItemMatrix.Specific;
                            string PackageType = null;
                            string PckCnt = null;
                            for (int i = 1; i < oMatrix.VisualRowCount; i++)
                            {
                                try
                                {
                                    oComboBox = (SAPbouiCOM.ComboBox)oMatrix.Columns.Item("U_PackageType").Cells.Item(i).Specific;
                                    PackageType = oComboBox.Selected.Value;
                                }
                                catch (Exception)
                                {
                                    PackageType = null;
                                }

                                oEditText = (SAPbouiCOM.EditText)oMatrix.Columns.Item("U_PckCnt").Cells.Item(i).Specific;
                                PckCnt = oEditText.Value.ToString();
                                if (PackageType == "SAP" && !String.IsNullOrEmpty(PckCnt))
                                {
                                    if (SuccessCancel == false)
                                    {
                                        SAPbobsCOM.Documents oDelivery;
                                        oDelivery = (SAPbobsCOM.Documents)_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oDeliveryNotes);
                                        if (oDelivery.GetByKey(int.Parse(U_Delivery)) == true)
                                        {
                                            SAPbobsCOM.Documents oCancelDelivery = oDelivery.CreateCancellationDocument();

                                            //Delivery target Series
                                            int Series = 0;
                                            string qSeries = $@" SELECT Series FROM NNM1
                                                                 WHERE ObjectCode = '15'
                                                                 AND SeriesName = N'ΑΚΑΝΑΛ'  ";

                                            SAPbobsCOM.Recordset rsSeries;
                                            rsSeries = (SAPbobsCOM.Recordset)_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                                            rsSeries.DoQuery(qSeries);
                                            Series = int.Parse(rsSeries.Fields.Item("Series").Value.ToString());
                                            rsSeries = null;

                                            oCancelDelivery.Series = Series;
                                            ResCode = oCancelDelivery.Add();
                                            if (ResCode != 0)
                                            {
                                                //den mporei na akirothei to delivery
                                                _Company.GetLastError(out ErrCode, out ErrMsg);
                                                break;
                                            }
                                            else
                                            {
                                                SuccessCancel = true;
                                            }
                                        }
                                    }
                                    if (SuccessCancel == true) //Update UDO PCK
                                    {
                                        try
                                        {
                                            SAPbobsCOM.GeneralService oGeneralService = null;
                                            SAPbobsCOM.GeneralData oGeneralData = null;
                                            SAPbobsCOM.GeneralDataParams oGeneralParams = null;
                                            SAPbobsCOM.CompanyService sCmp = null;
                                            sCmp = (SAPbobsCOM.CompanyService)_Company.GetCompanyService();

                                            oGeneralService = sCmp.GetGeneralService("PCK");
                                            oGeneralParams = (SAPbobsCOM.GeneralDataParams)oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams);

                                            oGeneralParams.SetProperty("Code", PckCnt);
                                            oGeneralData = oGeneralService.GetByParams(oGeneralParams);
                                            oGeneralData.SetProperty("U_PckUse", "N");
                                            oGeneralService.Update(oGeneralData);
                                        }
                                        catch (Exception)
                                        { }
                                    }
                                }
                            }
                        }
                        catch (Exception)
                        { }
                    }
                    //Remove Draft Delivery if Delivery does not exist yet
                    else if (!String.IsNullOrEmpty(U_DrfDeliv) && String.IsNullOrEmpty(U_Delivery))
                    {
                        try
                        {
                            //Remove Draft Delivery
                            SAPbobsCOM.Documents oRemoveDraftDelivery;
                            oRemoveDraftDelivery = (SAPbobsCOM.Documents)_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oDrafts);
                            oRemoveDraftDelivery.DocObjectCode = SAPbobsCOM.BoObjectTypes.oDeliveryNotes;
                            if (oRemoveDraftDelivery.GetByKey(int.Parse(U_DrfDeliv)) == true)
                            {
                                ResCode = oRemoveDraftDelivery.Remove();
                                if (ResCode != 0)
                                {
                                    //den mporei na akirothei to delivery
                                    _Company.GetLastError(out ErrCode, out ErrMsg);
                                }
                            }
                        }
                        catch (Exception ex) { }
                    }   
                    
                }
                oForm.Freeze(false);
            }
            catch (Exception ex)
            {
                oForm.Freeze(false);
            }
            return DocEntryTarget;
        }
        //when the user click Add button in Draft Delivery the field (Stage) be informed
        public static void AddDraftDelivery(string _FormUID, SAPbobsCOM.Company _Company)
        {
            SAPbouiCOM.Form oForm = null;
            SAPbouiCOM.Form oFormUDF = null;
            SAPbouiCOM.Item oItem = null;
            SAPbouiCOM.EditText oEditText = null;
            SAPbouiCOM.EditText oEditTextTarget = null;
            SAPbouiCOM.ComboBox oComboBox = null;
            SAPbouiCOM.Item oItemMatrix = null;
            SAPbouiCOM.Matrix oMatrix = null;
            string DescSeries;
            //string ValSeries;
            string Draft_U_Delivery = null;
            string Draft_U_DrfDeliv = null;
            //string Draft_DocNum = null;
            //string Delivery_U_Delivery = null;
            //string Delivery_U_DrfDeliv = null;

            try
            {
                oForm = Application.SBO_Application.Forms.Item(_FormUID);
                oForm.Freeze(true);
                oFormUDF = Application.SBO_Application.Forms.Item(oForm.UDFFormUID); //Get UDF Form

                oComboBox = (SAPbouiCOM.ComboBox)oForm.Items.Item("88").Specific;
                DescSeries = oComboBox.Selected.Description;
                ////ValSeries = oComboBox.Selected.Value;

                if (DescSeries == "ΑΝΑΛΩΣΗ" && oForm.Mode.ToString() == "fm_ADD_MODE")
                {
                    oItem = oFormUDF.Items.Item("U_Delivery");
                    oEditText = (SAPbouiCOM.EditText)oItem.Specific;
                    Draft_U_Delivery = oEditText.Value.ToString();

                    oItem = oFormUDF.Items.Item("U_DrfDeliv");
                    oEditText = (SAPbouiCOM.EditText)oItem.Specific;
                    Draft_U_DrfDeliv = oEditText.Value.ToString();

                    //oItem = oForm.Items.Item("8");
                    //oEditText = (SAPbouiCOM.EditText)oItem.Specific;
                    //Draft_DocNum = oEditText.Value.ToString();

                    //string qDraftExist = $@" SELECT T0.U_Delivery, T0.U_DrfDeliv FROM ODLN T0
                    //                         INNER JOIN ODRF T1 ON T1.DocEntry = T0.U_DrfDeliv AND T0.DocEntry = T1.U_Delivery
                    //                         WHERE T0.DocEntry = '{Draft_U_Delivery}'
                    //                         AND T1.DocNum = {Draft_DocNum}
                    //                         AND T0.CANCELED = 'N'";

                    //SAPbobsCOM.Recordset rsDraftExist;
                    //rsDraftExist = (SAPbobsCOM.Recordset)_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                    //rsDraftExist.DoQuery(qDraftExist);
                    //rsDraftExist.MoveFirst();
                    //if (rsDraftExist.RecordCount == 1)
                    //{
                    //    Delivery_U_Delivery = (rsDraftExist.Fields.Item("U_Delivery").Value.ToString() == "0") ? "" : rsDraftExist.Fields.Item("U_Delivery").Value.ToString();
                    //    Delivery_U_DrfDeliv = rsDraftExist.Fields.Item("U_DrfDeliv").Value.ToString();
                    //}
                    //rsDraftExist = null;

                    //if (String.IsNullOrEmpty(Draft_U_DrfDeliv) && String.IsNullOrEmpty(Delivery_U_Delivery) && !String.IsNullOrEmpty(Draft_U_Delivery) && !String.IsNullOrEmpty(Delivery_U_DrfDeliv))
                    if (String.IsNullOrEmpty(Draft_U_DrfDeliv) && !String.IsNullOrEmpty(Draft_U_Delivery))
                    {
                        oItemMatrix = oForm.Items.Item("38");
                        oMatrix = (SAPbouiCOM.Matrix)oItemMatrix.Specific;

                        for (int i = 1; i < oMatrix.VisualRowCount; i++)
                        {
                            oEditText = (SAPbouiCOM.EditText)oMatrix.Columns.Item("U_UniqueID").Cells.Item(i).Specific;

                            oEditTextTarget = (SAPbouiCOM.EditText)oMatrix.Columns.Item("254000386").Cells.Item(i).Specific;
                            try
                            {
                                oEditTextTarget.Value = oEditText.Value.ToString();
                            }
                            catch (Exception) { }
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
    }
}
