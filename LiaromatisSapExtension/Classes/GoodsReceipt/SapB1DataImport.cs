using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using SAPbouiCOM.Framework;
using System.Text.RegularExpressions;
using System.IO;

namespace LiaromatisSapExtension.Classes.GoodsReceipt
{
    class SapB1DataImport
    {
        public static void AddGoodsIssue(string _FormUID, SAPbobsCOM.Company _Company)
        {
            int DocNum = 0;
            SAPbouiCOM.Form oForm = null;
            SAPbouiCOM.Item oItem = null;
            SAPbouiCOM.EditText oEditText = null;
            SAPbouiCOM.ComboBox oComboBox = null;
            SAPbouiCOM.Item oItemMatrix = null;
            string DescSeries;
            //string ValSeries;

            try
            {
                oForm = Application.SBO_Application.Forms.Item(_FormUID);
                oForm.Freeze(true);

                oComboBox = (SAPbouiCOM.ComboBox)oForm.Items.Item("30").Specific;
                DescSeries = oComboBox.Selected.Description;
                //ValSeries = oComboBox.Selected.Value;

                if (DescSeries == "H" || DescSeries == "Η")
                {
                    oItem = oForm.Items.Item("7");
                    oEditText = (SAPbouiCOM.EditText)oItem.Specific;
                    DocNum = int.Parse(oEditText.Value.ToString());

                    //Goods Receipt DocEntry
                    int DocEntry = 0;
                    string qGRDocEntry = $@" SELECT DocEntry FROM OIGN
                                         WHERE DocNum = {DocNum}  ";

                    SAPbobsCOM.Recordset rsGRDocEntry;
                    rsGRDocEntry = (SAPbobsCOM.Recordset)_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                    rsGRDocEntry.DoQuery(qGRDocEntry);
                    DocEntry = int.Parse(rsGRDocEntry.Fields.Item("DocEntry").Value.ToString());
                    rsGRDocEntry = null;

                    //Goods Receipt Series
                    int Series = 0;
                    string qSeries = $@" SELECT Series FROM NNM1
                                     WHERE ObjectCode = '60'
                                     AND SeriesName = N'ΠΑΡΑΓΩΓΗ'  ";

                    SAPbobsCOM.Recordset rsSeries;
                    rsSeries = (SAPbobsCOM.Recordset)_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                    rsSeries.DoQuery(qSeries);
                    Series = int.Parse(rsSeries.Fields.Item("Series").Value.ToString());
                    rsSeries = null;

                    //Add Goods Issue
                    SAPbobsCOM.Documents oGoodsReceipt;
                    oGoodsReceipt = (SAPbobsCOM.Documents)_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInventoryGenEntry);
                    if (oGoodsReceipt.GetByKey(DocEntry) == true)
                    {
                        SAPbobsCOM.Documents oGoodsIssue;
                        oGoodsIssue = (SAPbobsCOM.Documents)_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInventoryGenExit);

                        //Header
                        oGoodsIssue.Series = Series;
                        oGoodsIssue.DocDate = oGoodsReceipt.DocDate;
                        oGoodsIssue.TaxDate = oGoodsReceipt.TaxDate;
                        oGoodsIssue.Reference2 = oGoodsReceipt.Reference2;
                        oGoodsIssue.UserFields.Fields.Item("U_GRDocEntry").Value = oGoodsReceipt.DocEntry;
                        oGoodsIssue.Comments = $@"Subcontracting Strumis";

                        //Lines
                        oItemMatrix = oForm.Items.Item("13");
                        SAPbouiCOM.Matrix oMatrix = ((SAPbouiCOM.Matrix)(oItemMatrix.Specific));
                        for (int i = 0; i < oMatrix.VisualRowCount - 1; i++)
                        {
                            //prepei na tsekarw kai sta item an einai hmietimo to idos sto group. an den einai den to kataxoro to sigkekrimeno
                            oGoodsReceipt.Lines.SetCurrentLine(i);

                            //If Item is "Ημιέτοιμα Προϊόντα" add to issue
                            string qExistItmCd = $@" SELECT T0.ItemCode, T1.ItmsGrpCod, T1.ItmsGrpNam FROM OITM T0
                                             INNER JOIN OITB T1 ON T0.ItmsGrpCod = T1.ItmsGrpCod
                                             WHERE T1.ItmsGrpNam = N'Ημιέτοιμα Προϊόντα'
                                             AND T0.ItemCode = N'{oGoodsReceipt.Lines.ItemCode}'  ";

                            SAPbobsCOM.Recordset rsExistItmCd;
                            rsExistItmCd = (SAPbobsCOM.Recordset)_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                            rsExistItmCd.DoQuery(qExistItmCd);
                            if (rsExistItmCd.RecordCount > 0)
                            {
                                oGoodsIssue.Lines.ItemCode = oGoodsReceipt.Lines.ItemCode;
                                oGoodsIssue.Lines.Quantity = oGoodsReceipt.Lines.Quantity;
                                oGoodsIssue.Lines.UnitPrice = oGoodsReceipt.Lines.UnitPrice;
                                oGoodsIssue.Lines.WarehouseCode = oGoodsReceipt.Lines.WarehouseCode;
                                oGoodsIssue.Lines.UoMEntry = oGoodsReceipt.Lines.UoMEntry;
                                oGoodsIssue.Lines.ProjectCode = oGoodsReceipt.Lines.ProjectCode;

                                oGoodsIssue.Lines.UserFields.Fields.Item("U_UniqueID").Value = oGoodsReceipt.Lines.UserFields.Fields.Item("U_UniqueID").Value;
                                oGoodsIssue.Lines.UserFields.Fields.Item("U_Quantity").Value = oGoodsReceipt.Lines.UserFields.Fields.Item("U_Quantity").Value;
                                oGoodsIssue.Lines.UserFields.Fields.Item("U_Comments").Value = oGoodsReceipt.Lines.UserFields.Fields.Item("U_Comments").Value;
                                oGoodsIssue.Lines.UserFields.Fields.Item("U_DrawingNo").Value = oGoodsReceipt.Lines.UserFields.Fields.Item("U_DrawingNo").Value;
                                oGoodsIssue.Lines.UserFields.Fields.Item("U_Area").Value = oGoodsReceipt.Lines.UserFields.Fields.Item("U_Area").Value;
                                oGoodsIssue.Lines.UserFields.Fields.Item("U_WeightperQty").Value = oGoodsReceipt.Lines.UserFields.Fields.Item("U_WeightperQty").Value;
                                oGoodsIssue.Lines.UserFields.Fields.Item("U_Length").Value = oGoodsReceipt.Lines.UserFields.Fields.Item("U_Length").Value;
                                oGoodsIssue.Lines.UserFields.Fields.Item("U_Quantity").Value = oGoodsReceipt.Lines.UserFields.Fields.Item("U_Quantity").Value;
                                oGoodsIssue.Lines.UserFields.Fields.Item("U_GRLineNum").Value = oGoodsReceipt.Lines.LineNum;
                                oGoodsIssue.Lines.Add();
                            }
                            rsExistItmCd = null;
                        }
                        oGoodsIssue.Add();
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
