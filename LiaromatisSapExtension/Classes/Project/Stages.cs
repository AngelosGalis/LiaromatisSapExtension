// ===============================================================================================
// 1.0.0
// ===============================================================================================
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using SAPbouiCOM.Framework;
using System.Text.RegularExpressions;

using System.IO;
//using System.Security.Permissions;


namespace LiaromatisSapExtension.Classes.Project
{
    class Stages
    {
        public static void SortColumns(string _FormUID, string _ItemUID, string _ColUID)
        {
            try
            {
                SAPbouiCOM.Form oForm = null;
                string RowNum = null;
                string RowNumNew = null;
                SAPbouiCOM.EditText oEditText = null;
                SAPbouiCOM.Item oItemMatrix = null;
                SAPbouiCOM.Matrix oMatrix = null;
                SAPbouiCOM.Column oColumn = null;

                oForm = Application.SBO_Application.Forms.Item(_FormUID);
                oItemMatrix = oForm.Items.Item(_ItemUID);
                oMatrix = (SAPbouiCOM.Matrix)oItemMatrix.Specific;

                oEditText = (SAPbouiCOM.EditText)oMatrix.Columns.Item("234000000").Cells.Item(1).Specific;
                RowNum = oEditText.Value.ToString();

                oColumn = (SAPbouiCOM.Column)oMatrix.Columns.Item(_ColUID);
                oColumn.TitleObject.Sort(SAPbouiCOM.BoGridSortType.gst_Ascending);
                oEditText = (SAPbouiCOM.EditText)oMatrix.Columns.Item("234000000").Cells.Item(1).Specific;
                RowNumNew = oEditText.Value.ToString();

                if (RowNum == RowNumNew)
                {
                    oColumn = (SAPbouiCOM.Column)oMatrix.Columns.Item(_ColUID);
                    oColumn.TitleObject.Sort(SAPbouiCOM.BoGridSortType.gst_Descending);
                }
            }
            catch (Exception)
            { }
        }

        public static void ExportPdfCSV(SAPbobsCOM.Company _Company, string _FormUID)
        {
            try
            {
                //Application.SBO_Application.StatusBar.SetText("Please, wait for the procedure to complete", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
                string UniqueID = null;
                string UniqueID2 = null;
                Boolean Flag = false;
                string Project = null;
                List<Models.PdfPathsModel> PDFFilesPaths = new List<Models.PdfPathsModel>();

                SAPbouiCOM.Form oForm = null;
                SAPbouiCOM.EditText oEditText = null;
                SAPbouiCOM.Item oItem = null;
                SAPbouiCOM.Item oItemMatrix = null;
                SAPbouiCOM.Matrix oMatrix = null;
                SAPbouiCOM.ComboBox oComboBox = null;

                oForm = Application.SBO_Application.Forms.Item(_FormUID);

                oItem = oForm.Items.Item("234000049");
                oEditText = (SAPbouiCOM.EditText)oItem.Specific;
                Project = oEditText.Value.ToString();

                oItemMatrix = oForm.Items.Item("234000062");
                oMatrix = (SAPbouiCOM.Matrix)oItemMatrix.Specific;

                for (int i = 1; i < oMatrix.VisualRowCount + 1; i++)
                {
                    oComboBox = (SAPbouiCOM.ComboBox)oMatrix.Columns.Item("U_ExStage").Cells.Item(i).Specific;
                    oComboBox.Value.ToString();

                    if (oComboBox.Value.ToString() == "Yes")
                    {
                        if (Flag == false)
                        {
                            oEditText = (SAPbouiCOM.EditText)oMatrix.Columns.Item("234000033").Cells.Item(i).Specific;
                            UniqueID = $@"N'{oEditText.Value.ToString()}'";
                            UniqueID2 = $@"{oEditText.Value.ToString()}";
                        }
                        else
                        {
                            oEditText = (SAPbouiCOM.EditText)oMatrix.Columns.Item("234000033").Cells.Item(i).Specific;
                            UniqueID = $@"{UniqueID},N'{oEditText.Value.ToString()}'";
                            UniqueID2 = $@"{UniqueID2}, {oEditText.Value.ToString()}";
                        }
                        Flag = true;
                    }
                }

                if (Flag == true)
                {
                    string ExportPath = null;

                    //Get all appropriate pdf file paths from "UDO" for the selected "project" and "stage"
                    PDFFilesPaths = GetPdfFilePaths(_Company, Project, UniqueID);

                    //Export selected Pdfs to client folder
                    ExportPath = ExportPdf(_Company, PDFFilesPaths);

                    //Export a CSV with sort info of UDOs from the selected "project" and "stage"
                    ExportCsv(PDFFilesPaths, ExportPath, Project, UniqueID2);
                }
                else
                {
                    Application.SBO_Application.StatusBar.SetText($@"Please, select a stage first", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
                }
            }
            catch (Exception)
            { }
        }

        public static List<Models.PdfPathsModel> GetPdfFilePaths(SAPbobsCOM.Company _Company, string _Project, string _UniqueID)
        {
            List<Models.PdfPathsModel> PdfsPaths = new List<Models.PdfPathsModel>();
            try
            {
                string PDFPaths = null;
                PDFPaths = $@"  SELECT T1.ItemCode, T1.Length1, T1.Width1, T2.U_OriginMill, T2.U_HeatNumber, T2.U_CertNo, T2.U_CeMark, T2.U_DoP, CAST(T2.U_Path AS NVARCHAR(max)) [Path], T0.CardCode, T2.U_CountryOrigin, ISNULL(CAST(T6.DocNum AS nvarchar),'Certificates without PO') [PODocNum], T0.DocNum, CONVERT(nvarchar, T0.DocDate, 3) [Date] FROM OPDN T0
                                INNER JOIN PDN1 T1 ON T0.DocEntry = T1.DocEntry
                                INNER JOIN [@DDS_PATHL] T2 ON T2.DocEntry = T1.U_Path
                                INNER JOIN OPMG T3 ON T3.FIPROJECT = T1.Project
                                INNER JOIN PMG1 T4 ON T4.AbsEntry = T3.AbsEntry
                                INNER JOIN PMG4 T5 ON T5.AbsEntry = T4.AbsEntry AND T4.POS = T5.StageID AND T1.DocEntry = T5.DocEntry AND T1.LineNum = T5.LineNum
                                LEFT JOIN OPOR T6 ON T6.DocEntry = T1.BaseEntry
                                INNER JOIN [@DDS_PATH] T7 ON T7.DocEntry = T2.DocEntry
                                WHERE T5.TYP = 20
                                AND T0.CANCELED = 'N'
                                AND T3.FIPROJECT = N'{_Project}'
                                AND T4.UniqueID IN ({_UniqueID})
                                AND ISNULL(CAST(T2.U_Path as NVARCHAR(max)),'') != ''
                                AND T2.U_ApprovedToUse = 'Yes'
                                AND (T1.BaseType = '22' OR T1.BaseType = '-1')
                                AND T7.U_AllInFull = 'Yes'
                                GROUP BY T1.ItemCode, T1.Length1, T1.Width1, T2.U_OriginMill, T2.U_HeatNumber, T2.U_CertNo, T2.U_CeMark, T2.U_DoP, CAST(T2.U_Path as NVARCHAR(max)), T0.CardCode, T2.U_CountryOrigin, T6.DocNum, T0.DocNum, T0.DocDate
                                ORDER BY T1.ItemCode ";

                SAPbobsCOM.Recordset rsPDFPaths;
                rsPDFPaths = (SAPbobsCOM.Recordset)_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                rsPDFPaths.DoQuery(PDFPaths);

                rsPDFPaths.MoveFirst();
                while (rsPDFPaths.EoF == false)
                {
                    int Length = 0;
                    int charLocation = 0;
                    string strPdfName = null;
                    try
                    {
                        strPdfName = rsPDFPaths.Fields.Item("Path").Value.ToString();
                        Length = strPdfName.Length;
                        charLocation = strPdfName.LastIndexOf("\\") + 1;
                        strPdfName = strPdfName.Substring(charLocation, Length - charLocation);

                        PdfsPaths.Add(new Models.PdfPathsModel()
                        {
                            Path = rsPDFPaths.Fields.Item("Path").Value.ToString(),
                            ItemCode = rsPDFPaths.Fields.Item("ItemCode").Value.ToString(),
                            Length1 = rsPDFPaths.Fields.Item("Length1").Value.ToString(),
                            Width1 = rsPDFPaths.Fields.Item("Width1").Value.ToString(),
                            OriginMill = rsPDFPaths.Fields.Item("U_OriginMill").Value.ToString(),
                            HeatNumber = rsPDFPaths.Fields.Item("U_HeatNumber").Value.ToString(),
                            CertNo = rsPDFPaths.Fields.Item("U_CertNo").Value.ToString(),
                            CeMark = rsPDFPaths.Fields.Item("U_CeMark").Value.ToString(),
                            DoP = rsPDFPaths.Fields.Item("U_DoP").Value.ToString(),
                            PdfName = strPdfName,
                            CardCode = rsPDFPaths.Fields.Item("CardCode").Value.ToString(),
                            CountryOrigin = rsPDFPaths.Fields.Item("U_CountryOrigin").Value.ToString(),
                            PODocNum = rsPDFPaths.Fields.Item("PODocNum").Value.ToString(),
                            DocNum = rsPDFPaths.Fields.Item("DocNum").Value.ToString(),
                            Date = rsPDFPaths.Fields.Item("Date").Value.ToString()
                        });

                        rsPDFPaths.MoveNext();
                    }
                    catch (Exception) { }
                }
                System.Runtime.InteropServices.Marshal.ReleaseComObject(rsPDFPaths);
            }
            catch (Exception ex)
            { }

            return PdfsPaths;
        }

        public static string ExportPdf(SAPbobsCOM.Company _Company, List<Models.PdfPathsModel> PdfFiles)
        {
            string AtchPath = null;
            string ExportPath = null;

            try
            {
                if (PdfFiles.Count > 0)
                {
                    AtchPath = $@" SELECT Code, Name 
                                   FROM [@DDS_ATTACHPATH]
                                   WHERE Code = 'ExportPath' ";
                    SAPbobsCOM.Recordset rsAttachPath2;
                    rsAttachPath2 = (SAPbobsCOM.Recordset)_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                    rsAttachPath2.DoQuery(AtchPath);

                    rsAttachPath2.MoveFirst();
                    ExportPath = rsAttachPath2.Fields.Item("Name").Value.ToString();
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(rsAttachPath2);

                    if (!String.IsNullOrEmpty(ExportPath) && !String.IsNullOrWhiteSpace(ExportPath))
                    {
                        // We use Double "for" and an "if" because we need to make distinct with less columns than them we need to display
                        foreach (var k in PdfFiles.Select(x => new { x.ItemCode, x.Length1, x.Width1, x.HeatNumber, x.CertNo, x.DocNum }).Distinct())
                        {
                            foreach (var j in PdfFiles.Select(x => new { x.Path, x.PdfName, x.ItemCode, x.Length1, x.Width1, x.HeatNumber, x.CertNo, x.DocNum }).Distinct())
                            {
                                // We choose only the first row with the same values
                                if (k.ItemCode == j.ItemCode && k.Length1 == j.Length1 && k.Width1 == j.Width1 && k.HeatNumber == j.HeatNumber && k.CertNo == j.CertNo && k.DocNum == j.DocNum)
                                {
                                    try
                                    {
                                        string sourceFile = j.Path;
                                        string ExportFile = System.IO.Path.Combine(ExportPath, j.PdfName);
                                        //Copy or overwrite the destination file if it already exists.
                                        System.IO.File.Copy(sourceFile, ExportFile, true);
                                    }
                                    catch (Exception ex)
                                    {
                                        Application.SBO_Application.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                                        return "Error";
                                    }
                                    break;
                                }
                            }
                        }
                    }
                    else
                    {
                        Application.SBO_Application.StatusBar.SetText($@"Export folder not defined, or Export folder has been changed or removed", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                        return "Error";
                    }
                }
                else
                {
                    Application.SBO_Application.StatusBar.SetText($@"There are no attachments for the selected stages", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
                    return "Error";
                }
            }
            catch (Exception)
            {
                return "Error";
            }
            return ExportPath;
        }

        public static void ExportCsv(List<Models.PdfPathsModel> PdfFiles, string _ExportPath, string Project, string UniqueID)
        {
            try
            {
                if (_ExportPath != "Error")
                {
                    var csv = new StringBuilder();
                    string filePath = $@"{_ExportPath}\{Project}.csv";

                    var newLine = string.Format("{0}", UniqueID);
                    csv.AppendLine(newLine);

                    newLine = string.Format("ItemCode, Length, Width, Origin Mill, Heat Number, Certificate Number, CE MARK, DOP, Vendor, Country, Purchase Order, Goods Receipt PO, Receipt Date");
                    csv.AppendLine(newLine);

                    // We use Double "for" and an "if" because we need to make distinct with less columns than them we need to display
                    string concatenateDate;
                    foreach (var j in PdfFiles.Select(x => new { x.ItemCode, x.Length1, x.Width1, x.HeatNumber, x.CertNo}).Distinct())
                    {
                        //get all dates, and from the rows that we exclude
                        concatenateDate = null;
                        foreach (var k in PdfFiles.Select(x => new { x.ItemCode, x.Length1, x.Width1, x.HeatNumber, x.CertNo, x.Date }).Distinct().Where(y => y.ItemCode == j.ItemCode && y.Length1 == j.Length1 && y.Width1 == j.Width1 && y.HeatNumber == j.HeatNumber && y.CertNo == j.CertNo))
                        {
                            if (String.IsNullOrEmpty(concatenateDate))
                                concatenateDate = k.Date;
                            else
                                concatenateDate = concatenateDate + "," + k.Date;
                        }

                        // We choose only the first row with the same values
                        foreach (var i in PdfFiles.Select(x => new { x.ItemCode, x.Length1, x.Width1, x.OriginMill, x.HeatNumber, x.CertNo, x.CeMark, x.DoP, x.CardCode, x.CountryOrigin, x.PODocNum, x.DocNum, x.Date }).Distinct().Where(y => y.ItemCode == j.ItemCode && y.Length1 == j.Length1 && y.Width1 == j.Width1 && y.HeatNumber == j.HeatNumber && y.CertNo == j.CertNo))
                        {
                            newLine = string.Format("{0},{1},{2},{3},{4},{5},{6},{7},{8},{9},{10},{11},{12}", i.ItemCode, i.Length1, i.Width1, i.OriginMill, i.HeatNumber, i.CertNo, i.CeMark, i.DoP, i.CardCode, i.CountryOrigin, i.PODocNum, i.DocNum, concatenateDate);
                            csv.AppendLine(newLine);
                            break;
                        }
                    }
                                                                                                         
                    File.WriteAllText(filePath, csv.ToString());
                    Application.SBO_Application.StatusBar.SetText("The procedure completed successfully", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);                    
                }
            }
            catch (Exception ex)
            {
                Application.SBO_Application.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
        }
    }
}
