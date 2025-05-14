using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using SAPbouiCOM.Framework;

namespace LiaromatisSapExtension.Classes.ItemMasterData
{
    class SapB1DataImportUI
    {
        public static void ImportUDO(string _FormUID, SAPbobsCOM.Company _Company)
        {
            SAPbouiCOM.Form oForm = null;
            SAPbouiCOM.Form oFormUDF = null;
            SAPbouiCOM.Item oItem = null;
            SAPbouiCOM.Item oItemMatrix = null;
            SAPbouiCOM.EditText oEditText = null;
            SAPbouiCOM.CheckBox oCheckBox = null;
            string ItemCode = null;
            string ItemGroup = null;
            string PropertyName = null;
            Boolean PropertyNameChecked = false;
            SAPbouiCOM.ComboBox oComboBox = null;

            try
            {
                oForm = Application.SBO_Application.Forms.Item(_FormUID);
                oFormUDF = Application.SBO_Application.Forms.Item(oForm.UDFFormUID); //Get UDF Form
                oForm.Freeze(true);

                if (oForm.Mode.ToString() == "fm_ADD_MODE") //only in add mode.
                {
                    oItem = oForm.Items.Item("5");
                    oEditText = (SAPbouiCOM.EditText)oItem.Specific;
                    ItemCode =  oEditText.Value.ToString();

                    oComboBox = (SAPbouiCOM.ComboBox)oForm.Items.Item("39").Specific;
                    ItemGroup = oComboBox.Selected.Description;

                    if (ItemGroup == $@"Πάγια" && !String.IsNullOrEmpty(ItemCode)) 
                    {
                        oItemMatrix = oForm.Items.Item("129");
                        SAPbouiCOM.Matrix oMatrix = ((SAPbouiCOM.Matrix)(oItemMatrix.Specific));
                        for (int i = 1; i < oMatrix.VisualRowCount + 1; i++)
                        {
                            oEditText = (SAPbouiCOM.EditText)oMatrix.Columns.Item("1").Cells.Item(i).Specific;
                            PropertyName = oEditText.Value.ToString();
                            oCheckBox = (SAPbouiCOM.CheckBox)oMatrix.Columns.Item("2").Cells.Item(i).Specific;
                            PropertyNameChecked = oCheckBox.Checked;

                            if (PropertyName == $@"Μηχανήματα" && PropertyNameChecked == true)
                            {
                                //Create two New UDOs
                                string UdoCode = null;
                                //Check if the UDO "SpareParts" already Exist
                                UdoCode = ItemCode;
                                string ExistanceSPARE = $@" SELECT * FROM [@DDS_SPARE]
                                                            WHERE Code = N'{ItemCode}'  ";
                                SAPbobsCOM.Recordset rsExistS;
                                rsExistS = (SAPbobsCOM.Recordset)_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                                rsExistS.DoQuery(ExistanceSPARE);
                                if (rsExistS.RecordCount == 0)
                                {
                                    UdoCode = AddUDO(_Company, ItemCode, "SpareParts");
                                }
                                ComObjectDisposer.ReleaseComObject(rsExistS, null, null, null, null);
                                //Add Udo Code
                                oItem = oFormUDF.Items.Item("U_SpareParts");
                                oEditText = (SAPbouiCOM.EditText)oItem.Specific;
                                oEditText.Value = UdoCode;
                                
                                //Check if the UDO "Consumables" already Exist
                                UdoCode = ItemCode;
                                string ExistanceCONSUMABLES = $@" SELECT * FROM [@DDS_CONSUMABLES]
                                                                  WHERE Code = N'{ItemCode}'  ";
                                SAPbobsCOM.Recordset rsExistC;
                                rsExistC = (SAPbobsCOM.Recordset)_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                                rsExistC.DoQuery(ExistanceCONSUMABLES);
                                if (rsExistC.RecordCount == 0)
                                {
                                    UdoCode = AddUDO(_Company, ItemCode, "Consumables");
                                }
                                ComObjectDisposer.ReleaseComObject(rsExistC, null, null, null, null);
                                //Add Udo Code
                                oItem = oFormUDF.Items.Item("U_Consumables");
                                oEditText = (SAPbouiCOM.EditText)oItem.Specific;
                                oEditText.Value = UdoCode;
                            }
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
        public static string AddUDO(SAPbobsCOM.Company _Company,string _ItemCode, string _UdoName)
        {
            string Code = null;
            try
            {
                SAPbobsCOM.GeneralService oGeneralService = null;
                SAPbobsCOM.GeneralData oGeneralData = null;
                SAPbobsCOM.GeneralDataParams oGeneralParams = null;
                SAPbobsCOM.CompanyService sCmp = null;

                sCmp = (SAPbobsCOM.CompanyService)_Company.GetCompanyService();
                oGeneralService = sCmp.GetGeneralService(_UdoName);
                oGeneralData = (SAPbobsCOM.GeneralData)oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData);

                oGeneralData.SetProperty("Code", _ItemCode);

                //Add Udo
                oGeneralParams = oGeneralService.Add(oGeneralData);
                //Get the Code of the Added Udo
                Code = oGeneralParams.GetProperty("Code").ToString();

                ComObjectDisposer.ReleaseComObject(null, oGeneralService, oGeneralData, oGeneralParams, sCmp);
            }
            catch (Exception ex)
            { }
            return Code;
        }
    }
}
