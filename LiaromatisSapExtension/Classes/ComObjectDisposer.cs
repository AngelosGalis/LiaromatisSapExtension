using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using SAPbobsCOM;

namespace LiaromatisSapExtension.Classes
{
    // release all the ComObjects.
    internal static class ComObjectDisposer
    {
        internal static void ReleaseComObject(SAPbobsCOM.Recordset _rs, SAPbobsCOM.GeneralService _GeneralService, SAPbobsCOM.GeneralData _GeneralData, SAPbobsCOM.GeneralDataParams _GeneralParams, SAPbobsCOM.CompanyService _sCmp)
        {
            try
            {
                if (_rs != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(_rs);
                    _rs = null;
                }

                if (_GeneralService != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(_GeneralService);
                    _GeneralService = null;
                }

                if (_GeneralData != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(_GeneralData);
                    _GeneralData = null;
                }

                if (_GeneralParams != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(_GeneralParams);
                    _GeneralParams = null;
                }

                if (_sCmp != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(_sCmp);
                    _sCmp = null;
                }
            }
            catch (Exception ex)
            { }
        }
    }
}
