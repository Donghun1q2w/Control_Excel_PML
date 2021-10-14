using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Aveva.Core.PMLNet;
using excel = Microsoft.Office.Interop.Excel;

namespace PML.Excel.Control
{
    [PMLNetCallable()]
    public class PMLExcelControl
    {
        [PMLNetCallable()]
        public PMLExcelControl()
        { }
        [PMLNetCallable()]
        public void Assign(PMLExcelControl that)
        { }
        [PMLNetCallable()]
        public void ReadExcel(string Path)
        {
            
            excel.Workbook wb = null;
            excel.Worksheet ws = null;
            excel.Application Application = new excel.Application();
            wb = Application.Workbooks.Open(Path);
            ws = wb.Worksheets.Item[1] as excel.Worksheet;
            ws.Range["C3"].Value = "TEST";
            ws.SaveAs(Path);
            

        }

    }

}
