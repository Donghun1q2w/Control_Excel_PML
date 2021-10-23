using System;
using System.IO;
using System.Data;
using System.Collections;
using System.Runtime.InteropServices;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Aveva.Core.PMLNet;
using excel = Microsoft.Office.Interop.Excel;

namespace PML.Excel.Control
{
    [PMLNetCallable()]
    public class PMLExcelControl
    {
        excel.Workbook wb = null;
        excel.Worksheet ws = null;
        excel.Application Application = new excel.Application();
        string InitialPath;
        
        private string otherPath = "";
        [PMLNetCallable()]
        public string OtherPath { get => otherPath; set => otherPath = value; }

        [PMLNetCallable()]
        public PMLExcelControl()
        { }
        [PMLNetCallable()]
        public void Assign(PMLExcelControl that)
        { }
        [PMLNetCallable()]
        public void ReadExcel(string Path , double SheetNumber)
        {
            InitialPath = Path;
            try
            {
                wb = Application.Workbooks.Open(Path);
                ws = wb.Worksheets.Item[SheetNumber] as excel.Worksheet;
                
            }
            catch (Exception e)
            {
                MessageBox.Show(e.ToString());
                SaveAndClose();
            }
            
        }
        [PMLNetCallable()]
        public Hashtable LoadArray(double Start_Row, double Start_Column , double Row_Length , double Column_Length )
        {
            object[,] array = ws.Range[ws.Cells[Start_Row, Start_Column], ws.Cells[Start_Row + Row_Length - 1, Start_Column + Column_Length - 1]];
            Hashtable TableTotal = new Hashtable();
            Hashtable Row = new Hashtable();
            for (int i = 1; i<=array.GetLength(0);i++)
            {
                for(int j = 1;j<=array.GetLength(1);j++)
                {
                    if (!Row.ContainsKey(j))
                        Row.Add(j, array[i - 1, j - 1]);
                    else
                        Row[j] = array[i - 1, j - 1];
                }
                TableTotal.Add(i, Row);
            }
            return TableTotal;


        }
        [PMLNetCallable()]
        public void SaveExcel(Hashtable input , string path , double Start_Row, double Start_Column)
        {
            otherPath = path;
            try
            {
                bool chkDoubleArray = false;
                //int Start_Row = 2;
                //int Start_Column = 1;
                int Row_Length = input.Count;
                int Column_Length = 1;
                object afc = input[1];
                MessageBox.Show("type is " + afc.GetType().ToString());
                if (afc.GetType().ToString().Contains("Hashtable"))
                {
                    Hashtable aa = (Hashtable)input[Start_Row];
                    Column_Length = aa.Count;
                    chkDoubleArray = true;
                }
                object[,] data = new object[Row_Length, Column_Length];
                for (int i = 1;i<=Row_Length;i++)
                {
                    if(chkDoubleArray)
                    {
                        Hashtable RowData = (Hashtable)input[i];
                        for(int j = 1; j<=Column_Length;j++)
                        {
                            data[i - 1, j - 1] = RowData[j];
                        }
                    }
                    else
                    {
                        data[i - 1, 0] = input[i];
                    }
                }
                ws.Range[ws.Cells[Start_Row, Start_Column], ws.Cells[Start_Row + Row_Length - 1, Start_Column + Column_Length - 1]].Value = data;
                ws.Range["C5"].Value = "TEST";
            }
            catch(Exception e)
            { MessageBox.Show(e.ToString()); }
            finally
            {
                SaveAndClose();
            }

        }
        [PMLNetCallable()]
        public void SaveAndClose()
        {
            if( OtherPath==null|| OtherPath=="")
                wb.Save();
            else
                wb.SaveAs(OtherPath);
            wb.Close(true);
            Application.Quit();
            ReleaseObject(wb);
            ReleaseObject(ws);
            ReleaseObject(Application);
        }
        void ReleaseObject(object obj)
        {
            try
            {
                if (obj != null)
                {
                    Marshal.ReleaseComObject(obj); // 액셀 객체 해제 obj = null; 
                }
            }
            catch 
            (Exception ex) 
            { obj = null; throw ex; }
            finally
            {
                GC.Collect(); // 가비지 수집 
            }
        }
    }

}
