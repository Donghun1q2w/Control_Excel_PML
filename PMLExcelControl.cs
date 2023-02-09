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
        excel.Range rg = null;
        excel.Application Application = new excel.Application();
        string InitialPath;
        
        private string otherPath = "";
        [PMLNetCallable()]
        public string OtherPath { get => otherPath; set => otherPath = value; }

        [PMLNetCallable()]
        public PMLExcelControl()
        { }
        [PMLNetCallable()]
        public PMLExcelControl(string Path)
        { ReadExcel(Path, 1); }
        [PMLNetCallable()]
        public void Assign(PMLExcelControl that)
        { }
        [PMLNetCallable()]
        public void ReadFromExcel(string Path)
        {
            ReadExcel(Path, 1);

        }
        [PMLNetCallable()]
        public void ReadFromExcel(string Path, string SheetName)
        {
            ReadExcel(Path, SheetName);

        }
        [PMLNetCallable()]
        public void ReadFromExcel(string Path , double SheetNumber)
        {
            ReadExcel( Path, SheetNumber.ToInt());
        }
        private  void ReadExcel(string Path, object Sheet)
        {
            InitialPath = Path;
            try
            {
                wb = Application.Workbooks.Open(Path);
                ws = wb.Worksheets.Item[Sheet] as excel.Worksheet;

            }
            catch (Exception e)
            {
                MessageBox.Show(e.ToString());
                SaveAndClose();
            }

        }
        [PMLNetCallable()]
        public void OpenWorkSheet(string SheetName)
        {
            OpenWorkSheet_pri(SheetName);
        }
        [PMLNetCallable()]
        public void OpenWorkSheet(double SheetNum)
        {
            try
            {
                OpenWorkSheet_pri(SheetNum.ToInt());
            }
            catch (Exception e)
            {
                MessageBox.Show(e.ToString());
                SaveAndClose();
            }
        }
        private void OpenWorkSheet_pri(object Sheet)
        {
            ws = wb.Worksheets.Item[Sheet] as excel.Worksheet;
        }
        [PMLNetCallable()]
        public Hashtable LoadArray(double Start_Row, double Start_Column, double Row_Length, double Column_Length)
        {
            return LoadArrayPri(ws.Range[ws.Cells[Start_Row, Start_Column], ws.Cells[Start_Row + Row_Length - 1, Start_Column + Column_Length - 1]] as excel.Range);
        }
        [PMLNetCallable()]
        public Hashtable LoadArray()
        {
            return LoadArrayPri(ws.UsedRange as excel.Range);
        }
        private Hashtable LoadArrayPri(excel.Range range )
        {
            Hashtable TableTotal = new Hashtable();
            try
            {
                object[,] array = range.Value;
                for (int i = 1; i <= array.GetLength(0); i++)
                {
                    Hashtable Row = new Hashtable();
                    for (int j = 1; j <= array.GetLength(1); j++)
                    {
                        if (!Row.ContainsKey(j.ToDouble()))
                            Row.Add(j.ToDouble(), array[i , j ].ToString());
                        else
                            Row[j.ToDouble()] = array[i , j ].ToString();
                    }
                    TableTotal.Add(i.ToDouble(), Row);
                }
                return TableTotal;
            }
            catch (Exception e)
            {
                MessageBox.Show(e.ToString());
                SaveAndClose();
                return TableTotal;
            }
        }
        [PMLNetCallable()]
        public void ResetWorkSheet()
        {
            wb = null;
            ws = null;
            Application = new excel.Application();
        }
        [PMLNetCallable()]
        public void ReadHashtable(Hashtable input)
        {

            MessageBox.Show(input.Count.ToString());
            int i = 1;
            for(int a = 1; a <= input.Count; a++)
            {
                Hashtable ht = (Hashtable)input[a.ToDouble()];
                for(int b=1;b<=ht.Count;b++)
                    MessageBox.Show( " val " + a.ToString() + " ddd " + " val " + ht[b.ToDouble()].ToString());
            }
            //foreach (DictionaryEntry inputitem in input)
            //{
            //    Hashtable subitems = (Hashtable)inputitem.Value;
            //    foreach (DictionaryEntry subitem in subitems)
            //        MessageBox.Show( "key" + inputitem.Key.ToString() + " val " + i.ToString() + " ddd key" + subitem.Key.ToString() + " val "+ subitem.Value.ToString());
            //    i++;

            //}
            //string a = input[1].GetType().Name;
        }
        [PMLNetCallable()]
        public void UploadtoWorkBook(Hashtable input  , double Start_Row, double Start_Column)
        {
            try
            {
                if(wb==null)
                {
                    wb = Application.Workbooks.Add();
                    ws = wb.Worksheets.Item[1] as excel.Worksheet;
                }
                bool chkDoubleArray = false;
                //int Start_Row = 2;
                //int Start_Column = 1;
                int Row_Length = input.Count;
                int Column_Length = 1;
                object afc = input[Column_Length.ToDouble()];
                MessageBox.Show("type is " + afc.GetType().ToString());
                if (afc.GetType().ToString().Contains("Hashtable"))
                {
                    Hashtable aa = (Hashtable)input[(1).ToDouble()];
                    Column_Length = aa.Count;
                    chkDoubleArray = true;
                }
                object[,] data = new object[Row_Length, Column_Length];
                for (int i = 1;i<=Row_Length;i++)
                {
                    if(chkDoubleArray)
                    {
                        Hashtable RowData = (Hashtable)input[i.ToDouble()];
                        for(int j = 1; j<=Column_Length;j++)
                        {
                            data[i - 1, j - 1] = RowData[j.ToDouble()];
                        }
                    }
                    else
                    {
                        data[i - 1, 0] = input[i.ToDouble()];
                    }
                }
                ws.Range[ws.Cells[Start_Row.ToInt(), Start_Column.ToInt()], ws.Cells[Start_Row.ToInt() + Row_Length - 1, Start_Column.ToInt() + Column_Length - 1]].Value = data;
                ws.Range["C5"].Value = "TEST";
                
            }
            catch(Exception e)
            { 
                MessageBox.Show(e.ToString());
                SaveAndClose();
            }

        }
        [PMLNetCallable()]
        public Hashtable GetSheetNames()
        {
            Hashtable sheetNames = new Hashtable();
            try
            {
                for (int i = 1; i <= wb.Worksheets.Count; i++)
                {
                    excel.Worksheet sheet = wb.Worksheets[i];
                    sheetNames[i.ToDouble()] = sheet.Name;
                }
                return sheetNames;
            }
            catch (Exception e)
            {
                MessageBox.Show(e.ToString());
                SaveAndClose();
                return sheetNames;
            }
        }
        [PMLNetCallable()]
        public void CopyFormat( Hashtable sourceRange , Hashtable DestinationRange )
        {
            excel.Range Source_Range = ws.Range[ws.Cells[sourceRange[(1).ToDouble()], sourceRange[(2).ToDouble()]], ws.Cells[sourceRange[(3).ToDouble()], sourceRange[(4).ToDouble()]]];
            excel.Range Destination_Range = ws.Range[ws.Cells[DestinationRange[(1).ToDouble()], DestinationRange[(2).ToDouble()]], ws.Cells[DestinationRange[(3).ToDouble()], DestinationRange[(4).ToDouble()]]];
            Source_Range.Copy();
            Destination_Range.PasteSpecial(excel.XlPasteType.xlPasteFormats, excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, false, false);

        }
        [PMLNetCallable()]
        public void Save( string path)
        {
            otherPath = path;
            SaveAndClose();
        }
        [PMLNetCallable()]
        public void Close()
        {
            wb.Close(true);
            Application.Quit();
            ReleaseObject(wb);
            ReleaseObject(ws);
            ReleaseObject(Application);
        }
        private void SaveAndClose()
        {
            if( OtherPath==null|| OtherPath=="")
                wb.Save();
            else
                wb.SaveAs(OtherPath);
            Close();
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
