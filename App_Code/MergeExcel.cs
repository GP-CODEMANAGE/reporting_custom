using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Web;
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel; 

/// <summary>
/// Summary description for ExcelMerge
/// </summary>
public class MergeExcel
{
    Excel.Application app = new Microsoft.Office.Interop.Excel.ApplicationClass();
    Excel.Workbook bookDest = null;
    Excel.Worksheet sheetDest = null;
    Excel.Workbook bookSource = null;
    Excel.Worksheet sheetSource = null;
    string[] _sourceFiles = null;
    string _destFile = string.Empty;
    string _columnEnd = string.Empty;
    int _headerRowCount = 0;
    int _currentRowCount = 0;

    public MergeExcel(string[] sourceFiles, string destFile, string columnEnd, int headerRowCount)
    {
        bookDest = (Excel.WorkbookClass)app.Workbooks.Add(Missing.Value);
        sheetDest = bookDest.Worksheets.Add(Missing.Value, Missing.Value, Missing.Value, Missing.Value) as Excel.Worksheet;
        sheetDest.Name = "Data";
        _sourceFiles = sourceFiles;
        _destFile = destFile;
        _columnEnd = columnEnd;
        _headerRowCount = headerRowCount;
    }
    void OpenBook(string fileName)
    {
        bookSource = app.Workbooks._Open(fileName, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value);
        sheetSource = bookSource.Worksheets[1] as Excel.Worksheet;
    }
    void CloseBook()
    {
        bookSource.Close(false, Missing.Value, Missing.Value);
    }

    void CopyHeader()
    {
        Excel.Range range = sheetSource.get_Range("A1", _columnEnd + _headerRowCount.ToString());
        range.Copy(sheetDest.get_Range("A1", Missing.Value));
        _currentRowCount += _headerRowCount;
    }
    void CopyData()
    {
        int sheetRowCount = sheetSource.UsedRange.Rows.Count;
        Excel.Range range = sheetSource.get_Range(string.Format("A{0}", _headerRowCount), _columnEnd + sheetRowCount.ToString());
        range.Copy(sheetDest.get_Range(string.Format("A{0}", _currentRowCount), Missing.Value));
        _currentRowCount += range.Rows.Count;
    }
    void Save()
    {
        bookDest.Saved = true;
        bookDest.SaveCopyAs(_destFile);
    }
    void Quit()
    {
        app.Quit();
    }
    void DoMerge()
    {
        bool b = false;
        foreach (string strFile in _sourceFiles)
        {
            OpenBook(strFile);
            if (b == false)
            {
                CopyHeader();
                b = true;
            }
            CopyData();
            CloseBook();
        }
        Save();
        Quit();
    }
}


