using Microsoft.CodeAnalysis.CSharp.Syntax;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms.VisualStyles;

namespace ExtractFuntion
{
    public class ExcelManager
    {
        private static ExcelManager excelMngInstance;

        public static ExcelManager Instance 
        { 
            get 
            {
                if (excelMngInstance == null)
                {
                    excelMngInstance = new ExcelManager();
                }

                return excelMngInstance; 
            } 
        }

        private Application excelApp = null;
        private Workbooks excelFiles = null;
        private Workbook excelFile = null;
        private Worksheet worksheet = null;

        public string solutionFileName = "";
        public string classFileNames;
        public string[] methodNames;

        public bool Init()
        {
            excelApp = new Application();

            //ture일 경우 메시지 발생.
            excelApp.DisplayAlerts = false;

            string ExcelFileName = "D:\\ExtractFunction.xlsx";

            //엑셀 기본 생성

            FileInfo NewFileInfo = new FileInfo(ExcelFileName);
            
            if(NewFileInfo.Exists)
            {
                excelFiles = excelApp.Workbooks;
                excelFile = excelFiles.Open(ExcelFileName);
            }
            else
            {
                excelFiles = excelApp.Workbooks;
                excelFile = excelFiles.Add();
                excelFile.SaveAs(ExcelFileName);
            }
            

            //Range NewCell_1 = NewWorkSheet.Cells[1, 1];
            //Range NewCell_2 = NewWorkSheet.Cells[1, 2];
            //Range NewCell_3 = NewWorkSheet.Cells[2, 1];
            //Range NewCell_4 = NewWorkSheet.Cells[2, 2];

            //NewCell_1.Value = "1번 셀 데이터";
            //NewCell_2.Value = "2번 셀 데이터";
            //NewCell_3.Value = "3번 셀 데이터";
            //NewCell_4.Value = "4번 셀 데이터";

            return true;
        }

        private void ReleaseExcelObject(object obj)
        {
            try
            {
                if (obj != null)
                {
                    Marshal.ReleaseComObject(obj);
                    obj = null;
                }
            }
            catch (Exception ex)
            {
                obj = null;
                throw ex;
            }
            finally
            {
                GC.Collect();
            }
        }

        public void Make_ExcelSheet(string ProjectFileName)
        {
            //기본 시트 후에 생성
            worksheet = excelFile.Worksheets.Add(Type.Missing, excelFile.Worksheets[1]);

            foreach(Worksheet sheet in excelFile.Sheets)
            {
                if(sheet.Name == ProjectFileName)
                {
                    return;
                }
            }

            worksheet.Name = ProjectFileName;

              
        }

        public void Make_CellValue(string ClassFiles, IEnumerable<MethodDeclarationSyntax> Method,int CellIndex)
        {
            Range ClassCell = worksheet.Cells[CellIndex + 1, 1];
            Range MethodCell = worksheet.Cells[CellIndex + 1, 2];

            while (true)
            {
                if (ClassCell.Value != null)
                {
                    return;
                }               

                break;
            }

            ClassCell.Value = ClassFiles.Split('\\').Last();
            MethodCell.Value = $"{Method.ElementAt(CellIndex).Modifiers} {Method.ElementAt(CellIndex).ReturnType} {Method.ElementAt(CellIndex).Identifier} {Method.ElementAt(CellIndex).ParameterList}";

            excelFile.Save();
        }

        public void ReleaseMemory()
        {
            excelFile.Close();
            ReleaseExcelObject(excelFile);

            excelFiles.Close();
            ReleaseExcelObject(excelFiles);

            excelApp.Quit();
            ReleaseExcelObject(excelApp);

            //메모리 해제를 위한 처리
            Marshal.FinalReleaseComObject(excelApp);
            Marshal.FinalReleaseComObject(excelFiles);
            Marshal.FinalReleaseComObject(excelFile);
            Marshal.FinalReleaseComObject(worksheet);
        }
    }
}
