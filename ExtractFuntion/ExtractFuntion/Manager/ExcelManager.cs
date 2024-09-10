using Microsoft.CodeAnalysis.CSharp.Syntax;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Drawing;
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

        private int cellIndex;
        public int CELL_INDEX { get { return cellIndex; } set { cellIndex = value; } }

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

        public void Make_Excel_UI()
        {
            Range ProjectNameCell = worksheet.Range["A1", "D1"];
            ProjectNameCell.Merge();

            worksheet.Rows[1].RowHeight = 40; // 행 1의 높이를 40포인트로 설정
            worksheet.Columns["A"].ColumnWidth = 50; // 열 A의 너비를 30으로 설정
            worksheet.Columns["B"].ColumnWidth = 50; // 열 B의 너비를 30으로 설정
            worksheet.Columns["C"].ColumnWidth = 50; // 열 C의 너비를 30으로 설정
            worksheet.Columns["D"].ColumnWidth = 50; // 열 D의 너비를 30으로 설정

            // Cell 의 프로젝트 명 기입
            ProjectNameCell.Value = worksheet.Name;

            // Cell 의 배경색 변경
            ProjectNameCell.Interior.Color = ColorTranslator.FromOle(Color.FromArgb(172, 185, 202).ToArgb());

            excelFile.Save();
        }

        public void Make_CellValue(string ClassFiles, IEnumerable<MethodDeclarationSyntax> Method,int CellIndex, int CurrentIndex)
        {
            Range ClassCell = worksheet.Cells[CellIndex + 2, 1];
            Range MethodCell = worksheet.Cells[CellIndex + 2, 2];

            while (true)
            {
                if (ClassCell.Value != null)
                {
                    return;
                }

                break;
            }

            ClassCell.Value = ClassFiles.Split('\\').Last();
            MethodCell.Value = $"{Method.ElementAt(CurrentIndex).Modifiers} {Method.ElementAt(CurrentIndex).ReturnType} {Method.ElementAt(CurrentIndex).Identifier} {Method.ElementAt(CurrentIndex).ParameterList}";

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
            if(excelApp != null) 
            {
                Marshal.FinalReleaseComObject(excelApp);
            }
            if (excelFiles != null)
            {
                Marshal.FinalReleaseComObject(excelFiles);
            }
            if (excelFile != null)
            {
                Marshal.FinalReleaseComObject(excelFile);
            }
            if (worksheet != null)
            {
                Marshal.FinalReleaseComObject(worksheet);
            }
        }
    }
}
