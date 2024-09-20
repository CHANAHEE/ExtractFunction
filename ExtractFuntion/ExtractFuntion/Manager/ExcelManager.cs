using Microsoft.CodeAnalysis.CSharp.Syntax;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
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

        public void Init_UI()
        {
            Range ProjectNameCell = worksheet.Range["A1", "D1"];
            ProjectNameCell.Merge();

            worksheet.Rows[1].RowHeight = 54; // 행 1의 높이를 40포인트로 설정
            worksheet.Columns["A"].ColumnWidth = 50; // 열 A의 너비를 30으로 설정
            worksheet.Columns["B"].ColumnWidth = 50; // 열 B의 너비를 30으로 설정
            worksheet.Columns["C"].ColumnWidth = 50; // 열 C의 너비를 30으로 설정
            worksheet.Columns["D"].ColumnWidth = 50; // 열 D의 너비를 30으로 설정

            // Cell 의 프로젝트 명 기입
            ProjectNameCell.Value = worksheet.Name;

            // Cell 의 배경색 변경
            ProjectNameCell.Interior.Color = ColorTranslator.ToOle(Color.FromArgb(172, 185, 202));

            // Cell 의 글씨체 변경
            ProjectNameCell.Font.Name = "맑은 고딕";
            ProjectNameCell.Font.Size = 36;
            ProjectNameCell.Font.Bold = true;

            // 가운데 정렬
            ProjectNameCell.HorizontalAlignment = XlHAlign.xlHAlignCenter;

            // 파일 이름 Cell
            Range FileNameCell = worksheet.Cells[3, 1];
            ChangeSetting_CategoryCell(FileNameCell, "파일 이름");

            // 함수 이름 Cell
            Range FunctionNameCell = worksheet.Cells[3, 2];
            ChangeSetting_CategoryCell(FunctionNameCell, "함수 이름");

            // 설명 Cell
            Range ClarificationCell = worksheet.Cells[3, 3];
            ChangeSetting_CategoryCell(ClarificationCell, "설명");

            // 비고 Cell
            Range NoteCell = worksheet.Cells[3, 4];
            ChangeSetting_CategoryCell(NoteCell, "비고");

            excelFile.Save();
        }

        public void ChangeSetting_CategoryCell(Range Cell, string CellName)
        {
            Cell.Value = CellName;
            Cell.Interior.Color = ColorTranslator.ToOle(Color.FromArgb(255, 230, 153));
            Cell.Font.Name = "맑은 고딕";
            Cell.Font.Size = 11;
            Cell.Font.Bold = true;
            Cell.HorizontalAlignment = XlHAlign.xlHAlignCenter;

            Borders NewBoards = Cell.Borders;
            // 상단 테두리 설정
            Border TopBorder = NewBoards[XlBordersIndex.xlEdgeTop];
            TopBorder.LineStyle = XlLineStyle.xlContinuous;
            TopBorder.ColorIndex = 0; // 검정색
            TopBorder.TintAndShade = 0;
            TopBorder.Weight = XlBorderWeight.xlThin;

            // 하단 테두리 설정
            Border BottomBorder = NewBoards[XlBordersIndex.xlEdgeBottom];
            BottomBorder.LineStyle = XlLineStyle.xlContinuous;
            BottomBorder.ColorIndex = 0; // 검정색
            BottomBorder.TintAndShade = 0;
            BottomBorder.Weight = XlBorderWeight.xlThin;

            // 좌측 테두리 설정
            Border LeftBorder = NewBoards[XlBordersIndex.xlEdgeLeft];
            LeftBorder.LineStyle = XlLineStyle.xlContinuous;
            LeftBorder.ColorIndex = 0; // 검정색
            LeftBorder.TintAndShade = 0;
            LeftBorder.Weight = XlBorderWeight.xlThin;

            // 우측 테두리 설정
            Border RightBorder = NewBoards[XlBordersIndex.xlEdgeRight];
            RightBorder.LineStyle = XlLineStyle.xlContinuous;
            RightBorder.ColorIndex = 0; // 검정색
            RightBorder.TintAndShade = 0;
            RightBorder.Weight = XlBorderWeight.xlThin;
        }

        public bool Make_ClassFile_CellValue(string ClassFileName, int MethodCount, int StartIndex)
        {
            if (MethodCount == 0)
            {
                return false;
            }

            Range ClassCell = worksheet.Range[worksheet.Cells[StartIndex + 4, 1], worksheet.Cells[StartIndex + (MethodCount - 1) + 4, 1]];

            //Console.WriteLine($"[{StartIndex} + 4, 1] , [({MethodInfo.Count()} - 1) + 4, 1] Range is Merge");
            ClassCell.Merge();
            ClassCell.Value = ClassFileName.Split('\\').Last();

            excelFile.Save();

            return true;
        }

        public void Make_Function_CellValue(IEnumerable<MethodDeclarationSyntax> Method, int CellIndex, int CurrentIndex)
        {
            Range MethodCell = worksheet.Cells[CellIndex + 4, 2];
            MethodCell.Value = $"{Method.ElementAt(CurrentIndex).ReturnType} {Method.ElementAt(CurrentIndex).Identifier} {Method.ElementAt(CurrentIndex).ParameterList}";

            excelFile.Save();
        }

        public void Make_UI(int Last_CellIndex)
        {
            Range ClassCell = worksheet.Range[worksheet.Cells[4, 1], worksheet.Cells[4 + (Last_CellIndex - 1), 1]];            

            ClassCell.Font.Name = "맑은 고딕";
            ClassCell.Font.Size = 13;
            ClassCell.HorizontalAlignment = XlHAlign.xlHAlignCenter;

            Range FunctionCell = worksheet.Range[worksheet.Cells[4, 2], worksheet.Cells[4 + (Last_CellIndex - 1), 2]];

            FunctionCell.Font.Name = "맑은 고딕";
            FunctionCell.Font.Size = 11;
            FunctionCell.Font.Bold = true;
            FunctionCell.HorizontalAlignment = XlHAlign.xlHAlignCenter;

            for (int i = 0; i < Last_CellIndex; i++)
            {
                Range FirstCell = worksheet.Cells[4 + i, 1];
                FirstCell.RowHeight = 34.5;
            }

            Range FunctionCol = worksheet.Columns[2];
            FunctionCol.AutoFit();

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
    }
}
