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
        private Workbook excelFile = null;

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
                excelFile = excelApp.Workbooks.Open(ExcelFileName);
            }
            else
            {
                excelFile = excelApp.Workbooks.Add();
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
            //엑셀 파일의 시트 관련 객체
            Worksheet NewWorkSheet = null;

            //기본 시트 후에 생성
            NewWorkSheet = excelFile.Worksheets.Add(Type.Missing, excelFile.Worksheets[1]);
            NewWorkSheet.Name = ProjectFileName;

            excelFile.Save();

            //메모리 해제를 위한 처리
            excelFile.Close();
            excelApp.Quit();
            ReleaseExcelObject(NewWorkSheet);
            ReleaseExcelObject(excelFile);
            ReleaseExcelObject(excelApp);
        }
    }
}
