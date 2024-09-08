using Microsoft.Build.Construction;
using Microsoft.CodeAnalysis.CSharp.Syntax;
using Microsoft.CodeAnalysis.CSharp;
using System;
using System.Collections.Generic;
using System.IO;
using System.Windows.Forms;
using Microsoft.CodeAnalysis;
using System.Linq;
using System.Diagnostics;

namespace ExtractFuntion
{
    public partial class MainForm : Form
    {
        public MainForm()
        {
            InitializeComponent();
            Init();
        }

        public void Init()
        {
            this.MaximizeBox = false;
        }

        private void button_FindSolution_Click(object sender, EventArgs e)
        {            
            using (OpenFileDialog FileDialog = new OpenFileDialog())
            {
                FileDialog.InitialDirectory = "D:\\Project\\ExtractFunctionProject\\ExtractFuntion";
                FileDialog.Filter = "txt files (*.txt)|*.txt|All files (*.*)|*.*";
                FileDialog.FilterIndex = 2;
                FileDialog.RestoreDirectory = true;

                if (FileDialog.ShowDialog() == DialogResult.OK)
                {
                    this.textBox_SolutionPath.Text = FileDialog.FileName;
                }
            }
        }

        private void button_StartExtract_Click(object sender, EventArgs e)
        {
            string SolutionPath = this.textBox_SolutionPath.Text;

            SolutionFile Solution = SolutionFile.Parse(SolutionPath);
            IEnumerable<ProjectInSolution> ProjectList = Solution.ProjectsInOrder;

            foreach (var Project in ProjectList)
            {
                string ProjectFileName = Project.RelativePath.Replace($"{Project.ProjectName}\\", "");
                string ProjectFilePath = Path.Combine(Path.GetDirectoryName(SolutionPath), Project.AbsolutePath);
                string ProjectFolderPath = ProjectFilePath.Replace($"\\{ProjectFileName}", "");

                // 시트 생성
                ExcelManager.Instance.Make_ExcelSheet(ProjectFileName.Replace(".csproj",""));

                ExtractClassFile_All(ProjectFolderPath);
            }

            // 엑셀 완료
        }

        private void ExtractClassFile_All(string ProjectPath)
        {
            var files = Directory.GetFiles(ProjectPath, "*.cs", SearchOption.AllDirectories).
                                                Where(s => s.Contains("\\bin\\") == false).
                                                Where(s => s.Contains("\\obj\\") == false).
                                                Where(s => s.Contains("\\Config\\") == false).
                                                Where(s => s.Contains(".Designer") == false).
                                                Where(s => s.Contains("\\Properties\\") == false);

            foreach (var file in files)
            {                
                ExtractMethod_All(file);
            }
        }

        private void ExtractMethod_All(string ClassFile)
        {
            string CodeScript = File.ReadAllText(ClassFile);
            SyntaxTree Tree = CSharpSyntaxTree.ParseText(CodeScript);

            try
            {
                var Method = Tree.GetRoot().DescendantNodes()
                         .OfType<MethodDeclarationSyntax>();

                // cs 파일의 정보와 해당 cs 파일의 모든 메소드의 정보를 엑셀에 기재
                for (int i = 0; i < Method.Count(); i++)
                {
                    ExcelManager.Instance.Make_CellValue(ClassFile, Method, i);
                    Console.WriteLine($"Method{i} :{Method.ElementAt(i).Modifiers} {Method.ElementAt(i).ReturnType} {Method.ElementAt(i).Identifier} {Method.ElementAt(i).ParameterList}");
                }
            }
            catch(Exception ex) 
            {
                Console.WriteLine($"[Error] {ex.Message}");
            }
        }

        private void MainForm_FormClosed(object sender, FormClosedEventArgs e)
        {
            ExcelManager.Instance.ReleaseMemory();
            Process.GetCurrentProcess().Kill();
        }
    }
}
