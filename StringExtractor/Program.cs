using ClosedXML.Excel;
using Mono.Cecil;
using Mono.Cecil.Cil;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Windows.Forms;

namespace StringExtractor
{
    public class Program
    {
        private static XLWorkbook xlsxObject;

        private static JArray jsonObject;

        private static int Variable0;



        [STAThread]
        private static void Main(string[] args)
        {
            string DLLPath = @"C:\Program Files (x86)\Steam\steamapps\common\From The Depths\From_The_Depths_Data\Managed\Ftd";
            string OutputPath = Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), "StringExtraction");



            string Text0 = "DLLファイルを選択してください";
            Console.WriteLine(Text0);

            OpenFileDialog OFD = new OpenFileDialog
            {
                Title = Text0,
                InitialDirectory = Path.GetDirectoryName(DLLPath),
                FileName = DLLPath,
                Filter = "dll files (*.dll)|*.dll|All files (*.*)|*.*"
            };

            if (OFD.ShowDialog() == DialogResult.OK)
            {
                DLLPath = OFD.FileName;
            }

            Console.WriteLine(DLLPath);
            OFD.Dispose();



            string Text1 = "保存先を指定してください";
            Console.WriteLine(Text1);

            SaveFileDialog SFD = new SaveFileDialog
            {
                Title = Text1,
                InitialDirectory = Path.GetDirectoryName(OutputPath),
                FileName = OutputPath,
                Filter = "xlsx files (*.xlsx)|*.xlsx|json files (*.json)|*.json|All files (*.*)|*.*"
            };

            if (SFD.ShowDialog() == DialogResult.OK)
            {
                OutputPath = SFD.FileName;
            }

            Console.WriteLine(OutputPath);
            SFD.Dispose();



            Console.WriteLine("キー入力で文字列の抽出を開始します");
            Console.ReadKey();



            string Extension = Path.GetExtension(OutputPath);

            if (Extension == ".xlsx")
            {
                xlsxObject = new XLWorkbook();
                xlsxObject.Worksheets.Add("Translation");

                Start(Extension, DLLPath);

                xlsxObject.SaveAs(OutputPath);
            }
            else
            {
                jsonObject = new JArray();

                Start(Extension, DLLPath);

                File.WriteAllText(OutputPath, jsonObject.ToString());
            }



            Console.WriteLine("\n文字列の抽出が完了しました　キー入力で終了します");
            Console.ReadKey();
        }

        private static void Start(string Extension, string DLLPath)
        {
            AssemblyDefinition AssemblyDef = AssemblyDefinition.ReadAssembly(DLLPath);
            List<TypeDefinition> TypeDefList = AssemblyDef.Modules.SelectMany(x => x.Types).ToList();

            foreach (TypeDefinition TypeDef in TypeDefList)
            {
                ShowString(TypeDef, Extension);
            }
        }

        private static void ShowString(TypeDefinition TypeDef, string Extension)
        {
            foreach (TypeDefinition NestedTypeDef in TypeDef.NestedTypes.ToList())
            {
                ShowString(NestedTypeDef, Extension);
            }

            foreach (MethodDefinition MethodDef in TypeDef.Methods)
            {
                if (!MethodDef.HasBody)
                {
                    continue;
                }

                string NameSpaceName = TypeDef.Namespace;
                string TypeName = TypeDef.Name;
                string ParameterName = string.Join(",", MethodDef.Parameters.Select(x => x.ParameterType.FullName));
                string MethodName = $"{MethodDef.Name}({ParameterName})";

                int Count0 = 0;

                if (Extension == ".xlsx")
                {
                    IXLWorksheet Worksheet = xlsxObject.Worksheets.Worksheet("Translation");
                    bool WritingStart = false;

                    foreach (Instruction Ins in MethodDef.Body.Instructions)
                    {
                        if (Ins.OpCode != OpCodes.Ldstr)
                        {
                            continue;
                        }

                        string Text = Ins.Operand.ToString();

                        if (Text == "")
                        {
                            ++Count0;
                            continue;
                        }

                        Console.WriteLine(Text);

                        if (!WritingStart)
                        {
                            WritingStart = true;
                            Worksheet.Cell(++Variable0, 1).Value = NameSpaceName;
                            Worksheet.Cell(++Variable0, 1).Value = TypeName;
                            Worksheet.Cell(++Variable0, 1).Value = MethodName;
                        }

                        int Count1 = ++Variable0;
                        Worksheet.Cell(Count1, 1).Value = Count0;
                        Worksheet.Cell(Count1, 2).Value = Text;
                        Worksheet.Cell(Count1, 3).Value = Text;

                        ++Count0;
                    }

                    if (WritingStart)
                    {
                        Worksheet.Cell(++Variable0, 1).Value = "--";
                    }
                }
                else
                {
                    JArray TextObject = null;

                    foreach (Instruction Ins in MethodDef.Body.Instructions)
                    {
                        if (Ins.OpCode != OpCodes.Ldstr)
                        {
                            continue;
                        }

                        string Text = Ins.Operand.ToString();

                        if (Text == "")
                        {
                            ++Count0;
                            continue;
                        }

                        Console.WriteLine(Text);

                        TextObject = TextObject ?? (TextObject = new JArray { NameSpaceName, TypeName, MethodName });
                        TextObject.Add(new JArray { Count0, Text, Text });

                        ++Count0;
                    }

                    if (TextObject != null)
                    {
                        jsonObject.Add(TextObject);
                    }
                }
            }
        }
    }
}
