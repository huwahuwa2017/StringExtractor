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
        private static string DLLPath = @"C:\Program Files (x86)\Steam\steamapps\common\From The Depths\From_The_Depths_Data\Managed\Ftd.dll";

        private static string OutputPath = Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), "StringExtraction.json");

        private static JArray MainJObject = new JArray();



        [STAThread]
        private static void Main(string[] args)
        {
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
                Filter = "json files (*.json)|*.json|All files (*.*)|*.*"
            };

            if (SFD.ShowDialog() == DialogResult.OK)
            {
                OutputPath = SFD.FileName;
            }

            Console.WriteLine(OutputPath);
            SFD.Dispose();



            Console.WriteLine("キー入力で文字列の抽出を開始します");
            Console.ReadKey();



            AssemblyDefinition AssemblyDef = AssemblyDefinition.ReadAssembly(DLLPath);
            List<TypeDefinition> TypeDefList = AssemblyDef.Modules.SelectMany(x => x.Types).ToList();

            foreach (TypeDefinition TypeDef in TypeDefList)
            {
                ShowString(TypeDef);
            }

            File.WriteAllText(OutputPath, MainJObject.ToString());



            Console.WriteLine("\n文字列の抽出が完了しました　キー入力で終了します");
            Console.ReadKey();
        }

        private static void ShowString(TypeDefinition TypeDef)
        {
            foreach (TypeDefinition NestedTypeDef in TypeDef.NestedTypes.ToList())
            {
                ShowString(NestedTypeDef);
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

                JArray TextObject = null;
                int Count0 = 0;

                foreach (Instruction Ins in MethodDef.Body.Instructions)
                {
                    if (Ins.OpCode != OpCodes.Ldstr)
                    {
                        continue;
                    }

                    string Text = Ins.Operand.ToString();
                    Console.WriteLine(Text);
                    TextObject = TextObject ?? (TextObject = new JArray { NameSpaceName, TypeName, MethodName });
                    TextObject.Add(new JArray { Count0, Text, Text });

                    ++Count0;
                }

                if (TextObject != null)
                {
                    MainJObject.Add(TextObject);
                }
            }
        }
    }
}
