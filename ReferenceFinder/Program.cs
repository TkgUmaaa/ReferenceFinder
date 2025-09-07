// このプログラムは指定された .sln (ソリューション) を Roslyn (MSBuildWorkspace) で解析し
// 各プロジェクト内の「public const フィールド」を列挙し、その参照元(クラス/メンバー)と使用箇所コード断片を列挙します。
// 出力はコンソールへ表示します。

using System;
using System.Linq;
using System.Threading.Tasks;
using System.Collections.Generic;
using System.Text;
using System.IO;
using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.CSharp;
using Microsoft.CodeAnalysis.CSharp.Syntax;
using Microsoft.CodeAnalysis.MSBuild;
using Microsoft.CodeAnalysis.FindSymbols;
using Microsoft.CodeAnalysis.Text;

var workspace = MSBuildWorkspace.Create();

// コマンドライン引数から .sln を取得
if (args.Length == 0)
{
    Console.WriteLine("使用方法: ReferenceFinder <solution.sln>");
    return;
}

var solutionPath = args[0];
if (!File.Exists(solutionPath))
{
    Console.WriteLine($"指定されたソリューションファイルが存在しません: {solutionPath}");
    return;
}
solutionPath = Path.GetFullPath(solutionPath);
Console.WriteLine($"解析対象ソリューション: {solutionPath}");

Solution solution;
try
{
    solution = await workspace.OpenSolutionAsync(solutionPath);
}
catch (Exception ex)
{
    Console.WriteLine($"ソリューションを開く際にエラーが発生しました: {ex.Message}");
    return;
}

// 全 public const フィールドシンボル収集
var constFieldSymbols = new List<IFieldSymbol>();
foreach (var project in solution.Projects)
{
    var compilation = await project.GetCompilationAsync();
    if (compilation == null) continue;

    foreach (var tree in compilation.SyntaxTrees)
    {
        var model = compilation.GetSemanticModel(tree);
        var root = await tree.GetRootAsync();
        var fields = root.DescendantNodes()
            .OfType<FieldDeclarationSyntax>()
            .Where(f => f.Modifiers.Any(m => m.IsKind(SyntaxKind.ConstKeyword)) &&
                        f.Modifiers.Any(m => m.IsKind(SyntaxKind.PublicKeyword)));

        foreach (var fd in fields)
        {
            foreach (var v in fd.Declaration.Variables)
            {
                if (model.GetDeclaredSymbol(v) is IFieldSymbol fs)
                {
                    constFieldSymbols.Add(fs);
                }
            }
        }
    }
}

constFieldSymbols = constFieldSymbols
    .Distinct<IFieldSymbol>(SymbolEqualityComparer.Default)
    .ToList();

Console.WriteLine($"public const フィールド数: {constFieldSymbols.Count}");

foreach (IFieldSymbol fieldSymbol in constFieldSymbols)
{
    Console.WriteLine($"定数: {fieldSymbol.ContainingType?.ToDisplayString(SymbolDisplayFormat.MinimallyQualifiedFormat)}.{fieldSymbol.Name} = {fieldSymbol.ConstantValue ?? "null"}");

    var references = await SymbolFinder.FindReferencesAsync(fieldSymbol, solution);
    var refCount = 0;

    foreach (var refResult in references)
    {
        foreach (var loc in refResult.Locations)
        {
            var document = solution.GetDocument(loc.Document.Id);
            if (document == null) continue;
            var syntaxRoot = await document.GetSyntaxRootAsync();
            var node = syntaxRoot?.FindNode(loc.Location.SourceSpan);
            if (node == null) continue;

            var semanticModel = await document.GetSemanticModelAsync();
            if (semanticModel == null) continue;

            var containingSymbol = semanticModel.GetEnclosingSymbol(loc.Location.SourceSpan.Start);
            var containingType = containingSymbol?.ContainingType ?? (containingSymbol as ITypeSymbol);
            string typeName = containingType?.ToDisplayString(SymbolDisplayFormat.MinimallyQualifiedFormat) ?? "(不明な型)";

            string memberName = containingSymbol switch
            {
                IMethodSymbol m when m.MethodKind == MethodKind.Constructor => m.ContainingType.Name + ".ctor",
                IMethodSymbol m when m.MethodKind == MethodKind.StaticConstructor => m.ContainingType.Name + ".cctor",
                IMethodSymbol m when m.MethodKind == MethodKind.LocalFunction => m.Name + " (local function)",
                IMethodSymbol m => m.Name,
                IPropertySymbol p => p.Name,
                IFieldSymbol f => f.Name + " (field init)",
                IEventSymbol e => e.Name + " (event)",
                _ => containingSymbol?.Name ?? "(無名)"
            };

            var lineSpan = loc.Location.GetLineSpan();
            var line = lineSpan.StartLinePosition.Line + 1;
            var filePath = lineSpan.Path;

            var sourceText = await document.GetTextAsync();
            var snippet = BuildSnippet(sourceText, lineSpan); // 使用行のみ

            Console.WriteLine($"   参照: {typeName}.{memberName} 行:{line} ファイル:{filePath}");
            Console.WriteLine(snippet);
            refCount++;
        }
    }

    if (refCount == 0)
    {
        Console.WriteLine("   参照: (なし)");
    }
}

// 使用行のみ出力
static string BuildSnippet(SourceText text, FileLinePositionSpan lineSpan)
{
    int lineNumber = lineSpan.StartLinePosition.Line; // 定数識別子開始行
    if (lineNumber < 0 || lineNumber >= text.Lines.Count) return string.Empty;

    var lineText = text.Lines[lineNumber].ToString();
    var sb = new StringBuilder();
    sb.AppendLine("      ---- 使用行 ----");
    sb.AppendLine($"      >> {lineNumber + 1,5}: {lineText}");
    sb.Append("      ---------------");
    return sb.ToString();
}