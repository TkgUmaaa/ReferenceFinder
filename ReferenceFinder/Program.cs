// このプログラムは指定された .sln (ソリューション) を Roslyn (MSBuildWorkspace) で解析し
// 各プロジェクト内の「public const フィールド」を列挙し、その参照元(クラス/メンバー)と使用箇所コード断片を列挙します。
// 出力はコンソールと実行ファイルと同じフォルダの CSV ファイルへ書き出します。

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

// コンソール出力蓄積 (ログ目的)
var output = new StringBuilder();
void Log(string s)
{
    Console.WriteLine(s);
    output.AppendLine(s);
}

// CSV 行蓄積
var csvRows = new List<string>();
// ヘッダー (定数宣言クラス追加)
csvRows.Add(string.Join(',', new[]{
    "FieldDeclaration",
    "FieldDeclaringType",
    "ReferenceType",
    "ReferenceMember",
    "LineNumber",
    "FilePath",
    "CodeLine"
}));

var workspace = MSBuildWorkspace.Create();

// コマンドライン引数から .sln を取得
if (args.Length == 0)
{
    Log("使用方法: ReferenceFinder <solution.sln>");
    WriteOutAndExit();
    return;
}

var solutionPath = args[0];
if (!File.Exists(solutionPath))
{
    Log($"指定されたソリューションファイルが存在しません: {solutionPath}");
    WriteOutAndExit();
    return;
}
solutionPath = Path.GetFullPath(solutionPath);
Log($"解析対象ソリューション: {solutionPath}");

Solution solution;
try
{
    solution = await workspace.OpenSolutionAsync(solutionPath);
}
catch (Exception ex)
{
    Log($"ソリューションを開く際にエラーが発生しました: {ex.Message}");
    WriteOutAndExit();
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

Log($"public const フィールド数: {constFieldSymbols.Count}");

foreach (IFieldSymbol fieldSymbol in constFieldSymbols)
{
    // 元のフィールド宣言 (単一変数ならソースそのまま) もしくは再構築した宣言テキスト
    string declarationText = string.Empty;
    var syntaxRef = fieldSymbol.DeclaringSyntaxReferences.FirstOrDefault();
    if (syntaxRef != null)
    {
        var syntaxNode = await syntaxRef.GetSyntaxAsync();
        if (syntaxNode is VariableDeclaratorSyntax vds)
        {
            if (vds.Parent?.Parent is FieldDeclarationSyntax fieldDecl)
            {
                // フィールドが 1 変数のみの場合はフィールド宣言全体をそのまま利用
                if (fieldDecl.Declaration.Variables.Count == 1)
                {
                    declarationText = fieldDecl.ToFullString().Trim();
                }
                else
                {
                    // 複数宣言から対象変数のみを再構築
                    var modifiers = string.Join(" ", fieldDecl.Modifiers.Select(m => m.Text));
                    if (!string.IsNullOrEmpty(modifiers)) modifiers += " ";
                    var typeText = fieldDecl.Declaration.Type.ToFullString().Trim();
                    var initText = vds.Initializer != null ? " = " + vds.Initializer.Value.ToFullString().Trim() : string.Empty;
                    declarationText = $"{modifiers}{typeText} {vds.Identifier.Text}{initText};";
                }
            }
        }
    }

    if (string.IsNullOrWhiteSpace(declarationText))
    {
        // フォールバック (必要最低限の情報)
        declarationText = $"public const {fieldSymbol.Type.ToDisplayString(SymbolDisplayFormat.MinimallyQualifiedFormat)} {fieldSymbol.Name} = {fieldSymbol.ConstantValue ?? "null"};";
    }

    var fieldDeclaringType = fieldSymbol.ContainingType?.ToDisplayString(SymbolDisplayFormat.MinimallyQualifiedFormat) ?? "(不明な型)";
    Log(declarationText);

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
            var lineText = GetSingleLine(sourceText, lineSpan).Trim();

            Log($"   参照: {typeName}.{memberName} 行:{line} ファイル:{filePath}");
            Log($"      >> {line,5}: {lineText}");
            refCount++;

            csvRows.Add(string.Join(',', new[]{
                CsvEscape(declarationText),
                CsvEscape(fieldDeclaringType),
                CsvEscape(typeName),
                CsvEscape(memberName),
                CsvEscape(line.ToString()),
                CsvEscape(filePath),
                CsvEscape(lineText)
            }));
        }
    }

    if (refCount == 0)
    {
        Log("   参照: (なし)");
        // 参照なしでもレコードを 1 行出す
        csvRows.Add(string.Join(',', new[]{
            CsvEscape(declarationText),
            CsvEscape(fieldDeclaringType),
            "","","","",""
        }));
    }
}

// CSV 出力
WriteOutAndExit();

// 使用行の単一行テキスト
static string GetSingleLine(SourceText text, FileLinePositionSpan lineSpan)
{
    int lineNumber = lineSpan.StartLinePosition.Line; // 定数識別子開始行
    if (lineNumber < 0 || lineNumber >= text.Lines.Count) return string.Empty;
    return text.Lines[lineNumber].ToString();
}

static string CsvEscape(string value)
{
    if (string.IsNullOrEmpty(value)) return string.Empty;
    if (value.Contains('"') || value.Contains(',') || value.Contains('\n') || value.Contains('\r'))
    {
        return "\"" + value.Replace("\"", "\"\"") + "\"";
    }
    return value;
}

// ローカル関数: CSV ファイルへ書き出し
void WriteOutAndExit()
{
    try
    {
        var exeDir = AppContext.BaseDirectory;
        var timestamp = DateTime.Now.ToString("yyyyMMddHHmmss");
        var outPath = Path.Combine(exeDir, $"ReferenceFinderResult_{timestamp}.csv");
        File.WriteAllLines(outPath, csvRows, new UTF8Encoding(false));
        Console.WriteLine($"結果ファイル: {outPath}");
    }
    catch (Exception ex)
    {
        Console.WriteLine($"結果ファイル書き込み中にエラー: {ex.Message}");
    }
}