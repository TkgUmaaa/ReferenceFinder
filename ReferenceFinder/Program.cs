// このプログラムは指定された .sln (ソリューション) を Roslyn (MSBuildWorkspace) で解析し
// VB プロジェクト内の Public Const フィールド と Public / Protected メソッドを列挙し
// それぞれの参照元(クラス/メンバー)と使用箇所コード断片を CSV 出力します。

using System;
using System.Linq;
using System.Threading.Tasks;
using System.Collections.Generic;
using System.Text;
using System.IO;
using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.MSBuild;
using Microsoft.CodeAnalysis.FindSymbols;
using Microsoft.CodeAnalysis.Text;
using Microsoft.CodeAnalysis.VisualBasic;
using Microsoft.CodeAnalysis.VisualBasic.Syntax;
using VBSyntaxKind = Microsoft.CodeAnalysis.VisualBasic.SyntaxKind;

// 追加: Shift_JIS 利用のためコードページプロバイダ登録 (.NET 8)
Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

// コンソール出力蓄積 (ログ目的)
var output = new StringBuilder();
void Log(string s)
{
    Console.WriteLine(s);
    output.AppendLine(s);
}

// CSV 行蓄積
var csvRows = new List<string>();
// ヘッダー (日本語) アクセスレベル列を追加
csvRows.Add(string.Join(',', new[]{
    "メンバー種別",         // MemberKind (ConstField / Method)
    "アクセスレベル",       // Accessibility (Public / Protected)
    "宣言名前空間",         // DeclaringNamespace
    "宣言クラス",           // DeclaringType
    "宣言",                 // Declaration
    "参照クラス名前空間",   // ReferenceTypeNamespace
    "参照クラス",           // ReferenceType
    "参照メンバー",         // ReferenceMember
    "行番号",               // LineNumber
    "コード行",             // CodeLine
    "ファイルパス"          // FilePath
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

// 収集用リスト
var constFieldSymbols = new List<IFieldSymbol>();
var methodSymbols = new List<IMethodSymbol>();

foreach (var project in solution.Projects.Where(p => p.Language == LanguageNames.VisualBasic))
{
    var compilation = await project.GetCompilationAsync();
    if (compilation == null) continue;

    foreach (var tree in compilation.SyntaxTrees)
    {
        if (tree.Options is not VisualBasicParseOptions) continue; // 念のため
        var model = compilation.GetSemanticModel(tree);
        var root = await tree.GetRootAsync();

        // Public Const フィールド
        var fieldDecls = root.DescendantNodes()
            .OfType<FieldDeclarationSyntax>()
            .Where(f => f.Modifiers.Any(m => m.IsKind(VBSyntaxKind.PublicKeyword)) &&
                        f.Modifiers.Any(m => m.IsKind(VBSyntaxKind.ConstKeyword)));

        foreach (var fd in fieldDecls)
        {
            foreach (var declarator in fd.Declarators)
            {
                foreach (var name in declarator.Names)
                {
                    if (model.GetDeclaredSymbol(name) is IFieldSymbol fs)
                    {
                        constFieldSymbols.Add(fs);
                    }
                }
            }
        }

        // Public / Protected 系 メソッド
        var methodStatements = root.DescendantNodes().OfType<MethodStatementSyntax>();
        foreach (var ms in methodStatements)
        {
            if (!(ms.Kind() == VBSyntaxKind.SubStatement || ms.Kind() == VBSyntaxKind.FunctionStatement))
                continue;
            var sym = model.GetDeclaredSymbol(ms) as IMethodSymbol;
            if (sym == null) continue;
            if (sym.MethodKind == MethodKind.PropertyGet || sym.MethodKind == MethodKind.PropertySet) continue;
            if (sym.MethodKind == MethodKind.EventAdd || sym.MethodKind == MethodKind.EventRemove || sym.MethodKind == MethodKind.EventRaise) continue;
            // アクセスレベル: Public / Protected のみ
            if (!(sym.DeclaredAccessibility == Accessibility.Public ||
                  sym.DeclaredAccessibility == Accessibility.Protected))
            {
                continue;
            }
            methodSymbols.Add(sym);
        }
    }
}

constFieldSymbols = constFieldSymbols.Distinct<IFieldSymbol>(SymbolEqualityComparer.Default).ToList();
methodSymbols = methodSymbols.Distinct<IMethodSymbol>(SymbolEqualityComparer.Default).ToList();

Log($"(VB) Public Const フィールド数: {constFieldSymbols.Count}");
Log($"(VB) Public/Protected メソッド数: {methodSymbols.Count}");

var memberEntries = new List<(ISymbol Symbol, string Kind)>();
memberEntries.AddRange(constFieldSymbols.Select(f => ((ISymbol)f, "ConstField")));
memberEntries.AddRange(methodSymbols.Select(m => ((ISymbol)m, "Method")));

foreach (var (symbol, kind) in memberEntries)
{
    string declarationText = GetDeclarationText(symbol, kind);

    var declaringTypeSymbol = symbol.ContainingType;
    var declaringNamespace = declaringTypeSymbol?.ContainingNamespace is { IsGlobalNamespace: true } ? "(グローバル)" : declaringTypeSymbol?.ContainingNamespace?.ToDisplayString() ?? "(不明な名前空間)";
    var declaringType = declaringTypeSymbol?.ToDisplayString(SymbolDisplayFormat.MinimallyQualifiedFormat) ?? "(不明な型)";
    string accessibility = symbol switch
    {
        IFieldSymbol => "Public", // ConstField は Public のみ
        IMethodSymbol ms => ms.DeclaredAccessibility == Accessibility.Public ? "Public" : "Protected",
        _ => string.Empty
    };

    Log($"[{kind}] {declarationText} ({accessibility})");

    var references = await SymbolFinder.FindReferencesAsync(symbol, solution);
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

            var referenceNamespace = containingType?.ContainingNamespace is { IsGlobalNamespace: true } ? "(グローバル)" : containingType?.ContainingNamespace?.ToDisplayString() ?? "(不明な名前空間)";

            csvRows.Add(string.Join(',', new[]{
                CsvEscape(kind),
                CsvEscape(accessibility),
                CsvEscape(declaringNamespace),
                CsvEscape(declaringType),
                CsvEscape(declarationText),
                CsvEscape(referenceNamespace),
                CsvEscape(typeName),
                CsvEscape(memberName),
                CsvEscape(line.ToString()),
                CsvEscape(lineText),
                CsvEscape(filePath)
            }));
        }
    }

    if (refCount == 0)
    {
        Log("   参照: (なし)");
        csvRows.Add(string.Join(',', new[]{
            CsvEscape(kind),
            CsvEscape(accessibility),
            CsvEscape(declaringNamespace),
            CsvEscape(declaringType),
            CsvEscape(declarationText),
            "","","","","",""
        }));
    }
}

// CSV 出力
WriteOutAndExit();

// 宣言テキスト取得
static string GetDeclarationText(ISymbol symbol, string kind)
{
    if (kind == "ConstField" && symbol is IFieldSymbol fieldSymbol)
    {
        string declarationText = string.Empty;
        var syntaxRef = fieldSymbol.DeclaringSyntaxReferences.FirstOrDefault();
        if (syntaxRef != null)
        {
            var syntaxNode = syntaxRef.GetSyntaxAsync().Result;
            if (syntaxNode is ModifiedIdentifierSyntax mis &&
                mis.Parent is VariableDeclaratorSyntax vbVarDecl &&
                vbVarDecl.Parent is FieldDeclarationSyntax vbFieldDecl)
            {
                bool single = vbFieldDecl.Declarators.Count == 1 && vbVarDecl.Names.Count == 1;
                if (single)
                {
                    declarationText = vbFieldDecl.ToFullString().Trim();
                }
                else
                {
                    var modifiers = string.Join(" ", vbFieldDecl.Modifiers.Select(m => m.Text));
                    if (!string.IsNullOrEmpty(modifiers)) modifiers += " ";
                    string? typeText = null;
                    if (vbVarDecl.AsClause is SimpleAsClauseSyntax simpleAs)
                    {
                        typeText = simpleAs.Type.ToFullString().Trim();
                    }
                    else if (vbVarDecl.AsClause is AsNewClauseSyntax asNew)
                    {
                        typeText = asNew.ToFullString().Trim();
                    }
                    if (string.IsNullOrEmpty(typeText))
                    {
                        typeText = fieldSymbol.Type.ToDisplayString(SymbolDisplayFormat.MinimallyQualifiedFormat);
                    }
                    var initValue = vbVarDecl.Initializer?.Value?.ToFullString().Trim();
                    string initText = !string.IsNullOrEmpty(initValue)
                        ? " = " + initValue
                        : " = " + FormatConstantValue(fieldSymbol);
                    if (!typeText!.StartsWith("As ", StringComparison.OrdinalIgnoreCase))
                    {
                        typeText = "As " + typeText;
                    }
                    declarationText = $"{modifiers}Const {mis.Identifier.Text} {typeText}{initText}";
                }
            }
        }
        if (string.IsNullOrWhiteSpace(declarationText))
        {
            declarationText = $"Public Const {fieldSymbol.Name} As {fieldSymbol.Type.ToDisplayString(SymbolDisplayFormat.MinimallyQualifiedFormat)} = {FormatConstantValue(fieldSymbol)}";
        }
        return declarationText;
    }
    else if (kind == "Method" && symbol is IMethodSymbol methodSymbol)
    {
        var syntaxRef = methodSymbol.DeclaringSyntaxReferences.FirstOrDefault();
        if (syntaxRef != null)
        {
            var sn = syntaxRef.GetSyntaxAsync().Result;
            if (sn is MethodStatementSyntax ms)
            {
                return ms.ToFullString().Trim();
            }
        }
        var returnType = methodSymbol.ReturnsVoid ? "Void" : methodSymbol.ReturnType.ToDisplayString(SymbolDisplayFormat.MinimallyQualifiedFormat);
        var parameters = string.Join(", ", methodSymbol.Parameters.Select(p => p.Name + " As " + p.Type.ToDisplayString(SymbolDisplayFormat.MinimallyQualifiedFormat)));
        var accessibility = methodSymbol.DeclaredAccessibility switch
        {
            Accessibility.Public => "Public ",
            Accessibility.Protected => "Protected ",
            _ => string.Empty
        };
        var keyword = methodSymbol.ReturnsVoid ? "Sub" : "Function";
        var ret = methodSymbol.ReturnsVoid ? string.Empty : " As " + returnType;
        return $"{accessibility}{keyword} {methodSymbol.Name}({parameters}){ret}".Trim();
    }
    return symbol.ToDisplayString();
}

// VB 用: 定数値を VB コード片として整形
static string FormatConstantValue(IFieldSymbol field)
{
    object? v = field.ConstantValue;
    if (v == null) return "Nothing";
    return v switch
    {
        string s => "\"" + s.Replace("\"", "\"\"") + "\"",
        char c => "\"" + c.ToString().Replace("\"", "\"\"") + "\"",
        bool b => b ? "True" : "False",
        float f => f.ToString(System.Globalization.CultureInfo.InvariantCulture) + "F",
        double d => d.ToString(System.Globalization.CultureInfo.InvariantCulture),
        decimal m => m.ToString(System.Globalization.CultureInfo.InvariantCulture) + "D",
        _ => Convert.ToString(v, System.Globalization.CultureInfo.InvariantCulture) ?? "Nothing"
    };
}

// 使用行の単一行テキスト
static string GetSingleLine(SourceText text, FileLinePositionSpan lineSpan)
{
    int lineNumber = lineSpan.StartLinePosition.Line; // 識別子開始行
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
        // Shift_JIS (932) で出力
        var sjis = Encoding.GetEncoding(932);
        File.WriteAllLines(outPath, csvRows, sjis);
        Console.WriteLine($"結果ファイル: {outPath}");
    }
    catch (Exception ex)
    {
        Console.WriteLine($"結果ファイル書き込み中にエラー: {ex.Message}");
    }
}