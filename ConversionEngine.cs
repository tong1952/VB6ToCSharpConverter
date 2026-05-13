using System.Text;
using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.CSharp;
using Microsoft.CodeAnalysis.CSharp.Syntax;
using Microsoft.CodeAnalysis.Formatting;
using VB6ToCSharp.Lexer;
using VB6ToCSharp.Parser;
using VB6ToCSharp.Transformer;
using static Microsoft.CodeAnalysis.CSharp.SyntaxFactory;

namespace VB6ToCSharp;

public static class ConversionEngine
{
    // ─── Public API ───────────────────────────────────────────────────────

    public static (string CsCode, IReadOnlyList<string> Diagnostics) Convert(
        string vb6Source,
        string moduleName,
        string? filePath = null,
        string namespaceName = "Converted")
    {
        vb6Source = StripRcsIfNeeded(vb6Source);
        var lexer = new VB6Lexer(vb6Source);
        var tokens = lexer.Tokenize();

        var moduleKind = DetectModuleKind(vb6Source, filePath);
        var parser = new VB6Parser(tokens);
        var module = parser.Parse(moduleName, moduleKind);

        var header = new StringBuilder();
        header.AppendLine($"// Converted from {Path.GetFileName(filePath ?? moduleName)} by vb6cs");
        header.AppendLine($"// Date: {DateTime.Now:yyyy-MM-dd HH:mm}");
        header.AppendLine();

        var transformer = new VB6ToCSharpTransformer(namespaceName);
        var compilationUnit = transformer.Transform(module);

        compilationUnit = (CompilationUnitSyntax)new CommentEmptyStatementRewriter().Visit(compilationUnit);

        using var workspace = new AdhocWorkspace();
        var formatted = Formatter.Format(compilationUnit, workspace).ToFullString();

        var diagnostics = new List<string>(parser.Diagnostics);
        diagnostics.AddRange(transformer.Diagnostics);
        diagnostics.AddRange(DetectComPatterns(tokens));
        return (header + formatted, diagnostics);
    }

    // Converts one file; calls log(message) after each file with a status line.
    // Returns true on success. warningCount is set before log is called.
    public static bool ConvertFile(
        string inputPath,
        string outputPath,
        string ns,
        out int warningCount,
        Action<string> log)
    {
        warningCount = 0;
        try
        {
            string source = File.ReadAllText(inputPath, Encoding.Default);
            var (csSource, diagnostics) = Convert(
                source, Path.GetFileNameWithoutExtension(inputPath), inputPath, ns);

            File.WriteAllText(outputPath, csSource,
                new UTF8Encoding(encoderShouldEmitUTF8Identifier: true));

            warningCount = diagnostics.Count;
            if (diagnostics.Count > 0)
            {
                string logPath = Path.ChangeExtension(outputPath, ".log");
                WriteLog(logPath, inputPath, diagnostics);
                log($"{Path.GetFileName(inputPath)} → {Path.GetFileName(outputPath)}: " +
                    $"OK ({diagnostics.Count} warning(s) → {Path.GetFileName(logPath)})");
            }
            else
            {
                log($"{Path.GetFileName(inputPath)} → {Path.GetFileName(outputPath)}: OK");
            }
            return true;
        }
        catch (Exception ex)
        {
            log($"{Path.GetFileName(inputPath)}: FAILED — {ex.Message}");
            return false;
        }
    }

    // ─── Helpers ──────────────────────────────────────────────────────────

    private static readonly (TokenKind Kind, string? Text, string Message)[] ComPatterns =
    [
        // Implements: fully converted to a C# base type — no warning needed.
        // WithEvents: handled by the transformer; unmatched vars reported via transformer diagnostics.
        (TokenKind.Identifier,  "GetObject",     "Note: GetObject — COM moniker binding; replace with an appropriate .NET API"),
    ];

    private static IEnumerable<string> DetectComPatterns(List<Token> tokens)
    {
        var seen = new HashSet<string>(); // one warning per pattern per file
        for (int i = 0; i < tokens.Count; i++)
        {
            var tok = tokens[i];
            foreach (var (kind, text, message) in ComPatterns)
            {
                if (tok.Kind != kind) continue;
                if (text != null && !string.Equals(tok.Text, text, StringComparison.OrdinalIgnoreCase)) continue;
                // Skip identifiers that are method calls on an object (obj.GetObject)
                if (i > 0 && tokens[i - 1].Kind == TokenKind.Dot) continue;
                if (seen.Add(message))
                    yield return $"Line {tok.Line}: {message}";
            }
        }
    }

    // If the source is an RCS file, extract only the head revision's VB6 source.
    // RCS files start with "head\t" and embed each revision's text in @...@ blocks.
    // The first text block is the head revision (full source); subsequent blocks are deltas.
    public static string StripRcsIfNeeded(string source)
    {
        if (!source.StartsWith("head\t", StringComparison.OrdinalIgnoreCase) &&
            !source.StartsWith("head ", StringComparison.OrdinalIgnoreCase))
            return source;

        // Find the first "\ntext\n@" or "\ntext\r\n@" — that's the head revision's text section.
        int textIdx = -1;
        foreach (var marker in new[] { "\ntext\n@", "\ntext\r\n@" })
        {
            int idx = source.IndexOf(marker, StringComparison.Ordinal);
            if (idx >= 0 && (textIdx < 0 || idx < textIdx))
                textIdx = idx;
        }
        if (textIdx < 0) return source;

        int atStart = source.IndexOf('@', textIdx + 1);
        if (atStart < 0) return source;
        atStart++; // step past the opening '@'

        var sb = new StringBuilder();
        int i = atStart;
        while (i < source.Length)
        {
            if (source[i] == '@')
            {
                if (i + 1 < source.Length && source[i + 1] == '@')
                {
                    sb.Append('@');
                    i += 2;
                }
                else break; // closing '@'
            }
            else
            {
                sb.Append(source[i++]);
            }
        }
        return sb.ToString();
    }

    private static void WriteLog(string logPath, string inputPath, IReadOnlyList<string> diagnostics)
    {
        var sb = new StringBuilder();
        sb.AppendLine($"Conversion notes for {Path.GetFileName(inputPath)}");
        sb.AppendLine($"Date: {DateTime.Now:yyyy-MM-dd HH:mm}");
        sb.AppendLine();
        foreach (var d in diagnostics)
            sb.AppendLine($"  {d}");
        File.WriteAllText(logPath, sb.ToString(),
            new UTF8Encoding(encoderShouldEmitUTF8Identifier: true));
    }

    private static ModuleKind DetectModuleKind(string source, string? filePath)
    {
        string ext = filePath != null ? Path.GetExtension(filePath).ToLowerInvariant() : "";
        if (ext is ".frm" or ".ctl" or ".pag") return ModuleKind.Form;
        if (ext == ".cls") return ModuleKind.Class;
        return ModuleKind.Standard;
    }

    // ─── Syntax rewriter ─────────────────────────────────────────────────

    // Promotes comments from comment-only EmptyStatements onto the next real statement,
    // eliminating the "; // remark" anti-pattern produced by Roslyn.
    private sealed class CommentEmptyStatementRewriter : CSharpSyntaxRewriter
    {
        public override SyntaxNode VisitBlock(BlockSyntax node)
        {
            node = (BlockSyntax)base.VisitBlock(node)!;

            var statements = node.Statements;
            if (statements.Count == 0) return node;

            var result = new List<StatementSyntax>();
            var pending = new List<SyntaxTrivia>();

            foreach (var stmt in statements)
            {
                if (IsCommentOnlyEmpty(stmt, out var trivia))
                {
                    pending.AddRange(trivia);
                }
                else
                {
                    if (pending.Count > 0)
                    {
                        result.Add(stmt.WithLeadingTrivia(
                            TriviaList(pending).AddRange(stmt.GetLeadingTrivia())));
                        pending.Clear();
                    }
                    else
                    {
                        result.Add(stmt);
                    }
                }
            }

            if (pending.Count > 0)
            {
                var close = node.CloseBraceToken;
                node = node.WithCloseBraceToken(
                    close.WithLeadingTrivia(TriviaList(pending).AddRange(close.LeadingTrivia)));
            }

            return node.WithStatements(SyntaxList(result));
        }

        private static bool IsCommentOnlyEmpty(StatementSyntax stmt, out IEnumerable<SyntaxTrivia> trivia)
        {
            trivia = Enumerable.Empty<SyntaxTrivia>();
            if (stmt is not EmptyStatementSyntax) return false;

            var comments = stmt.GetLeadingTrivia()
                .Concat(stmt.GetTrailingTrivia())
                .Where(t => t.IsKind(SyntaxKind.SingleLineCommentTrivia))
                .ToList();

            if (comments.Count == 0) return false;

            trivia = comments.SelectMany(c => new[]
            {
                c,
                SyntaxFactory.EndOfLine(Environment.NewLine)
            });
            return true;
        }

        private static SyntaxTriviaList TriviaList(IEnumerable<SyntaxTrivia> items) =>
            Microsoft.CodeAnalysis.CSharp.SyntaxFactory.TriviaList(items);

        private static SyntaxList<StatementSyntax> SyntaxList(IEnumerable<StatementSyntax> stmts) =>
            Microsoft.CodeAnalysis.CSharp.SyntaxFactory.List(stmts);
    }
}
