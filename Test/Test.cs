using VB6ToCSharp;

string[] exts = ["*.bas", "*.cls", "*.frm"];
var files = exts
    .SelectMany(e => Directory.GetFiles("samples", e, SearchOption.AllDirectories))
    .OrderBy(f => f)
    .ToArray();

int totalTodos = 0, totalWarnings = 0;
foreach (var f in files)
{
    Console.Write($"  {f.Replace("samples\\", "")}... ");
    Console.Out.Flush();
    string src = File.ReadAllText(f, System.Text.Encoding.Default);
    src = ConversionEngine.StripRcsIfNeeded(src);
    Console.Write("lex... "); Console.Out.Flush();
    var lexer = new VB6ToCSharp.Lexer.VB6Lexer(src);
    var tokens = lexer.Tokenize();
    Console.Write($"({tokens.Count} tokens) parse... "); Console.Out.Flush();
    var parser = new VB6ToCSharp.Parser.VB6Parser(tokens);
    var module = parser.Parse(Path.GetFileNameWithoutExtension(f), VB6ToCSharp.Parser.ModuleKind.Standard);
    Console.Write("transform... "); Console.Out.Flush();
    var (cs, diags) = ConversionEngine.Convert(src, Path.GetFileNameWithoutExtension(f), f, "Converted");
    Console.WriteLine("done");
    int todos = cs.Split("// TODO:").Length - 1;
    totalTodos += todos;
    totalWarnings += diags.Count;
    if (todos > 0 || diags.Count > 0)
    {
        Console.WriteLine($"{Path.GetFileName(f)}: {todos} TODO(s), {diags.Count} warning(s)");
        foreach (var line in cs.Split('\n').Where(l => l.Contains("// TODO:")).Take(3))
            Console.WriteLine($"  {line.Trim()}");
        foreach (var d in diags.Take(5))
            Console.WriteLine($"  WARN: {d}");
    }
}
Console.WriteLine($"\nTotal: {files.Length} files, {totalTodos} TODO(s), {totalWarnings} warning(s)");
return totalTodos > 0 ? 1 : 0;
