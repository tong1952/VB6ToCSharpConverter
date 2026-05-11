using Microsoft.CodeAnalysis.CSharp;
using Microsoft.CodeAnalysis.CSharp.Syntax;
using static Microsoft.CodeAnalysis.CSharp.SyntaxFactory;

namespace VB6ToCSharp.Transformer;

/// <summary>
/// Maps VB6 built-in functions to C# .NET 8 equivalents.
/// Returns null when the caller should emit a regular call (unknown function).
/// </summary>
public static class BuiltinMapper
{
    /// <summary>
    /// Try to convert a VB6 built-in call to a C# expression.
    /// Returns null if not a recognised built-in (caller emits a plain method call).
    /// </summary>
    public static ExpressionSyntax? TryMap(
        string funcName,
        IReadOnlyList<ExpressionSyntax> args)
    {
        return funcName.ToUpperInvariant() switch
        {
            // ── String functions ──────────────────────────────────────────
            "LEN"   => OneArg(args, a => Member(a, "Length")),
            "UCASE" => OneArg(args, a => Invoke(Member(a, "ToUpper"))),
            "LCASE" => OneArg(args, a => Invoke(Member(a, "ToLower"))),
            "TRIM"  => OneArg(args, a => Invoke(Member(a, "Trim"))),
            "LTRIM" => OneArg(args, a => Invoke(Member(a, "TrimStart"))),
            "RTRIM" => OneArg(args, a => Invoke(Member(a, "TrimEnd"))),

            "LEFT"  => TwoArgs(args, (s, n) =>
                Invoke(Member(s, "Substring"), Literal(0), n)),

            "RIGHT" => TwoArgs(args, (s, n) =>
                Invoke(Member(s, "Substring"),
                    BinaryExpression(SyntaxKind.SubtractExpression, Member(s, "Length"), n))),

            "MID" => args.Count >= 3
                ? Call("Microsoft.VisualBasic.Strings.Mid", args)
                : TwoArgs(args, (s, start) =>
                    Invoke(Member(s, "Substring"),
                        BinaryExpression(SyntaxKind.SubtractExpression, start, Literal(1)))),

            "INSTR" => args.Count == 2
                ? TwoArgs(args, (haystack, needle) =>
                    BinaryExpression(SyntaxKind.AddExpression,
                        Invoke(Member(haystack, "IndexOf"), needle,
                            MemberAccessExpression(SyntaxKind.SimpleMemberAccessExpression,
                                IdentifierName("StringComparison"), IdentifierName("Ordinal"))),
                        Literal(1)))
                : null,

            "INSTRREV" => null, // complex — emit as Strings.InStrRev

            "REPLACE" => args.Count >= 3
                ? ThreeArgs(args, (s, find, repl) =>
                    Invoke(Member(s, "Replace"), find, repl))
                : null,

            "SPLIT" => TwoArgs(args, (s, delim) =>
                Invoke(Member(s, "Split"), delim)),

            "JOIN" => TwoArgs(args, (arr, delim) =>
                Invoke(StaticMember("string", "Join"), delim, arr)),

            "STRING" => TwoArgs(args, (count, ch) =>
                ObjectCreation("string", ch, count)),

            "SPACE" => OneArg(args, n =>
                ObjectCreation("string", Literal(' '), n)),

            "ASC"  => OneArg(args, s => Cast("short",
                Invoke(Member(s, "FirstOrDefault")))),

            "CHR"  => OneArg(args, n => Cast("char", n)),

            "HEX"  => OneArg(args, n =>
                Invoke(StaticMember("Convert", "ToString"), n, Literal(16))),

            "OCT"  => OneArg(args, n =>
                Invoke(StaticMember("Convert", "ToString"), n, Literal(8))),

            // ── Conversion functions ──────────────────────────────────────
            "CSTR"  => OneArg(args, a => Invoke(Member(a, "ToString"))),
            "CINT"  => OneArg(args, a => Cast("short", a)),
            "CLNG"  => OneArg(args, a => Cast("int", a)),
            "CSNG"  => OneArg(args, a => Cast("float", a)),
            "CDBL"  => OneArg(args, a => Cast("double", a)),
            "CCUR"  => OneArg(args, a => Cast("decimal", a)),
            "CBOOL" => OneArg(args, a => Cast("bool", a)),
            "CBYTE" => OneArg(args, a => Cast("byte", a)),
            "CDATE" => OneArg(args, a =>
                Invoke(StaticMember("DateTime", "Parse"),
                    Invoke(Member(a, "ToString")))),
            "CVAR"  => OneArg(args, a => a), // identity — already object
            "VAL"   => OneArg(args, a =>
                Invoke(StaticMember("double", "Parse"),
                    Invoke(Member(a, "ToString")))),
            "STR"   => OneArg(args, a => Invoke(Member(a, "ToString"))),

            // ── Math functions ────────────────────────────────────────────
            "ABS"   => OneArg(args, a => Invoke(StaticMember("Math", "Abs"), a)),
            "INT"   => OneArg(args, a => Invoke(StaticMember("Math", "Floor"), a)),
            "FIX"   => OneArg(args, a => Invoke(StaticMember("Math", "Truncate"), a)),
            "RND"   => OneArg(args, a => Invoke(StaticMember("Random", "Shared.NextDouble"))),
            "SGN"   => OneArg(args, a => Invoke(StaticMember("Math", "Sign"), a)),
            "SQR"   => OneArg(args, a => Invoke(StaticMember("Math", "Sqrt"), a)),
            "LOG"   => OneArg(args, a => Invoke(StaticMember("Math", "Log"), a)),
            "EXP"   => OneArg(args, a => Invoke(StaticMember("Math", "Exp"), a)),
            "SIN"   => OneArg(args, a => Invoke(StaticMember("Math", "Sin"), a)),
            "COS"   => OneArg(args, a => Invoke(StaticMember("Math", "Cos"), a)),
            "TAN"   => OneArg(args, a => Invoke(StaticMember("Math", "Tan"), a)),
            "ATN"   => OneArg(args, a => Invoke(StaticMember("Math", "Atan"), a)),

            // ── Array / collection helpers ────────────────────────────────
            "UBOUND" => args.Count == 1
                ? OneArg(args, a =>
                    BinaryExpression(SyntaxKind.SubtractExpression,
                        Member(a, "Length"), Literal(1)))
                : null,
            "LBOUND" => Literal(0),
            "ARRAY"  => null, // new object[] { ... } — handled in transformer

            // ── Type checking ─────────────────────────────────────────────
            "ISNUMERIC" => OneArg(args, a =>
                Invoke(StaticMember("double", "TryParse"),
                    Invoke(Member(a, "ToString")),
                    OutDiscard())),
            "ISNULL"  => OneArg(args, a =>
                BinaryExpression(SyntaxKind.EqualsExpression, a,
                    LiteralExpression(SyntaxKind.NullLiteralExpression))),
            "ISEMPTY" => OneArg(args, a =>
                BinaryExpression(SyntaxKind.EqualsExpression, a,
                    LiteralExpression(SyntaxKind.NullLiteralExpression))),
            "ISOBJECT" => OneArg(args, a =>
                BinaryExpression(SyntaxKind.NotEqualsExpression, a,
                    LiteralExpression(SyntaxKind.NullLiteralExpression))),
            "ISARRAY" => OneArg(args, a =>
                IsPatternExpression(a,
                    TypePattern(ArrayType(PredefinedType(Token(SyntaxKind.ObjectKeyword)),
                        SingletonList(ArrayRankSpecifier()))))),

            // ── Date / time ───────────────────────────────────────────────
            "NOW"   => args.Count == 0 ? StaticMember("DateTime", "Now") : null,
            "DATE"  => args.Count == 0 ? StaticMember("DateTime", "Today") : null,
            "TIME"  => args.Count == 0 ? StaticMember("DateTime", "Now") : null,
            "YEAR"  => OneArg(args, d => Member(d, "Year")),
            "MONTH" => OneArg(args, d => Member(d, "Month")),
            "DAY"   => OneArg(args, d => Member(d, "Day")),
            "HOUR"  => OneArg(args, d => Member(d, "Hour")),
            "MINUTE"=> OneArg(args, d => Member(d, "Minute")),
            "SECOND"=> OneArg(args, d => Member(d, "Second")),
            "WEEKDAY" => OneArg(args, d =>
                Cast("int", Member(Member(d, "DayOfWeek")))),
            "DATEADD" => args.Count >= 3
                ? DateAdd(args)
                : null,
            "DATEDIFF" => args.Count >= 3
                ? DateDiff(args)
                : null,
            "DATEPART" => null, // complex — left as-is

            // ── File I/O helpers ──────────────────────────────────────────
            "FREEFILE" => Literal(1), // returns next free file number; 1 is a placeholder
            // EOF(n) / LOF(n): n is the file variable (StreamReader/Writer) in translated code
            "EOF" => OneArg(args, n => Member(n, "EndOfStream")),
            "LOF" => OneArg(args, n => Member(Member(n, "BaseStream"), "Length")),
            "FILELEN" => OneArg(args, path =>
                Member(ObjectCreation("FileInfo", path), "Length")),
            "DIR" => args.Count >= 1
                ? OneArg(args, path => Invoke(StaticMember("Directory", "Exists"), path))
                : StaticMember("string", "Empty"),
            "CURDIR" => args.Count == 0
                ? Invoke(StaticMember("Directory", "GetCurrentDirectory"))
                : null,

            // ── I/O (approximate) ─────────────────────────────────────────
            "MSGBOX" => args.Count >= 1
                ? Invoke(StaticMember("Console", "WriteLine"), args[0])
                : null,
            "INPUTBOX" => args.Count >= 1
                ? Invoke(StaticMember("Console", "ReadLine"))
                : null,

            // ── Debug ─────────────────────────────────────────────────────
            "DEBUG.PRINT" => args.Count >= 1
                ? Invoke(StaticMember("Debug", "WriteLine"), args[0])
                : null,

            // ── Environment ───────────────────────────────────────────────
            "ENVIRON" => OneArg(args, s =>
                Invoke(StaticMember("Environment", "GetEnvironmentVariable"), s)),

            _ => null
        };
    }

    // ── Private helpers ──────────────────────────────────────────────────

    private static ExpressionSyntax? OneArg(
        IReadOnlyList<ExpressionSyntax> args,
        Func<ExpressionSyntax, ExpressionSyntax> build)
        => args.Count >= 1 ? build(args[0]) : null;

    private static ExpressionSyntax? TwoArgs(
        IReadOnlyList<ExpressionSyntax> args,
        Func<ExpressionSyntax, ExpressionSyntax, ExpressionSyntax> build)
        => args.Count >= 2 ? build(args[0], args[1]) : null;

    private static ExpressionSyntax? ThreeArgs(
        IReadOnlyList<ExpressionSyntax> args,
        Func<ExpressionSyntax, ExpressionSyntax, ExpressionSyntax, ExpressionSyntax> build)
        => args.Count >= 3 ? build(args[0], args[1], args[2]) : null;

    private static ExpressionSyntax Member(ExpressionSyntax obj, string name) =>
        MemberAccessExpression(SyntaxKind.SimpleMemberAccessExpression, obj, IdentifierName(name));

    private static ExpressionSyntax Member(ExpressionSyntax obj) => obj;

    private static ExpressionSyntax StaticMember(string type, string member)
    {
        // Handle dotted members like "Random.Shared.NextDouble"
        var parts = (type + "." + member).Split('.');
        ExpressionSyntax expr = IdentifierName(parts[0]);
        for (int i = 1; i < parts.Length; i++)
            expr = MemberAccessExpression(SyntaxKind.SimpleMemberAccessExpression,
                expr, IdentifierName(parts[i]));
        return expr;
    }

    private static InvocationExpressionSyntax Invoke(ExpressionSyntax target, params ExpressionSyntax[] args) =>
        InvocationExpression(target,
            ArgumentList(SeparatedList(args.Select(Argument))));

    private static ExpressionSyntax Call(string fullName, IReadOnlyList<ExpressionSyntax> args)
    {
        var parts = fullName.Split('.');
        ExpressionSyntax expr = IdentifierName(parts[0]);
        for (int i = 1; i < parts.Length; i++)
            expr = MemberAccessExpression(SyntaxKind.SimpleMemberAccessExpression,
                expr, IdentifierName(parts[i]));
        return InvocationExpression(expr,
            ArgumentList(SeparatedList(args.Select(Argument))));
    }

    private static ExpressionSyntax Cast(string type, ExpressionSyntax expr) =>
        CastExpression(ParseTypeName(type), expr);

    private static LiteralExpressionSyntax Literal(int value) =>
        LiteralExpression(SyntaxKind.NumericLiteralExpression,
            Microsoft.CodeAnalysis.CSharp.SyntaxFactory.Literal(value));

    private static LiteralExpressionSyntax Literal(char value) =>
        LiteralExpression(SyntaxKind.CharacterLiteralExpression,
            Microsoft.CodeAnalysis.CSharp.SyntaxFactory.Literal(value));

    private static ExpressionSyntax ObjectCreation(string type, params ExpressionSyntax[] args) =>
        ObjectCreationExpression(IdentifierName(type))
            .WithArgumentList(ArgumentList(SeparatedList(args.Select(Argument))));

    private static ExpressionSyntax OutDiscard() =>
        Argument(DeclarationExpression(
            IdentifierName("_"),
            SingleVariableDesignation(Identifier("_"))))
        .WithRefKindKeyword(Token(SyntaxKind.OutKeyword))
        .Expression;

    private static ExpressionSyntax? DateAdd(IReadOnlyList<ExpressionSyntax> args)
    {
        // DateAdd(interval, number, date) → date.Add(TimeSpan.FromX(number))
        // Simplified: always AddDays
        return BinaryExpression(SyntaxKind.AddExpression,
            args[2],
            Invoke(StaticMember("TimeSpan", "FromDays"), args[1]));
    }

    private static ExpressionSyntax? DateDiff(IReadOnlyList<ExpressionSyntax> args)
    {
        // DateDiff(interval, date1, date2) → (date2 - date1).TotalDays
        return Member(
            ParenthesizedExpression(
                BinaryExpression(SyntaxKind.SubtractExpression, args[2], args[1])),
            "TotalDays");
    }
}
