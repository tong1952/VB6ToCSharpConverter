using Microsoft.CodeAnalysis.CSharp;
using Microsoft.CodeAnalysis.CSharp.Syntax;
using static Microsoft.CodeAnalysis.CSharp.SyntaxFactory;

namespace VB6ToCSharp.Transformer;

/// <summary>
/// Maps VB6 type names to their C# .NET 8 equivalents.
/// </summary>
public static class TypeMapper
{
    private static readonly Dictionary<string, string> PrimitiveMap =
        new(StringComparer.OrdinalIgnoreCase)
        {
            ["Integer"] = "short",
            ["Long"]    = "int",
            ["LongLong"] = "long",
            ["Single"]  = "float",
            ["Double"]  = "double",
            ["Currency"] = "decimal",
            ["Decimal"] = "decimal",
            ["Byte"]    = "byte",
            ["Boolean"] = "bool",
            ["String"]  = "string",
            ["Date"]    = "DateTime",
            ["Variant"] = "object",
            ["Object"]  = "object",
            ["Nothing"] = "object",
            ["Empty"]   = "object",
            ["Null"]    = "object",
            ["Any"]     = "IntPtr",   // Declare Lib Any
        };

    public static TypeSyntax ToCSharp(string vb6Type, bool isArray = false)
    {
        string baseType = PrimitiveMap.TryGetValue(vb6Type, out var mapped) ? mapped : vb6Type;
        TypeSyntax ts = baseType switch
        {
            "short"    => PredefinedType(Token(SyntaxKind.ShortKeyword)),
            "int"      => PredefinedType(Token(SyntaxKind.IntKeyword)),
            "long"     => PredefinedType(Token(SyntaxKind.LongKeyword)),
            "float"    => PredefinedType(Token(SyntaxKind.FloatKeyword)),
            "double"   => PredefinedType(Token(SyntaxKind.DoubleKeyword)),
            "decimal"  => PredefinedType(Token(SyntaxKind.DecimalKeyword)),
            "byte"     => PredefinedType(Token(SyntaxKind.ByteKeyword)),
            "bool"     => PredefinedType(Token(SyntaxKind.BoolKeyword)),
            "string"   => PredefinedType(Token(SyntaxKind.StringKeyword)),
            "object"   => PredefinedType(Token(SyntaxKind.ObjectKeyword)),
            "void"     => PredefinedType(Token(SyntaxKind.VoidKeyword)),
            _ => IdentifierName(SanitizeName(baseType))
        };

        return isArray ? ArrayType(ts, SingletonList(ArrayRankSpecifier())) : ts;
    }

    public static string SanitizeName(string name)
    {
        // Avoid C# keyword collisions
        var reserved = new HashSet<string>(StringComparer.Ordinal)
        {
            "class", "base", "object", "string", "int", "long", "float", "double",
            "bool", "byte", "short", "decimal", "void", "event", "type", "is",
            "operator", "params", "ref", "out", "checked", "lock", "switch",
            "delegate", "interface", "namespace", "abstract", "override", "virtual",
            "sealed", "readonly", "static", "const", "new", "null", "true", "false",
        };

        if (reserved.Contains(name, StringComparer.OrdinalIgnoreCase))
            return "@" + name;

        return name;
    }

    public static bool IsValueType(string vb6Type) =>
        PrimitiveMap.TryGetValue(vb6Type, out var cs) && cs is
            "short" or "int" or "long" or "float" or "double" or "decimal" or
            "byte" or "bool" or "DateTime";

    public static ExpressionSyntax DefaultValue(string vb6Type)
    {
        if (!PrimitiveMap.TryGetValue(vb6Type, out var cs)) cs = vb6Type;
        return cs switch
        {
            "bool" => LiteralExpression(SyntaxKind.FalseLiteralExpression),
            "string" => LiteralExpression(SyntaxKind.StringLiteralExpression, Literal("")),
            "object" => LiteralExpression(SyntaxKind.NullLiteralExpression),
            "DateTime" => MemberAccessExpression(SyntaxKind.SimpleMemberAccessExpression,
                IdentifierName("DateTime"), IdentifierName("MinValue")),
            _ => LiteralExpression(SyntaxKind.NumericLiteralExpression, Literal(0))
        };
    }
}
