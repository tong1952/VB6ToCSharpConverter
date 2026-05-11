using System.Globalization;
using System.Text;

namespace VB6ToCSharp.Lexer;

public class VB6Lexer(string source)
{
    private int _pos;
    private int _line = 1;
    private int _column = 1;

    private static readonly Dictionary<string, TokenKind> Keywords =
        new(StringComparer.OrdinalIgnoreCase)
        {
            ["Dim"] = TokenKind.Dim,
            ["As"] = TokenKind.As,
            ["Sub"] = TokenKind.Sub,
            ["Function"] = TokenKind.Function,
            ["Property"] = TokenKind.Property,
            ["Get"] = TokenKind.Get,
            ["Let"] = TokenKind.Let,
            ["Set"] = TokenKind.Set,
            ["If"] = TokenKind.If,
            ["Then"] = TokenKind.Then,
            ["Else"] = TokenKind.Else,
            ["ElseIf"] = TokenKind.ElseIf,
            ["End"] = TokenKind.End,
            ["For"] = TokenKind.For,
            ["Each"] = TokenKind.Each,
            ["In"] = TokenKind.In,
            ["Next"] = TokenKind.Next,
            ["While"] = TokenKind.While,
            ["Wend"] = TokenKind.Wend,
            ["Do"] = TokenKind.Do,
            ["Loop"] = TokenKind.Loop,
            ["Until"] = TokenKind.Until,
            ["To"] = TokenKind.To,
            ["Step"] = TokenKind.Step,
            ["Select"] = TokenKind.Select,
            ["Case"] = TokenKind.Case,
            ["Is"] = TokenKind.Is,
            ["With"] = TokenKind.With,
            ["Exit"] = TokenKind.Exit,
            ["GoTo"] = TokenKind.GoTo,
            ["GoSub"] = TokenKind.GoSub,
            ["Return"] = TokenKind.Return,
            ["On"] = TokenKind.On,
            ["Error"] = TokenKind.Error,
            ["Resume"] = TokenKind.Resume,
            ["RaiseEvent"] = TokenKind.RaiseEvent,
            ["Event"] = TokenKind.Event,
            ["ReDim"] = TokenKind.ReDim,
            ["Erase"] = TokenKind.Erase,
            ["Preserve"] = TokenKind.Preserve,
            ["New"] = TokenKind.New,
            ["Nothing"] = TokenKind.Nothing,
            ["Null"] = TokenKind.Null,
            ["Empty"] = TokenKind.Empty,
            ["True"] = TokenKind.True,
            ["False"] = TokenKind.False,
            ["Not"] = TokenKind.Not,
            ["And"] = TokenKind.And,
            ["Or"] = TokenKind.Or,
            ["Xor"] = TokenKind.Xor,
            ["Mod"] = TokenKind.Mod,
            ["Eqv"] = TokenKind.Eqv,
            ["Imp"] = TokenKind.Imp,
            ["Like"] = TokenKind.Like,
            ["TypeOf"] = TokenKind.TypeOf,
            ["AddressOf"] = TokenKind.AddressOf,
            ["Public"] = TokenKind.Public,
            ["Private"] = TokenKind.Private,
            ["Friend"] = TokenKind.Friend,
            ["Static"] = TokenKind.Static,
            ["ByRef"] = TokenKind.ByRef,
            ["ByVal"] = TokenKind.ByVal,
            ["Optional"] = TokenKind.Optional,
            ["ParamArray"] = TokenKind.ParamArray,
            ["Class"] = TokenKind.Class,
            ["Module"] = TokenKind.Module,
            ["Type"] = TokenKind.Type,
            ["Const"] = TokenKind.Const,
            ["Enum"] = TokenKind.Enum,
            ["Call"] = TokenKind.Call,
            ["Me"] = TokenKind.Me,
            ["MyBase"] = TokenKind.MyBase,
            ["MyClass"] = TokenKind.MyClass,
            ["Implements"] = TokenKind.Implements,
            ["Attribute"] = TokenKind.Attribute,
            ["Option"] = TokenKind.Option,
            ["Explicit"] = TokenKind.Explicit,
            ["Base"] = TokenKind.Base,
            ["Compare"] = TokenKind.Compare,
            ["Declare"] = TokenKind.Declare,
            ["Lib"] = TokenKind.Lib,
            ["Alias"] = TokenKind.Alias,
            ["WithEvents"] = TokenKind.WithEvents,
            ["String"] = TokenKind.StringType,
            ["Object"] = TokenKind.ObjectType,
            ["Step"] = TokenKind.Step,
        };

    private char Current => _pos < source.Length ? source[_pos] : '\0';
    private char PeekAt(int offset) => _pos + offset < source.Length ? source[_pos + offset] : '\0';

    private char Advance()
    {
        char ch = Current;
        _pos++;
        if (ch == '\n') { _line++; _column = 1; }
        else _column++;
        return ch;
    }

    public List<Token> Tokenize()
    {
        var tokens = new List<Token>();
        while (true)
        {
            var tok = NextToken();
            tokens.Add(tok);
            if (tok.Kind == TokenKind.EndOfFile) break;
        }
        return tokens;
    }

    private Token NextToken()
    {
        while (Current is ' ' or '\t') Advance();

        if (_pos >= source.Length)
            return MakeAt(TokenKind.EndOfFile, "", null, _line, _column);

        int startLine = _line, startCol = _column;
        char ch = Current;

        // Comment
        if (ch == '\'')
        {
            while (_pos < source.Length && Current != '\n' && Current != '\r') Advance();
            return NextToken();
        }

        // Rem keyword as comment (only when at statement start — approximated here)
        if (ch == 'R' && IsWordAt("Rem") && (PeekAt(3) is ' ' or '\t' or '\r' or '\n' or '\0'))
        {
            while (_pos < source.Length && Current != '\n' && Current != '\r') Advance();
            return NextToken();
        }

        // Newlines
        if (ch == '\r' || ch == '\n')
        {
            if (ch == '\r' && PeekAt(1) == '\n') Advance();
            Advance();
            return MakeAt(TokenKind.NewLine, "\n", null, startLine, startCol);
        }

        // Line continuation: _ followed by optional spaces, optional ' comment, then newline
        if (ch == '_')
        {
            int ahead = 1;
            while (_pos + ahead < source.Length && source[_pos + ahead] is ' ' or '\t') ahead++;
            char next = _pos + ahead < source.Length ? source[_pos + ahead] : '\0';
            if (next == '\'')
            {
                // skip past the comment to the newline, then treat as continuation
                while (_pos + ahead < source.Length && source[_pos + ahead] is not '\r' and not '\n')
                    ahead++;
                next = _pos + ahead < source.Length ? source[_pos + ahead] : '\0';
            }
            if (next is '\r' or '\n' or '\0')
            {
                while (Current != '\n' && Current != '\r' && _pos < source.Length) Advance();
                if (Current == '\r') Advance();
                if (Current == '\n') Advance();
                return NextToken();
            }
        }

        // String literal
        if (ch == '"') return ReadString(startLine, startCol);

        // Date literal
        if (ch == '#') return ReadDate(startLine, startCol);

        // Hex literal: &H...
        if (ch == '&' && (PeekAt(1) is 'H' or 'h')) return ReadHex(startLine, startCol);

        // Octal literal: &O...
        if (ch == '&' && (PeekAt(1) is 'O' or 'o')) return ReadOctal(startLine, startCol);

        // Number
        if (char.IsDigit(ch) || (ch == '.' && char.IsDigit(PeekAt(1))))
            return ReadNumber(startLine, startCol);

        // Identifier / keyword
        if (char.IsLetter(ch) || ch == '_')
            return ReadIdentifier(startLine, startCol);

        // Two-character operators
        Advance();
        if (ch == '<')
        {
            if (Current == '>') { Advance(); return MakeAt(TokenKind.NotEqual, "<>", null, startLine, startCol); }
            if (Current == '=') { Advance(); return MakeAt(TokenKind.LessEqual, "<=", null, startLine, startCol); }
            return MakeAt(TokenKind.LessThan, "<", null, startLine, startCol);
        }
        if (ch == '>')
        {
            if (Current == '=') { Advance(); return MakeAt(TokenKind.GreaterEqual, ">=", null, startLine, startCol); }
            return MakeAt(TokenKind.GreaterThan, ">", null, startLine, startCol);
        }

        // Single-character tokens
        var kind = ch switch
        {
            '+' => TokenKind.Plus,
            '-' => TokenKind.Minus,
            '*' => TokenKind.Star,
            '/' => TokenKind.Slash,
            '\\' => TokenKind.BackSlash,
            '^' => TokenKind.Caret,
            '&' => TokenKind.Ampersand,
            '=' => TokenKind.Equals,
            '(' => TokenKind.LeftParen,
            ')' => TokenKind.RightParen,
            ',' => TokenKind.Comma,
            ':' => TokenKind.Colon,
            '.' => TokenKind.Dot,
            '!' => TokenKind.Bang,
            _ => (TokenKind)(-1)
        };

        if ((int)kind != -1)
            return MakeAt(kind, ch.ToString(), null, startLine, startCol);

        return NextToken(); // skip unrecognised
    }

    private bool IsWordAt(string word)
    {
        for (int i = 0; i < word.Length; i++)
            if (_pos + i >= source.Length || char.ToUpperInvariant(source[_pos + i]) != char.ToUpperInvariant(word[i]))
                return false;
        return true;
    }

    private Token ReadString(int line, int col)
    {
        Advance(); // "
        var sb = new StringBuilder();
        while (_pos < source.Length && Current != '\n' && Current != '\r')
        {
            if (Current == '"')
            {
                Advance();
                if (Current == '"') { sb.Append('"'); Advance(); }
                else break;
            }
            else sb.Append(Advance());
        }
        return MakeAt(TokenKind.StringLiteral, sb.ToString(), sb.ToString(), line, col);
    }

    private Token ReadDate(int line, int col)
    {
        Advance(); // consume leading #

        // Lookahead: does a closing '#' appear before the end of this line?
        // If yes → real date literal (#12/31/1999#).
        // If no  → VB6 file-channel reference (#1, #fileNum) — read only the
        //           identifier/number chars so the trailing comma etc. remain tokens.
        bool hasClosingHash = false;
        for (int i = 0; _pos + i < source.Length; i++)
        {
            char c = source[_pos + i];
            if (c == '\n' || c == '\r') break;
            if (c == '#') { hasClosingHash = true; break; }
        }

        var sb = new StringBuilder();
        if (hasClosingHash)
        {
            while (_pos < source.Length && Current != '#' && Current != '\n')
                sb.Append(Advance());
            if (Current == '#') Advance();
        }
        else
        {
            // File-channel number: read only word characters (digit, letter, underscore)
            while (_pos < source.Length && (char.IsLetterOrDigit(Current) || Current == '_'))
                sb.Append(Advance());
        }

        var ds = sb.ToString().Trim();
        return MakeAt(TokenKind.DateLiteral, ds, ds, line, col);
    }

    private Token ReadHex(int line, int col)
    {
        Advance(); Advance(); // &H
        var sb = new StringBuilder();
        while (IsHexDigit(Current)) sb.Append(Advance());
        if (Current == '&') Advance();
        long val = sb.Length > 0 ? Convert.ToInt64(sb.ToString(), 16) : 0;
        return MakeAt(TokenKind.LongLiteral, "&H" + sb, val, line, col);
    }

    private Token ReadOctal(int line, int col)
    {
        Advance(); Advance(); // &O
        var sb = new StringBuilder();
        while (Current >= '0' && Current <= '7') sb.Append(Advance());
        long val = sb.Length > 0 ? Convert.ToInt64(sb.ToString(), 8) : 0;
        return MakeAt(TokenKind.LongLiteral, "&O" + sb, val, line, col);
    }

    private Token ReadNumber(int line, int col)
    {
        var sb = new StringBuilder();
        bool isFloat = false;

        while (char.IsDigit(Current)) sb.Append(Advance());

        if (Current == '.' && char.IsDigit(PeekAt(1)))
        {
            isFloat = true;
            sb.Append(Advance());
            while (char.IsDigit(Current)) sb.Append(Advance());
        }

        if (Current is 'E' or 'e')
        {
            isFloat = true;
            sb.Append(Advance());
            if (Current is '+' or '-') sb.Append(Advance());
            while (char.IsDigit(Current)) sb.Append(Advance());
        }

        char suffix = Current;
        if (suffix is '%' or '&' or '!' or '#' or '@') Advance();

        var text = sb.ToString();

        if (!isFloat && suffix is not '!' and not '#')
        {
            if (long.TryParse(text, out long lv))
            {
                if (suffix != '&' && lv is >= short.MinValue and <= short.MaxValue)
                    return MakeAt(TokenKind.IntegerLiteral, text, (short)lv, line, col);
                return MakeAt(TokenKind.LongLiteral, text, lv, line, col);
            }
        }

        if (suffix == '!' && float.TryParse(text, NumberStyles.Float, CultureInfo.InvariantCulture, out float fv))
            return MakeAt(TokenKind.SingleLiteral, text, fv, line, col);

        if (double.TryParse(text, NumberStyles.Float, CultureInfo.InvariantCulture, out double dv))
            return MakeAt(TokenKind.DoubleLiteral, text, dv, line, col);

        return MakeAt(TokenKind.DoubleLiteral, text, 0.0, line, col);
    }

    private Token ReadIdentifier(int line, int col)
    {
        var sb = new StringBuilder();
        while (char.IsLetterOrDigit(Current) || Current == '_') sb.Append(Advance());
        if (Current is '%' or '&' or '!' or '#' or '$' or '@') Advance(); // type-declaration char
        var text = sb.ToString();
        if (Keywords.TryGetValue(text, out var kw))
            return MakeAt(kw, text, null, line, col);
        return MakeAt(TokenKind.Identifier, text, null, line, col);
    }

    private static bool IsHexDigit(char c) =>
        char.IsDigit(c) || c is >= 'A' and <= 'F' || c is >= 'a' and <= 'f';

    private static Token MakeAt(TokenKind kind, string text, object? value, int line, int col) =>
        new(kind, text, value, line, col);
}
