using VB6ToCSharp.Lexer;

namespace VB6ToCSharp.Parser;

/// <summary>
/// Recursive-descent parser for VB6 source.
/// Produces a VB6Module AST covering the full VB6 language grammar.
/// </summary>
public class VB6Parser(List<Token> tokens)
{
    private int _pos;
    private readonly List<string> _diagnostics = [];

    public IReadOnlyList<string> Diagnostics => _diagnostics;

    // ─── Token navigation ─────────────────────────────────────────────────

    private Token Current => _pos < tokens.Count ? tokens[_pos] : tokens[^1];
    private Token Peek(int offset = 1) =>
        _pos + offset < tokens.Count ? tokens[_pos + offset] : tokens[^1];

    private Token Advance()
    {
        var t = Current;
        if (_pos < tokens.Count - 1) _pos++;
        return t;
    }

    private bool Check(TokenKind kind) => Current.Kind == kind;

    private bool Match(TokenKind kind)
    {
        if (Current.Kind == kind) { Advance(); return true; }
        return false;
    }

    private Token Consume(TokenKind kind, string ctx = "")
    {
        if (Current.Kind == kind) return Advance();
        _diagnostics.Add($"Line {Current.Line}: expected {kind}{(ctx.Length > 0 ? " in " + ctx : "")}, got '{Current.Text}'");
        return Current;
    }

    private void SkipNewlines() { while (Check(TokenKind.NewLine)) Advance(); }

    private void EndLine()
    {
        while (Check(TokenKind.Colon)) Advance();
        Match(TokenKind.NewLine);
    }

    private bool IsEol() =>
        Current.Kind is TokenKind.NewLine or TokenKind.EndOfFile or TokenKind.Colon;

    private void SkipRestOfLine()
    {
        while (!IsEol()) Advance();
        EndLine();
    }

    // Checks if current+next form "End <keyword>"
    private bool IsEnd(string keyword) =>
        Check(TokenKind.End) && string.Equals(Peek().Text, keyword, StringComparison.OrdinalIgnoreCase);

    // ─── Entry point ──────────────────────────────────────────────────────

    public VB6Module Parse(string moduleName, ModuleKind kind = ModuleKind.Standard)
    {
        var implements = new List<string>();
        var members = new List<VB6MemberNode>();

        while (!Check(TokenKind.EndOfFile))
        {
            SkipNewlines();
            if (Check(TokenKind.EndOfFile)) break;

            if (Check(TokenKind.Option) || Check(TokenKind.Attribute)) { SkipRestOfLine(); continue; }

            if (Check(TokenKind.Implements))
            {
                Advance();
                if (Current.Kind == TokenKind.Identifier || Current.Kind == TokenKind.ObjectType)
                    implements.Add(Advance().Text);
                EndLine();
                continue;
            }

            var member = ParseMember();
            if (member != null) members.Add(member);
        }

        return new VB6Module(kind, moduleName, implements, members);
    }

    // ─── Member parsing ───────────────────────────────────────────────────

    private VB6MemberNode? ParseMember()
    {
        var access = ParseAccess();
        bool isStatic = Match(TokenKind.Static);

        switch (Current.Kind)
        {
            case TokenKind.Declare:
                return ParseDeclare(access);

            case TokenKind.Sub:
                Advance();
                return ParseSubBody(access, isStatic);

            case TokenKind.Function:
                Advance();
                return ParseFunctionBody(access, isStatic);

            case TokenKind.Property:
                Advance();
                return ParsePropertyBody(access);

            case TokenKind.Event:
                Advance();
                var evtName = ExpectIdent();
                var evtParms = ParseParamList();
                EndLine();
                return new EventDecl(evtName, access, evtParms);

            case TokenKind.Type:
                Advance();
                return ParseUserType(access);

            case TokenKind.Enum:
                Advance();
                return ParseEnum(access);

            case TokenKind.Const:
                Advance();
                return ParseModuleConst(access);

            case TokenKind.Dim:
                Advance();
                var decls = ParseDeclaratorList();
                EndLine();
                return new FieldDecl(access, isStatic, decls);

            default:
                // Public/Private field without Dim keyword
                if (access != AccessModifier.Default || isStatic)
                {
                    var fDecls = ParseDeclaratorList();
                    EndLine();
                    return new FieldDecl(access, isStatic, fDecls);
                }
                // Unknown / skip
                if (!Check(TokenKind.EndOfFile))
                    SkipRestOfLine();
                return null;
        }
    }

    private SubDecl ParseSubBody(AccessModifier access, bool isStatic)
    {
        string name = ExpectIdent();
        var parms = ParseParamList();
        EndLine();
        var body = ParseBodyUntil("Sub");
        return new SubDecl(name, access, isStatic, parms, body);
    }

    private FunctionDecl ParseFunctionBody(AccessModifier access, bool isStatic)
    {
        string name = ExpectIdent();
        var parms = ParseParamList();
        var ret = new VB6TypeRef("Variant");
        if (Match(TokenKind.As)) ret = ParseTypeRef();
        EndLine();
        var body = ParseBodyUntil("Function");
        return new FunctionDecl(name, access, isStatic, parms, ret, body);
    }

    private PropertyDecl ParsePropertyBody(AccessModifier access)
    {
        PropertyKind pk = Current.Kind switch
        {
            TokenKind.Get => PropertyKind.Get,
            TokenKind.Let => PropertyKind.Let,
            TokenKind.Set => PropertyKind.Set,
            _ => PropertyKind.Get
        };
        Advance();
        string name = ExpectIdent();
        var parms = ParseParamList();
        var propType = new VB6TypeRef("Variant");
        if (Match(TokenKind.As)) propType = ParseTypeRef();
        EndLine();
        var body = ParseBodyUntil("Property");
        return new PropertyDecl(name, pk, access, parms, propType, body);
    }

    private DeclareDecl ParseDeclare(AccessModifier access)
    {
        Advance(); // Declare
        bool isFn = Match(TokenKind.Function);
        if (!isFn) Match(TokenKind.Sub);
        string name = ExpectIdent();
        Consume(TokenKind.Lib, "Declare");
        string lib = Current.Kind == TokenKind.StringLiteral
            ? (string)Advance().Value!
            : ExpectIdent();
        string? alias = null;
        if (Match(TokenKind.Alias))
            alias = Current.Kind == TokenKind.StringLiteral
                ? (string)Advance().Value!
                : ExpectIdent();
        var parms = ParseParamList();
        VB6TypeRef? ret = null;
        if (isFn && Match(TokenKind.As)) ret = ParseTypeRef();
        EndLine();
        return new DeclareDecl(name, access, lib, alias, parms, ret);
    }

    private UserTypeDecl ParseUserType(AccessModifier access)
    {
        string name = ExpectIdent();
        EndLine();
        var fields = new List<VB6Declarator>();
        while (!Check(TokenKind.EndOfFile) && !IsEnd("Type"))
        {
            SkipNewlines();
            if (IsEnd("Type")) break;
            int before = _pos;
            fields.Add(ParseDeclarator());
            EndLine();
            if (_pos == before) Advance(); // stall guard
        }
        if (Check(TokenKind.End)) { Advance(); Advance(); }
        EndLine();
        return new UserTypeDecl(name, access, fields, false);
    }

    private UserTypeDecl ParseEnum(AccessModifier access)
    {
        string name = ExpectIdent();
        EndLine();
        var fields = new List<VB6Declarator>();
        long counter = 0;
        while (!Check(TokenKind.EndOfFile) && !IsEnd("Enum"))
        {
            SkipNewlines();
            if (IsEnd("Enum")) break;
            int before = _pos;
            string fname = ExpectIdent();
            VB6Expression? valExpr = null;
            if (Match(TokenKind.Equals))
            {
                valExpr = ParseExpression();
                if (valExpr is LiteralExpr { Kind: LiteralKind.Integer or LiteralKind.Long } lit)
                    counter = Convert.ToInt64(lit.Value) + 1;
            }
            else
            {
                valExpr = new LiteralExpr(counter, LiteralKind.Long);
                counter++;
            }
            fields.Add(new VB6Declarator(fname, new VB6TypeRef("Long"), false, [valExpr]));
            EndLine();
            if (_pos == before) Advance(); // stall guard
        }
        if (Check(TokenKind.End)) { Advance(); Advance(); }
        EndLine();
        return new UserTypeDecl(name, access, fields, IsEnum: true);
    }

    private ConstMember ParseModuleConst(AccessModifier access)
    {
        string name = ExpectIdent();
        VB6TypeRef? type = null;
        if (Match(TokenKind.As)) type = ParseTypeRef();
        Consume(TokenKind.Equals);
        var val = ParseExpression();
        EndLine();
        return new ConstMember(name, type, val, access);
    }

    // ─── Statement body helpers ───────────────────────────────────────────

    private List<VB6Statement> ParseBodyUntil(string endKeyword)
    {
        var stmts = new List<VB6Statement>();
        while (!Check(TokenKind.EndOfFile) && !IsEnd(endKeyword))
        {
            SkipNewlines();
            if (Check(TokenKind.EndOfFile) || IsEnd(endKeyword)) break;
            int before = _pos;
            ParseStatements(stmts);
            if (_pos == before) Advance(); // stall guard: dangling token (e.g. unmatched ')')
        }
        if (Check(TokenKind.End)) { Advance(); Advance(); }
        EndLine();
        return stmts;
    }

    // Parses one logical line which may contain colon-separated statements
    private void ParseStatements(List<VB6Statement> stmts)
    {
        var s = ParseStatement();
        if (s is not EmptyStmt) stmts.Add(s);
        while (Check(TokenKind.Colon))
        {
            Advance();
            if (IsEol()) break;
            s = ParseStatement();
            if (s is not EmptyStmt) stmts.Add(s);
        }
        Match(TokenKind.NewLine);
    }

    // ─── Statements ───────────────────────────────────────────────────────

    private VB6Statement ParseStatement()
    {
        if (IsEol()) return new EmptyStmt();

        var tok = Current;

        // Label: Identifier (or keyword-as-ident such as "Error:") followed by ':'
        if ((tok.Kind == TokenKind.Identifier || IsKeywordAsIdent(tok.Kind)) && Peek().Kind == TokenKind.Colon
            && (Peek(2).Kind == TokenKind.NewLine || Peek(2).Kind == TokenKind.EndOfFile))
        {
            Advance(); Advance();
            return new LabelStmt(tok.Text);
        }

        // VB6 file I/O and file system context-keywords.
        // These are lexed as Identifier but have special statement syntax that conflicts
        // with keyword parsing (e.g. "Open ... For ... As" trips up the For-loop parser).
        if (tok.Kind == TokenKind.Identifier)
        {
            switch (tok.Text.ToUpperInvariant())
            {
                case "OPEN":  return ParseOpen();
                case "CLOSE": return ParseClose();
                case "NAME" when Peek().Kind != TokenKind.Equals: return ParseNameFile();
                case "KILL":  return ParseSingleArgFileOp(t => new KillStmt(t));
                case "MKDIR": return ParseSingleArgFileOp(t => new MkDirStmt(t));
                case "RMDIR": return ParseSingleArgFileOp(t => new RmDirStmt(t));
                case "CHDIR": return ParseSingleArgFileOp(t => new ChDirStmt(t));
                case "CHDRIVE": return ParseSingleArgFileOp(t => new ChDriveStmt(t));
                case "FILECOPY": return ParseFileCopy();
                // "Line Input" is two tokens; detect by lookahead
                case "LINE" when Peek().Kind == TokenKind.Identifier &&
                    string.Equals(Peek().Text, "Input", StringComparison.OrdinalIgnoreCase):
                    return ParseLineInput();
                // File-channel statements detected by "#n" (DateLiteral) immediately after keyword
                case "PRINT" when Peek().Kind == TokenKind.DateLiteral: return ParsePrintFile();
                case "WRITE" when Peek().Kind == TokenKind.DateLiteral: return ParseWriteFile();
                case "INPUT" when Peek().Kind == TokenKind.DateLiteral: return ParseInputFile();
                case "GET"   when Peek().Kind == TokenKind.DateLiteral: return ParseGetFile();
                case "PUT"   when Peek().Kind == TokenKind.DateLiteral: return ParsePutFile();
            }
        }

        // Get is a keyword token — handle file I/O Get #n separately
        if (tok.Kind == TokenKind.Get && Peek().Kind == TokenKind.DateLiteral)
            return ParseGetFile();

        switch (tok.Kind)
        {
            case TokenKind.If:        return ParseIf();
            case TokenKind.For:       return ParseFor();
            case TokenKind.While:     return ParseWhile();
            case TokenKind.Do:        return ParseDo();
            case TokenKind.Select:    return ParseSelect();
            case TokenKind.With:      return ParseWith();
            case TokenKind.Exit:      return ParseExit();
            case TokenKind.GoTo:      Advance(); var gtLbl = Current.Text; Advance(); return new GoToStmt(gtLbl);
            case TokenKind.GoSub:     Advance(); var gsLbl = Current.Text; Advance(); return new GoSubStmt(gsLbl);
            case TokenKind.Return:    Advance(); return new ReturnStmt();
            case TokenKind.On:        return ParseOnError();
            case TokenKind.Resume:    return ParseResume();
            case TokenKind.RaiseEvent: return ParseRaiseEvent();
            case TokenKind.ReDim:     return ParseReDim();
            case TokenKind.Erase:     return ParseErase();
            case TokenKind.Dim:       Advance(); return new LocalDimStmt(ParseDeclaratorList(), false);
            case TokenKind.Static:    Advance(); return new LocalDimStmt(ParseDeclaratorList(), true);
            case TokenKind.Const:     Advance(); return ParseLocalConst();
            case TokenKind.Call:      return ParseCall();
            case TokenKind.Set:       Advance(); return ParseSetAssign();
            default:                  return ParseAssignOrCall();
        }
    }

    // ─── If ───────────────────────────────────────────────────────────────

    private VB6Statement ParseIf()
    {
        Advance(); // If
        var cond = ParseExpression();
        Match(TokenKind.Then);

        // Single-line If
        if (!Check(TokenKind.NewLine) && !Check(TokenKind.EndOfFile) && !Check(TokenKind.Colon))
        {
            var s = ParseStatement();
            List<VB6Statement>? elseB = null;
            if (Match(TokenKind.Else)) elseB = [ParseStatement()];
            return new IfStmt(cond, [s], [], elseB);
        }
        EndLine();

        var then = new List<VB6Statement>();
        var elseIfs = new List<(VB6Expression, List<VB6Statement>)>();
        List<VB6Statement>? elseBody = null;

        while (!Check(TokenKind.EndOfFile) && !IsEnd("If"))
        {
            SkipNewlines();
            if (Check(TokenKind.ElseIf) || (Check(TokenKind.Else) && Peek().Kind == TokenKind.If))
            {
                Advance(); if (Check(TokenKind.If)) Advance();
                var ec = ParseExpression();
                Match(TokenKind.Then);
                EndLine();
                var eb = ParseBlockUntil(() => Check(TokenKind.Else) || Check(TokenKind.ElseIf) || IsEnd("If"));
                elseIfs.Add((ec, eb));
            }
            else if (Check(TokenKind.Else))
            {
                Advance(); EndLine();
                elseBody = ParseBlockUntil(() => IsEnd("If"));
                break;
            }
            else if (IsEnd("If")) break;
            else
            {
                ParseStatements(then);
            }
        }

        if (IsEnd("If")) { Advance(); Advance(); }
        return new IfStmt(cond, then, elseIfs, elseBody);
    }

    // ─── For ──────────────────────────────────────────────────────────────

    private VB6Statement ParseFor()
    {
        Advance(); // For
        if (Check(TokenKind.Each))
        {
            Advance();
            string v = ExpectIdent();
            VB6TypeRef? vt = null;
            if (Match(TokenKind.As)) vt = ParseTypeRef();
            Consume(TokenKind.In, "For Each");
            var coll = ParseExpression();
            EndLine();
            var body = ParseBlockUntil(() => Check(TokenKind.Next));
            Advance(); // Next
            if (Current.Kind == TokenKind.Identifier) Advance(); // optional var name
            return new ForEachStmt(v, vt, coll, body);
        }
        else
        {
            string v = ExpectIdent();
            var eq = Consume(TokenKind.Equals);
            if (eq.Kind != TokenKind.Equals) { SkipRestOfLine(); return new EmptyStmt(); }
            var start = ParseExpression();
            Consume(TokenKind.To, "For");
            var end = ParseExpression();
            VB6Expression? step = null;
            if (Match(TokenKind.Step)) step = ParseExpression();
            EndLine();
            var body = ParseBlockUntil(() => Check(TokenKind.Next));
            Advance(); // Next
            if (Current.Kind == TokenKind.Identifier) Advance();
            return new ForStmt(v, start, end, step, body);
        }
    }

    // ─── While ────────────────────────────────────────────────────────────

    private VB6Statement ParseWhile()
    {
        Advance(); // While
        var cond = ParseExpression();
        EndLine();
        var body = ParseBlockUntil(() => Check(TokenKind.Wend));
        Advance(); // Wend
        return new WhileStmt(cond, body);
    }

    // ─── Do ───────────────────────────────────────────────────────────────

    private VB6Statement ParseDo()
    {
        Advance(); // Do
        DoKind kind;
        VB6Expression? cond = null;

        if (Check(TokenKind.While)) { Advance(); cond = ParseExpression(); kind = DoKind.WhileTop; EndLine(); }
        else if (Check(TokenKind.Until)) { Advance(); cond = ParseExpression(); kind = DoKind.UntilTop; EndLine(); }
        else { kind = DoKind.Infinite; EndLine(); }

        var body = ParseBlockUntil(() => Check(TokenKind.Loop));
        Advance(); // Loop

        if (kind == DoKind.Infinite)
        {
            if (Check(TokenKind.While)) { Advance(); cond = ParseExpression(); kind = DoKind.WhileBottom; }
            else if (Check(TokenKind.Until)) { Advance(); cond = ParseExpression(); kind = DoKind.UntilBottom; }
        }
        return new DoStmt(kind, cond, body);
    }

    // ─── Select Case ──────────────────────────────────────────────────────

    private VB6Statement ParseSelect()
    {
        Advance(); // Select
        Consume(TokenKind.Case, "Select Case");
        var expr = ParseExpression();
        EndLine();
        SkipNewlines();

        var cases = new List<VB6CaseClause>();
        List<VB6Statement>? elseBody = null;

        while (!Check(TokenKind.EndOfFile) && !IsEnd("Select"))
        {
            SkipNewlines();
            if (IsEnd("Select")) break;
            if (!Check(TokenKind.Case)) { SkipRestOfLine(); continue; }
            Advance(); // Case

            if (Check(TokenKind.Else)) { Advance(); EndLine(); elseBody = ParseBlockUntil(() => IsEnd("Select")); break; }

            var caseExprs = ParseCaseExprList();
            EndLine();
            var caseBody = ParseBlockUntil(() => Check(TokenKind.Case) || IsEnd("Select"));
            cases.Add(new VB6CaseClause(caseExprs, caseBody));
        }

        if (IsEnd("Select")) { Advance(); Advance(); }
        return new SelectStmt(expr, cases, elseBody);
    }

    private List<VB6CaseExpr> ParseCaseExprList()
    {
        var list = new List<VB6CaseExpr> { ParseCaseExpr() };
        while (Match(TokenKind.Comma)) list.Add(ParseCaseExpr());
        return list;
    }

    private VB6CaseExpr ParseCaseExpr()
    {
        if (Check(TokenKind.Is))
        {
            Advance();
            string op = ParseRelOp();
            return new CaseIsExpr(op, ParseExpression());
        }
        var from = ParseExpression();
        if (Match(TokenKind.To)) return new CaseRangeExpr(from, ParseExpression());
        return new CaseValueExpr(from);
    }

    private string ParseRelOp()
    {
        string op = Current.Kind switch
        {
            TokenKind.Equals => "=",
            TokenKind.NotEqual => "<>",
            TokenKind.LessThan => "<",
            TokenKind.GreaterThan => ">",
            TokenKind.LessEqual => "<=",
            TokenKind.GreaterEqual => ">=",
            _ => "="
        };
        if (Current.Kind is TokenKind.Equals or TokenKind.NotEqual or TokenKind.LessThan
            or TokenKind.GreaterThan or TokenKind.LessEqual or TokenKind.GreaterEqual) Advance();
        return op;
    }

    // ─── With ─────────────────────────────────────────────────────────────

    private VB6Statement ParseWith()
    {
        Advance(); // With
        var obj = ParseExpression();
        EndLine();
        var body = ParseBlockUntil(() => IsEnd("With"));
        if (IsEnd("With")) { Advance(); Advance(); }
        return new WithStmt(obj, body);
    }

    // ─── Exit ─────────────────────────────────────────────────────────────

    private VB6Statement ParseExit()
    {
        Advance(); // Exit
        ExitKind ek = Current.Kind switch
        {
            TokenKind.Sub => ExitKind.Sub,
            TokenKind.Function => ExitKind.Function,
            TokenKind.For => ExitKind.For,
            TokenKind.Do => ExitKind.Do,
            TokenKind.While => ExitKind.While,
            TokenKind.Property => ExitKind.Property,
            _ => ExitKind.Sub
        };
        Advance();
        return new ExitStmt(ek);
    }

    // ─── On Error / Resume ────────────────────────────────────────────────

    private VB6Statement ParseOnError()
    {
        Advance(); // On
        Consume(TokenKind.Error, "On Error");
        if (Check(TokenKind.Resume)) { Advance(); Match(TokenKind.Next); return new OnErrorStmt(OnErrorKind.ResumeNext, null); }
        if (Check(TokenKind.GoTo))
        {
            Advance();
            if (Current.Text == "0") { Advance(); return new OnErrorStmt(OnErrorKind.Zero, null); }
            string lbl = Current.Text; Advance();
            return new OnErrorStmt(OnErrorKind.GoTo, lbl);
        }
        return new OnErrorStmt(OnErrorKind.Zero, null);
    }

    private VB6Statement ParseResume()
    {
        Advance(); // Resume
        if (Check(TokenKind.Next)) { Advance(); return new ResumeStmt(ResumeKind.Next, null); }
        if (IsEol()) return new ResumeStmt(ResumeKind.Default, null);
        string lbl = Current.Text; Advance();
        return new ResumeStmt(ResumeKind.Label, lbl);
    }

    // ─── RaiseEvent / ReDim / Erase ──────────────────────────────────────

    private VB6Statement ParseRaiseEvent()
    {
        Advance(); // RaiseEvent
        string name = ExpectIdent();
        var args = new List<VB6Expression>();
        if (Match(TokenKind.LeftParen))
        {
            if (!Check(TokenKind.RightParen)) args = ParseArgList();
            Consume(TokenKind.RightParen);
        }
        return new RaiseEventStmt(name, args);
    }

    private VB6Statement ParseReDim()
    {
        Advance(); // ReDim
        bool preserve = Match(TokenKind.Preserve);
        string name = ExpectIdent();
        Consume(TokenKind.LeftParen);
        var dims = new List<VB6Expression> { ParseArrayDim() };
        while (Match(TokenKind.Comma)) dims.Add(ParseArrayDim());
        Consume(TokenKind.RightParen);
        VB6TypeRef? type = null;
        if (Match(TokenKind.As)) type = ParseTypeRef();
        return new ReDimStmt(name, dims, preserve, type);
    }

    private VB6Statement ParseErase()
    {
        Advance(); // Erase
        var vars = new List<string> { ExpectIdent() };
        while (Match(TokenKind.Comma)) vars.Add(ExpectIdent());
        return new EraseStmt(vars);
    }

    // ─── File I/O statement parsing ──────────────────────────────────────

    // Reads a VB6 file-channel argument: "#n" is tokenised as DateLiteral("n")
    // or may be a plain integer/variable.
    private VB6Expression ParseFileNumber()
    {
        if (Current.Kind == TokenKind.DateLiteral)
        {
            var text = Current.Text.Trim();
            Advance();
            return int.TryParse(text, out int n)
                ? new LiteralExpr(n, LiteralKind.Integer)
                : new NameExpr(text);
        }
        return ParseExpression();
    }

    // Open path For {Input|Output|Append|Binary|Random} As [#]n [Len=n]
    private VB6Statement ParseOpen()
    {
        Advance(); // Open
        var path = ParseExpression();
        Consume(TokenKind.For, "Open");
        string mode = "Input";
        if (Current.Kind == TokenKind.Identifier || Current.Kind != TokenKind.EndOfFile)
        {
            mode = Current.Text;
            Advance();
        }
        Consume(TokenKind.As, "Open");
        var fileNum = ParseFileNumber();
        // Optional: Len = n
        if (Current.Kind == TokenKind.Identifier &&
            string.Equals(Current.Text, "Len", StringComparison.OrdinalIgnoreCase))
        {
            Advance(); Match(TokenKind.Equals); ParseExpression(); // consume Len=n, discard
        }
        return new OpenStmt(path, mode, fileNum);
    }

    // Close [#n [, #m ...]]
    private VB6Statement ParseClose()
    {
        Advance(); // Close
        var nums = new List<VB6Expression>();
        while (!IsEol())
        {
            nums.Add(ParseFileNumber());
            if (!Match(TokenKind.Comma)) break;
        }
        return new CloseStmt(nums);
    }

    // Line Input #n, var
    private VB6Statement ParseLineInput()
    {
        Advance(); Advance(); // Line, Input
        var fileNum = ParseFileNumber();
        Consume(TokenKind.Comma, "Line Input");
        var target = ParseExpression();
        return new LineInputStmt(fileNum, target);
    }

    // Print #n [, expr ...]  (semicolons treated as commas)
    private VB6Statement ParsePrintFile()
    {
        Advance(); // Print
        var fileNum = ParseFileNumber();
        var values = new List<VB6Expression>();
        // Consume optional separator (, or ;) then parse each value
        while (!IsEol() && (Current.Kind == TokenKind.Comma ||
               (Current.Kind == TokenKind.Identifier && Current.Text == ";")))
        {
            Advance(); // skip separator
            if (IsEol()) break;
            values.Add(ParseExpression());
        }
        return new PrintFileStmt(fileNum, values);
    }

    // Write #n [, expr ...]
    private VB6Statement ParseWriteFile()
    {
        Advance(); // Write
        var fileNum = ParseFileNumber();
        var values = new List<VB6Expression>();
        if (Match(TokenKind.Comma))
        {
            if (!IsEol()) values.Add(ParseExpression());
            while (Match(TokenKind.Comma)) { if (!IsEol()) values.Add(ParseExpression()); }
        }
        return new WriteFileStmt(fileNum, values);
    }

    // Input #n, var [, var ...]
    private VB6Statement ParseInputFile()
    {
        Advance(); // Input
        var fileNum = ParseFileNumber();
        Consume(TokenKind.Comma, "Input #");
        var vars = new List<VB6Expression> { ParseExpression() };
        while (Match(TokenKind.Comma)) vars.Add(ParseExpression());
        return new InputFileStmt(fileNum, vars);
    }

    // Get [#]n [, [pos]], var
    private VB6Statement ParseGetFile()
    {
        Advance(); // Get
        var fileNum = ParseFileNumber();
        Consume(TokenKind.Comma, "Get #");
        VB6Expression? pos = null;
        if (!Check(TokenKind.Comma)) pos = ParseExpression();
        Consume(TokenKind.Comma, "Get #");
        var variable = ParseExpression();
        return new GetFileStmt(fileNum, pos, variable);
    }

    // Put [#]n [, [pos]], expr
    private VB6Statement ParsePutFile()
    {
        Advance(); // Put
        var fileNum = ParseFileNumber();
        Consume(TokenKind.Comma, "Put #");
        VB6Expression? pos = null;
        if (!Check(TokenKind.Comma)) pos = ParseExpression();
        Consume(TokenKind.Comma, "Put #");
        var value = ParseExpression();
        return new PutFileStmt(fileNum, pos, value);
    }

    // Name oldpath As newpath
    private VB6Statement ParseNameFile()
    {
        Advance(); // Name
        var oldPath = ParseExpression();
        Consume(TokenKind.As, "Name");
        var newPath = ParseExpression();
        return new NameFileStmt(oldPath, newPath);
    }

    // Kill/MkDir/RmDir/ChDir/ChDrive — single expression argument
    private VB6Statement ParseSingleArgFileOp(Func<VB6Expression, VB6Statement> factory)
    {
        Advance(); // keyword
        var arg = ParseExpression();
        return factory(arg);
    }

    // FileCopy source, dest
    private VB6Statement ParseFileCopy()
    {
        Advance(); // FileCopy
        var src = ParseExpression();
        Consume(TokenKind.Comma, "FileCopy");
        var dst = ParseExpression();
        return new FileCopyStmt(src, dst);
    }

    // ─── Call / Set / assign-or-call ──────────────────────────────────────

    private VB6Statement ParseCall()
    {
        Advance(); // Call
        var target = ParsePostfix();
        // Arguments already consumed by ParsePostfix if parens present
        return new CallStmt(target, []);
    }

    private VB6Statement ParseSetAssign()
    {
        var target = ParsePostfix();
        Consume(TokenKind.Equals);
        var val = ParseExpression();
        return new AssignStmt(target, val, IsSet: true);
    }

    private VB6Statement ParseAssignOrCall()
    {
        var lhs = ParsePostfix();

        if (Check(TokenKind.Equals))
        {
            Advance();
            var rhs = ParseExpression();
            return new AssignStmt(lhs, rhs);
        }

        // Sub-call without parens: SubName arg1, arg2
        var args = new List<VB6Expression>();
        if (!IsEol())
        {
            args.Add(ParseExpression());
            while (Match(TokenKind.Comma)) args.Add(ParseExpression());
        }
        return new CallStmt(lhs, args);
    }

    private VB6Statement ParseLocalConst()
    {
        string name = ExpectIdent();
        VB6TypeRef? type = null;
        if (Match(TokenKind.As)) type = ParseTypeRef();
        Consume(TokenKind.Equals);
        var val = ParseExpression();
        return new LocalConstStmt(name, type, val);
    }

    // ─── Block helper ─────────────────────────────────────────────────────

    private List<VB6Statement> ParseBlockUntil(Func<bool> stop)
    {
        var stmts = new List<VB6Statement>();
        while (!Check(TokenKind.EndOfFile) && !stop())
        {
            SkipNewlines();
            if (Check(TokenKind.EndOfFile) || stop()) break;
            int before = _pos;
            ParseStatements(stmts);
            if (_pos == before) Advance(); // stall guard
        }
        return stmts;
    }

    // ─── Declarators ──────────────────────────────────────────────────────

    private List<VB6Declarator> ParseDeclaratorList()
    {
        var list = new List<VB6Declarator> { ParseDeclarator() };
        while (Match(TokenKind.Comma)) list.Add(ParseDeclarator());
        return list;
    }

    private VB6Declarator ParseDeclarator()
    {
        Match(TokenKind.WithEvents);
        string name = ExpectIdent();
        bool isArray = false;
        List<VB6Expression>? dims = null;

        if (Check(TokenKind.LeftParen))
        {
            isArray = true;
            Advance();
            dims = [];
            if (!Check(TokenKind.RightParen))
            {
                dims.Add(ParseArrayDim());
                while (Match(TokenKind.Comma)) dims.Add(ParseArrayDim());
            }
            Consume(TokenKind.RightParen);
        }

        VB6TypeRef? type = null;
        if (Match(TokenKind.As))
        {
            Match(TokenKind.New);
            type = ParseTypeRef();
        }

        return new VB6Declarator(name, type, isArray, dims);
    }

    // ─── Parameters ───────────────────────────────────────────────────────

    private List<VB6Parameter> ParseParamList()
    {
        var list = new List<VB6Parameter>();
        if (!Match(TokenKind.LeftParen)) return list;
        if (!Check(TokenKind.RightParen))
        {
            list.Add(ParseParam());
            while (Match(TokenKind.Comma)) list.Add(ParseParam());
        }
        Consume(TokenKind.RightParen);
        return list;
    }

    private VB6Parameter ParseParam()
    {
        bool optional = false, paramArray = false, byRef = true;
        bool done = false;
        while (!done)
        {
            switch (Current.Kind)
            {
                case TokenKind.Optional: optional = true; Advance(); break;
                case TokenKind.ParamArray: paramArray = true; Advance(); break;
                case TokenKind.ByRef: byRef = true; Advance(); break;
                case TokenKind.ByVal: byRef = false; Advance(); break;
                default: done = true; break;
            }
        }

        string name = ExpectIdent();
        bool isArr = false;
        if (Check(TokenKind.LeftParen)) { Advance(); Consume(TokenKind.RightParen); isArr = true; }

        var type = new VB6TypeRef("Variant");
        if (Match(TokenKind.As))
        {
            Match(TokenKind.New);
            type = ParseTypeRef();
        }
        if (isArr) type = type with { IsArray = true };

        VB6Expression? defVal = null;
        if (Match(TokenKind.Equals)) defVal = ParseExpression();

        return new VB6Parameter(name, type, byRef, optional, paramArray, defVal);
    }

    // ─── Type reference ───────────────────────────────────────────────────

    private VB6TypeRef ParseTypeRef()
    {
        string name = Current.Kind switch
        {
            TokenKind.StringType => "String",
            TokenKind.ObjectType => "Object",
            _ => Current.Text
        };
        Advance();
        while (Check(TokenKind.Dot))
        {
            Advance();
            name += "." + Current.Text;
            Advance();
        }
        return new VB6TypeRef(name);
    }

    // ─── Access modifier ──────────────────────────────────────────────────

    private AccessModifier ParseAccess()
    {
        switch (Current.Kind)
        {
            case TokenKind.Public: Advance(); return AccessModifier.Public;
            case TokenKind.Private: Advance(); return AccessModifier.Private;
            case TokenKind.Friend: Advance(); return AccessModifier.Friend;
            default: return AccessModifier.Default;
        }
    }

    // ─── Expression parsing (Pratt) ───────────────────────────────────────

    private VB6Expression ParseExpression() => ParseOr();

    private VB6Expression ParseOr()
    {
        var left = ParseAnd();
        while (Current.Kind is TokenKind.Or or TokenKind.Xor or TokenKind.Eqv or TokenKind.Imp)
        {
            string op = Advance().Text;
            left = new BinaryExpr(op, left, ParseAnd());
        }
        return left;
    }

    private VB6Expression ParseAnd()
    {
        var left = ParseNot();
        while (Check(TokenKind.And)) { Advance(); left = new BinaryExpr("And", left, ParseNot()); }
        return left;
    }

    private VB6Expression ParseNot()
    {
        if (Check(TokenKind.Not)) { Advance(); return new UnaryExpr("Not", ParseNot()); }
        return ParseComparison();
    }

    private VB6Expression ParseComparison()
    {
        var left = ParseConcat();
        while (Current.Kind is TokenKind.Equals or TokenKind.NotEqual or TokenKind.LessThan
               or TokenKind.GreaterThan or TokenKind.LessEqual or TokenKind.GreaterEqual
               or TokenKind.Is or TokenKind.Like)
        {
            string op = Current.Kind == TokenKind.Is ? "Is" : Current.Text;
            Advance();
            if (op == "Is" && Check(TokenKind.Not)) { Advance(); op = "IsNot"; }
            left = new BinaryExpr(op, left, ParseConcat());
        }
        return left;
    }

    private VB6Expression ParseConcat()
    {
        var left = ParseAddSub();
        while (Check(TokenKind.Ampersand)) { Advance(); left = new BinaryExpr("&", left, ParseAddSub()); }
        return left;
    }

    private VB6Expression ParseAddSub()
    {
        var left = ParseMulDiv();
        while (Current.Kind is TokenKind.Plus or TokenKind.Minus)
        {
            string op = Advance().Text;
            left = new BinaryExpr(op, left, ParseMulDiv());
        }
        return left;
    }

    private VB6Expression ParseMulDiv()
    {
        var left = ParseIntDiv();
        while (Current.Kind is TokenKind.Star or TokenKind.Slash)
        {
            string op = Advance().Text;
            left = new BinaryExpr(op, left, ParseIntDiv());
        }
        return left;
    }

    private VB6Expression ParseIntDiv()
    {
        var left = ParseMod();
        while (Check(TokenKind.BackSlash)) { Advance(); left = new BinaryExpr("\\", left, ParseMod()); }
        return left;
    }

    private VB6Expression ParseMod()
    {
        var left = ParseExp();
        while (Check(TokenKind.Mod)) { Advance(); left = new BinaryExpr("Mod", left, ParseExp()); }
        return left;
    }

    private VB6Expression ParseExp()
    {
        var left = ParseUnary();
        while (Check(TokenKind.Caret)) { Advance(); left = new BinaryExpr("^", left, ParseUnary()); }
        return left;
    }

    private VB6Expression ParseUnary()
    {
        if (Check(TokenKind.Minus)) { Advance(); return new UnaryExpr("-", ParseUnary()); }
        if (Check(TokenKind.Plus)) { Advance(); return ParseUnary(); }
        if (Check(TokenKind.TypeOf))
        {
            Advance();
            var obj = ParsePostfix();
            Consume(TokenKind.Is, "TypeOf");
            return new TypeOfExpr(obj, ParseTypeRef());
        }
        if (Check(TokenKind.New)) { Advance(); return new NewExpr(ParseTypeRef()); }
        if (Check(TokenKind.AddressOf)) { Advance(); return new AddressOfExpr(ParsePostfix()); }
        return ParsePostfix();
    }

    private VB6Expression ParsePostfix()
    {
        var expr = ParsePrimary();
        while (true)
        {
            if (Check(TokenKind.Dot))
            {
                Advance();
                string member = Current.Kind is TokenKind.Identifier or TokenKind.Get or TokenKind.Let or TokenKind.Set
                    ? Advance().Text
                    : ExpectIdent();
                if (Check(TokenKind.LeftParen))
                {
                    Advance();
                    var args = Check(TokenKind.RightParen) ? new List<VB6Expression>() : ParseArgList();
                    Consume(TokenKind.RightParen);
                    expr = new CallExpr(new MemberExpr(expr, member), args);
                }
                else expr = new MemberExpr(expr, member);
            }
            else if (Check(TokenKind.Bang))
            {
                Advance();
                string member = ExpectIdent();
                expr = new MemberExpr(expr, member);
            }
            else if (Check(TokenKind.LeftParen))
            {
                Advance();
                var args = Check(TokenKind.RightParen) ? new List<VB6Expression>() : ParseArgList();
                Consume(TokenKind.RightParen);
                expr = new CallExpr(expr, args);
            }
            else break;
        }
        return expr;
    }

    private VB6Expression ParsePrimary()
    {
        var tok = Current;
        switch (tok.Kind)
        {
            case TokenKind.IntegerLiteral: Advance(); return new LiteralExpr(tok.Value, LiteralKind.Integer);
            case TokenKind.LongLiteral:    Advance(); return new LiteralExpr(tok.Value, LiteralKind.Long);
            case TokenKind.SingleLiteral:  Advance(); return new LiteralExpr(tok.Value, LiteralKind.Single);
            case TokenKind.DoubleLiteral:  Advance(); return new LiteralExpr(tok.Value, LiteralKind.Double);
            case TokenKind.StringLiteral:  Advance(); return new LiteralExpr(tok.Value, LiteralKind.String);
            case TokenKind.DateLiteral:    Advance(); return new LiteralExpr(tok.Value, LiteralKind.Date);
            case TokenKind.True:           Advance(); return new LiteralExpr(true, LiteralKind.Boolean);
            case TokenKind.False:          Advance(); return new LiteralExpr(false, LiteralKind.Boolean);
            case TokenKind.Nothing:        Advance(); return new LiteralExpr(null, LiteralKind.Nothing);
            case TokenKind.Null:           Advance(); return new LiteralExpr(null, LiteralKind.Null);
            case TokenKind.Empty:          Advance(); return new LiteralExpr(null, LiteralKind.Empty);
            case TokenKind.Me:             Advance(); return new MeExpr();
            case TokenKind.Dot:
                Advance();
                string wm = Current.Text; Advance();
                return new WithMemberExpr(wm);
            case TokenKind.LeftParen:
                Advance();
                var inner = ParseExpression();
                Consume(TokenKind.RightParen);
                return inner;
            case TokenKind.Identifier:
                Advance();
                return new NameExpr(tok.Text);
            case TokenKind.StringType:
                Advance();
                return new NameExpr("String");
            case TokenKind.ObjectType:
                Advance();
                return new NameExpr("Object");
            default:
                // Keywords used as identifiers (e.g., Error, Name, Type)
                if (tok.Kind != TokenKind.EndOfFile && tok.Kind != TokenKind.NewLine
                    && tok.Kind != TokenKind.RightParen && !IsEol())
                {
                    Advance();
                    return new NameExpr(tok.Text);
                }
                return new LiteralExpr(null, LiteralKind.Nothing);
        }
    }

    private List<VB6Expression> ParseArgList()
    {
        var args = new List<VB6Expression>();
        do
        {
            if (Check(TokenKind.Comma))
            {
                args.Add(new LiteralExpr(null, LiteralKind.Empty)); // missing optional arg
            }
            else
            {
                // Named argument: name:=expr — skip name and :=
                // (name may be a keyword used as an argument name, e.g. Module:=, FileName:=)
                if ((Current.Kind == TokenKind.Identifier || IsKeywordAsIdent(Current.Kind))
                    && Peek().Kind == TokenKind.Colon && Peek(2).Kind == TokenKind.Equals)
                {
                    Advance(); Advance(); Advance();
                }
                args.Add(ParseExpression());
            }
        }
        while (Match(TokenKind.Comma));
        return args;
    }

    // Parses one array-dimension slot, handling the "lowerBound To upperBound" form.
    // The lower bound is discarded; the upper bound (or the sole value) is returned.
    private VB6Expression ParseArrayDim()
    {
        var expr = ParseExpression();
        if (Match(TokenKind.To)) expr = ParseExpression();
        return expr;
    }

    // ─── Helpers ──────────────────────────────────────────────────────────

    private string ExpectIdent()
    {
        if (Current.Kind == TokenKind.Identifier || IsKeywordAsIdent(Current.Kind))
            return Advance().Text;
        _diagnostics.Add($"Line {Current.Line}: expected identifier, got '{Current.Text}'");
        return "__missing";
    }

    private static bool IsKeywordAsIdent(TokenKind k) => k is
        TokenKind.Get or TokenKind.Let or TokenKind.Set or TokenKind.Error or
        TokenKind.Resume or TokenKind.Return or TokenKind.Event or TokenKind.Type or
        TokenKind.Property or TokenKind.Class or TokenKind.Module or TokenKind.Lib or
        TokenKind.Alias or TokenKind.Compare or TokenKind.Base or TokenKind.Explicit or
        TokenKind.Implements or TokenKind.ObjectType or TokenKind.StringType or
        TokenKind.Step or TokenKind.In;
}
