using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.CSharp;
using Microsoft.CodeAnalysis.CSharp.Syntax;
using VB6ToCSharp.Parser;
using static Microsoft.CodeAnalysis.CSharp.SyntaxFactory;

namespace VB6ToCSharp.Transformer;

/// <summary>
/// Transforms a VB6Module AST into a Roslyn CompilationUnitSyntax (C# .NET 8).
/// Uses SyntaxFactory throughout so the output is a proper Roslyn green tree
/// that can be further analysed, formatted, or emitted.
/// </summary>
public class VB6ToCSharpTransformer(string namespaceName = "Converted")
{
    // Stack used to resolve .Member expressions inside With blocks
    private readonly Stack<ExpressionSyntax> _withStack = new();

    // Name of the current function being parsed (for VB6 "FuncName = val" returns)
    private string _currentFunctionName = "";

    // WithEvents event subscriptions to inject into Class_Initialize (or a constructor)
    private List<StatementSyntax> _pendingEventSubscriptions = [];
    private bool _hasClassInitialize = false;

    public List<string> Diagnostics { get; } = [];

    // Return type of the current function (used to emit __result local)
    private TypeSyntax? _currentReturnType;

    // Name of the __result local variable used inside functions
    private const string ResultVar = "__result";

    // Name of the exception variable in generated catch blocks
    private const string CatchVar = "ex";

    // Counter for unique temporary variable names
    private int _tempSeq;

    // ─── Entry point ──────────────────────────────────────────────────────

    public CompilationUnitSyntax Transform(VB6Module module)
    {
        var usings = List(new[]
        {
            UsingDirective(ParseName("System")),
            UsingDirective(ParseName("System.Collections.Generic")),
            UsingDirective(ParseName("System.Diagnostics")),
            UsingDirective(ParseName("System.IO")),
            UsingDirective(ParseName("System.Linq")),
            UsingDirective(ParseName("System.Text")),
            UsingDirective(ParseName("System.Runtime.InteropServices")),
        });

        var classDecl = BuildClass(module);

        MemberDeclarationSyntax topLevel = string.IsNullOrWhiteSpace(namespaceName)
            ? classDecl
            : FileScopedNamespaceDeclaration(ParseName(namespaceName))
                .WithMembers(SingletonList<MemberDeclarationSyntax>(classDecl));

        return CompilationUnit()
            .WithUsings(usings)
            .WithMembers(SingletonList(topLevel))
            .NormalizeWhitespace();
    }

    // ─── Class / struct building ──────────────────────────────────────────

    private ClassDeclarationSyntax BuildClass(VB6Module module)
    {
        _pendingEventSubscriptions = BuildEventSubscriptions(module);
        _hasClassInitialize = false;

        var members = new List<MemberDeclarationSyntax>();

        foreach (var m in module.Members)
        {
            var converted = TransformMember(m);
            if (converted != null) members.Add(converted);
        }

        // If there are WithEvents subscriptions but no Class_Initialize, emit a constructor
        if (_pendingEventSubscriptions.Count > 0 && !_hasClassInitialize)
        {
            var ctor = ConstructorDeclaration(TypeMapper.SanitizeName(module.Name))
                .WithModifiers(TokenList(Token(SyntaxKind.PublicKeyword)))
                .WithBody(Block(List(_pendingEventSubscriptions)));
            members.Insert(0, ctor);
        }

        var modifiers = module.Kind == ModuleKind.Standard
            ? TokenList(Token(SyntaxKind.PublicKeyword), Token(SyntaxKind.StaticKeyword))
            : TokenList(Token(SyntaxKind.PublicKeyword));

        if (module.Kind == ModuleKind.Standard)
            members = members.Select(EnsureStaticMember).ToList();

        var cls = ClassDeclaration(TypeMapper.SanitizeName(module.Name))
            .WithModifiers(modifiers)
            .WithMembers(List(members));

        // Partial interface implementations
        if (module.Implements.Count > 0)
        {
            cls = cls.WithBaseList(BaseList(SeparatedList(
                module.Implements.Select(i =>
                    (BaseTypeSyntax)SimpleBaseType(IdentifierName(TypeMapper.SanitizeName(i)))))));
        }

        return cls;
    }

    // Ensures every member of a Standard module (static class) carries the static modifier.
    private static MemberDeclarationSyntax EnsureStaticMember(MemberDeclarationSyntax m)
    {
        if (m.Modifiers.Any(SyntaxKind.StaticKeyword) ||
            m.Modifiers.Any(SyntaxKind.ConstKeyword) ||
            m is TypeDeclarationSyntax or EnumDeclarationSyntax or StructDeclarationSyntax)
            return m;
        return m switch
        {
            MethodDeclarationSyntax md  => md.AddModifiers(Token(SyntaxKind.StaticKeyword)),
            FieldDeclarationSyntax fd   => fd.AddModifiers(Token(SyntaxKind.StaticKeyword)),
            PropertyDeclarationSyntax pd => pd.AddModifiers(Token(SyntaxKind.StaticKeyword)),
            EventFieldDeclarationSyntax ed => ed.AddModifiers(Token(SyntaxKind.StaticKeyword)),
            _ => m
        };
    }

    // ─── Member transformation ────────────────────────────────────────────

    private MemberDeclarationSyntax? TransformMember(VB6MemberNode node)
    {
        return node switch
        {
            SubDecl sd => TransformSub(sd),
            FunctionDecl fd => TransformFunction(fd),
            PropertyDecl pd => TransformProperty(pd),
            FieldDecl fd => TransformField(fd),
            ConstMember cm => TransformConstMember(cm),
            EventDecl ed => TransformEvent(ed),
            UserTypeDecl ud => ud.IsEnum ? TransformEnum(ud) : TransformStruct(ud),
            DeclareDecl dd => TransformDeclare(dd),
            _ => null
        };
    }

    // Sub → void method
    private MethodDeclarationSyntax TransformSub(SubDecl sd)
    {
        _currentFunctionName = "";
        var bodyStmts = TransformMethodStatements(sd.Body).ToList();

        if (string.Equals(sd.Name, "Class_Initialize", StringComparison.OrdinalIgnoreCase)
            && _pendingEventSubscriptions.Count > 0)
        {
            _hasClassInitialize = true;
            bodyStmts.InsertRange(0, _pendingEventSubscriptions);
        }

        var body = Block(List(bodyStmts));
        return MethodDeclaration(PredefinedType(Token(SyntaxKind.VoidKeyword)),
                TypeMapper.SanitizeName(sd.Name))
            .WithModifiers(BuildModifiers(sd.Access, sd.IsStatic))
            .WithParameterList(BuildParams(sd.Parameters))
            .WithBody(body);
    }

    // Recursively yields every statement in a body, descending into all nested blocks.
    private static IEnumerable<VB6Statement> FlattenStmts(IEnumerable<VB6Statement> stmts)
    {
        foreach (var s in stmts)
        {
            yield return s;
            IEnumerable<VB6Statement> children = s switch
            {
                IfStmt ifs      => ifs.Then
                                    .Concat(ifs.ElseIfs.SelectMany(e => e.Body))
                                    .Concat(ifs.Else ?? []),
                ForStmt fs      => fs.Body,
                ForEachStmt fes => fes.Body,
                WhileStmt ws    => ws.Body,
                DoStmt ds       => ds.Body,
                WithStmt wts    => wts.Body,
                SelectStmt ss   => ss.Cases.SelectMany(c => c.Body)
                                    .Concat(ss.ElseBody ?? []),
                _               => []
            };
            foreach (var child in FlattenStmts(children))
                yield return child;
        }
    }

    // Returns the set of WithEvents variable names that use "Set varName.Notify = Me"
    // (COM sink pattern) anywhere in the module body — these don't need += wiring.
    private static HashSet<string> FindNotifySinkVars(VB6Module module)
    {
        var sinks = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var allStmts = FlattenStmts(
            module.Members.OfType<SubDecl>().SelectMany(s => s.Body)
            .Concat(module.Members.OfType<FunctionDecl>().SelectMany(f => f.Body)));
        foreach (var stmt in allStmts)
            if (stmt is AssignStmt { IsSet: true, Value: MeExpr,
                    Target: MemberExpr { Object: NameExpr objExpr } })
                sinks.Add(objExpr.Name);
        return sinks;
    }

    private List<StatementSyntax> BuildEventSubscriptions(VB6Module module)
    {
        var result = new List<StatementSyntax>();

        var withEventsVars = module.Members
            .OfType<FieldDecl>()
            .SelectMany(f => f.Declarators)
            .Where(d => d.IsWithEvents)
            .Select(d => d.Name)
            .ToList();

        if (withEventsVars.Count == 0) return result;

        var handlerNames = module.Members
            .OfType<SubDecl>()
            .Select(s => s.Name)
            .ToList();

        var notifySinks = FindNotifySinkVars(module);

        foreach (var varName in withEventsVars)
        {
            string prefix = varName + "_";
            var matched = new List<StatementSyntax>();
            foreach (var handlerName in handlerNames)
            {
                if (!handlerName.StartsWith(prefix, StringComparison.OrdinalIgnoreCase)) continue;
                string eventName = handlerName[prefix.Length..];
                matched.Add(ExpressionStatement(
                    AssignmentExpression(
                        SyntaxKind.AddAssignmentExpression,
                        MemberAccessExpression(SyntaxKind.SimpleMemberAccessExpression,
                            IdentifierName(varName),
                            IdentifierName(eventName)),
                        IdentifierName(handlerName))));
            }

            if (matched.Count > 0)
                result.AddRange(matched);
            else if (!notifySinks.Contains(varName))
            {
                result.Add(EmptyStatement()
                    .WithLeadingTrivia(TriviaList(
                        Comment($"// TODO: {varName}.??? += {varName}_???;"),
                        EndOfLine(Environment.NewLine))));
            }
        }

        return result;
    }

    // Function → typed method (VB6 "FuncName = val" → local __result + return __result)
    private MethodDeclarationSyntax TransformFunction(FunctionDecl fd)
    {
        _currentFunctionName = fd.Name;
        _currentReturnType = TypeMapper.ToCSharp(fd.ReturnType.Name, fd.ReturnType.IsArray);

        var innerStmts = TransformMethodStatements(fd.Body).ToList();

        // Prepend: T __result = <typed-default>;
        var resultDecl = LocalDeclarationStatement(
            VariableDeclaration(_currentReturnType)
                .WithVariables(SingletonSeparatedList(
                    VariableDeclarator(ResultVar)
                        .WithInitializer(EqualsValueClause(
                            TypeMapper.DefaultValue(fd.ReturnType.Name))))));

        // Append: return __result;
        var returnStmt = ReturnStatement(IdentifierName(ResultVar));

        // Remove trailing bare returns that were emitted for Exit Function
        var stmts = new List<StatementSyntax> { resultDecl };
        stmts.AddRange(innerStmts);
        // Only add trailing return if the last statement is not already a return
        if (stmts.Count == 0 || stmts[^1] is not ReturnStatementSyntax)
            stmts.Add(returnStmt);

        var body = Block(List(stmts));

        _currentFunctionName = "";
        _currentReturnType = null;

        return MethodDeclaration(TypeMapper.ToCSharp(fd.ReturnType.Name, fd.ReturnType.IsArray),
                TypeMapper.SanitizeName(fd.Name))
            .WithModifiers(BuildModifiers(fd.Access, fd.IsStatic))
            .WithParameterList(BuildParams(fd.Parameters))
            .WithBody(body);
    }

    // Property Get/Let/Set → C# property or method
    private MemberDeclarationSyntax TransformProperty(PropertyDecl pd)
    {
        var type = TypeMapper.ToCSharp(pd.Type.Name, pd.Type.IsArray);

        if (pd.Kind == PropertyKind.Get)
        {
            // Property Get uses the same return pattern as functions
            _currentFunctionName = pd.Name;
            _currentReturnType = type;

            var innerStmts = TransformMethodStatements(pd.Body).ToList();
            var resultDecl = LocalDeclarationStatement(
                VariableDeclaration(type)
                    .WithVariables(SingletonSeparatedList(
                        VariableDeclarator(ResultVar)
                            .WithInitializer(EqualsValueClause(
                                TypeMapper.DefaultValue(pd.Type.Name))))));
            var returnStmt = ReturnStatement(IdentifierName(ResultVar));
            var stmts = new List<StatementSyntax> { resultDecl };
            stmts.AddRange(innerStmts);
            if (stmts.Count == 0 || stmts[^1] is not ReturnStatementSyntax)
                stmts.Add(returnStmt);

            _currentFunctionName = "";
            _currentReturnType = null;

            return PropertyDeclaration(type, TypeMapper.SanitizeName(pd.Name))
                .WithModifiers(BuildModifiers(pd.Access, false))
                .WithAccessorList(AccessorList(SingletonList(
                    AccessorDeclaration(SyntaxKind.GetAccessorDeclaration)
                        .WithBody(Block(List(stmts))))));
        }
        else
        {
            _currentFunctionName = "";
            _currentReturnType = null;
            var body = TransformBody(pd.Body);
            // Let/Set → emit as setter method; a future pass can merge Get+Let into a full C# property
            string suffix = pd.Kind == PropertyKind.Set ? "Set" : "Let";
            return MethodDeclaration(PredefinedType(Token(SyntaxKind.VoidKeyword)),
                    TypeMapper.SanitizeName(pd.Name) + "_" + suffix)
                .WithModifiers(BuildModifiers(pd.Access, false))
                .WithParameterList(BuildParams(pd.Parameters))
                .WithBody(body);
        }
    }

    // Field / module-level variable
    private FieldDeclarationSyntax TransformField(FieldDecl fd)
    {
        var variables = SeparatedList(fd.Declarators.Select(d =>
        {
            var typeName = d.Type?.Name ?? "Variant";
            var varDeclarator = VariableDeclarator(TypeMapper.SanitizeName(d.Name));
            if (d.IsArray)
            {
                ExpressionSyntax? init = null;
                if (d.Dimensions != null && d.Dimensions.Count > 0)
                {
                    var sizes = d.Dimensions.Select(dim =>
                        BinaryExpression(SyntaxKind.AddExpression,
                            TransformExpr(dim), LiteralExpression(SyntaxKind.NumericLiteralExpression, Literal(1))))
                        .ToList();

                    init = ArrayCreationExpression(
                        ArrayType(TypeMapper.ToCSharp(typeName))
                            .WithRankSpecifiers(SingletonList(ArrayRankSpecifier(
                                SeparatedList<ExpressionSyntax>(sizes)))));
                }
                if (init != null)
                    varDeclarator = varDeclarator.WithInitializer(EqualsValueClause(init));
            }
            return varDeclarator;
        }));

        var firstType = fd.Declarators.FirstOrDefault()?.Type;
        var csType = firstType != null
            ? TypeMapper.ToCSharp(firstType.Name, firstType.IsArray || (fd.Declarators.FirstOrDefault()?.IsArray ?? false))
            : PredefinedType(Token(SyntaxKind.ObjectKeyword));

        var varDecl = VariableDeclaration(csType).WithVariables(variables);
        return FieldDeclaration(varDecl).WithModifiers(BuildModifiers(fd.Access, fd.IsStatic));
    }

    // Module-level Const → C# const field
    private FieldDeclarationSyntax TransformConstMember(ConstMember cm)
    {
        var type = cm.Type != null
            ? TypeMapper.ToCSharp(cm.Type.Name)
            : PredefinedType(Token(SyntaxKind.ObjectKeyword));

        var varDecl = VariableDeclaration(type)
            .WithVariables(SingletonSeparatedList(
                VariableDeclarator(TypeMapper.SanitizeName(cm.Name))
                    .WithInitializer(EqualsValueClause(TransformExpr(cm.Value)))));

        var mods = BuildModifiers(cm.Access, false);
        mods = mods.Add(Token(SyntaxKind.ConstKeyword));

        return FieldDeclaration(varDecl).WithModifiers(mods);
    }

    // Event declaration
    private EventFieldDeclarationSyntax TransformEvent(EventDecl ed)
    {
        var delegateType = ed.Parameters.Count == 0
            ? (TypeSyntax)IdentifierName("Action")
            : GenericName("Action")
                .WithTypeArgumentList(TypeArgumentList(SeparatedList(
                    ed.Parameters.Select(p => TypeMapper.ToCSharp(p.Type.Name, p.Type.IsArray)))));

        return EventFieldDeclaration(
            VariableDeclaration(delegateType)
                .WithVariables(SingletonSeparatedList(
                    VariableDeclarator(TypeMapper.SanitizeName(ed.Name)))))
            .WithModifiers(BuildModifiers(ed.Access, false));
    }

    // User-defined Type → struct
    private StructDeclarationSyntax TransformStruct(UserTypeDecl ud)
    {
        var fields = ud.Fields.Select(f =>
        {
            var ft = f.Type != null ? TypeMapper.ToCSharp(f.Type.Name, f.IsArray) : PredefinedType(Token(SyntaxKind.ObjectKeyword));
            return (MemberDeclarationSyntax)FieldDeclaration(
                VariableDeclaration(ft)
                    .WithVariables(SingletonSeparatedList(VariableDeclarator(TypeMapper.SanitizeName(f.Name)))))
                .WithModifiers(TokenList(Token(SyntaxKind.PublicKeyword)));
        });

        return StructDeclaration(TypeMapper.SanitizeName(ud.Name))
            .WithModifiers(BuildModifiers(ud.Access, false))
            .WithMembers(List(fields));
    }

    // Enum
    private EnumDeclarationSyntax TransformEnum(UserTypeDecl ud)
    {
        var members = ud.Fields.Select(f =>
        {
            var em = EnumMemberDeclaration(TypeMapper.SanitizeName(f.Name));
            if (f.Dimensions?.Count == 1)
                em = em.WithEqualsValue(EqualsValueClause(TransformExpr(f.Dimensions[0])));
            return em;
        });

        return EnumDeclaration(TypeMapper.SanitizeName(ud.Name))
            .WithModifiers(BuildModifiers(ud.Access, false))
            .WithMembers(SeparatedList(members));
    }

    // Declare → DllImport extern method
    private MethodDeclarationSyntax TransformDeclare(DeclareDecl dd)
    {
        var retType = dd.ReturnType != null
            ? TypeMapper.ToCSharp(dd.ReturnType.Name)
            : PredefinedType(Token(SyntaxKind.VoidKeyword));

        var dllImport = Attribute(IdentifierName("DllImport"))
            .WithArgumentList(AttributeArgumentList(SeparatedList(new[]
            {
                AttributeArgument(LiteralExpression(SyntaxKind.StringLiteralExpression, Literal(dd.Lib))),
                AttributeArgument(
                    NameEquals("EntryPoint"),
                    null,
                    LiteralExpression(SyntaxKind.StringLiteralExpression,
                        Literal(dd.Alias ?? dd.Name)))
            })));

        return MethodDeclaration(retType, TypeMapper.SanitizeName(dd.Name))
            .WithAttributeLists(SingletonList(
                AttributeList(SingletonSeparatedList(dllImport))))
            .WithModifiers(TokenList(
                Token(SyntaxKind.PublicKeyword),
                Token(SyntaxKind.StaticKeyword),
                Token(SyntaxKind.ExternKeyword)))
            .WithParameterList(BuildParams(dd.Parameters))
            .WithSemicolonToken(Token(SyntaxKind.SemicolonToken));
    }

    // ─── Statement transformation ─────────────────────────────────────────

    private BlockSyntax TransformBody(List<VB6Statement> stmts)
    {
        var csharpStmts = stmts
            .SelectMany(TransformStatement)
            .Where(s => s != null)
            .Select(s => s!);
        return Block(List(csharpStmts));
    }

    // Top-level method body transformation — applies On Error GoTo / Resume Next restructuring.
    // Not used for nested blocks (if/for/while bodies) to avoid incorrect partial-wrapping.
    private IEnumerable<StatementSyntax> TransformMethodStatements(List<VB6Statement> stmts)
    {
        int onErrorIdx = stmts.FindIndex(
            s => s is OnErrorStmt { Kind: OnErrorKind.GoTo });
        if (onErrorIdx >= 0)
            return TransformWithGoToHandler(stmts, onErrorIdx);

        int resumeNextIdx = stmts.FindIndex(
            s => s is OnErrorStmt { Kind: OnErrorKind.ResumeNext });
        if (resumeNextIdx >= 0)
            return TransformWithResumeNext(stmts, resumeNextIdx);

        return stmts.SelectMany(TransformStatement).Where(s => s != null).Select(s => s!);
    }

    // Restructures a method body containing "On Error GoTo <label>" into a try/catch.
    //
    // VB6 pattern:
    //   [pre-stmts]
    //   On Error GoTo ErrH
    //   [protected-stmts]
    //   Exit Sub / Exit Function
    // ErrH:
    //   [handler-stmts]
    //
    // C# output:
    //   [pre-stmts]
    //   try { [protected-stmts] }
    //   catch (Exception ex) { [handler-stmts] }
    private IEnumerable<StatementSyntax> TransformWithGoToHandler(List<VB6Statement> stmts, int onErrorIdx)
    {
        var onError = (OnErrorStmt)stmts[onErrorIdx];
        string handlerLabel = onError.Label!;

        int labelIdx = stmts.FindIndex(
            s => s is LabelStmt ls &&
                 string.Equals(ls.Name, handlerLabel, StringComparison.OrdinalIgnoreCase));

        if (labelIdx < 0)
        {
            // Label not in this scope — emit comment and fall through normally
            yield return EmptyStatement().WithLeadingTrivia(
                Comment($"// TODO: On Error GoTo {handlerLabel} — handler label not found in this scope"));
            foreach (var s in stmts.Skip(onErrorIdx + 1)
                         .SelectMany(TransformStatement).Where(x => x != null).Select(x => x!))
                yield return s;
            yield break;
        }

        // Statements before "On Error GoTo" — no error handling active yet
        foreach (var s in stmts.Take(onErrorIdx)
                     .SelectMany(TransformStatement).Where(x => x != null).Select(x => x!))
            yield return s;

        // Try body: everything between "On Error GoTo" and the handler label
        var tryStmts = stmts.Skip(onErrorIdx + 1).Take(labelIdx - onErrorIdx - 1).ToList();

        // Catch body: everything after the handler label
        var catchStmts = stmts.Skip(labelIdx + 1).ToList();

        // Detect the "GoTo Done / Error: handler / Resume Done / Done: cleanup" pattern.
        // If a Resume <label> or GoTo <label> in the catch body targets a LabelStmt that
        // also appears in the catch body, the code at/after that label is "cleanup" that
        // must run after both the normal and error paths — it belongs OUTSIDE the try/catch.
        List<VB6Statement> afterStmts = [];
        string? cleanupLabel = null;
        int cleanupLabelIdx = -1;

        foreach (var s in catchStmts)
        {
            string? candidate = s switch
            {
                ResumeStmt rs when rs.Kind == ResumeKind.Label => rs.Label,
                GoToStmt gts => gts.Label,
                _ => null
            };
            if (candidate == null) continue;
            int idx = catchStmts.FindIndex(c => c is LabelStmt ls &&
                string.Equals(ls.Name, candidate, StringComparison.OrdinalIgnoreCase));
            if (idx >= 0) { cleanupLabelIdx = idx; cleanupLabel = candidate; break; }
        }

        if (cleanupLabelIdx >= 0)
        {
            // Split: cleanup label and everything after it moves outside the try/catch
            afterStmts = catchStmts.Skip(cleanupLabelIdx).ToList();
            catchStmts = catchStmts.Take(cleanupLabelIdx).ToList();

            // Strip trailing Resume <cleanupLabel> from the catch body (implicit fall-through)
            while (catchStmts.Count > 0 && catchStmts[^1] is ResumeStmt)
                catchStmts.RemoveAt(catchStmts.Count - 1);

            // Strip the final GoTo <cleanupLabel> from the try body (implicit fall-through)
            if (tryStmts.Count > 0 && tryStmts[^1] is GoToStmt lastGoto &&
                string.Equals(lastGoto.Label, cleanupLabel, StringComparison.OrdinalIgnoreCase))
                tryStmts.RemoveAt(tryStmts.Count - 1);
        }
        else
        {
            while (catchStmts.Count > 0 && catchStmts[^1] is ResumeStmt)
                catchStmts.RemoveAt(catchStmts.Count - 1);
        }

        var tryBlock = Block(List(
            tryStmts.SelectMany(TransformStatement).Where(x => x != null).Select(x => x!)));

        var catchBodyStmts = catchStmts.SelectMany(s =>
        {
            // Resume / Resume Next inside handler: suppress (already in catch flow)
            if (s is ResumeStmt rs)
            {
                return rs.Kind == ResumeKind.Next
                    ? [(StatementSyntax)EmptyStatement().WithLeadingTrivia(Comment("// Resume Next"))]
                    : Array.Empty<StatementSyntax>();
            }
            return TransformStatement(s).Where(x => x != null).Select(x => x!);
        }).ToList();

        var catchBlock = Block(List(catchBodyStmts));

        var catchClause = CatchClause()
            .WithDeclaration(CatchDeclaration(
                IdentifierName("Exception"),
                Identifier(CatchVar)))
            .WithBlock(catchBlock);

        yield return TryStatement(tryBlock, SingletonList(catchClause), null);

        // Emit cleanup code (the Done: label and whatever follows) outside the try/catch
        foreach (var s in afterStmts.SelectMany(TransformStatement).Where(x => x != null).Select(x => x!))
            yield return s;
    }

    // Restructures "On Error Resume Next" into try { remaining } catch { /* suppressed */ }.
    private IEnumerable<StatementSyntax> TransformWithResumeNext(List<VB6Statement> stmts, int resumeNextIdx)
    {
        // Statements before "On Error Resume Next" — translate normally
        foreach (var s in stmts.Take(resumeNextIdx)
                     .SelectMany(TransformStatement).Where(x => x != null).Select(x => x!))
            yield return s;

        var remaining = stmts.Skip(resumeNextIdx + 1).ToList();
        if (remaining.Count == 0) yield break;

        var tryBlock = Block(List(
            remaining.SelectMany(TransformStatement).Where(x => x != null).Select(x => x!)));

        var catchBlock = Block(SingletonList<StatementSyntax>(
            EmptyStatement().WithLeadingTrivia(
                Comment("// On Error Resume Next: exceptions suppressed"))));

        yield return TryStatement(tryBlock, SingletonList(CatchClause().WithBlock(catchBlock)), null);
    }

    private IEnumerable<StatementSyntax> TransformStatement(VB6Statement stmt)
    {
        switch (stmt)
        {
            case EmptyStmt:
                yield break;

            case LabelStmt ls:
                yield return LabeledStatement(
                    TypeMapper.SanitizeName(ls.Name),
                    EmptyStatement());
                break;

            case LocalDimStmt lds:
                foreach (var s in TransformLocalDim(lds)) yield return s;
                break;

            case LocalConstStmt lcs:
                yield return TransformLocalConst(lcs);
                break;

            case AssignStmt ass:
                foreach (var s in TransformAssignStmt(ass)) yield return s;
                break;

            case CallStmt cs when IsErrMethod(cs.Target, "Raise"):
            {
                // Err.Raise number[, source[, description]] → throw new Exception(description)
                var args = cs.Arguments;
                var msg = args.Count >= 3
                    ? TransformExpr(args[2])
                    : args.Count >= 1
                        ? TransformExpr(args[0])
                        : (ExpressionSyntax)LiteralExpression(SyntaxKind.StringLiteralExpression, Literal("Error"));
                yield return ThrowStatement(
                    ObjectCreationExpression(IdentifierName("Exception"))
                        .WithArgumentList(ArgumentList(SingletonSeparatedList(Argument(msg)))));
                break;
            }

            case CallStmt cs when IsErrMethod(cs.Target, "Clear"):
                yield return EmptyStatement().WithLeadingTrivia(Comment("// Err.Clear()"));
                break;

            case CallStmt cs:
                yield return ExpressionStatement(TransformCall(cs.Target, cs.Arguments));
                break;

            case IfStmt ifs:
                yield return TransformIf(ifs);
                break;

            case ForStmt fs:
                yield return TransformFor(fs);
                break;

            case ForEachStmt fes:
                yield return TransformForEach(fes);
                break;

            case WhileStmt ws:
                yield return TransformWhile(ws);
                break;

            case DoStmt ds:
                yield return TransformDo(ds);
                break;

            case SelectStmt ss:
                yield return TransformSelect(ss);
                break;

            case WithStmt wts:
                foreach (var s in TransformWith(wts)) yield return s;
                break;

            case ExitStmt es:
                yield return TransformExit(es);
                break;

            case GoToStmt gt:
                yield return GotoStatement(SyntaxKind.GotoStatement,
                    IdentifierName(TypeMapper.SanitizeName(gt.Label)));
                break;

            case GoSubStmt gss:
                // GoSub not directly available in C# — emit a method call placeholder
                yield return ExpressionStatement(
                    InvocationExpression(IdentifierName(TypeMapper.SanitizeName(gss.Label))));
                break;

            case ReturnStmt:
                yield return ReturnStatement();
                break;

            case OnErrorStmt oe:
                foreach (var s in TransformOnError(oe)) yield return s;
                break;

            case ResumeStmt re:
                // Resume is handled by try/catch in OnError; emit goto if label given
                if (re.Kind == ResumeKind.Label && re.Label != null)
                    yield return GotoStatement(SyntaxKind.GotoStatement,
                        IdentifierName(TypeMapper.SanitizeName(re.Label)));
                break;

            case RaiseEventStmt rev:
                yield return TransformRaiseEvent(rev);
                break;

            case ReDimStmt rdm:
                yield return TransformReDim(rdm);
                break;

            case EraseStmt er:
                foreach (var s in er.Variables.Select(v =>
                    (StatementSyntax)ExpressionStatement(
                        AssignmentExpression(SyntaxKind.SimpleAssignmentExpression,
                            IdentifierName(TypeMapper.SanitizeName(v)),
                            LiteralExpression(SyntaxKind.NullLiteralExpression)))))
                    yield return s;
                break;

            case OpenStmt os:
                yield return TransformOpen(os);
                break;

            case CloseStmt cls:
                foreach (var s in TransformClose(cls)) yield return s;
                break;

            case LineInputStmt lis:
                yield return ExpressionStatement(
                    AssignmentExpression(SyntaxKind.SimpleAssignmentExpression,
                        TransformExpr(lis.Target),
                        BinaryExpression(SyntaxKind.CoalesceExpression,
                            InvocationExpression(MemberAccessExpression(
                                SyntaxKind.SimpleMemberAccessExpression,
                                IdentifierName(FileNumToVar(lis.FileNum)),
                                IdentifierName("ReadLine"))),
                            LiteralExpression(SyntaxKind.StringLiteralExpression, Literal("")))));
                break;

            case PrintFileStmt pfs:
                foreach (var s in TransformPrintFile(pfs)) yield return s;
                break;

            case WriteFileStmt wfs:
                foreach (var s in TransformWriteFile(wfs)) yield return s;
                break;

            case InputFileStmt ifs2:
            {
                var fvar = FileNumToVar(ifs2.FileNum);
                foreach (var v in ifs2.Variables)
                    yield return ExpressionStatement(
                        AssignmentExpression(SyntaxKind.SimpleAssignmentExpression,
                            TransformExpr(v),
                            BinaryExpression(SyntaxKind.CoalesceExpression,
                                InvocationExpression(MemberAccessExpression(
                                    SyntaxKind.SimpleMemberAccessExpression,
                                    IdentifierName(fvar), IdentifierName("ReadLine"))),
                                LiteralExpression(SyntaxKind.StringLiteralExpression, Literal("")))));
                break;
            }

            case GetFileStmt gfs:
                yield return EmptyStatement().WithLeadingTrivia(
                    Comment($"// TODO: Get #{FileNumToVar(gfs.FileNum)} (binary read) into {gfs.Variable}"));
                break;

            case PutFileStmt putfs:
                yield return EmptyStatement().WithLeadingTrivia(
                    Comment($"// TODO: Put #{FileNumToVar(putfs.FileNum)} (binary write)"));
                break;

            case KillStmt ks:
                yield return ExpressionStatement(InvocationExpression(
                    MemberAccessExpression(SyntaxKind.SimpleMemberAccessExpression,
                        IdentifierName("File"), IdentifierName("Delete")),
                    ArgumentList(SingletonSeparatedList(Argument(TransformExpr(ks.Path))))));
                break;

            case NameFileStmt nfs:
                yield return ExpressionStatement(InvocationExpression(
                    MemberAccessExpression(SyntaxKind.SimpleMemberAccessExpression,
                        IdentifierName("File"), IdentifierName("Move")),
                    ArgumentList(SeparatedList(new[]
                    {
                        Argument(TransformExpr(nfs.OldPath)),
                        Argument(TransformExpr(nfs.NewPath))
                    }))));
                break;

            case MkDirStmt mds:
                yield return ExpressionStatement(InvocationExpression(
                    MemberAccessExpression(SyntaxKind.SimpleMemberAccessExpression,
                        IdentifierName("Directory"), IdentifierName("CreateDirectory")),
                    ArgumentList(SingletonSeparatedList(Argument(TransformExpr(mds.Path))))));
                break;

            case RmDirStmt rds:
                yield return ExpressionStatement(InvocationExpression(
                    MemberAccessExpression(SyntaxKind.SimpleMemberAccessExpression,
                        IdentifierName("Directory"), IdentifierName("Delete")),
                    ArgumentList(SingletonSeparatedList(Argument(TransformExpr(rds.Path))))));
                break;

            case FileCopyStmt fcs:
                yield return ExpressionStatement(InvocationExpression(
                    MemberAccessExpression(SyntaxKind.SimpleMemberAccessExpression,
                        IdentifierName("File"), IdentifierName("Copy")),
                    ArgumentList(SeparatedList(new[]
                    {
                        Argument(TransformExpr(fcs.Source)),
                        Argument(TransformExpr(fcs.Dest))
                    }))));
                break;

            case ChDirStmt cds:
                yield return ExpressionStatement(InvocationExpression(
                    MemberAccessExpression(SyntaxKind.SimpleMemberAccessExpression,
                        IdentifierName("Directory"), IdentifierName("SetCurrentDirectory")),
                    ArgumentList(SingletonSeparatedList(Argument(TransformExpr(cds.Path))))));
                break;

            case ChDriveStmt:
                yield return EmptyStatement().WithLeadingTrivia(
                    Comment("// ChDrive — no .NET equivalent; set via Directory.SetCurrentDirectory"));
                break;

            default:
                yield return EmptyStatement().WithTrailingTrivia(
                    Comment($"// TODO: {stmt.GetType().Name}"));
                break;
        }
    }

    // Dim → local variable declaration(s)
    private IEnumerable<StatementSyntax> TransformLocalDim(LocalDimStmt lds)
    {
        foreach (var d in lds.Declarators)
        {
            var typeName = d.Type?.Name ?? "Variant";
            var csType = TypeMapper.ToCSharp(typeName, d.IsArray);

            ExpressionSyntax? initExpr = null;
            if (d.IsArray && d.Dimensions != null && d.Dimensions.Count > 0)
            {
                // new T[dim+1] (VB6 arrays are 0-based by default with upper bound inclusive)
                var sizes = d.Dimensions.Select(dim =>
                    BinaryExpression(SyntaxKind.AddExpression,
                        TransformExpr(dim),
                        LiteralExpression(SyntaxKind.NumericLiteralExpression, Literal(1))))
                    .ToList();

                initExpr = ArrayCreationExpression(
                    ArrayType(TypeMapper.ToCSharp(typeName))
                        .WithRankSpecifiers(SingletonList(ArrayRankSpecifier(
                            SeparatedList<ExpressionSyntax>(sizes)))));
                csType = ArrayType(TypeMapper.ToCSharp(typeName),
                    SingletonList(ArrayRankSpecifier()));
            }

            var varDeclarator = VariableDeclarator(TypeMapper.SanitizeName(d.Name));
            if (initExpr != null)
                varDeclarator = varDeclarator.WithInitializer(EqualsValueClause(initExpr));
            else if (!d.IsArray)
                varDeclarator = varDeclarator.WithInitializer(
                    EqualsValueClause(TypeMapper.DefaultValue(typeName)));

            var varDecl = VariableDeclaration(csType)
                .WithVariables(SingletonSeparatedList(varDeclarator));

            var localDecl = LocalDeclarationStatement(varDecl);
            if (lds.IsStatic)
            {
                // C# has no static locals — move to a class-level static field manually
                yield return EmptyStatement().WithLeadingTrivia(
                    Comment($"// TODO: Static local '{d.Name}' — move to a class-level static field"));
            }

            yield return localDecl;
        }
    }

    private LocalDeclarationStatementSyntax TransformLocalConst(LocalConstStmt lcs)
    {
        var type = lcs.Type != null
            ? TypeMapper.ToCSharp(lcs.Type.Name)
            : PredefinedType(Token(SyntaxKind.ObjectKeyword));
        var varDecl = VariableDeclaration(type)
            .WithVariables(SingletonSeparatedList(
                VariableDeclarator(TypeMapper.SanitizeName(lcs.Name))
                    .WithInitializer(EqualsValueClause(TransformExpr(lcs.Value)))));
        return LocalDeclarationStatement(varDecl)
            .WithModifiers(TokenList(Token(SyntaxKind.ConstKeyword)));
    }

    private IEnumerable<StatementSyntax> TransformAssignStmt(AssignStmt ass)
    {
        // target = CreateObject("ProgID") → two-statement COM expansion
        if (ass.Value is CallExpr { Target: NameExpr { Name: var coName }, Arguments: var coArgs }
            && coName.Equals("CreateObject", StringComparison.OrdinalIgnoreCase)
            && coArgs.Count >= 1)
        {
            string tmpName = $"__comType{(++_tempSeq == 1 ? "" : _tempSeq.ToString())}";
            var progId     = TransformExpr(coArgs[0]);
            var typeLocal  = LocalDeclarationStatement(
                VariableDeclaration(IdentifierName("Type"))
                    .WithVariables(SingletonSeparatedList(
                        VariableDeclarator(tmpName)
                            .WithInitializer(EqualsValueClause(
                                InvocationExpression(
                                    MemberAccessExpression(SyntaxKind.SimpleMemberAccessExpression,
                                        IdentifierName("Type"), IdentifierName("GetTypeFromProgID")))
                                .WithArgumentList(ArgumentList(SingletonSeparatedList(Argument(progId)))))))));

            var activate = InvocationExpression(
                MemberAccessExpression(SyntaxKind.SimpleMemberAccessExpression,
                    IdentifierName("Activator"), IdentifierName("CreateInstance")))
                .WithArgumentList(ArgumentList(SingletonSeparatedList(Argument(IdentifierName(tmpName)))));

            yield return typeLocal;
            yield return ExpressionStatement(
                AssignmentExpression(SyntaxKind.SimpleAssignmentExpression,
                    TransformExpr(ass.Target), activate));
            yield break;
        }

        yield return TransformAssign(ass);
    }

    private ExpressionStatementSyntax TransformAssign(AssignStmt ass)
    {
        // VB6 "FuncName = value" sets the function's return value → assign to __result
        ExpressionSyntax target;
        if (!string.IsNullOrEmpty(_currentFunctionName)
            && ass.Target is NameExpr ne
            && string.Equals(ne.Name, _currentFunctionName, StringComparison.OrdinalIgnoreCase))
        {
            target = IdentifierName(ResultVar);
        }
        else
        {
            target = TransformExpr(ass.Target);
        }
        var value = TransformExpr(ass.Value);
        return ExpressionStatement(
            AssignmentExpression(SyntaxKind.SimpleAssignmentExpression, target, value));
    }

    private ExpressionSyntax TransformCall(VB6Expression target, List<VB6Expression> args)
    {
        // Check if it is a known built-in
        if (target is NameExpr ne)
        {
            var csArgs = args.Select(TransformExpr).ToList();
            var mapped = BuiltinMapper.TryMap(ne.Name, csArgs);
            if (mapped != null) return mapped;
        }

        var targetExpr = TransformExpr(target);
        var argList = ArgumentList(SeparatedList(args.Select(a => Argument(TransformExpr(a)))));
        return InvocationExpression(targetExpr, argList);
    }

    private IfStatementSyntax TransformIf(IfStmt ifs)
    {
        var cond = TransformExpr(ifs.Condition);
        var then = TransformBody(ifs.Then);

        ElseClauseSyntax? elseClause = null;

        // Build else-if chain from right to left
        if (ifs.Else != null)
            elseClause = ElseClause(TransformBody(ifs.Else));

        foreach (var (ec, eb) in ifs.ElseIfs.AsEnumerable().Reverse())
        {
            var eif = IfStatement(TransformExpr(ec), TransformBody(eb));
            elseClause = ElseClause(elseClause != null ? eif.WithElse(elseClause) : eif);
        }

        return IfStatement(cond, then, elseClause);
    }

    private ForStatementSyntax TransformFor(ForStmt fs)
    {
        var varName = TypeMapper.SanitizeName(fs.Variable);

        // Use expression-style initializer to avoid redeclaring a variable already Dim'd
        var init = AssignmentExpression(SyntaxKind.SimpleAssignmentExpression,
            IdentifierName(varName), TransformExpr(fs.Start));

        var cond = BinaryExpression(SyntaxKind.LessThanOrEqualExpression,
            IdentifierName(varName), TransformExpr(fs.End));

        var step = fs.Step != null
            ? (ExpressionSyntax)AssignmentExpression(SyntaxKind.AddAssignmentExpression,
                IdentifierName(varName), TransformExpr(fs.Step))
            : PostfixUnaryExpression(SyntaxKind.PostIncrementExpression,
                IdentifierName(varName));

        return ForStatement(TransformBody(fs.Body))
            .WithInitializers(SingletonSeparatedList<ExpressionSyntax>(init))
            .WithCondition(cond)
            .WithIncrementors(SingletonSeparatedList<ExpressionSyntax>(step));
    }

    private ForEachStatementSyntax TransformForEach(ForEachStmt fes)
    {
        var type = fes.VariableType != null
            ? TypeMapper.ToCSharp(fes.VariableType.Name)
            : IdentifierName("var");
        return ForEachStatement(type,
            TypeMapper.SanitizeName(fes.Variable),
            TransformExpr(fes.Collection),
            TransformBody(fes.Body));
    }

    private WhileStatementSyntax TransformWhile(WhileStmt ws) =>
        WhileStatement(TransformExpr(ws.Condition), TransformBody(ws.Body));

    private StatementSyntax TransformDo(DoStmt ds)
    {
        var body = TransformBody(ds.Body);
        return ds.Kind switch
        {
            DoKind.WhileTop => WhileStatement(TransformExpr(ds.Condition!), body),
            DoKind.UntilTop => WhileStatement(
                PrefixUnaryExpression(SyntaxKind.LogicalNotExpression,
                    ParenthesizedExpression(TransformExpr(ds.Condition!))), body),
            DoKind.WhileBottom => DoStatement(body, TransformExpr(ds.Condition!)),
            DoKind.UntilBottom => DoStatement(body,
                PrefixUnaryExpression(SyntaxKind.LogicalNotExpression,
                    ParenthesizedExpression(TransformExpr(ds.Condition!)))),
            _ => WhileStatement(LiteralExpression(SyntaxKind.TrueLiteralExpression), body)
        };
    }

    private SwitchStatementSyntax TransformSelect(SelectStmt ss)
    {
        var expr = TransformExpr(ss.Expr);
        var sections = new List<SwitchSectionSyntax>();

        foreach (var c in ss.Cases)
        {
            var labels = new List<SwitchLabelSyntax>();
            foreach (var ce in c.Conditions)
            {
                labels.Add(ce switch
                {
                    CaseValueExpr cv => CaseSwitchLabel(TransformExpr(cv.Value)),
                    CaseIsExpr ci => CaseSwitchLabel(TransformExpr(ci.Value)), // simplified
                    CaseRangeExpr _ => DefaultSwitchLabel(),                   // ranges → default
                    _ => DefaultSwitchLabel()
                });
            }

            var stmts = c.Body.SelectMany(TransformStatement).ToList<StatementSyntax>();
            stmts.Add(BreakStatement());

            sections.Add(SwitchSection(List(labels), List(stmts)));
        }

        if (ss.ElseBody != null)
        {
            var elseStmts = ss.ElseBody.SelectMany(TransformStatement).ToList<StatementSyntax>();
            elseStmts.Add(BreakStatement());
            sections.Add(SwitchSection(
                SingletonList<SwitchLabelSyntax>(DefaultSwitchLabel()),
                List(elseStmts)));
        }

        return SwitchStatement(expr, List(sections));
    }

    private IEnumerable<StatementSyntax> TransformWith(WithStmt wts)
    {
        _withStack.Push(TransformExpr(wts.Obj));
        foreach (var s in wts.Body.SelectMany(TransformStatement))
            yield return s;
        _withStack.Pop();
    }

    private StatementSyntax TransformExit(ExitStmt es)
    {
        switch (es.Kind)
        {
            case ExitKind.Function:
            case ExitKind.Property:
                // Return __result if inside a function with a return type
                if (!string.IsNullOrEmpty(_currentFunctionName) && _currentReturnType != null)
                    return ReturnStatement(IdentifierName(ResultVar));
                return ReturnStatement();
            case ExitKind.Sub:
                return ReturnStatement();
            case ExitKind.For:
            case ExitKind.Do:
            case ExitKind.While:
                return BreakStatement();
            default:
                return ReturnStatement();
        }
    }

    // Called only for OnErrorStmt found inside nested blocks (if/for/while bodies).
    // Top-level handlers are restructured by TransformMethodStatements before reaching here.
    private IEnumerable<StatementSyntax> TransformOnError(OnErrorStmt oe)
    {
        if (oe.Kind == OnErrorKind.Zero) yield break;
        if (oe.Kind == OnErrorKind.ResumeNext)
            yield return EmptyStatement().WithLeadingTrivia(
                Comment("// On Error Resume Next — nested handler; see enclosing try/catch"));
        else if (oe.Label != null)
            yield return EmptyStatement().WithLeadingTrivia(
                Comment($"// On Error GoTo {oe.Label} — nested handler not restructured"));
    }

    private ExpressionStatementSyntax TransformRaiseEvent(RaiseEventStmt rev)
    {
        var invocation = ConditionalAccessExpression(
            IdentifierName(TypeMapper.SanitizeName(rev.EventName)),
            InvocationExpression(
                MemberBindingExpression(IdentifierName("Invoke")),
                ArgumentList(SeparatedList(rev.Arguments.Select(a => Argument(TransformExpr(a)))))));
        return ExpressionStatement(invocation);
    }

    private ExpressionStatementSyntax TransformReDim(ReDimStmt rdm)
    {
        var sizes = rdm.Dimensions.Select(dim =>
            BinaryExpression(SyntaxKind.AddExpression,
                TransformExpr(dim),
                LiteralExpression(SyntaxKind.NumericLiteralExpression, Literal(1))))
            .ToList();

        var typeName = rdm.Type?.Name ?? "object";
        ExpressionSyntax rhs = ArrayCreationExpression(
            ArrayType(TypeMapper.ToCSharp(typeName))
                .WithRankSpecifiers(SingletonList(ArrayRankSpecifier(
                    SeparatedList<ExpressionSyntax>(sizes)))));

        return ExpressionStatement(
            AssignmentExpression(SyntaxKind.SimpleAssignmentExpression,
                IdentifierName(TypeMapper.SanitizeName(rdm.Variable)), rhs));
    }

    // ─── Expression transformation ────────────────────────────────────────

    private ExpressionSyntax TransformExpr(VB6Expression expr)
    {
        return expr switch
        {
            LiteralExpr le => TransformLiteral(le),
            NameExpr ne when ne.Name.Equals("vbObjectError", StringComparison.OrdinalIgnoreCase) =>
                CastExpression(
                    PredefinedType(Token(SyntaxKind.IntKeyword)),
                    CheckedExpression(SyntaxKind.UncheckedExpression,
                        LiteralExpression(SyntaxKind.NumericLiteralExpression,
                            Literal("0x80040000", unchecked((int)0x80040000))))),
            NameExpr ne => IdentifierName(TypeMapper.SanitizeName(ne.Name)),
            MeExpr => ThisExpression(),
            // Err.Number / Err.Description / Err.Source → caught exception properties
            MemberExpr meErr when meErr.Object is NameExpr errObj &&
                string.Equals(errObj.Name, "Err", StringComparison.OrdinalIgnoreCase) =>
                MapErrMember(meErr.Member),
            MemberExpr me => MemberAccessExpression(SyntaxKind.SimpleMemberAccessExpression,
                TransformExpr(me.Object), IdentifierName(TypeMapper.SanitizeName(me.Member))),
            WithMemberExpr wme => _withStack.Count > 0
                ? MemberAccessExpression(SyntaxKind.SimpleMemberAccessExpression,
                    _withStack.Peek(), IdentifierName(TypeMapper.SanitizeName(wme.Member)))
                : IdentifierName(TypeMapper.SanitizeName(wme.Member)),
            CallExpr ce => TransformCallExpr(ce),
            BinaryExpr be => TransformBinary(be),
            UnaryExpr ue => TransformUnary(ue),
            NewExpr ne2 => ObjectCreationExpression(
                    IdentifierName(TypeMapper.SanitizeName(ne2.Type.Name)))
                .WithArgumentList(ArgumentList()),
            TypeOfExpr toe => IsPatternExpression(TransformExpr(toe.Obj),
                TypePattern(TypeMapper.ToCSharp(toe.Type.Name))),
            AddressOfExpr aoe => TransformExpr(aoe.Target),
            _ => IdentifierName("__unknown")
        };
    }

    private ExpressionSyntax TransformCallExpr(CallExpr ce)
    {
        // Check for VB6 built-in
        if (ce.Target is NameExpr nameRef)
        {
            var csArgs = ce.Arguments.Select(TransformExpr).ToList();
            var mapped = BuiltinMapper.TryMap(nameRef.Name, csArgs);
            if (mapped != null) return mapped;
        }
        var target = TransformExpr(ce.Target);
        var args = ArgumentList(SeparatedList(ce.Arguments.Select(a => Argument(TransformExpr(a)))));
        return InvocationExpression(target, args);
    }

    private ExpressionSyntax TransformBinary(BinaryExpr be)
    {
        var left = TransformExpr(be.Left);
        var right = TransformExpr(be.Right);

        return be.Op switch
        {
            "+"  => BinaryExpression(SyntaxKind.AddExpression, left, right),
            "-"  => BinaryExpression(SyntaxKind.SubtractExpression, left, right),
            "*"  => BinaryExpression(SyntaxKind.MultiplyExpression, left, right),
            "/"  => BinaryExpression(SyntaxKind.DivideExpression, left, right),
            "\\" => BinaryExpression(SyntaxKind.DivideExpression,
                CastExpression(PredefinedType(Token(SyntaxKind.IntKeyword)), left),
                CastExpression(PredefinedType(Token(SyntaxKind.IntKeyword)), right)),
            "^"  => InvocationExpression(
                MemberAccessExpression(SyntaxKind.SimpleMemberAccessExpression,
                    IdentifierName("Math"), IdentifierName("Pow")),
                ArgumentList(SeparatedList(new[] { Argument(left), Argument(right) }))),
            "Mod" => BinaryExpression(SyntaxKind.ModuloExpression, left, right),
            "&"  => InvocationExpression(
                MemberAccessExpression(SyntaxKind.SimpleMemberAccessExpression,
                    IdentifierName("string"), IdentifierName("Concat")),
                ArgumentList(SeparatedList(new[] { Argument(left), Argument(right) }))),
            "="  => BinaryExpression(SyntaxKind.EqualsExpression, left, right),
            "<>" => BinaryExpression(SyntaxKind.NotEqualsExpression, left, right),
            "<"  => BinaryExpression(SyntaxKind.LessThanExpression, left, right),
            ">"  => BinaryExpression(SyntaxKind.GreaterThanExpression, left, right),
            "<=" => BinaryExpression(SyntaxKind.LessThanOrEqualExpression, left, right),
            ">=" => BinaryExpression(SyntaxKind.GreaterThanOrEqualExpression, left, right),
            "And" => BinaryExpression(SyntaxKind.BitwiseAndExpression, left, right),
            "Or"  => BinaryExpression(SyntaxKind.BitwiseOrExpression, left, right),
            "Xor" => BinaryExpression(SyntaxKind.ExclusiveOrExpression, left, right),
            "Is"  => BinaryExpression(SyntaxKind.EqualsExpression, left,
                LiteralExpression(SyntaxKind.NullLiteralExpression)),
            "IsNot" => BinaryExpression(SyntaxKind.NotEqualsExpression, left,
                LiteralExpression(SyntaxKind.NullLiteralExpression)),
            "Like" => InvocationExpression(
                MemberAccessExpression(SyntaxKind.SimpleMemberAccessExpression,
                    IdentifierName("Regex"), IdentifierName("IsMatch")),
                ArgumentList(SeparatedList(new[] { Argument(left), Argument(right) }))),
            "Eqv" => BinaryExpression(SyntaxKind.EqualsExpression, left, right),
            "Imp" => BinaryExpression(SyntaxKind.LogicalOrExpression,
                PrefixUnaryExpression(SyntaxKind.LogicalNotExpression, left), right),
            _ => BinaryExpression(SyntaxKind.AddExpression, left, right)
        };
    }

    private ExpressionSyntax TransformUnary(UnaryExpr ue)
    {
        var operand = TransformExpr(ue.Operand);
        return ue.Op switch
        {
            "-" => PrefixUnaryExpression(SyntaxKind.UnaryMinusExpression, operand),
            "Not" => PrefixUnaryExpression(SyntaxKind.LogicalNotExpression, operand),
            _ => operand
        };
    }

    private ExpressionSyntax TransformLiteral(LiteralExpr le)
    {
        return le.Kind switch
        {
            LiteralKind.Boolean => le.Value is true
                ? LiteralExpression(SyntaxKind.TrueLiteralExpression)
                : LiteralExpression(SyntaxKind.FalseLiteralExpression),
            LiteralKind.Nothing or LiteralKind.Null =>
                LiteralExpression(SyntaxKind.NullLiteralExpression),
            LiteralKind.Empty =>
                LiteralExpression(SyntaxKind.NullLiteralExpression),
            LiteralKind.String =>
                LiteralExpression(SyntaxKind.StringLiteralExpression,
                    Literal(le.Value?.ToString() ?? "")),
            LiteralKind.Integer =>
                LiteralExpression(SyntaxKind.NumericLiteralExpression,
                    Literal(unchecked((short)Convert.ToInt64(le.Value)))),
            LiteralKind.Long =>
                LiteralExpression(SyntaxKind.NumericLiteralExpression,
                    Literal(unchecked((int)Convert.ToInt64(le.Value)))),
            LiteralKind.Single =>
                LiteralExpression(SyntaxKind.NumericLiteralExpression,
                    Literal(Convert.ToSingle(le.Value))),
            LiteralKind.Double =>
                LiteralExpression(SyntaxKind.NumericLiteralExpression,
                    Literal(Convert.ToDouble(le.Value))),
            LiteralKind.Date => InvocationExpression(
                MemberAccessExpression(SyntaxKind.SimpleMemberAccessExpression,
                    IdentifierName("DateTime"), IdentifierName("Parse")),
                ArgumentList(SingletonSeparatedList(Argument(
                    LiteralExpression(SyntaxKind.StringLiteralExpression,
                        Literal(le.Value?.ToString() ?? "")))))),
            _ => LiteralExpression(SyntaxKind.NullLiteralExpression)
        };
    }

    // ─── File I/O helpers ─────────────────────────────────────────────────

    // Maps a VB6 file-number expression to a C# local variable name.
    // #1 → _file1,  #fileNum → fileNum (variable)
    private static string FileNumToVar(VB6Expression fileNum) => fileNum switch
    {
        LiteralExpr { Kind: LiteralKind.Integer, Value: var n } => $"_file{n}",
        LiteralExpr { Kind: LiteralKind.Long,    Value: var n } => $"_file{n}",
        NameExpr ne => ne.Name,
        _ => "_file"
    };

    // Open path For mode As #n  →  var _fileN = new StreamReader/Writer(path);
    private LocalDeclarationStatementSyntax TransformOpen(OpenStmt os)
    {
        var varName = FileNumToVar(os.FileNum);
        var path = TransformExpr(os.FilePath);

        ExpressionSyntax init = os.Mode.ToUpperInvariant() switch
        {
            "OUTPUT" => ObjectCreationExpression(IdentifierName("StreamWriter"))
                            .WithArgumentList(ArgumentList(
                                SingletonSeparatedList(Argument(path)))),
            "APPEND" => ObjectCreationExpression(IdentifierName("StreamWriter"))
                            .WithArgumentList(ArgumentList(SeparatedList(new[]
                            {
                                Argument(path),
                                Argument(LiteralExpression(SyntaxKind.TrueLiteralExpression))
                            }))),
            "BINARY" or "RANDOM" => InvocationExpression(
                                MemberAccessExpression(SyntaxKind.SimpleMemberAccessExpression,
                                    IdentifierName("File"), IdentifierName("Open")),
                                ArgumentList(SeparatedList(new[]
                                {
                                    Argument(path),
                                    Argument(MemberAccessExpression(
                                        SyntaxKind.SimpleMemberAccessExpression,
                                        IdentifierName("FileMode"),
                                        IdentifierName("OpenOrCreate")))
                                }))),
            _ => // INPUT (default)
                ObjectCreationExpression(IdentifierName("StreamReader"))
                    .WithArgumentList(ArgumentList(
                        SingletonSeparatedList(Argument(path))))
        };

        return LocalDeclarationStatement(
            VariableDeclaration(IdentifierName("var"))
                .WithVariables(SingletonSeparatedList(
                    VariableDeclarator(Identifier(varName))
                        .WithInitializer(EqualsValueClause(init)))))
            .WithUsingKeyword(Token(SyntaxKind.UsingKeyword));
    }

    private IEnumerable<StatementSyntax> TransformClose(CloseStmt cls)
    {
        if (cls.FileNums.Count == 0)
        {
            yield return EmptyStatement().WithLeadingTrivia(
                Comment("// Close — dispose all open file variables manually"));
            yield break;
        }
        foreach (var fn in cls.FileNums)
            yield return ExpressionStatement(InvocationExpression(
                MemberAccessExpression(SyntaxKind.SimpleMemberAccessExpression,
                    IdentifierName(FileNumToVar(fn)), IdentifierName("Dispose"))));
    }

    private IEnumerable<StatementSyntax> TransformPrintFile(PrintFileStmt pfs)
    {
        var fvar = FileNumToVar(pfs.FileNum);
        if (pfs.Values.Count == 0)
        {
            yield return ExpressionStatement(InvocationExpression(
                MemberAccessExpression(SyntaxKind.SimpleMemberAccessExpression,
                    IdentifierName(fvar), IdentifierName("WriteLine"))));
            yield break;
        }
        foreach (var v in pfs.Values)
            yield return ExpressionStatement(InvocationExpression(
                MemberAccessExpression(SyntaxKind.SimpleMemberAccessExpression,
                    IdentifierName(fvar), IdentifierName("Write")),
                ArgumentList(SingletonSeparatedList(Argument(TransformExpr(v))))));
    }

    private IEnumerable<StatementSyntax> TransformWriteFile(WriteFileStmt wfs)
    {
        var fvar = FileNumToVar(wfs.FileNum);
        if (wfs.Values.Count == 0)
        {
            yield return ExpressionStatement(InvocationExpression(
                MemberAccessExpression(SyntaxKind.SimpleMemberAccessExpression,
                    IdentifierName(fvar), IdentifierName("WriteLine"))));
            yield break;
        }
        foreach (var v in wfs.Values)
            yield return ExpressionStatement(InvocationExpression(
                MemberAccessExpression(SyntaxKind.SimpleMemberAccessExpression,
                    IdentifierName(fvar), IdentifierName("WriteLine")),
                ArgumentList(SingletonSeparatedList(Argument(TransformExpr(v))))));
    }

    // ─── Err object helpers ───────────────────────────────────────────────

    // Maps VB6 Err.xxx → the caught exception variable's equivalent property.
    // The generated catch clause always binds the exception to CatchVar ("ex").
    private static ExpressionSyntax MapErrMember(string member) =>
        member.ToLowerInvariant() switch
        {
            "number"      => MemberAccessExpression(SyntaxKind.SimpleMemberAccessExpression,
                                 IdentifierName(CatchVar), IdentifierName("HResult")),
            "description" => MemberAccessExpression(SyntaxKind.SimpleMemberAccessExpression,
                                 IdentifierName(CatchVar), IdentifierName("Message")),
            "source"      => MemberAccessExpression(SyntaxKind.SimpleMemberAccessExpression,
                                 IdentifierName(CatchVar), IdentifierName("Source")),
            "lastdllerror" => InvocationExpression(
                                 MemberAccessExpression(SyntaxKind.SimpleMemberAccessExpression,
                                     IdentifierName("Marshal"), IdentifierName("GetLastWin32Error"))),
            _ => MemberAccessExpression(SyntaxKind.SimpleMemberAccessExpression,
                     IdentifierName(CatchVar), IdentifierName(TypeMapper.SanitizeName(member)))
        };

    private static bool IsErrMethod(VB6Expression target, string method) =>
        target is MemberExpr me &&
        me.Object is NameExpr { Name: var objName } &&
        string.Equals(objName, "Err", StringComparison.OrdinalIgnoreCase) &&
        string.Equals(me.Member, method, StringComparison.OrdinalIgnoreCase);

    // ─── Parameter / modifier helpers ─────────────────────────────────────

    private ParameterListSyntax BuildParams(List<VB6Parameter> parms)
    {
        var csParams = parms.Select(p =>
        {
            var type = TypeMapper.ToCSharp(p.Type.Name, p.Type.IsArray);
            var param = Parameter(Identifier(TypeMapper.SanitizeName(p.Name)))
                .WithType(type);

            if (p.IsParamArray)
                param = param.WithModifiers(TokenList(Token(SyntaxKind.ParamsKeyword)));
            else if (p.IsByRef && !p.IsOptional)
                param = param.WithModifiers(TokenList(Token(SyntaxKind.RefKeyword)));

            if (p.IsOptional || p.DefaultValue != null)
            {
                var defVal = p.DefaultValue != null
                    ? TransformExpr(p.DefaultValue)
                    : TypeMapper.DefaultValue(p.Type.Name);
                param = param.WithDefault(EqualsValueClause(defVal));
            }

            return param;
        });

        return ParameterList(SeparatedList(csParams));
    }

    private static SyntaxTokenList BuildModifiers(AccessModifier access, bool isStatic)
    {
        var tokens = new List<SyntaxToken>
        {
            access switch
            {
                AccessModifier.Public => Token(SyntaxKind.PublicKeyword),
                AccessModifier.Private => Token(SyntaxKind.PrivateKeyword),
                AccessModifier.Friend => Token(SyntaxKind.InternalKeyword),
                _ => Token(SyntaxKind.PublicKeyword)
            }
        };
        if (isStatic) tokens.Add(Token(SyntaxKind.StaticKeyword));
        return TokenList(tokens);
    }
}
