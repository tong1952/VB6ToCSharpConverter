namespace VB6ToCSharp.Parser;

// ─── Root ──────────────────────────────────────────────────────────────────

public record VB6Module(
    ModuleKind Kind,
    string Name,
    List<string> Implements,
    List<VB6MemberNode> Members);

public enum ModuleKind { Standard, Class, Form }
public enum AccessModifier { Public, Private, Friend, Default }

// ─── Member nodes ──────────────────────────────────────────────────────────

public abstract record VB6MemberNode;

public record SubDecl(
    string Name,
    AccessModifier Access,
    bool IsStatic,
    List<VB6Parameter> Parameters,
    List<VB6Statement> Body) : VB6MemberNode;

public record FunctionDecl(
    string Name,
    AccessModifier Access,
    bool IsStatic,
    List<VB6Parameter> Parameters,
    VB6TypeRef ReturnType,
    List<VB6Statement> Body) : VB6MemberNode;

public record PropertyDecl(
    string Name,
    PropertyKind Kind,
    AccessModifier Access,
    List<VB6Parameter> Parameters,
    VB6TypeRef Type,
    List<VB6Statement> Body) : VB6MemberNode;

public enum PropertyKind { Get, Let, Set }

public record FieldDecl(
    AccessModifier Access,
    bool IsStatic,
    List<VB6Declarator> Declarators) : VB6MemberNode;

public record ConstMember(
    string Name,
    VB6TypeRef? Type,
    VB6Expression Value,
    AccessModifier Access) : VB6MemberNode;

public record EventDecl(
    string Name,
    AccessModifier Access,
    List<VB6Parameter> Parameters) : VB6MemberNode;

public record UserTypeDecl(
    string Name,
    AccessModifier Access,
    List<VB6Declarator> Fields,
    bool IsEnum = false) : VB6MemberNode;

public record DeclareDecl(
    string Name,
    AccessModifier Access,
    string Lib,
    string? Alias,
    List<VB6Parameter> Parameters,
    VB6TypeRef? ReturnType) : VB6MemberNode;

// ─── Parameter / declarator ────────────────────────────────────────────────

public record VB6Parameter(
    string Name,
    VB6TypeRef Type,
    bool IsByRef,
    bool IsOptional,
    bool IsParamArray,
    VB6Expression? DefaultValue);

public record VB6Declarator(
    string Name,
    VB6TypeRef? Type,
    bool IsArray,
    List<VB6Expression>? Dimensions,
    bool IsWithEvents = false);

// ─── Type reference ────────────────────────────────────────────────────────

public record VB6TypeRef(string Name, bool IsArray = false, int ArrayRank = 1);

// ─── Statements ────────────────────────────────────────────────────────────

public abstract record VB6Statement;

public record AssignStmt(VB6Expression Target, VB6Expression Value, bool IsSet = false) : VB6Statement;
public record CallStmt(VB6Expression Target, List<VB6Expression> Arguments) : VB6Statement;
public record IfStmt(
    VB6Expression Condition,
    List<VB6Statement> Then,
    List<(VB6Expression Cond, List<VB6Statement> Body)> ElseIfs,
    List<VB6Statement>? Else) : VB6Statement;
public record ForStmt(
    string Variable,
    VB6Expression Start,
    VB6Expression End,
    VB6Expression? Step,
    List<VB6Statement> Body) : VB6Statement;
public record ForEachStmt(
    string Variable,
    VB6TypeRef? VariableType,
    VB6Expression Collection,
    List<VB6Statement> Body) : VB6Statement;
public record WhileStmt(VB6Expression Condition, List<VB6Statement> Body) : VB6Statement;
public record DoStmt(DoKind Kind, VB6Expression? Condition, List<VB6Statement> Body) : VB6Statement;
public enum DoKind { Infinite, WhileTop, UntilTop, WhileBottom, UntilBottom }
public record SelectStmt(VB6Expression Expr, List<VB6CaseClause> Cases, List<VB6Statement>? ElseBody) : VB6Statement;
public record WithStmt(VB6Expression Obj, List<VB6Statement> Body) : VB6Statement;
public record ExitStmt(ExitKind Kind) : VB6Statement;
public enum ExitKind { Sub, Function, For, Do, While, Property }
public record GoToStmt(string Label) : VB6Statement;
public record GoSubStmt(string Label) : VB6Statement;
public record ReturnStmt : VB6Statement;
public record LabelStmt(string Name) : VB6Statement;
public record OnErrorStmt(OnErrorKind Kind, string? Label) : VB6Statement;
public enum OnErrorKind { GoTo, ResumeNext, Zero }
public record ResumeStmt(ResumeKind Kind, string? Label) : VB6Statement;
public enum ResumeKind { Default, Next, Label }
public record RaiseEventStmt(string EventName, List<VB6Expression> Arguments) : VB6Statement;
public record ReDimStmt(string Variable, List<VB6Expression> Dimensions, bool Preserve, VB6TypeRef? Type) : VB6Statement;
public record EraseStmt(List<string> Variables) : VB6Statement;
public record LocalDimStmt(List<VB6Declarator> Declarators, bool IsStatic) : VB6Statement;
public record LocalConstStmt(string Name, VB6TypeRef? Type, VB6Expression Value) : VB6Statement;
public record EmptyStmt : VB6Statement;

// ─── File I/O statements ────────────────────────────────────────────────────

// Open path For {Input|Output|Append|Binary|Random} As [#]n
public record OpenStmt(VB6Expression FilePath, string Mode, VB6Expression FileNum) : VB6Statement;

// Close [#n [, #m ...]]   (empty list = close all)
public record CloseStmt(List<VB6Expression> FileNums) : VB6Statement;

// Line Input #n, var
public record LineInputStmt(VB6Expression FileNum, VB6Expression Target) : VB6Statement;

// Print #n [, expr ...]
public record PrintFileStmt(VB6Expression FileNum, List<VB6Expression> Values) : VB6Statement;

// Write #n [, expr ...]
public record WriteFileStmt(VB6Expression FileNum, List<VB6Expression> Values) : VB6Statement;

// Input #n, var [, var ...]
public record InputFileStmt(VB6Expression FileNum, List<VB6Expression> Variables) : VB6Statement;

// Get [#]n [, [pos]], var
public record GetFileStmt(VB6Expression FileNum, VB6Expression? Position, VB6Expression Variable) : VB6Statement;

// Put [#]n [, [pos]], expr
public record PutFileStmt(VB6Expression FileNum, VB6Expression? Position, VB6Expression Value) : VB6Statement;

// ─── File system statements ──────────────────────────────────────────────────

public record KillStmt(VB6Expression Path) : VB6Statement;              // Delete file
public record NameFileStmt(VB6Expression OldPath, VB6Expression NewPath) : VB6Statement; // Rename
public record MkDirStmt(VB6Expression Path) : VB6Statement;
public record RmDirStmt(VB6Expression Path) : VB6Statement;
public record FileCopyStmt(VB6Expression Source, VB6Expression Dest) : VB6Statement;
public record ChDirStmt(VB6Expression Path) : VB6Statement;
public record ChDriveStmt(VB6Expression Drive) : VB6Statement;

// ─── Case clause ───────────────────────────────────────────────────────────

public record VB6CaseClause(List<VB6CaseExpr> Conditions, List<VB6Statement> Body);

public abstract record VB6CaseExpr;
public record CaseValueExpr(VB6Expression Value) : VB6CaseExpr;
public record CaseRangeExpr(VB6Expression From, VB6Expression To) : VB6CaseExpr;
public record CaseIsExpr(string Op, VB6Expression Value) : VB6CaseExpr;

// ─── Expressions ───────────────────────────────────────────────────────────

public abstract record VB6Expression;

public record LiteralExpr(object? Value, LiteralKind Kind) : VB6Expression;
public enum LiteralKind { Integer, Long, Single, Double, String, Boolean, Date, Nothing, Null, Empty }

public record NameExpr(string Name) : VB6Expression;
public record MeExpr : VB6Expression;
public record MemberExpr(VB6Expression Object, string Member) : VB6Expression;
public record CallExpr(VB6Expression Target, List<VB6Expression> Arguments) : VB6Expression;
public record BinaryExpr(string Op, VB6Expression Left, VB6Expression Right) : VB6Expression;
public record UnaryExpr(string Op, VB6Expression Operand) : VB6Expression;
public record NewExpr(VB6TypeRef Type) : VB6Expression;
public record TypeOfExpr(VB6Expression Obj, VB6TypeRef Type) : VB6Expression;
public record WithMemberExpr(string Member) : VB6Expression;
public record AddressOfExpr(VB6Expression Target) : VB6Expression;
