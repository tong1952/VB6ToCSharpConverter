namespace VB6ToCSharp.Lexer;

public enum TokenKind
{
    // Literals
    StringLiteral,
    IntegerLiteral,
    LongLiteral,
    SingleLiteral,
    DoubleLiteral,
    DateLiteral,

    // Declaration keywords
    Dim, As, Sub, Function, Property,
    Get, Let, Set,
    Public, Private, Friend, Static,
    ByRef, ByVal, Optional, ParamArray,
    Const, Enum, Type, Class, Module,
    Declare, Lib, Alias,
    Implements, Event, RaiseEvent,
    Attribute, Option, Explicit, Base, Compare,
    WithEvents,

    // Control flow keywords
    If, Then, Else, ElseIf, End,
    For, Each, In, Next,
    While, Wend,
    Do, Loop, Until, To, Step,
    Select, Case, Is,
    With,
    Exit,
    GoTo, GoSub, Return,
    Call,

    // Error handling
    On, Error, Resume,

    // Memory keywords
    ReDim, Erase, Preserve, New,

    // Literal keywords
    True, False,
    Nothing, Null, Empty,
    Me, MyBase, MyClass,

    // Operator keywords
    Not, And, Or, Xor, Mod, Eqv, Imp, Like,
    TypeOf, AddressOf,

    // Built-in type names
    StringType,   // String (keyword and type name)
    ObjectType,   // Object

    // Identifier
    Identifier,

    // Arithmetic operators
    Plus,           // +
    Minus,          // -
    Star,           // *
    Slash,          // /
    BackSlash,      // \
    Caret,          // ^
    Ampersand,      // &

    // Comparison operators
    Equals,         // =
    LessThan,       // <
    GreaterThan,    // >
    NotEqual,       // <>
    LessEqual,      // <=
    GreaterEqual,   // >=

    // Punctuation
    LeftParen,      // (
    RightParen,     // )
    Comma,          // ,
    Colon,          // :
    Dot,            // .
    Bang,           // !

    // Structure
    NewLine,
    EndOfFile,
}

public record Token(TokenKind Kind, string Text, object? Value, int Line, int Column);
