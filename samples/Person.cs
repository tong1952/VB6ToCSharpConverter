// Converted from Person.cls by vb6cs
// Date: 2026-05-07 23:25

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Runtime.InteropServices;

namespace Converted;
public class Person
{
    private string mName;
    private short mAge;
    private string mEmail;
    public event Action<string, string> NameChanged;
    public event Action<short> AgeChanged;
    public string Name
    {
        get
        {
            string __result = default;
            __result = mName;
            return __result;
        }
    }

    public void Name_Let(string value)
    {
        string oldName = "";
        oldName = mName;
        if (value.Length == 0)
        {
            Err.Raise(5, "Person", "Name cannot be empty");
        }

        mName = value;
        NameChanged?.Invoke(oldName, mName);
    }

    public short Age
    {
        get
        {
            short __result = default;
            __result = mAge;
            return __result;
        }
    }

    public void Age_Let(short value)
    {
        if (value < 0 | value > 150)
        {
            Err.Raise(5, "Person", "Invalid age");
        }

        mAge = value;
        AgeChanged?.Invoke(mAge);
    }

    public string Email
    {
        get
        {
            string __result = default;
            __result = mEmail;
            return __result;
        }
    }

    public void Email_Let(string value)
    {
        mEmail = value;
    }

    public string Greeting()
    {
        string __result = default;
        __result = string.Concat(string.Concat(string.Concat(string.Concat("Hello, my name is ", mName), " and I am "), mAge.ToString()), " years old.");
        return __result;
    }

    public void CopyFrom(ref Person other)
    {
        mName = other.Name;
        mAge = other.Age;
        mEmail = other.Email;
    }

    public bool IsAdult()
    {
        bool __result = default;
        __result = mAge >= 18;
        return __result;
    }

    public string ToString()
    {
        string __result = default;
        string sb = "";
        sb = string.Concat("Person{Name=", mName);
        sb = string.Concat(string.Concat(sb, ", Age="), mAge.ToString());
        sb = string.Concat(string.Concat(string.Concat(sb, ", Email="), mEmail), "}");
        __result = sb;
        return __result;
    }
}