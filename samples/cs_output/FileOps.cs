// Converted from FileOps.bas by vb6cs
// Date: 2026-05-11 03:23

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Runtime.InteropServices;

namespace Converted;
public static class FileOps
{
    public static void ReadFile(string path)
    {
        short fileNum = 0;
        string line = "";
        fileNum = 1;
        try
        {
            using var fileNum = new StreamReader(path);
            while (!fileNum.EndOfStream)
            {
                line = fileNum.ReadLine() ?? "";
                Debug.Print(line);
            }

            fileNum.Dispose();
            return;
        }
        catch (Exception ex)
        {
            Console.WriteLine(string.Concat("Error reading file: ", ex.Message));
            fileNum.Dispose();
        }
    }

    public static void WriteFile(string path, string content)
    {
        short fileNum = 0;
        fileNum = 1;
        using var _file1 = new StreamWriter(path);
        _file1.Write(content);
        _file1.Dispose();
    }

    public static void AppendFile(string path, string line)
    {
        using var _file2 = new StreamWriter(path, true);
        _file2.Write(line);
        _file2.Dispose();
    }

    public static void DeleteFile(string path)
    {
        try
        {
            File.Delete(path);
        }
        catch
        {
            // On Error Resume Next: exceptions suppressed
        }
    }

    public static void RenameFile(string oldPath, string newPath)
    {
        File.Move(oldPath, newPath);
    }

    public static void CreateFolder(string path)
    {
        Directory.CreateDirectory(path);
    }

    public static void DeleteFolder(string path)
    {
        Directory.Delete(path);
    }

    public static void CopyFile(string src, string dst)
    {
        File.Copy(src, dst);
    }
}