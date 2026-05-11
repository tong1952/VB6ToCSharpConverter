// Converted from FileOps.bas by vb6cs
// Date: 2026-05-10 02:16

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Runtime.InteropServices;

namespace Converted;
public static class FileOps
{
    public void ReadFile(string path)
    {
        short fileNum = 0;
        string line = "";
        fileNum = 1;
        try
        {
            var fileNum = new StreamReader(path);
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

    public void WriteFile(string path, string content)
    {
        short fileNum = 0;
        fileNum = 1;
        var _file1 = new StreamWriter(path);
        _file1.Write(content);
        _file1.Dispose();
    }

    public void AppendFile(string path, string line)
    {
        var _file2 = new StreamWriter(path, true);
        _file2.Write(line);
        _file2.Dispose();
    }

    public void DeleteFile(string path)
    {
        try
        {
            File.Delete(path);
        }
        catch
        {
            // On Error Resume Next: exceptions suppressed
            ;
        }
    }

    public void RenameFile(string oldPath, string newPath)
    {
        File.Move(oldPath, newPath);
    }

    public void CreateFolder(string path)
    {
        Directory.CreateDirectory(path);
    }

    public void DeleteFolder(string path)
    {
        Directory.Delete(path);
    }

    public void CopyFile(string src, string dst)
    {
        File.Copy(src, dst);
    }
}