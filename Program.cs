// See https://aka.ms/new-console-template for more information
using System;
using System.IO;
using System.Collections;
using System.Security.Cryptography;
using System.Text;
using System.Reflection.Metadata.Ecma335;
using Xceed.Words.NET;
using Markdig;


public class RecursiveFileProcessor
{
    public static void Main(string[] args)
    {
        Console.WriteLine("MarkdownToWord - Converts a folder structure filled with markdown files to a new folder structure mirror with docx/word files.");
        Console.WriteLine("(C) Daniel 'btk' Kirstenpfad 2024");

        if (args == null || args.Length != 2)
        {
            // no arguments given
            Console.WriteLine("Syntax:");
            Console.WriteLine("markdown2docx <input directory> <output directory>");
            return;
        }
        else
        {
            // enough arguments given
            var inputpath = args[0];
            var outputpath = args[1];
            Console.WriteLine("InputPath: " + inputpath);
            Console.WriteLine("OutputPath: " + outputpath);

            if (File.Exists(inputpath))
            {
                // This path is a file
                Console.WriteLine("Input cannot be a file.");
                return;
            }
            else if (Directory.Exists(inputpath))
            {
                // This path is a directory
                ProcessDirectory(inputpath, args[1], args[0]);
            }
            else
            {
                Console.WriteLine("{0} is not a valid file or directory.", inputpath);
            }
        }
    }

    // Process all files in the directory passed in, recurse on any directories 
    // that are found, and process the files they contain.
    public static void ProcessDirectory(string targetDirectory, string outputpath, string originalinputpath)
    {
        // Process the list of files found in the directory.
        string[] fileEntries = Directory.GetFiles(targetDirectory);
        foreach (string fileName in fileEntries)
            ProcessFile(fileName.Remove(targetDirectory.Length+1),fileName.Remove(0, targetDirectory.Length + 1),outputpath, originalinputpath);

        // Recurse into subdirectories of this directory.
        string[] subdirectoryEntries = Directory.GetDirectories(targetDirectory);
        foreach (string subdirectory in subdirectoryEntries)
        {
            // only consider those without a leading .
            if (subdirectory.Remove(0,targetDirectory.Length+1).StartsWith(".") == false)
            {
                ProcessDirectory(subdirectory,outputpath, originalinputpath);
            }            
        }            
    }
    internal static string GetStringSha256Hash(string text)
    {
        if (String.IsNullOrEmpty(text))
            return String.Empty;

        return Convert.ToBase64String(SHA256.HashData(Encoding.UTF8.GetBytes(text)));
    }

    // Insert logic for processing found files here.
    public static void ProcessFile(string inputpath, string filename, string outputpath, string originalinputpath)
    {
        // only consider .md files
        try
        {
            if (filename.EndsWith(".md"))
            {
                // remove...
                String InputFileRawContents = File.ReadAllText(inputpath + Path.DirectorySeparatorChar + filename);
                DateTime CreationDateTime = File.GetCreationTime(inputpath+ Path.DirectorySeparatorChar + filename);
                DateTime ModifiedDateTime = File.GetLastWriteTime(filename+ Path.DirectorySeparatorChar + filename);
                String FullOutputPath = outputpath+inputpath.Remove(0,originalinputpath.Length);
                String HashFileName = FullOutputPath + filename.Replace(".md",".hash");
                String FullOutputPathFilenameDocx = FullOutputPath + filename.Replace(".md", ".docx");

                // create target directory if it does not already exist
                if (!Directory.Exists(FullOutputPath))
                {
                    Directory.CreateDirectory(FullOutputPath);
                }

                // check if file exists and has changed...
                if (File.Exists(HashFileName))
                {
                    // check hash
                    string SourceFileHash = GetStringSha256Hash((string)InputFileRawContents);
                    string TargetFileHash = (string)File.ReadAllText(HashFileName);
                    if (SourceFileHash.Equals(TargetFileHash))
                    {
                        //Console.WriteLine("Did not change: " + HashFileName);
                        return;
                    } else
                    {
                        Console.WriteLine("Updating: "+ FullOutputPathFilenameDocx);
                        WriteFile(InputFileRawContents, HashFileName, FullOutputPathFilenameDocx, CreationDateTime, ModifiedDateTime);
                    }
                }
                else // we have not seen this file yet
                {
                    Console.WriteLine("Processing: " + inputpath + filename + " ===> " + FullOutputPathFilenameDocx);
                    WriteFile(InputFileRawContents, HashFileName, FullOutputPathFilenameDocx, CreationDateTime, ModifiedDateTime);
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine(ex.ToString()); 
        }
        //Console.ReadLine();
    }

    public static void WriteFile(string InputFileRawContents, string HashFileName, string FullOutputPathFilenameDocx, DateTime CreationTime, DateTime ModifiedDateTime)
    {
        String MarkdownRemovedContents = Markdown.ToPlainText(InputFileRawContents).Replace("[[", "").Replace("]]", "");
        File.WriteAllText(HashFileName, GetStringSha256Hash((string)InputFileRawContents));
        var doc = DocX.Create(FullOutputPathFilenameDocx);
        var p = doc.InsertParagraph((string)MarkdownRemovedContents);
        doc.Save();
        File.SetCreationTime(FullOutputPathFilenameDocx, CreationTime);
        File.SetLastWriteTime(FullOutputPathFilenameDocx, ModifiedDateTime);
    }
}
