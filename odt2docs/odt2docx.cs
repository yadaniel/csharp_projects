using System;
using System.Linq;
using System.IO;
using System.IO.Compression;
using Microsoft.Office.Interop.Word;
using Word = Microsoft.Office.Interop.Word;

class App {

    static string[] required_files = {
        "content.xml",
        "meta.xml",
        "mimetype",
        "settings.xml",
        "styles.xml",
    };

    static string[] required_dirs = {
        "META-INF"
    };

    static string[] optional_dirs = {
        "media"
    };

    static void Main(string[] args) {

        if(args.Length != 2) {
            Console.WriteLine(@"usage: odt2docx.exe <odt-folder-input> <docx-output>");
        } else {

            string cwd = Directory.GetCurrentDirectory();
            string infolder = args[0];
            string outfile = args[1];
            string infolderPath = $"{cwd}\\{infolder}";
            string outfilePath = $"{cwd}\\{outfile}";
            string outzip = outfile.Replace(".docx", "");
            string outzipPath = $"{cwd}\\{outzip}";
            string zippedDocumentPath = $"{cwd}\\{outzip}.zip";

            Console.WriteLine($"infolder => {infolder}");
            Console.WriteLine($"infolderPath => {infolderPath}");
            Console.WriteLine($"outfile => {outfile}");
            Console.WriteLine($"outfilePath => {outfilePath}");
            Console.WriteLine($"outzip => {outzip}");
            Console.WriteLine($"outzipPath => {outzipPath}");

            // verify odt-output
            if(File.Exists(outfilePath)) {
                Console.WriteLine($"output file already exists => {outfilePath}");
                Environment.Exit(1);
            }

            // verify odt-output
            if(outfile.EndsWith(".docx") == false) {
                Console.WriteLine($"output file not DOCX => {outfile}");
                Environment.Exit(2);
            }

            // verify odt-folder-input
            if(Directory.Exists(infolderPath) == false) {
                Console.WriteLine($"input folder does not exist => {infolderPath}");
                Environment.Exit(3);
            }


            foreach(string fname in required_files) {
                if(File.Exists($"{infolderPath}\\{fname}") == false) {
                    Console.WriteLine($"infolder file missing => {fname}");
                    Environment.Exit(4);
                }
            }

            foreach(string dname in required_dirs) {
                if(Directory.Exists($"{infolderPath}\\{dname}") == false) {
                    Console.WriteLine($"infolder dir missing => {dname}");
                    Environment.Exit(4);
                }
            }

            Directory.CreateDirectory(outzipPath);
            foreach(string fname in required_files) {
                File.Copy($"{infolderPath}\\{fname}", $"{outzipPath}\\{fname}", true);
            }
            foreach(string dname in required_dirs) {
                Directory.CreateDirectory($"{outzipPath}\\{dname}");
                CopyFilesRecursively($"{infolderPath}\\{dname}", $"{outzipPath}\\{dname}");
            }
            foreach(string dname in optional_dirs) {
                if(Directory.Exists($"{infolderPath}\\{dname}")) {
                    Directory.CreateDirectory($"{outzipPath}\\{dname}");
                    CopyFilesRecursively($"{infolderPath}\\{dname}", $"{outzipPath}\\{dname}");
                }
            }

            // do not overwrite existing .zip
            if(File.Exists(zippedDocumentPath)) {
                Console.WriteLine($"outfile exists => {zippedDocumentPath}");
                Environment.Exit(5);
            }

            // do not overwrite existing .odt
            string odtDocumentPath = zippedDocumentPath.Replace(".zip", ".odt");
            if(File.Exists(odtDocumentPath)) {
                Console.WriteLine($"outfile exists => {odtDocumentPath}");
                Environment.Exit(5);
            }

            // do not overwrite existing .docx
            string docxDocumentPath = zippedDocumentPath.Replace(".zip", ".docx");
            if(File.Exists(docxDocumentPath)) {
                Console.WriteLine($"outfile exists => {docxDocumentPath}");
                Environment.Exit(5);
            }

            ZipFile.CreateFromDirectory(outzipPath, zippedDocumentPath);
            File.Move(zippedDocumentPath, odtDocumentPath);
            // Convert(odtDocumentPath, docxDocumentPath, WdSaveFormat.wdFormatXMLTemplate);
            // Convert(odtDocumentPath, docxDocumentPath, WdSaveFormat.wdFormatXML);
            Convert(odtDocumentPath, docxDocumentPath, WdSaveFormat.wdFormatDocumentDefault);

            // cleanup
            Directory.Delete(outzipPath, true);
            File.Delete(zippedDocumentPath);
            // File.Delete(odtDocumentPath);

        }

    }

    // Convert method
    public static void Convert(string input, string output, WdSaveFormat format) {
        // Create an instance of Word.exe
        _Application oWord = new Word.Application {

            // Make this instance of word invisible (Can still see it in the taskmgr).
            Visible = false
        };

        // Interop requires objects.
        object oMissing = System.Reflection.Missing.Value;
        object isVisible = true;
        object readOnly = true;     // Does not cause any word dialog to show up
        //object readOnly = false;  // Causes a word object dialog to show at the end of the conversion
        object oInput = input;
        object oOutput = output;
        object oFormat = format;

        // Load a document into our instance of word.exe
        _Document oDoc = oWord.Documents.Open(
                             ref oInput, ref oMissing, ref readOnly, ref oMissing, ref oMissing,
                             ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing,
                             ref oMissing, ref isVisible, ref oMissing, ref oMissing, ref oMissing, ref oMissing
                         );

        // Make this document the active document.
        oDoc.Activate();

        // Save this document using Word
        oDoc.SaveAs(ref oOutput, ref oFormat, ref oMissing, ref oMissing,
                    ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing,
                    ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing
                   );

        // Always close Word.exe.
        oWord.Quit(ref oMissing, ref oMissing, ref oMissing);
    }

    private static void CopyFilesRecursively(string sourcePath, string targetPath) {
        //Now Create all of the directories
        foreach (string dirPath in Directory.GetDirectories(sourcePath, "*", SearchOption.AllDirectories)) {
            Directory.CreateDirectory(dirPath.Replace(sourcePath, targetPath));
        }

        //Copy all the files & Replaces any files with the same name
        foreach (string newPath in Directory.GetFiles(sourcePath, "*.*", SearchOption.AllDirectories)) {
            File.Copy(newPath, newPath.Replace(sourcePath, targetPath), true);
        }
    }

}

