open System;
open System.Collections.Generic
open System.IO;
 
// let applicationDirectory = Environment.CurrentDirectory;
let applicationDirectory = "F:\swissedu"
 
let filtered = fun(s:string)->s.EndsWith("docx")
let rec getAllFilesNames directory = 
    seq { yield! (set(Directory.EnumerateFiles(directory))|> Set.filter filtered)
          for d in Directory.EnumerateDirectories(directory) do yield! getAllFilesNames d}
 
let outFile = new StreamWriter(applicationDirectory + @"\swissedu_documents.txt")
 
seq{for d in Directory.EnumerateDirectories(applicationDirectory) do yield! getAllFilesNames d} 
|> Seq.iter (fun (x:string) -> outFile.WriteLine(x.Substring(applicationDirectory.Length+1))) 
 
outFile.Close()