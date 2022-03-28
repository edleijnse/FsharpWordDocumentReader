open System;
open System.Collections.Generic
open System.IO;
open System.IO.Packaging;
open System.Xml;
open Spire.Doc
open System.Text;
open Spire.Doc.Documents;
// let applicationDirectory = Environment.CurrentDirectory;
let applicationDirectory = "F:\swissedu"
 
let filtered = fun(s:string)->s.EndsWith("docx")
let rec getAllFilesNames directory = 
    seq { yield! (set(Directory.EnumerateFiles(directory))|> Set.filter filtered)
          for d in Directory.EnumerateDirectories(directory) do yield! getAllFilesNames d}
 
let outFile = new StreamWriter(applicationDirectory + @"\swissedu_documents.txt")
let getDocxContent (path: string) =
    use package = Package.Open(path, FileMode.Open)
    let stream = package.GetPart(new Uri("/word/document.xml", UriKind.Relative)).GetStream()
    stream.Seek(0L, SeekOrigin.Begin) |> ignore
    let xmlDoc = new XmlDocument()
    xmlDoc.Load(stream)
    xmlDoc.DocumentElement.InnerText

let getDocxContentSpire (path: string) =
      let document = new Document()
      document.LoadFromFile(path)
      let stringBuilder = new StringBuilder()
      for section in document.Sections do
          for paragraph in section.Paragraphs do
              stringBuilder.AppendLine(paragraph.Text)
      stringBuilder
 
seq{for d in Directory.EnumerateDirectories(applicationDirectory) do yield! getAllFilesNames d} 
|> Seq.iter (fun (x:string) ->
    outFile.WriteLine("----------------------------------------------------------------------")
    outFile.WriteLine(x.Substring(applicationDirectory.Length+1))
    outFile.WriteLine("----------------------------------------------------------------------")
    outFile.WriteLine(getDocxContentSpire(x))
    
    )
 
outFile.Close()
    
// printfn "%s" (getDocxContent @"F:\swissedu\swissedu_attachments\2020-09-21_Lesson 1- notes.docx")
 