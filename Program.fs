open System.IO
open Spire.Doc
open System.Text
let applicationDirectory = "F:\swissedu"

let filtered = fun (s: string) -> s.EndsWith("docx")

let rec getAllFilesNames directory =
    seq {
        yield!
            (set (Directory.EnumerateFiles(directory))
             |> Set.filter filtered)

        for d in Directory.EnumerateDirectories(directory) do
            yield! getAllFilesNames d
    }

let outFile =
    new StreamWriter(applicationDirectory + @"\swissedu_documents.txt")

let outWordsFile =
    new StreamWriter(
        applicationDirectory
        + @"\swissedu_documents_words.txt"
    )

let outSummaryFile =
    new StreamWriter(
        applicationDirectory
        + @"\swissedu_documents_summary.txt"
    )

let getDocxContentSpire (path: string) =
    let document = new Document()
    document.LoadFromFile(path)
    let stringBuilder = new StringBuilder()

    for section in document.Sections do
        for paragraph in section.Paragraphs do
            stringBuilder.AppendLine(paragraph.Text)

    stringBuilder

let getDocxContentWordsSpire (path: string) =
    let document = new Document()
    document.LoadFromFile(path)
    let stringBuilder = new StringBuilder()

    for section in document.Sections do
        for paragraph in section.Paragraphs do
            let words = paragraph.Text.Split(" ")

            for word in words do
                stringBuilder.AppendLine(word)

    stringBuilder

// read all documents and write them in 2 files (outWordsFile is for the summary)
seq {
    for d in Directory.EnumerateDirectories(applicationDirectory) do
        yield! getAllFilesNames d
}
|> Seq.iter
    (fun (x: string) ->
        outFile.WriteLine("----------------------------------------------------------------------")
        outFile.WriteLine(x.Substring(applicationDirectory.Length + 1))
        outFile.WriteLine("----------------------------------------------------------------------")
        outFile.WriteLine(getDocxContentSpire (x))
        outWordsFile.WriteLine(getDocxContentWordsSpire (x))

        )

outFile.Close()
outWordsFile.Close()

let lines =
    File.ReadLines(applicationDirectory + @"\swissedu_documents.txt")
// To check
lines |> Seq.iter (fun x -> printfn "%s" x)

// prepare summary Word: count
let lineWords =
    File.ReadLines(
        applicationDirectory
        + @"\swissedu_documents_words.txt"
    )
// https://stackoverflow.com/questions/19069835/how-do-i-deal-with-ienumerable-in-f

// Equivalent for Python Counter
// https://stackoverflow.com/questions/52166170/is-there-an-f-equivalent-to-pythons-counter-collection
let counts =
    lineWords
    |> Seq.cast<string>
    |> List.ofSeq
    |> List.countBy id
    |> Map.ofList
// write summary
for count in counts do
    printfn $"{count.Key}: {count.Value}"
    outSummaryFile.WriteLine($"{count.Key}: {count.Value}")

outSummaryFile.Close()
