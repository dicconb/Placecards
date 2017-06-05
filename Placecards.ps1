$folder = 'C:\temp\wedding'

$peopleCSV = "$folder\Wedding place names.csv"
$cardsDocTemplate = "$folder\Place Cards.docx"
$outputfile = "$folder\Generated Place Cards - $(Get-Date -UFormat "%Y-%m-%d_%H%M%S").docx"

$cardspersheet = 6
$people = import-csv $peopleCSV | ?{!($_."not attending")}
$requiredpages = [math]::Ceiling($people.Count / $cardspersheet)

$word = New-Object -ComObject Word.Application
$word.Visible = $true

$doc = $word.Documents.Open($cardsDocTemplate)
$Templaterange = $doc.range()
$Templaterange.Copy()
$doc.close()

$cardsdoc = $word.Documents.Add()

$i=0
while ($i -lt $requiredpages) {
    $targetrange = $cardsdoc.Content
    $targetrange.Collapse([microsoft.office.interop.word.wdcollapsedirection]::wdCollapseEnd)
    $targetrange.paste()
    $i++
}

$wdReplaceOne = [microsoft.office.interop.word.wdreplace]::wdReplaceOne
$wdReplaceAll = [microsoft.office.interop.word.wdreplace]::wdReplaceAll 
$wdFindContinue = 1 
 
$FindText = "Text" 
$MatchCase = $False 
$MatchWholeWord = $True 
$MatchWildcards = $False 
$MatchSoundsLike = $False 
$MatchAllWordForms = $False 
$Forward = $True 
$Wrap = $wdFindContinue 
$Format = $False 

foreach ($person in $people) {
$a = $cardsdoc.content.Find.Execute($FindText,$MatchCase,$MatchWholeWord, ` 
    $MatchWildcards,$MatchSoundsLike,$MatchAllWordForms,$Forward,` 
    $Wrap,$Format,$person.name,$wdReplaceOne) 
}

$a = $cardsdoc.content.Find.Execute($FindText,$MatchCase,$MatchWholeWord, ` 
    $MatchWildcards,$MatchSoundsLike,$MatchAllWordForms,$Forward,` 
    $Wrap,$Format,"",$wdReplaceAll)

$cardsdoc.SaveAs($outputfile)