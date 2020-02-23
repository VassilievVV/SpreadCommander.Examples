$schost = Get-SCHost;                                
$schost.Silent = $true;

Clear-Book;

Invoke-SCScript '~\..\Common\InitBookStyles.ps1';

Set-BookSectionHeader '<b>Spread Commander</b> - <i>Examples: Book</i>' -Html;
Set-BookSectionFooter 'Page {PAGE} of {NUMPAGES}' -ExpandFields;

Write-Text -ParagraphStyle:'Header1' 'BOOK';

Write-Html -ParagraphStyle:'Description' @'
<p><b>Book</b> is one of main document types in Spread Commander.
It exists as standalone document type, and also is used as
console for scripts (PowerShell, R, Python).</p>
'@;

Write-Html -ParagraphStyle:'Text' @'
<p align=justify>For PowerShell scripts Spread Commander provides multiple 
cmdlets that allow customize output.</p>
'@;

Write-Html -ParagraphStyle:'Description' @'
<p>Current script is accompanied with script file ReadMe.ps1 with 
source code.</p>
'@;

Write-Html -ParagraphStyle:'Text' @'
<ul>
<li><h4>Write-Text</h4>
<p align=justify>Outputs text into console similar to Write-Host cmdlet.
Allows to specify additional parameters such as character properties,
paragraph styles and others.</p></li>

<li><h4>Write-Html</h4>
<p align=justify>Outputs HTML-formatted string into console.</p></li>

<li><h4>Write-Markdown</h4>
<p align=justify>Outputs markdown into console.</p></li>

<li><h4>Write-Latex</h4>
<p align=justify>Converts Latex-formatted text into image and
outputs it into console.</p></li>

<li><h4>Write-SyntaxText</h4>
<p align=justify>Outputs text and highlights syntax.</p></li>

<li><h4>Write-ErrorMessage</h4>
<p align=justify>Outputs text formatted as error. Can be used to 
output non-terminating errors. SpreadCommander sets
ErrorActionPreference to Stop, to make errors thrown
in its cmdlets to be terminating, so cmdlet Write-ErrorMessage
is recommended to output error/warning messages.</p></li>

<li><h4>Write-Image</h4>
<p align=justify>Outputs image into Book.</p></li>

<li><h4>Write-Content</h4>
<p align=justify>Outputs content of existing file into Book. MS Word, RTF,
TXT, HTML, markdown (.md, .markdown, .mdown) files.</p></li>

<li><h4>Write-DataTable</h4>
<p align=justify>Outputs content of list or DataTable into Book. Internally
list is exported into spreadsheet, formatted, and then copied
into Book.</p></li>

<li><h4>Write-SpreadTable</h4>
<p align=justify>Outputs existing spreadsheet table into Book.</p></li>

<li><h4>Write-Chart</h4>
<p align=justify>Generates and outputs chart.</p></li>

<li><h4>Write-Map</h4>
<p align=justify>Generates and outputs geo map.</p></li>

<li></li>
</ul>
'@;

Write-Html -ParagraphStyle:'Text' @'
<p align=justify>Also multiple cmdlets allow to customize output
using rich-text capabilities.</p>
'@;

Write-Html -ParagraphStyle:'Text' @'
<ul>
<li><h4>Add-BookCharacterStyle</h4>
<p align=justify>Adds character style. Same character style can be 
shared in multiple paragraph styles.</p></li>

<li><h4>Add-BookParagraphStyle</h4>
<p align=justify>Adds paragraph styles. Paragraph style can
be specified in cmdlets Write-Text, Write-Html and others.</p></li>

<li><h4>Add-BookSection</h4>
<p align=justify>Add new section into book.</p></li>

<li><h4>Add-BookShape</h4>
<p align=justify>Adds a shape, either text or image.</p></li>

<li><h4>Clear-Book</h4>
<p align=justify>Clears book content.</p></li>

<li><h4>New-Book</h4>
<p align=justify>Create new Book. This Book can be uses same way as
console's Book by setting property Book in cmdlets. This Book
is not displayed in UI but it can be saved into file.
Using Books created with cmdlet New-Book is more effective and
these Books can be used in parallel processing.</p></li>

<li><h4>Merge-Book</h4>
<p align=justify>Merges Book created using cmdlet New-Book into another Book.</p></li>

<li><h4>Save-Book</h4>
<p align=justify>Saves Book into file. MS Word (.docx, .doc, .rtf), 
Text (.txt, .text), HTML (.html, .htm), MHT (.mht), 
OpenOffice (.odt), ePub (.epub), PDF (.pdf) formats are supported.</p></li>

<li><h4>Set-BookDefaultCharacterProperties</h4>
<p align=justify>Sets parameters of default character style.</p></li>

<li><h4>Set-BookDefaultParagraphProperties</h4>
<p align=justify>Sets parameters of default paragraph style.</p></li>

<li><h4>Set-BookSection</h4>
<p align=justify>Configures Book section.</p></li>

<li><h4>Set-BookSectionHeader</h4>
Setups Book's section header.</p></li>

<li><h4>Set-BookSectionFooter</h4>
<p align=justify>Setups Book's section footer.</p></li>

<li></li>
</ul>
'@;

Write-Text -ParagraphStyle:'Header4' 'Sample use of some cmdlets';

Write-Text -ParagraphStyle:'Text' '';
Write-Html -ParagraphStyle:'Text' '<b>Write-Image</b>';
Write-Image '~\..\Common\SpreadCommander.png';

Write-Text -ParagraphStyle:'Text' '';
Write-Html -ParagraphStyle:'Text' '<b>Write-Latex</b>';
Write-Latex @(
	'B''=-\nabla \times E', 
	'E''=\nabla \times B - 4\pi j',
	'e^{ix} = \cos{x} + i \sin{x}');

$sqlLandUse = @'
--#table LandUse
select c.Region, round(sum(a.Value)/1e6, 2) as LandUse
from [AG.LND.TOTL.K2] a
join Countries c on c.[Country Code] = a.[Country Code]
where a.Year = 2018 and c.Region > ''
group by c.Region
order by c.Region
'@;

$dsLandUse = Invoke-SqlScript 'sqlite:~\..\Data\WorldData.db' -Query:$sqlLandUse;

Write-Text -ParagraphStyle:'Text' '';
Write-Html -ParagraphStyle:'Text' '<b>Write-DataTable</b>';
$dsLandUse.Tables['LandUse'] | 
	Write-DataTable -TableStyle:Medium25 `
		-Formatting:"format column 'LandUse' with ColorScale='Red,Blue', ForeColor='White'";
		
Write-Text -ParagraphStyle:'Text' '';
Write-Html -ParagraphStyle:'Text' '<b>Write-Chart</b>';
$dsLandUse.Tables['LandUse'] | 
	New-Chart pie Region LandUse -TextPattern:'{A}: {VP:P2} ({V:F1}M km²)' -NoLegend -LegendTextPattern:'{A}: {VP:P2} ({V:F1} km²)' | 
	Add-ChartTitle 'Regions Land Use' -Font:'Tahoma,18,Italic' |
	Set-ChartLegend -ShadowColor:Yellow -TitleText:'Area by regions' |
	Write-Chart;
	
Write-Html -ParagraphStyle:'Text' @'
Cmdlets that output into <b>Book</b> support comments.
'@ -Comment:'This paragraph has comment.';

Write-Text -ParagraphStyle:'Text' @'
Comments can have HTML formatting.
'@ -Comment:'This paragraph has <b>HTML</b>-<i>formatted</i> comment.' -CommentHtml;

Write-Text '';
Write-Text -ParagraphStyle:'Header5' 'Fields';

Write-Html -ParagraphStyle:'Text' @'
SpreadCommander allows to use fields to specify codes that instructs <b>Book</b> to insert 
text or graphics automatically. Press Ctrl+F9 to add new field and Alt+F9 to switch between 
field view result view. Press Ctrl+Alt+F9 to recalculate fields. 

Fields are mostly used in standalone <b>Book</b>. In console <b>Book</b> it is usually 
easier and more correct to use PowerShell (or another) script to generate needed output.

Available field codes:
<table>
<tr>
    <td>AUTHOR</td>
    <td><p align=justify>Inserts contents of the Author box from the Summary tab in the Document Properties dialog.</p></td>
</tr>
<tr>
    <td>COMMENTS</td>
    <td><p align=justify>Inserts contents of the Comments box from the Summary tab in the Document Properties dialog.</p></td>
</tr>
<tr>
    <td>CREATEDATE</td>
    <td><p align=justify>The date and time when the document was first created. After a mail merge the field is replaced with the date and time of the mail merge operation.</p></td>
</tr>
<tr>
    <td>DATE</td>
    <td><p align=justify>Inserts the current date and time.</p></td>
</tr>
<tr>
    <td>DOCPROPERTY</td>
    <td><p align=justify>Inserts the value of the document property specified by the field parameter.</p></td>
</tr>
<tr>
    <td>HYPERLINK</td>
    <td><p align=justify>Enables you to navigate to another location or to a bookmark.</p></td>
</tr>
<tr>
    <td>IF</td>
    <td><p align=justify>Compares two values and inserts the text according to the results of the comparison.</p></td>
</tr>
<tr>
    <td>INCLUDEPICTURE</td>
    <td><p align=justify>Inserts the specified image.</p></td>
</tr>
<tr>
    <td>KEYWORDS</td>
    <td><p align=justify>Inserts contents of the Keywords box from the Summary tab in the Document Properties dialog.</p></td>
</tr>
<tr>
    <td>LASTSAVEDBY</td>
    <td><p align=justify>Inserts the name of the last person who modified and saved the document.</p></td>
</tr>
<tr>
    <td>MERGEFIELD</td>
    <td><p align=justify>Retrieves a value from the bound data source.</p></td>
</tr>
<tr>
    <td>NUMPAGES</td>
    <td><p align=justify>Inserts the total number of pages.</p></td>
</tr>
<tr>
    <td>PAGE</td>
    <td><p align=justify>Inserts the number of the page containing the field.</p></td>
</tr>
<tr>
    <td>PRINTDATE</td>
    <td><p align=justify>Inserts the date and time that a document was last printed.</p></td>
</tr>
<tr>
    <td>REVNUM</td>
    <td><p align=justify>Inserts the number of document revisions.</p></td>
</tr>
<tr>
    <td>SAVEDATE</td>
    <td><p align=justify>Inserts the date and time a document was last saved.</p></td>
</tr>
<tr>
    <td>SEQ</td>
    <td><p align=justify>Provides sequential numbering in the document.</p></td>
</tr>
<tr>
    <td>SUBJECT</td>
    <td><p align=justify>Inserts contents of the Subject box from the Summary tab in the Document Properties dialog.</p></td>
</tr>
<tr>
    <td>SYMBOL</td>
    <td><p align=justify>Inserts a symbol.</p></td>
</tr>
<tr>
    <td>TC</td>
    <td><p align=justify>Defines entries for the table of contents.</p></td>
</tr>
<tr>
    <td>TIME</td>
    <td><p align=justify>Inserts the current time.</p></td>
</tr>
<tr>
    <td>TITLE</td>
    <td><p align=justify>Inserts contents of the Title box from the Summary tab in the Document Properties dialog.</p></td>
</tr>
<tr>
    <td>TOC</td>
    <td><p align=justify>Builds a table of contents.</p></td>
</tr>
</table>

Also custom fields are available using field DOCVARIABLE. Syntax is:
DOCVARIABLE VariableName Parameters. Parameters is single string formatted
as ConnectionString, i.e. in form "Parameter1=Value1;Parameter2=Value2;Parameter3=Value3".
Characters can be escaped with '\', '\' itself shall be escaped as '\\'.

VariableName can be one of following:
<br>
<table>
<tr>
	<td>Document, File</td>
	<td>Insert content of another document (Book/Word file, RTF file, HTML file, ePub file, Markdown file). 
		First argument is filename, also arugment "recalculate" (synonyms "rebuild" and "recalc") is
		supported to recalculate fields in source document if it is Book/Word file.</td>
</tr>
<tr>
	<td>Image, Picture</td>
	<td>Image file. First parameter is filename. Also parameters "dpi", "scale", "scaleX", "scaleY" are 
		supported.</td>
</tr>
<tr>
	<td>SVG</td>
	<td>SVG file. First parameter is filename. Also parameters "dpi", "scale", "scaleX", "scaleY", "size" (in form 100x100) are 
		supported.</td>
</tr>
<tr>
	<td>Latex, Formula</td>
	<td>LATEX-formatted text. Supported parameters are "dpi", "scale", "scaleX", "scaleY", "FontSize".</td>
</tr>
<tr>
	<td>SpreadTable</td>
	<td>Spreadsheet table. First parameter is filename. Second parameter is table name, defined range or range.
		Also parameter "recalculate" (with synonyms "rebuild" and "recalc") is supported, to recalculate spreadsheet.</td>
</tr>
<tr>
	<td>SpreadChart</td>
	<td>SpreadsheetChart. First parameter is filename. Second parameter is chart sheet name or chart name (in form "Worksheet!ChartName").
		Also parameters "recalculate" (with synonyms "rebuild" and "recalc"), "size" (in form 100x200), "scale", "scaleX", "scaleY"
		are supported.</td>
</tr>
<tr>
	<td>SpreadPivot</td>
	<td>Spreadsheet pivot table. First parameter is filename. Second parameter is pivot sheet name or pivot table name (in form "Worksheet!PivotTable").
		Also parameters "recalculate" (with synonyms "rebuild" and "recalc"), "dataonly" are supported.</td>
</tr>
</table>
<br>
To add fields from PowerShell script use switch <b>ExpandFields</b> that is available in cmdlets <i>Write-Text</i> 
	and <i>Write-Html</i>.
'@;

$imagePath  = $schost.MapPath('~\..\Common\SpreadCommander.png').Replace('\', '\\');

Write-Text -ParagraphStyle:'Text' -ExpandFields @"
Today is {DATE} {TIME}.
Current document contains {NUMPAGES} pages.

{INCLUDEPICTURE "$imagePath"}
"@;

Add-BookSection -ContinuePageNumbering -LinkHeaderToPrevious -LinkFooterToPrevious;

Write-Text -ParagraphStyle:'Header2' 'Book cmdlets';

. $schost.MapPath('~\..\Common\CmdletHelp.ps1');

$cmdlets = [string[]]@(
	'Add-BookCharacterStyle',
	'Add-BookPageBreak',
	'Add-BookParagraphStyle',
	'Add-BookSection',
	'Add-BookShape',
	'Add-BookTOC',
	'Clear-Book',
	'Merge-Book',
	'New-Book',
	'Save-Book',
	'Set-BookDefaultCharacterProperties',
	'Set-BookDefaultParagraphProperties',
	'Set-BookDefaultProperties',
	'Set-BookSection',
	'Set-BookSectionFooter',
	'Set-BookSectionHeader',
	'Write-Content',
	'Write-DataTable',
	'Write-ErrorMessage',
	'Write-HTML',
	'Write-Image',
	'Write-Latex',
	'Write-Markdown',
	'Write-SpreadTable',
	'Write-SyntaxText',
	'Write-Text'
);

$firstCmdlet = $true;
$cmdlets |
%{
	if (-not $firstCmdlet) { Add-BookPageBreak; }

	$help = GenerateCmdletHelp($_);
	
	Write-Html -ParagraphStyle:'Header3' "Cmdlet <i>$_</i>";
	Write-Html $help -ParagraphStyle:'CmdletHelp';
	
	$firstCmdlet = $false;
};

Add-BookSection -ContinuePageNumbering -LinkHeaderToPrevious -LinkFooterToPrevious;
Write-Text -ParagraphStyle:'Header2' 'Table of Contents';
Add-BookTOC;

Save-Book '~\ReadMe.docx' -Replace;
