$schost = Get-SCHost;                              
$schost.Silent = $true;

Clear-Book;
Clear-Data;

Invoke-SCScript '~\..\Common\InitBookStyles.ps1';

Set-BookSectionHeader '<b>Spread Commander</b> - <i>Examples: Text Import-Export</i>' -Html;
Set-BookSectionFooter 'Page {PAGE} of {NUMPAGES}' -ExpandFields;

Write-Text -ParagraphStyle:'Header1' 'Text Import-Export';

Write-Html -ParagraphStyle:'Description' @'
<p align=justify><b>SpreadCommander</b> provides built-in way to import and export text
files, both delimited and fixed length. Data output goes into <b>Spreadsheet</b>.</p>
'@;


#Retrieve sample data
$sqlData = @'
--#table "Energy use - by Regions"
select c.Region, 
	min(pp.Value) as MinValue,
	max(pp.Value) as MaxValue,
	avg(pp.Value) as AverageValue,
	median(pp.Value) as MedianValue,
	stdev(pp.Value) as StdDev
from [EG.GDP.PUSE.KO.PP] pp
join Countries c on c.[Country Code] = pp.[Country Code]
where Year = 2014 and Value is not null and c.Region > ''
group by c.Region
order by Region;

--#table "Energy use"
select pp.ID, pp.[Country Code], c.[Table Name] as Country, c.Region, 
	c.[Income Group], pp.Year, pp.Value
from [EG.GDP.PUSE.KO.PP] pp
join Countries c on c.[Country Code] = pp.[Country Code]
where Year = 2014 and Value is not null and c.Region > ''
order by Country;
'@;

$dataSet = Invoke-SqlScript 'sqlite:~\..\Data\WorldData.db' -Query:$sqlData;


Write-Html -ParagraphStyle:'Text' @'
First sample shows import/export of delimited file using default settings.
'@;

$dataSet.Tables['Energy use - by Regions'] |
	Export-DelimitedText '~\Data\EnergyUse_ByRegions.csv' -Overwrite;

Import-DelimitedText '~\Data\EnergyUse_ByRegions.csv' |
	Out-SpreadTable -SheetName:'Delimited001' `
		-TableName:'Delimited001' -TableStyle:Medium25 -Replace;


Write-Html -ParagraphStyle:'Text' @'
Second sample shows how to customize columns. Each column has to be listed.
If column needs to be skipped - set <b>ColumnType</b> to <i>$null</i>.
'@;

Write-Html -ParagraphStyle:'Text' @'
<b>Columns</b> supports following properties:
'@;

Write-Html -ParagraphStyle:'Text' @'
<ul>
<li><h4>ColumnName (string)</h4>
<p align=justify>Column name.</p></li>

<li><h4>ColumnType (Type)</h4>
<p align=justify>Type of the values in the column.</p></li>

<li><h4>ColumnLength (integer)</h4>
<p align=justify>Length of the column in fixed-length format.</p></li>

<li><h4>Caption (string)</h4>
<p align=justify>Column name to use in output.</p></li>

<li><h4>IsNullable (boolean)</h4>
<p align=justify>Whether nulls are allowed for the column.</p></li>

<li><h4>DefaulValue (object)</h4>
<p align=justify>Default value to use when a null is encountered on a non-nullable.</p></li>

<li><h4>Culture (string)</h4>
<p align=justify>Culture used as format provider to use to parse the value.</p></li>

<li><h4>NumberStyles (NumberStyles?)</h4>
<p align=justify>Number styles to use when parsing the value..</p></li>

<li><h4>InputFormat (string)</h4>
<p align=justify>Format string to use when parsing the date and time.</p></li>

<li><h4>OutputFormat (string)</h4>
<p align=justify>Formatting to use when converting the value to a string..</p></li>

<li><h4>TrueString (string)</h4>
<p align=justify>Value representing True.</p></li>

<li><h4>FalseString (string)</h4>
<p align=justify>value representing False.</p></li>

<li><h4>AllowTrailing (boolean)</h4>
<p align=justify>.</p></li>

<li><h4>Trim (boolean, default True)</h4>
<p align=justify>Whether the value should be trimmed prior to parsing.</p></li>

<li><h4>FillCharacter (char?)</h4>
<p align=justify>.</p></li>

<li><h4>Alignment (FixedAlignment?)</h4>
<p align=justify>Alignment of a fixed width column. 
Possible values: <i>$null</i>, <i>LeftAligned</i>, <i>RightAligned</i>.</p></li>

<li><h4>TruncationPolicy (OverflowTruncationPolicy?)</h4>
<p align=justify>How to truncate columns when the data exceeds to maximum width. 
Possible values: <i>$null</i>, <i>TruncateLeading</i>, <i>TruncateTrailing</i>.</p></li>

<li></li>
</ul>
'@;

$dataSet.Tables['Energy use - by Regions'] |
	Export-DelimitedText '~\Data\EnergyUse_ByRegions_2.csv' -Overwrite `
	    -Columns: @(
	        @{ ColumnName = 'Region';      ColumnType = [string] },
	        @{ ColumnName = 'MinValue';    ColumnType = [double] },
	        @{ ColumnName = 'MaxValue';    ColumnType = [double] },
	        @{ ColumnName = 'MedianValue'; ColumnType = [double] },
	        @{ ColumnName = 'StdDev';      ColumnType = [double] }
	    );

Import-DelimitedText '~\Data\EnergyUse_ByRegions.csv' `
        -Columns: @(
	        @{ ColumnName = 'Region';      ColumnType = [string] },
	        @{ ColumnName = 'MinValue';    ColumnType = [double] },
	        @{ ColumnName = 'MaxValue';    ColumnType = [double] },
	        @{ ColumnName = 'MedianValue'; ColumnType = [double] },
	        @{ ColumnName = 'StdDev';      ColumnType = [double] }
	      ) |
	Out-SpreadTable -SheetName:'Delimited002' `
		-TableName:'Delimited002' -TableStyle:Medium25 -Replace;


Write-Html -ParagraphStyle:'Text' @'
Import and export of text file with fixed-length formatting requires to specify columns.
'@;


$dataSet.Tables['Energy use'] | 
	Export-FixedLengthText '~\Data\EnergyUse_Data.txt' -Overwrite `
		-Columns: @(
			@{ ColumnName = 'ID'; 			ColumnType = [long];   ColumnLength = 5 },
			@{ ColumnName = 'Country Code'; Caption = 'CODE';	   ColumnType = [string]; ColumnLength = 4 },
			@{ ColumnName = 'Country'; 		ColumnType = [string]; ColumnLength = 30 },
			@{ ColumnName = 'Region';		ColumnType = [string]; ColumnLength = 30 },
			@{ ColumnName = 'Income Group';	ColumnType = [string]; ColumnLength = 30 },
			@{ ColumnName = 'Year'; 		ColumnType = [long];   ColumnLength = 5 },
			@{ ColumnName = 'Value'; 		ColumnType = [double]; ColumnLength = 10 }
		);

Import-FixedLengthText '~\Data\EnergyUse_Data.txt' `
		-Columns: @(
			@{ ColumnName = 'ID'; 			ColumnType = [int];    ColumnLength = 5 },
			@{ ColumnName = 'CountryCode'; 	ColumnType = [string]; ColumnLength = 3 },
			@{ ColumnName = 'Country'; 		ColumnType = [string]; ColumnLength = 30 },
			@{ ColumnName = 'Region';		ColumnType = [string]; ColumnLength = 30 },
			@{ ColumnName = 'IncomeGroup';	ColumnType = [string]; ColumnLength = 30 },
			@{ ColumnName = 'Year'; 		ColumnType = [int];    ColumnLength = 5 },
			@{ ColumnName = 'Value'; 		ColumnType = [double]; ColumnLength = 10 }
		) |
	Out-SpreadTable -SheetName:'Fixed001' `
		-TableName:'Fixed001' -TableStyle:Medium25 -Replace;


Write-Html -ParagraphStyle:'Text' @'
Final sample shows import and export of <b>DBF</b> file. Support for <b>DBF</b> is
added mostly to read files that comes with ESRI <b>SHP</b> map files. Export to
<b>DBF</b> is very simple and exists mostly as companion for <b>DBF</b> import.
'@;

$dataSet.Tables['Energy use - by Regions'] |
    #Select-Object Region, MinValue, MaxValue |
	Export-Dbf '~\Data\EnergyUse_ByRegions.dbf' -Overwrite;

Import-Dbf '~\Data\EnergyUse_ByRegions.dbf' |
	Out-SpreadTable -SheetName:'Dbf001' `
		-TableName:'Dbf001' -TableStyle:Medium25 -Replace;


Add-BookSection -ContinuePageNumbering -LinkHeaderToPrevious -LinkFooterToPrevious;

Write-Text -ParagraphStyle:'Header2' 'Text Import-Export cmdlets';

Write-Html -ParagraphStyle:'Text' @'
This section contains help for <i>cmdlets</i> that allow to import and export text files.
'@;

. $schost.MapPath('~\..\Common\CmdletHelp.ps1');

$cmdlets = [string[]]@(
	'Import-DelimitedText',
	'Export-DelimitedText',
	'Import-FixedLengthText',
	'Export-FixedLengthText',
	'Import-Dbf',
	'Export-Dbf'
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

Save-Book '~\TextImportExport.docx' -Replace;
Save-Spreadsheet '~\TextImportExport.xlsx' -Replace;
