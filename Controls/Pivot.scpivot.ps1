$schost = Get-SCHost;
$schost.Silent = $true;

Clear-Book;
Clear-Data;

Invoke-SCScript '~#\..\Common\InitBookStyles.ps1';

Set-BookSectionHeader '<b>Spread Commander</b> - <i>Examples: Controls - Pivot</i>' -Html;
Set-BookSectionFooter 'Page {PAGE} of {NUMPAGES}' -ExpandFields;

Write-Text -ParagraphStyle:'Header1' 'CONTROLS - PIVOT';

Write-Html -ParagraphStyle:'Description' @'
<p align=justify><b>Spreadsheet</b> provides some controls for data visualization.
These controls include <b>Chart</b>, <b>Pivot</b> and <b>Dashboard</b>.
Current sample shows how to use <b>Pivot</b> control.</p>
'@;

Write-Html -ParagraphStyle:'Text' @'
<p align=justify>For <i>data visualization controls</i> <b>SpreadCommander</b>
adds new tab in <b>Console</b>; this tab contains <i>data visualization control</i>.
Data shall be added with cmdlets <i>Out-Data</i> or <i>Out-DataSet</i>.
Data will become a <i>data source</i> for <i>data visualization control</i> and
will be displayed on console tab <b>Data</b>; this allows to review raw data.
Script output will be sent into console tab <b>Book</b>. Console tab 
<b>Spreadsheet</b> can be used the same way as in <i>PowerShell scripts</i>
for more thorough data analysis. <i>Data visualization control</i> 
can be designed in User Interface. For <b>Pivot</b> control use 
customization control at left side to put columns into <i>Row Area</i>, 
<i>Column Area</i> and <i>Data Area</i>.</p>
'@;

#Retrieve sample data
$sqlData = @'
--#table Countries
select c.[Country Code], c.[Short Name] as Country, c.[Special Notes]
from Countries c 
where c.[Country Code] in ('ARB', 'CEB', 'EAS', 'ECS', 'LCN', 'MEA', 'NAC', 'SAS', 'SSF') 
order by c.[Country Code];

--#table Data
select c.[Country Code], c.[Short Name] as Country,
	gdp.Year, (gdp.Value / 1E9) as GDP, epc.Value as EPC
from Countries c 
join [NY.GDP.MKTP.CD] gdp on gdp.[Country Code] = c.[Country Code]
join [EG.USE.ELEC.KH.PC] epc on epc.[Country Code] = c.[Country Code] and epc.Year = gdp.Year
where c.[Country Code] in ('ARB', 'CEB', 'EAS', 'ECS', 'LCN', 'MEA', 'NAC', 'SAS', 'SSF') 
order by c.[Country Code], gdp.Year;
'@;

$dataSet = Invoke-SqlScript 'sqlite:~#\..\Data\WorldData.db' -Query:$sqlData;

$dataSet.Tables['Data'] |
	Out-Data -TableName:'Data' -Replace;

Write-Html -ParagraphStyle:'Header2' 'Creating a <b>Pivot</b>';

Write-Html -ParagraphStyle:'Text' @'
<ul>
	<li><p align=justify>Write a script that generates data and send data into console
		tab <b>Data</b> using cmdlet <b>Out-Data</b> or <b>Out-DataSet</b>.</p></li>
	<li><p align=justify>Select console tab <b>Pivot</b> and expand control <b>Customization</b>
		at left side.</p></li>
	<li><p align=justify>Put columns into corresponding areas - in this case - <b>County</b> int
		<i>Column Area</i>, <b>Year</b> into <i>Row Area</i>, <b>GDP</b> and <b>EPC</b> into
		<i>Data Area</i>.</p></li>
	<li><p align=justify>Make other changes as needed. For example - add <i>Format Conditions</i>
		using button <i>Format Conditions</i> on ribbon.</p></li>
</ul>

<p align=justify>When <b>Pivot</b> is opened next time - it will be empty.
This happens because data are not stored with <b>Pivot</b>. Re-execute
script that provides data, and new data will be bound to the <b>Pivot</b>.</p>

<p align=justify><b>Pivot</b> also includes tab <b>Chart</b>. Data source for
the <b>Chart</b> is always selection in <b>Pivot</b> control. Please see
sample for <b>Chart</b> in this project for details on using <i>Chart Designer</i>.</p>
'@;


Add-BookSection -ContinuePageNumbering -LinkHeaderToPrevious -LinkFooterToPrevious;

Write-Text -ParagraphStyle:'Header2' 'Controls cmdlets';

. $schost.MapPath('~#\..\Common\CmdletHelp.ps1');

$cmdlets = [string[]]@(
	'Out-Data',
	'Out-DataSet'
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

Save-Book '~#\Pivot.docx' -Replace;
