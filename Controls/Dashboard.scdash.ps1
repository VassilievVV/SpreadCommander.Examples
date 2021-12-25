$schost = Get-SCHost;
$schost.Silent = $true;

Clear-Book;
Clear-Data;

Invoke-SCScript '~#\..\Common\InitBookStyles.ps1';

Set-BookSectionHeader '<b>Spread Commander</b> - <i>Examples: Controls - Dashboard</i>' -Html;
Set-BookSectionFooter 'Page {PAGE} of {NUMPAGES}' -ExpandFields;

Write-Text -ParagraphStyle:'Header1' 'CONTROLS - DASHBOARD';

Write-Html -ParagraphStyle:'Description' @'
<p align=justify><b>Spreadsheet</b> provides some controls for data visualization.
These controls include <b>Chart</b>, <b>Pivot</b> and <b>Dashboard</b>.
Current sample shows how to use <b>Dashboard</b> control.</p>
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
can be designed in User Interface. <b>Dashboard</b> control allows to
add <i>Grids</i>, <i>Charts</i>, <i>Pivots</i>, <i>Maps</i> and other controls 
and organize them for better presentation of the <i>data source</i>.</p>
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


Write-Html -ParagraphStyle:'Header2' 'Creating a <b>Dashboard</b>';

Write-Html -ParagraphStyle:'Text' @'
<p align=justify>To add new control to the <b>Dashboard</b> - select it 
on ribbon's tab <i>Home</i> and place to the <b>Dashboard</b>. At left 
side select <i>Data Source</i> and configure <b>Control</b>, usually
by dragging <i>Fields</i> onto corresponding places in area <i>Data Items</i>.</p>

<p align=justify>When <b>Dashboard</b> is opened next time - it will be empty.
This happens because data are not stored with <b>Dashboard</b>. Re-execute
script that provides data, and new data will be bound to the <b>Dashboard</b>.</p>
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

Save-Book '~#\Dashboard.docx' -Replace;
