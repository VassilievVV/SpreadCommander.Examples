$schost = Get-SCHost;                                
$schost.Silent = $true;

Clear-Book;
Clear-Spreadsheet;
Clear-Data;

Invoke-SCScript '~\..\Common\InitBookStyles.ps1';

Set-BookSectionHeader '<b>Spread Commander</b> - <i>Examples: Getting Started</i>' -Html;
Set-BookSectionFooter 'Page {PAGE} of {NUMPAGES}' -ExpandFields;

Write-Text -ParagraphStyle:'Header1' 'GETTING STARTED';

Write-Html -ParagraphStyle:'Description' @'
<p align=justify>This example provides general introduction
into <b>SpreadCommander</b>. For more detailed samples on using
specific features please check other examples.</p>
'@;


Write-Html -ParagraphStyle:'Text' @'
<p align=justify>
    <b>Spread Commander</b> is an application for office-style
    data analysis that uses <i>PowerShell</i> script for data manipulation.
</p>
<p align=justify>
    <b>SpreadCommander</b> contains advanced <i>PowerShell console</i>
    with output into <i>Book</i> (rich text editor), <i>Spreadsheet</i>,
    powerfull <i>DataGrid</i> and ability to generate <i>Charts</i> and
    <i>Pivot tables</i>. Combination with standalone <i>Book</i>, 
    <i>Spreadsheet</i>, <i>SQL</i>, <i>Chart</i>, <i>Pivot</i>,
    <i>Dashboard</i> controls makes <b>SpreadCommander</b>
    a powerfull <i>data analysis</i> application for work with
    various kind of data.
</p>
<p align=justify>
    Combination of <i>PowerShell</i> script and advanced 
    <i>data visualization</i> controls makes work with data
    very comfortable.
</p>
'@;

Write-Html -ParagraphStyle:'Text' @'
<p align=justify>One of main features of <b>SpreadCommander</b> is
executing <i>PowerShell</i> scripts. Output of a script can be 
sent into <b>Book</b>, <b>Spreadsheet</b> and <b>Data</b> controls.
This allows to generate formatted output and send tabular data
into <b>Spreadsheet</b> or <b>Data</b> grids, with conditional
formatting, for further processing.</p>
'@;

Write-Html -ParagraphStyle:'Text' @'
<p align=justify><b>Console</b> is set of controls where <i>Script</i>
can send output. <b>Console</b> contains controls:</p>
<ul>
	<li><p align=justify><b>Book</b> gets output of the <i>Script</i>,
		including <i>text</i>, <i>images</i>, <i>HTML</i>, <i>Markdown</i>, 
		<i>LaTeX</i> (for formulas). <i>Images</i> can be loaded
		from disk or generated in <i>Script</i>, for example
		<b>SpreadCommander</b> contains cmdlets to load <i>tables</i>,
		<i>charts</i>, <i>pivot tables</i> from <i>spreadsheets</i> and to
		generate <i>charts</i> and <i>maps</i>.</p></li>
	<li><p align=justify><b>Spreadsheet</b> allows to show
		tabular data sent from the <i>Script</i>. <i>Data Source</i>
		can be <i>ADO.Net DataTable</i> or list of objects. <b>Spreadsheet</b>
		allows to add <i>conditional formatting</i>, <i>charts</i> and 
		<i>pivot tables</i>.</p></li>
	<li><p align=justify><b>Data</b> tab contains multiple <b>Grid</b>
		controls. It accepts same data as <b>Spreadsheet</b> - 
		<i>ADO.Net DataTable or DataSet</i> and list of objects.</p></li>
	<li><p align=justify><b>Heap</b> is a file browser and viewer for
		common files - <i>images</i>, <i>text files (including Word files)</i>,
		<i>spreadsheets</i> etc.</p></li>
</ul>
<p align=justify></p>
'@;

Write-Html -ParagraphStyle:'Text' @'
<p align=justify>Main script language in <b>SpreadCommander</b> is
<i>PowerShell</i>. <b>SpreadCommander</b> includes multiple
<i>Cmdlets</i> for rich text output, generating <b>Charts</b> and 
<b>Maps</b>, output formatted table data etc.</p>
<p align=justify><b>SpreadCommander</b> also allows to execute
<i>Python</i> and <i>R</i> script, if corresponding script engine 
is installed. At this moment these script engine support only
basic functionality; their output is redirected into <b>Book</b> console.</p>
'@;

Add-BookPageBreak;
Write-Html -ParagraphStyle:'Header4' 'Sample use of <b>Book</b>';

Write-Text -ParagraphStyle:'Text' '';
Write-Html -ParagraphStyle:'Text' '<b>Image</b>';
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
Write-Html -ParagraphStyle:'Text' '<b>DataTable</b>';
$dsLandUse.Tables['LandUse'] | 
	Write-DataTable -TableStyle:Medium25 `
		-Formatting:"format column 'LandUse' with ColorScale='Red,Blue', ForeColor='White'";
		
Write-Text -ParagraphStyle:'Text' '';
Write-Html -ParagraphStyle:'Text' '<b>Chart</b>';
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


Write-Html -ParagraphStyle:'Header4' 'Sample use of <b>Spreadsheet</b>';

$sqlData = @'
--#table Regions
select ID, [Country Code], [Short Name], [Table Name], [Long Name],
	[2-alpha code], [Currency Unit], [Special Notes], [WB-2 code], 
	[National accounts base year],
	[Balance of Payments Manual in use], [External debt Reporting status],
	[Latest agricultural census], [Latest industrial data], [Latest trade data]
from Countries
where ifnull(Region, '') = '';

--#table Countries
select ID, [Country Code], [Short Name], [Table Name], [Long Name],
	[2-alpha code], [Currency Unit], [Special Notes], [Region], 
	[Income group], [WB-2 code], [National accounts base year],
	[National accounts reference year], [SNA price valuation],
	[Lending category], [Other groups], [System of National Accounts],
	[Alternative conversion factor], [PPP survey year],
	[Balance of Payments Manual in use], [External debt Reporting status],
	[System of trade], [Government Accounting concept],
	[IMF data dissemination standard], [Latest population census],
	[Latest household survey], [Source of most recent Income and expenditure data],
	[Vital registration complete], [Latest agricultural census],
	[Latest industrial data], [Latest trade data] 
from Countries
where Region > '';

--#table Series
select ID, [Series Code], [Topic], [Indicator Name], [Short definition],
	[Long definition], [Unit of measure], [Periodicity], [Base period],
	[Other notes], [Aggregation method], [Limitations and exceptions], 
	[Notes from original source], [General comments], [Source],
	[Statistical concept and methodology], [Development relevance],
	[Related source links], [Other web links], [Related indicators],
	[License Type]
from Series;
go

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

--#table "Energy use - by Income"
select c.[Income Group], 
	min(pp.Value) as MinValue,
	max(pp.Value) as MaxValue,
	avg(pp.Value) as AverageValue,
	median(pp.Value) as MedianValue,
	stdev(pp.Value) as StdDev
from [EG.GDP.PUSE.KO.PP] pp
join Countries c on c.[Country Code] = pp.[Country Code]
where Year = 2014 and Value is not null and c.Region > ''
group by c.[Income Group]
order by case [Income Group] 
		when 'Low Income' then 1
		when 'Lower middle income' then 2
		when 'Upper middle income' then 3
		when 'High income' then 4 end;

--#table "Energy use"
select pp.ID, pp.[Country Code], c.[Table Name] as Country, c.Region, 
	c.[Income Group], pp.Year, pp.Value
from [EG.GDP.PUSE.KO.PP] pp
join Countries c on c.[Country Code] = pp.[Country Code]
where Year = 2014 and Value is not null and c.Region > ''
order by Country;
'@;

$dataSet = Invoke-SqlScript 'sqlite:~\..\Data\WorldData.db' -Query:$sqlData;

$dataSet.Tables['Energy use - by Regions'] |
	Out-SpreadTable -SheetName:'Energy use - by Regions' `
		-TableName:'Energy use - by Regions' `
		-TableStyle:Medium25 -Replace;
		
Write-Html -ParagraphStyle:'Text' @'
Cmdlet <i>Out-SpreadTable</i> can also copy output into <b>Book</b>.
Do not use this for large tables.
'@;

$dataSet.Tables['Energy use - by Income'] |
	Out-SpreadTable -SheetName:'Energy use - by Income' `
		-TableName:'Energy use - by Income' `
		-TableStyle:Medium8 -CopyToBook -Replace;


$dataSet.Tables['Energy use'] | Sort-Object -Property:'Region' |
	Out-SpreadTable -SheetName:'Energy use - Subtotal Regions' -AsRange `
		-SubtotalGroupBy:'Region' -SubtotalColumns:@('Value') `
		-SubtotalFunction:Average -TableStyle:Medium25 -Replace;
		
$dataSet.Tables['Energy use'] | Sort-Object -Property:'Income Group' |
	Out-SpreadTable -SheetName:'Energy use - Subtotal Income' -AsRange `
		-SubtotalGroupBy:'Income Group' -SubtotalColumns:@('Value') `
		-SubtotalFunction:Average -TableStyle:Medium25 -Replace;
		
$dataSet.Tables['Energy use'] | 
	Out-SpreadTable -SheetName:'Energy use - Data' -TableName:'Energy use - Data' `
		-TableStyle:Medium25 `
		-Formatting:'format column [Value] with DataBar=''Red'', Gradient=true' `
		-Replace;

New-SimpleSpreadChart -DataSheetName:'Energy use - by Regions' -DataTableName:'Energy use - by Regions' `
	-ChartSheetName:'Energy use - SimpleChart' -Replace -ChartType:LineMarker `
	-Arguments:'Region' -Values:'MinValue','AverageValue','MaxValue' `
	-Style:ColorGradient -VaryColors `
	-Title:'Energy use - Chart' -TitleFont:'Tahoma,12,Bold,Italic' `
	-SeriesTypes:ColumnClustered -CopyToBook;
	
New-SpreadChart -DataSheetName:'Energy use - by Regions' -DataTableName:'Energy use - by Regions' `
	-ChartSheetName:'Energy use - Chart' -Replace `
	-ChartType:LineMarker -Style:ColorGradient -VaryColors -BackColor:Yellow `
	-Title:'Energy use - by Regions' -TitleFont:'Segoe,12,Bold,Italic,Blue' `
	-Series:@(@{Arguments='Region'; Values='AverageValue'; Type='ColumnStacked'}, `
		@{Name='Min Values'; Arguments='Region'; Values='MinValue'; Type='LineMarker'}, `
		@{Name='Max Values'; Arguments='Region'; Values='MaxValue'; Type='LineMarker'}) `
	-PrimaryAxes:@(@{Title='Regions'}, @{Title='Value'});

New-SpreadPivot -DataSheetName:'Energy use - Data' -DataTableName:'Energy use - Data' `
	-PivotSheetName:'Energy use - Pivot' -PivotTableName:'Energy use - Pivot' `
	-RowFields:'Region' -ColumnFields:'Income Group' -DataFields:'Value' `
	-Layout:Outline -MergeTitles -SummarizeValuesBy:Average -Style:Medium25 `
	-Formatting:'format with DataBar=Red, Gradient=true' -Replace -CopyToBook;


Write-Html -ParagraphStyle:'Header4' 'Sample use of <b>Chart</b>';

$sqlChart = @'
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

$dataChart = Invoke-SqlScript 'sqlite:~\..\Data\WorldData.db' -Query:$sqlChart;

$dataChart.Tables['Data'] | 
	New-Chart StackedArea -SeriesField:'Country' 'Year' 'GDP' -TextPattern:'{V:F0}' `
		-ResolveOverlappingMode:JustifyAllAroundPoint -NoLabels |
	Add-ChartTitle 'Regional GDP' -Font:'Tahoma,18,Italic' |
	Set-ChartLegend -ShadowColor:Gray -AlignmentHorizontal:Center `
		-AlignmentVertical:BottomOutside -Direction:LeftToRight `
		-TitleText:'Regional GDP, billions USD' |
	Write-Chart -Width:2000 -Height:1600;
	
$dataChart.Tables['Data'] | ?{ $_.Year -eq 2014 } |
	New-Chart Line -NoLegend -GridLayout -RowDefinitions:'2,1' -ColumnDefinitions:'1,1' |
	Add-ChartTitle 'Regional GPD and Electric Power Consumption, 2014' -Font:'Tahoma,18,Italic' |
	#Add first series
	Add-ChartSeries Line 'Country' 'GDP' |
		Set-ChartSeriesLabel -TextPattern:'{V:F0}' -ResolveOverlappingMode:JustifyAllAroundPoint |
	Add-ChartAxis Y 'Y_EPC' |
	Add-ChartSeries Area 'Country' 'EPC' -AxisY:'Y_EPC' |
		Set-ChartSeriesLabel -TextPattern:'{V:F0}' -ResolveOverlappingMode:JustifyAllAroundPoint |
	#Have to have a series before setting default pane
	Set-ChartDefaultPane -Row:0 -Column:0 -ColumnSpan:2 |
	#Add second series
	Add-ChartPane 'Pane_GDP' -Row:1 -Column:0 |
	Add-ChartSeries Line 'Country' 'GDP' -Pane:'Pane_GDP' |
		Set-ChartSeriesLabel -TextPattern:'{V:F0}' -ResolveOverlappingMode:JustifyAllAroundPoint |
	#Add third series
	Add-ChartPane 'Pane_EPC' -Row:1 -Column:1 |
	Add-ChartSeries Area 'Country' 'EPC' -Pane:'Pane_EPC' |
		Set-ChartSeriesLabel -TextPattern:'{V:F0}' -ResolveOverlappingMode:JustifyAllAroundPoint |	
	#Output chart into Book
	Write-Chart -Width:2000 -Height:2400;


Write-Html -ParagraphStyle:'Header4' 'Sample use of <b>Map</b>';

$mapItems = @(
	[PSCustomObject] @{ Latitude = 55.7496; Longitude = 37.6237;   Name = 'Moscow' },
	[PSCustomObject] @{ Latitude = 39.904;  Longitude = 116.4075;  Name = 'Beijing' },
	[PSCustomObject] @{ Latitude = 35.6895; Longitude = 139.6917;  Name = 'Tokyo' },
	[PSCustomObject] @{ Latitude = 37.7749; Longitude = -122.4194; Name = 'San-Francisco' },
	[PSCustomObject] @{ Latitude = 40.7144; Longitude = -74.006;   Name = 'New-York' },
	[PSCustomObject] @{ Latitude = 48.8566; Longitude = 2.3522;    Name = 'Paris' }
);

$map = New-Map -BackColor:White | 
	Add-MapLayerImage Bing Road | 
	Add-MapLayerVectorItems;

$mapItems |
	%{ $map | Add-MapItem Pushpin @($_.Latitude, $_.Longitude) $_.Name | Out-Null };

$map | 
	Add-MapItem Line @(55.7496, 37.6237) @(39.904, 116.4075) -StrokeColor:Gold -StrokeWidth:10 -Geodesic | 
	Add-MapItem Line @(39.904, 116.4075) @(35.6895, 139.6917) -StrokeColor:Gold -StrokeWidth:10 -Geodesic | 
	Add-MapItem Line @(35.6895, 139.6917) @(37.7749, -122.4194) -StrokeColor:Gold -StrokeWidth:10 -Geodesic | 
	Add-MapItem Line @(37.7749, -122.4194) @(40.7144, -74.006) -StrokeColor:Gold -StrokeWidth:10 -Geodesic | 
	Add-MapItem Line @(40.7144, -74.006) @(48.8566, 2.3522) -StrokeColor:Gold -StrokeWidth:10 -Geodesic | 
	Add-MapItem Line @(48.8566, 2.3522) @(55.7496, 37.6237) -StrokeColor:Gold -StrokeWidth:10 -Geodesic | 
	Write-Map -CenterPoint:@(0.0, 0.0) -Width:2000 -Height:1600 -ZoomLevel:1;


Write-Html -ParagraphStyle:'Header4' 'Sample use of <b>SQL script</b> and <b>Data</b>';

Write-Html -ParagraphStyle:'Text' @'
<p align=justify><b>SpreadCommander</b> provides advanced features to work
with <i>SQL script</i> and display results of executing of <i>SQL queries</i>.
Check output on console tab <b>Data</b>.</p>
'@;

$dataSql = Invoke-SqlScript 'sqlite:~\..\Data\WorldData.db' -ScriptFile:'~\SampleData.sql';
$dataSql | Out-DataSet;


Write-Html -ParagraphStyle:'Header4' 'Sample use of <b>Math symbolic calculations</b>';

$x = [MathNet.Symbolics.SymbolicExpression]::Variable('x');
$x.Cos().Pow(4).TrigonometricContract().ToLaTeX() | Write-LaTeX -FontSize:24;

Write-Html -ParagraphStyle:'Text' @'
<p align=justify><b>SpreadCommander</b> allows advanced data visualization
controls - <i>Chart</i>, <i>Pivot table</i> and <i>Dashboard</i>. Check 
example <b>Controls</b> for more details.</p>
'@;


Add-BookSection -ContinuePageNumbering -LinkHeaderToPrevious -LinkFooterToPrevious;
Write-Text -ParagraphStyle:'Header2' 'Table of Contents';
Add-BookTOC;

Save-Book '~\ReadMe.docx' -Replace;
Save-Spreadsheet '~\ReadMe.xlsx' -Replace;