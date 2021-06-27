$schost = Get-SCHost;                              
$schost.Silent = $true;

Clear-Book;

Invoke-SCScript '~\..\Common\InitBookStyles.ps1';

Set-BookSectionHeader '<b>Spread Commander</b> - <i>Examples: Chart</i>' -Html;
Set-BookSectionFooter 'Page {PAGE} of {NUMPAGES}' -ExpandFields;

Write-Text -ParagraphStyle:'Header1' 'CHART';

Write-Html -ParagraphStyle:'Description' @'
<p><b>Chart</b> is a powerful tool that allows to visualize data. <b>Chart</b>
allows to output into a <b>Book</b> document or can be be saved in file for use
in other applications.</p>
'@;

Write-Html -ParagraphStyle:'Text' @'
<p align=justify>Creating new <b>Chart</b> starts with cmdlet </i>New-Chart</i>,
ends with cmdlet <i>Write-Chart</i> or <i>Save-Chart</i>. Multiple other cmdlets
(help on <b>Chart</b> cmdlets is provided at end of this <b>Book</b>) allow
to customize charts, add multiple series, title, labels, customize axis etc.
Examples below demonstrate basic operations with <b>Chart</b>.</p>
'@;

Write-Html -ParagraphStyle:'Description' @'
<p>Current script is accompanied with script file ReadMe.ps1 with 
source code.</p>
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

$dataSet = Invoke-SqlScript 'sqlite:~\..\Data\WorldData.db' -Query:$sqlData;

$dtSankey = [Data.DataTable]::new('Sankey');
[void]$dtSankey.Columns.Add('Source', [string]);
[void]$dtSankey.Columns.Add('Target', [string]);
[void]$dtSankey.Columns.Add('Weight', [double]);
[void]$dtSankey.Columns.Add('Selected', [boolean]);

[void]$dtSankey.Rows.Add('France', 'UK', 53, $true);
[void]$dtSankey.Rows.Add('Australia', 'UK', 72, $false);
[void]$dtSankey.Rows.Add('France', 'Canada', 81, $true);
[void]$dtSankey.Rows.Add('China', 'Canada', 96, $false);
[void]$dtSankey.Rows.Add('UK', 'France', 61, $false);
[void]$dtSankey.Rows.Add('Canada', 'France', 89, $false);

$dataSet.Tables.Add($dtSankey);

$dataSet.Tables['Countries'] |
	Select-Object -Property:@{Name='Code'; Expression='Country Code'}, 'Country', 'Special Notes' |
	Out-SpreadTable -TableName:'Countries' -SheetName:'Countries' `
		-TableStyle:Medium1 -WrapText `
		-ColumnWidths: @{ 'Code'=15; 'Country'=25; 'Special Notes'=40 } `
		-VerticalAlignment:Top -Replace;
		
$dataSet.Tables['Data'] | ?{ $_.Year -eq 2014 } |
	Select-Object -Property:@{Name='Code'; Expression='Country Code'}, 'GDP', 'EPC' |
	Sort-Object -Property: 'Code' |
	Out-SpreadTable -TableName:'Data' -SheetName:'Data' `
		-ColumnNumberFormats:@{ 'GDP'= '#,##0'; 'EPC'='#,##0' } `
		-ColumnWidths: @{ 'Code'=5; 'GDP'=10; 'EPC'=10 } `
		-TableStyle:Medium1 -Replace;
		
$dataSet.Tables['Sankey'] |
    Out-SpreadTable -SheetName:'Sankey' -TableName:'Sankey' `
        -TableStyle:Medium1 -Replace;

Add-BookPageBreak;
Write-Text -ParagraphStyle:'Header3' 'BAR CHART';
Write-Html -ParagraphStyle:'Text' @'
<p align=justify>First chart shows basic <i>Bar</i> chart with main elements - 
<b>Chart</b> itself, <b>Title</b>, <b>Legend</b>.</p>
'@;

$dataSet.Tables['Data'] | ?{ $_.Year % 3 -eq 0 } |
	New-Chart Bar -SeriesField:'Country' 'Year' 'GDP' -NoLabels |
	Add-ChartTitle 'Regional GDP' -Font:'Tahoma,18,Italic' |
	Set-ChartLegend -ShadowColor:Gray -AlignmentHorizontal:Center `
		-AlignmentVertical:BottomOutside -Direction:LeftToRight `
		-TitleText:'Regional GDP, billions USD' |
	Write-Chart -Width:2000 -Height:1600;
	
Add-BookPageBreak;
Write-Text -ParagraphStyle:'Header3' 'LINE CHART';
Write-Html -ParagraphStyle:'Text' @'
<p align=justify>Another basic <b>Chart</b> - with <i>Line</i> series.</p>
'@;
$dataSet.Tables['Data'] |
	New-Chart Line -SeriesField:'Country' 'Year' 'GDP' -TextPattern:'{V:F0}' `
		-NoLabels |
	Add-ChartTitle 'Regional GDP' -Font:'Tahoma,18,Italic' |
	Set-ChartLegend -ShadowColor:Gray -AlignmentHorizontal:Center `
		-AlignmentVertical:BottomOutside -Direction:LeftToRight `
		-TitleText:'Regional GDP, billions USD' |
	Write-Chart -Width:2000 -Height:2000;
	
Add-BookPageBreak;
Write-Text -ParagraphStyle:'Header3' 'STACKED AREA CHART';
Write-Html -ParagraphStyle:'Text' @'
<p align=justify>Another basic <b>Chart</b> - with <i>StackedArea</i> series.</p>
'@;
$dataSet.Tables['Data'] | 
	New-Chart StackedArea -SeriesField:'Country' 'Year' 'GDP' -TextPattern:'{V:F0}' `
		-ResolveOverlappingMode:JustifyAllAroundPoint -NoLabels |
	Add-ChartTitle 'Regional GDP' -Font:'Tahoma,18,Italic' |
	Set-ChartLegend -ShadowColor:Gray -AlignmentHorizontal:Center `
		-AlignmentVertical:BottomOutside -Direction:LeftToRight `
		-TitleText:'Regional GDP, billions USD' |
	Write-Chart -Width:2000 -Height:1600;
	
Add-BookPageBreak;
Write-Text -ParagraphStyle:'Header3' 'FULL STACKED AREA CHART';
Write-Html -ParagraphStyle:'Text' @'
<p align=justify>Another basic <b>Chart</b> - with <i>FullStackedArea</i> series.
This <b>Chart</b> shows data for every 5-years period (previous chart - yearly data),
and shows <i>labels</i>.</p>
'@;
$dataSet.Tables['Data'] | ?{ $_.Year % 5 -eq 0 } |
	New-Chart FullStackedArea -SeriesField:'Country' 'Year' 'GDP' -TextPattern:'{V:F0}' |
	Add-ChartTitle 'Regional GDP' -Font:'Tahoma,18,Italic' |
	Set-ChartLegend -ShadowColor:Gray -AlignmentHorizontal:Center `
		-AlignmentVertical:BottomOutside -Direction:LeftToRight `
		-TitleText:'Regional GDP, billions USD' |
	Write-Chart -Width:2000 -Height:1600;
	
Add-BookPageBreak;
Write-Text -ParagraphStyle:'Header3' 'AREA 3D CHART';
Write-Html -ParagraphStyle:'Text' @'
<p align=justify>3D charts are supported too. To make <b>Chart</b> provide more
information it is painted to higher size and <i>Scale=0.5</i> is applied.</p>
'@;
$dataSet.Tables['Data'] | 
	New-Chart Area3D -SeriesField:'Country' 'Year' 'GDP' -NoLabels `
		-DiagramFillMode:Gradient -DiagramBackColor:DarkGray -DiagramBackColor2:White `
		-RotationAngleX:10 -RotationAngleY:20 -RotationAngleZ:5 |
	Add-ChartTitle 'Regional GDP' -Font:'Tahoma,18,Italic' |
	Set-ChartLegend -ShadowColor:Gray -AlignmentHorizontal:RightOutside `
		-AlignmentVertical:Center -Direction:TopToBottom `
		-TitleText:'Regional GDP, billions USD' |
	Write-Chart -Width:4000 -Height:3200 -Scale:0.5;
	
Add-BookPageBreak;
Write-Text -ParagraphStyle:'Header3' 'PIE CHART';
Write-Html -ParagraphStyle:'Text' @'
<p align=justify>Another basic <b>Chart</b> - with <i>Pie</i> series.</p>
'@;
$dataSet.Tables['Data'] | ?{ $_.Year -eq 2018 } | Sort-Object -Property: 'GDP' |
	New-Chart Pie 'Country' 'GDP' -TextPattern:'{A}, {V:F0}' |
	Add-ChartTitle 'Regional GDP (2018)' -Font:'Tahoma,30,Italic' |
	Set-ChartLegend -ShadowColor:Gray -AlignmentHorizontal:Center `
		-AlignmentVertical:BottomOutside -Direction:LeftToRight |
	Set-ChartSeriesLabel -Font:'Tahoma,7' |
	Write-Chart -Width:2000 -Height:1600;

Add-BookPageBreak;
Write-Text -ParagraphStyle:'Header3' 'DOUGHNUT CHART';
Write-Html -ParagraphStyle:'Text' @'
<p align=justify>Another basic <b>Chart</b> - with <i>Doughnut</i> series.</p>
'@;
$dataSet.Tables['Data'] | ?{ $_.Year -eq 2018 } | Sort-Object -Property: 'GDP' |
	New-Chart Doughnut 'Country' 'GDP' -TextPattern:'{A}, {V:F0}' |
	Add-ChartTitle 'Regional GDP (2018)' -Font:'Tahoma,18,Italic' |
	Set-ChartLegend -ShadowColor:Gray -AlignmentHorizontal:Center `
		-AlignmentVertical:BottomOutside -Direction:LeftToRight |
	Set-ChartSeriesLabel -Font:'Tahoma,7' |
	Write-Chart -Width:2000 -Height:1600;
	
Add-BookPageBreak;
Write-Text -ParagraphStyle:'Header3' 'PIE 3D CHART';
Write-Html -ParagraphStyle:'Text' @'
<p align=justify>Another basic <b>Chart</b> - with <i>Pie3D</i> series.</p>
'@;
$dataSet.Tables['Data'] | ?{ $_.Year -eq 2018 } | Sort-Object -Property: 'GDP' |
	New-Chart Pie3D 'Country' 'GDP' -TextPattern:'{A}, {V:F0}' |
	Add-ChartTitle 'Regional GDP (2018)' -Font:'Tahoma,18,Italic' |
	Set-ChartLegend -ShadowColor:Gray -AlignmentHorizontal:Center `
		-AlignmentVertical:BottomOutside -Direction:LeftToRight |
	Set-ChartSeriesLabel -Font:'Tahoma,7' |
	Write-Chart -Width:2000 -Height:1600;

Add-BookPageBreak;
Write-Text -ParagraphStyle:'Header3' 'FUNNEL CHART';
Write-Html -ParagraphStyle:'Text' @'
<p align=justify>Another basic <b>Chart</b> - with <i>Funnel</i> series.</p>
'@;
$dataSet.Tables['Data'] | ?{ $_.Year -eq 2018 } |
	Sort-Object -Property:'GDP' -Descending |
	New-Chart Funnel 'Country' 'GDP' -TextPattern:'{A}, {V:F0}' |
	Add-ChartTitle 'Regional GDP (2018)' -Font:'Tahoma,18,Italic' |
	Set-ChartLegend -ShadowColor:Gray -AlignmentHorizontal:Center `
		-AlignmentVertical:BottomOutside -Direction:LeftToRight |
	Set-ChartSeriesLabel -Font:'Tahoma,7' |
	Write-Chart -Width:2000 -Height:1600;

Add-BookPageBreak;
Write-Text -ParagraphStyle:'Header3' 'FUNNEL 3D CHART';
Write-Html -ParagraphStyle:'Text' @'
<p align=justify>Another basic <b>Chart</b> - with <i>Funnel3D</i> series.</p>
'@;
$dataSet.Tables['Data'] | ?{ $_.Year -eq 2018 } | 
	Sort-Object -Property:'GDP' -Descending |
	New-Chart Funnel3D 'Country' 'GDP' -TextPattern:'{A}, {V:F0}' |
	Add-ChartTitle 'Regional GDP (2018)' -Font:'Tahoma,18,Italic' |
	Set-ChartLegend -ShadowColor:Gray -AlignmentHorizontal:Center `
		-AlignmentVertical:BottomOutside -Direction:LeftToRight |
	Set-ChartSeriesLabel -Font:'Tahoma,7' |
	Write-Chart -Width:2000 -Height:1600;
	
Add-BookPageBreak;	
Write-Text -ParagraphStyle:'Header3' 'SANKEY DIAGRAM';
Write-Html -ParagraphStyle:'Text' @'
<p align=justify>Special <b>Chart</b> created with own cmdlets.</p>
'@;
$dataSet.Tables['Sankey'] |
    Write-SankeyDiagram -Source:'Source' -Target:'Target' `
        -Weight:'Weight' -Selected:'Selected' `
        -SelectedNodes:@('France', 'China') -Palette:NorthernLights `
        -BackColor:Beige -TitleText:'Sankey Diagram' `
        -Width:500 -Height:300 -Scale:4;
	
Add-BookPageBreak;
Write-Text -ParagraphStyle:'Header3' 'MULTIPLE SERIES';
Write-Html -ParagraphStyle:'Text' @'
<p align=justify>Next <b>Chart</b> shows how to place multiple series on the same
<i>Diagram</i> (<i>Diagram</i> is <b>Chart</b>'s surface). <i>Series</i> are
adding individually. In <i>New-Chart</i> cmdlet series type is required and 
must be compatible with every <i>Series</i> that will be adding later.
For example, <i>Bar</i>, <i>Line</i>, <i>Area</i> series are compatible;
<i>Bar</i> and <i>Pie</i> are not compatible, as well as 2D and 3D series.</p>
'@;
$dataSet.Tables['Data'] | ?{ $_.Year -eq 2014 } |
	New-Chart Line -NoLegend |
	Add-ChartTitle 'Regional GPD and Electric Power Consumption, 2014' -Font:'Tahoma,18,Italic' |
	Add-ChartSeries Line 'Country' 'GDP' |
		Set-ChartSeriesLabel -TextPattern:'{V:F0}' -ResolveOverlappingMode:JustifyAllAroundPoint |
	Add-ChartAxis Y 'Y_EPC' |
	Add-ChartSeries Area 'Country' 'EPC' -AxisY:'Y_EPC' |
		Set-ChartSeriesLabel -TextPattern:'{V:F0}' -ResolveOverlappingMode:JustifyAllAroundPoint |
	Write-Chart -Width:2000 -Height:1600;
	
Add-BookPageBreak;
Write-Text -ParagraphStyle:'Header3' 'MULTIPLE PANES';
Write-Html -ParagraphStyle:'Text' @'
<p align=justify>Next <b>Chart</b> shows how to place multiple series on different
<i>Panes</i> (<i>Pane></i> is a part of <i>Diagram</i> that is <b>Chart</b>'s surface).</p>
'@;
$dataSet.Tables['Data'] | ?{ $_.Year -eq 2014 } |
	New-Chart Line -NoLegend -GridLayout -RowDefinitions:'2,1' -ColumnDefinitions:'1' |
	Add-ChartTitle 'Regional GPD and Electric Power Consumption, 2014' -Font:'Tahoma,18,Italic' |
	Add-ChartSeries Line 'Country' 'GDP' |
		Set-ChartSeriesLabel -TextPattern:'{V:F0}' -ResolveOverlappingMode:JustifyAllAroundPoint |
	Add-ChartPane 'Pane2' -Row:1 -Column:0 |
	Add-ChartSeries Area 'Country' 'EPC' -Pane:'Pane2' |
		Set-ChartSeriesLabel -TextPattern:'{V:F0}' -ResolveOverlappingMode:JustifyAllAroundPoint |
	Write-Chart -Width:2000 -Height:2000;
	
Add-BookPageBreak;
Write-Text -ParagraphStyle:'Header3' 'MULTIPLE SERIES AND PANES';
Write-Html -ParagraphStyle:'Text' @'
<p align=justify>Next <b>Chart</b> shows combination of multiple <i>Series</i> on 
multiple <i>Panes</i>.</p>
'@;
$dataSet.Tables['Data'] | ?{ $_.Year -eq 2014 } |
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
	
Add-BookPageBreak;
Write-Text -ParagraphStyle:'Header3' 'CHART CONTEXT';
Write-Html -ParagraphStyle:'Text' @'
<p align=justify>Same <b>Chart</b> as previous one, but built using <i>ChartContext</i>
variable. This technique is useful when it is not known how much <i>Series</i> will be
adding.</p>
'@;
$chart = $dataSet.Tables['Data'] | ?{ $_.Year -eq 2014 } |
	New-Chart Line -NoLegend -GridLayout -RowDefinitions:'2,1' -ColumnDefinitions:'1,1' |
	Add-ChartTitle 'Regional GPD and Electric Power Consumption, 2014' -Font:'Tahoma,18,Italic';
	
$chart = $chart |
	Add-ChartSeries Line 'Country' 'GDP' |
		Set-ChartSeriesLabel -TextPattern:'{V:F0}' -ResolveOverlappingMode:JustifyAllAroundPoint |
	Add-ChartAxis Y 'Y_EPC' |
	Add-ChartSeries Area 'Country' 'EPC' -AxisY:'Y_EPC' |
		Set-ChartSeriesLabel -TextPattern:'{V:F0}' -ResolveOverlappingMode:JustifyAllAroundPoint |
	Set-ChartDefaultPane -Row:0 -Column:0 -ColumnSpan:2;

$chart = $chart |
	#Add second series
	Add-ChartPane 'Pane_GDP' -Row:1 -Column:0 |
	Add-ChartSeries Line 'Country' 'GDP' -Pane:'Pane_GDP' |
		Set-ChartSeriesLabel -TextPattern:'{V:F0}' -ResolveOverlappingMode:JustifyAllAroundPoint;

$chart = $chart |
	Add-ChartPane 'Pane_EPC' -Row:1 -Column:1 |
	Add-ChartSeries Area 'Country' 'EPC' -Pane:'Pane_EPC' |
		Set-ChartSeriesLabel -TextPattern:'{V:F0}' -ResolveOverlappingMode:JustifyAllAroundPoint;

$chart |
	Write-Chart -Width:2000 -Height:2400;
	
Add-BookPageBreak;
Write-Text -ParagraphStyle:'Header3' 'CHART BITMAPS AND HTML TEMPLATE';
Write-Html -ParagraphStyle:'Text' @'
<p align=justify><b>Charts</b> can be saved into variables (as <i>System.Drawing.Bitmap</i>
objects) and written into a <b>Book</b> using cmdlet <i>Write-Html</i> with embedded <i>Fields</i>.</p>
'@;
$chart1 = $dataSet.Tables['Data'] | ?{ $_.Year -eq 2014 } |
	New-Chart Line -NoLegend |
	Add-ChartTitle 'Regional GPD and Electric Power Consumption, 2014' -Font:'Tahoma,1,Italic' |
	Add-ChartSeries Line 'Country' 'GDP' |
		Set-ChartSeriesLabel -HideLabels |
	Add-ChartAxis Y 'Y_EPC' |
	Add-ChartSeries Area 'Country' 'EPC' -AxisY:'Y_EPC' |
		Set-ChartSeriesLabel -HideLabels |
	Save-Chart -InMemory -Width:1000 -Height:1000;
	
$chart2 = $dataSet.Tables['Data'] | ?{ $_.Year -eq 2014 } |
	New-Chart Line -NoLegend |
	Add-ChartTitle 'Regional GDP, 2014' -Font:'Tahoma,12,Italic' |
	Add-ChartSeries Line 'Country' 'GDP' |
		Set-ChartSeriesLabel -TextPattern:'{V:F0}' -ResolveOverlappingMode:JustifyAllAroundPoint |
	Save-Chart -InMemory -Width:1000 -Height:1000;

$chart3 = $dataSet.Tables['Data'] | ?{ $_.Year -eq 2014 } |
	New-Chart Area -NoLegend |
	Add-ChartTitle 'Electric Power Consumption, 2014' -Font:'Tahoma,12,Italic' |
	Add-ChartSeries Area 'Country' 'EPC' |
		Set-ChartSeriesLabel -TextPattern:'{V:F0}' -ResolveOverlappingMode:JustifyAllAroundPoint |
	Save-Chart -InMemory -Width:1000 -Height:1000;
	
Write-Html @'
<table>
	<tr>
		<td>{#IMAGE chart1}</td>
		<td>{#SPREADTABLE $SPREAD Data}</td>
	</tr>
	<tr>
		<td>{#IMAGE chart2}</td>
		<td>{#IMAGE chart3}</td>
	</tr>
	<tr>
		<td colspan=2>{#SPREADTABLE $SPREAD Countries}</td>
	</tr>
</table>
'@ -ExpandFields `
	-Snippets:@{ 'chart1'=$chart1; 'chart2'=$chart2; 'chart3'=$chart3 };

Add-BookPageBreak;
Write-Text -ParagraphStyle:'Header3' 'CHART INDICATORS';
Write-Html -ParagraphStyle:'Text' @'
<p align=justify>Different <i>Indicators</i> such as <b>Regression</b>, <b>Moving Average</b>, 
<b>Error Bars</b> can be added to the <b>Chart</b>.</p>
'@;
	
$dataSet.Tables['Data'] | ?{ $_.Country -eq 'Europe & Central Asia' } |
	New-Chart Line |
	Set-ChartLegend -ShadowColor:Gray -AlignmentHorizontal:Center `
		-AlignmentVertical:BottomOutside -Direction:LeftToRight `
		-TitleText:'Europe & Central Asia GDP, billions USD' |
	Add-ChartSeries -Name:'GDP' Line 'Year' 'GDP' -NoLabels |
	Add-ChartIndicator RegressionLine -SeriesName:'GDP' -Name:'Regression' `
		-Color:'Red' -ShowInLegend |
	Add-ChartIndicator SMA -SeriesName:'GDP' -Name:'Simple Moving Average' `
		-Color:'Blue' -ShowInLegend |
	Add-ChartIndicator TMA -SeriesName:'GDP' -Name:'Triangular Moving Average' `
		-Color:'SteelBlue' -ShowInLegend |
	Add-ChartIndicator EMA -SeriesName:'GDP' -Name:'Exponential Moving Average' `
		-Color:'SteelBlue' -ShowInLegend |
	Add-ChartIndicator StandardErrorBars -SeriesName:'GDP' -Name:'Standard Error Bars' `
		-Color:'Brown' -ShowInLegend |
	Write-Chart -Width:2000 -Height:2000;
	
$dataSet.Dispose();


Add-BookSection -ContinuePageNumbering -LinkHeaderToPrevious -LinkFooterToPrevious;

Write-Text -ParagraphStyle:'Header2' 'Chart cmdlets';

. $schost.MapPath('~\..\Common\CmdletHelp.ps1');

$cmdlets = [string[]]@(
	'Add-ChartAnnotation',
	'Add-chartAxis',
	'Add-ChartAxisCustomLabel',
	'Add-ChartConstantLine',
	'Add-ChartIndicator|Help for cmdlet <i>Add-ChartIndicator</i> with parameter sets for different indicators (regression, trend, median etc) is provided in later section of this <b>Book</b>',
	'Add-ChartLegend',
	'Add-ChartPane',
	'Add-ChartScaleBreak',
	'Add-ChartSegmentColorizer|Help for cmdlet <i>Add-ChartSegmentColorizer</i> with parameter sets for different colorizers (point, range, trend) is provided in later section of this <b>Book</b>',
	'Add-ChartSeries|Help for cmdlet <i>Add-ChartSeries</i> with parameter sets for different series types (line, bar, area, 3D series etc) is provided in later section of this <b>Book</b>',
	'Add-ChartSeriesColorizer|Help for cmdlet <i>Add-ChartSeriesColorizer</i> with parameter sets for different colorizers (key, object, range) is provided in later section of this <b>Book</b>',
	'Add-ChartSeriesTitle',
	'New-Chart|Help for cmdlet <i>New-Chart</i> with parameter sets for different diagrams (XY, XY3D, Pie etc) is provided in later section of this <b>Book</b>',
	'Save-Chart',
	'Save-ChartTemplate',
	'Save-SankeyDiagram',
	'Set-ChartAxis',
	'Set-ChartAxisLabel',
	'Set-ChartAxisTitle',
	'Set-ChartDefaultPane',
	'Set-ChartLegend',
	'Set-ChartSeriesLabel',
	'Set-ChartTotalLabel',
	'Write-Chart',
	'Write-SankeyDiagram'
);

$firstCmdlet = $true;
$cmdlets |
%{
	if (-not $firstCmdlet) { Add-BookPageBreak; }
	
	[string]$cmdlet = $_;
	[string]$description = $null;
	
	if ($_.Contains('|') -eq $true)
	{
		$pos         = $_.IndexOf('|');
		$cmdlet      = $_.Substring(0, $pos);
		$description = $_.Substring($pos + 1);
	}

	$help = GenerateCmdletHelp($cmdlet);
	
	Write-Html -ParagraphStyle:'Header3' "Cmdlet <i>$cmdlet</i>";
	
	if ($description -ne $null)
	{
		Write-Html $description -ParagraphStyle:'Description';
	}
	
	Write-Html $help -ParagraphStyle:'CmdletHelp';
	
	$firstCmdlet = $false;
};

Add-BookPageBreak;

Write-Html -ParagraphStyle:'Header3' 'Cmdlet <i>New-Chart</i>: dynamic parameter sets';
Write-Html -ParagraphStyle:'Description' 'Table of Contents is at end of the <b>Book</b>.';

$firstParameterSet = $true;
@(
	'Pie|Works for series <i>Pie</i>, <i>Doughnut</i>, <i>NestedDoughnut</i>, <i>Funnel</i>.',
	'Pie3D|Works for series <i>Pie3D</i>, <i>Doughnut3D</i>.',
	'Funnel3D|Works for series <i>Funnel3D</i>',
	'Bar|Works for series <i>Bar</i>, <i>StackedBar</i>, <i>FullStackedBar</i>, 
		<i>SideBySideStackedBar</i>, <i>SideBySideFullStackedBar</i>, <i>Point</i>, 
		<i>Bubble</i>, <i>Line</i>, <i>StackedLine</i>, <i>FullStackedLine</i>, 
		<i>StepLine</i>, <i>Spline</i>, <i>ScatterLine</i>, <i>Area</i>, <i>StepArea</i>, 
		<i>SplineArea</i>, <i>StackedArea</i>, <i>StackedStepArea</i>, <i>StackedSplineArea</i>, 
		<i>FullStackedArea</i>, <i>FullStackedSplineArea</i>, <i>FullStackedStepArea</i>, 
		<i>RangeArea</i>, <i>Stock</i>, <i>CandleStick</i>, <i>SideBySideRangeBar</i>, 
		<i>RangeBar</i>, <i>BoxPlot</i>, <i>Waterfall</i>.',
	'Bar3D|Works for series <i>Bar3D</i>, <i>StackedBar3D</i>, <i>FullStackedBar3D</i>, 
		<i>ManhattanBar</i>, <i>SideBySideStackedBar3D</i>, <i>SideBySideFullStackedBar3D</i>, 
		<i>Line3D</i>, <i>StackedLine3D</i>, <i>FullStackedLine3D</i>, <i>StepLine3D</i>, 
		<i>Area3D</i>, <i>StackedArea3D</i>, <i>FullStackedArea3D</i>, <i>StepArea3D</i>, 
		<i>Spline3D</i>, <i>SplineArea3D</i>, <i>StackedSplineArea3D</i>, 
		<i>FullStackedSplineArea3D</i>, <i>RangeArea3D</i>',
	'Gantt|Works for series <i>Gantt</i>, <i>SideBySideGantt</i>',
	'PolarPoint|Works for series <i>PolarPoint</i>, <i>PolarLine</i>, <i>ScatterPolarLine</i>, 
		<i>PolarArea</i>, <i>PolarRangeArea</i>, <i>RadarPoint</i>, <i>RadarLine</i>, 
		<i>ScatterRadarLine</i>, <i>RadarArea</i>, <i>RadarRangeArea</i>',
	'Swift|Works for series <i>SwiftPlot</i>'
) |
%{
	if (-not $firstParameterSet) { Add-BookPageBreak; }
	
	[string]$parameterSet = $_;
	[string]$description = $null;
	
	if ($_.Contains('|') -eq $true)
	{
		$pos          = $_.IndexOf('|');
		$parameterSet = $_.Substring(0, $pos);
		$description  = $_.Substring($pos + 1);
	}
	
	$help = GenerateCmdletParametersHelp 'New-Chart' $parameterSet;
	
	Write-Html -ParagraphStyle:'Header4' "Parameters <i>$parameterSet</i>";
	
	if ($description -ne $null)
	{
		Write-Html $description -ParagraphStyle:'Description';
	}
	
	Write-Html $help -ParagraphStyle:'CmdletHelp';
	
	$firstParameterSet = $false;
};

Add-BookPageBreak;

Write-Html -ParagraphStyle:'Header3' 'Cmdlet <i>Add-ChartSeries</i>: dynamic parameter sets';
Write-Html -ParagraphStyle:'Description' 'Table of Contents is at end of the <b>Book</b>.';

$firstParameterSet = $true;
@(
	'Bar',
	'StackedBar',
	'FullStackedBar',
	'SideBySideStackedBar',
	'SideBySideFullStackedBar',
	'Pie',
	'Doughnut',
	'NestedDoughnut',
	'Funnel',
	'Point',
	'Bubble',
	'Line',
	'StackedLine',
	'FullStackedLine',
	'StepLine',
	'Spline',
	'ScatterLine',
	'SwiftPlot',
	'Area',
	'StepArea',
	'SplineArea',
	'StackedArea',
	'StackedStepArea',
	'StackedSplineArea',
	'FullStackedArea',
	'FullStackedSplineArea',
	'FullStackedStepArea',
	'RangeArea',
	'Stock',
	'CandleStick',
	'SideBySideRangeBar',
	'RangeBar',
	'SideBySideGantt',
	'Gantt',
	'PolarPoint',
	'PolarLine',
	'ScatterPolarLine',
	'PolarArea',
	'PolarRangeArea',
	'RadarPoint',
	'RadarLine',
	'ScatterRadarLine',
	'RadarArea',
	'RadarRangeArea',
	'Bar3D',
	'StackedBar3D',
	'FullStackedBar3D',
	'ManhattanBar',
	'SideBySideStackedBar3D',
	'SideBySideFullStackedBar3D',
	'Pie3D',
	'Doughnut3D',
	'Funnel3D',
	'Line3D',
	'StackedLine3D',
	'FullStackedLine3D',
	'StepLine3D',
	'Area3D',
	'StackedArea3D',
	'FullStackedArea3D',
	'StepArea3D',
	'Spline3D',
	'SplineArea3D',
	'StackedSplineArea3D',
	'FullStackedSplineArea3D',
	'RangeArea3D',
	'BoxPlot',
	'Waterfall'
) |
%{
	if (-not $firstParameterSet) { Add-BookPageBreak; }
	
	$help = GenerateCmdletParametersHelp 'Add-ChartSeries' $_;
	
	Write-Html -ParagraphStyle:'Header4' "Parameters <i>$_</i>";
	Write-Html $help -ParagraphStyle:'CmdletHelp';
	
	$firstParameterSet = $false;
};

Add-BookPageBreak;

Write-Html -ParagraphStyle:'Header3' 'Cmdlet <i>Add-ChartIndicator</i>: dynamic parameter sets';
Write-Html -ParagraphStyle:'Description' 'Table of Contents is at end of the <b>Book</b>.';

$firstParameterSet = $true;
@(
	'RegressionLine', 
	'TrendLine', 
	'MedianPrice', 
	'TypicalPrice', 
	'WeightedClose', 
	'Fibonacci', 
	'EMA', 
	'ExponentialMovingAverage', 
	'SMA', 
	'SimpleMovingAverage', 
	'TMA', 
	'TriangularMovingAverage', 
	'TEMA', 
	'TripleExponentialMovingAverageTema', 
	'WMA', 
	'WeightedMovingAverage', 
	'BollingerBands', 
	'MassIndex', 
	'StandardDeviation', 
	'ATR', 
	'AverageTrueRange', 
	'CHV', 
	'ChaikinsVolatility', 
	'CCI', 
	'CommodityChannelIndex', 
	'DPO', 
	'DetrendedPriceOscillator', 
	'MACD', 
	'MovingAverageConvergenceDivergence', 
	'ROC', 
	'RateOfChange', 
	'RSI', 
	'RelativeStrengthIndex', 
	'TRIX', 
	'TripleExponentialMovingAverageTrix', 
	'WilliamsR', 
	'DataSourceBasedErrorBars', 
	'FixedValueErrorBars', 
	'PercentageErrorBars', 
	'StandardDeviationErrorBars', 
	'StandardErrorBars' 
) |
%{
	if (-not $firstParameterSet) { Add-BookPageBreak; }
	
	$help = GenerateCmdletParametersHelp 'Add-ChartIndicator' $_;
	
	Write-Html -ParagraphStyle:'Header4' "Parameters <i>$_</i>";
	Write-Html $help -ParagraphStyle:'CmdletHelp';
	
	$firstParameterSet = $false;
};

Add-BookPageBreak;

Write-Html -ParagraphStyle:'Header3' 'Cmdlet <i>Add-ChartSeriesColorizer</i>: dynamic parameter sets';
Write-Html -ParagraphStyle:'Description' 'Table of Contents is at end of the <b>Book</b>.';

$firstParameterSet = $true;
@(
	'Object',
	'Key',
	'Range'
) |
%{
	if (-not $firstParameterSet) { Add-BookPageBreak; }
	
	$help = GenerateCmdletParametersHelp 'Add-ChartSeriesColorizer' $_;
	
	Write-Html -ParagraphStyle:'Header4' "Parameters <i>$_</i>";
	Write-Html $help -ParagraphStyle:'CmdletHelp';
	
	$firstParameterSet = $false;
};

Add-BookPageBreak;

Write-Html -ParagraphStyle:'Header3' 'Cmdlet <i>Add-ChartSegmentColorizer</i>: dynamic parameter sets';
Write-Html -ParagraphStyle:'Description' 'Table of Contents is at end of the <b>Book</b>.';

$firstParameterSet = $true;
@(
	'Range',
	'Trend',
	'Point'
) |
%{
	if (-not $firstParameterSet) { Add-BookPageBreak; }
	
	$help = GenerateCmdletParametersHelp 'Add-ChartSegmentColorizer' $_;
	
	Write-Html -ParagraphStyle:'Header4' "Parameters <i>$_</i>";
	Write-Html $help -ParagraphStyle:'CmdletHelp';
	
	$firstParameterSet = $false;
};


Add-BookSection -ContinuePageNumbering -LinkHeaderToPrevious -LinkFooterToPrevious;
Write-Text -ParagraphStyle:'Header2' 'Table of Contents';
Add-BookTOC;

Save-Book '~\ReadMe.docx' -Replace;
