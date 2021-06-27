$schost = Get-SCHost;                                
$schost.Silent = $true;

Clear-Book;
Clear-Spreadsheet;

Invoke-SCScript '~\..\Common\InitBookStyles.ps1';

Set-BookSectionHeader '<b>Spread Commander</b> - <i>Examples: Spreadsheet</i>' -Html;
Set-BookSectionFooter 'Page {PAGE} of {NUMPAGES}' -ExpandFields;

Write-Text -ParagraphStyle:'Header1' 'SPREADSHEET';

Write-Html -ParagraphStyle:'Description' @'
<p align=justify><b>Spreadsheet</b> is one of main document types in Spread Commander.
It exists as standalone document type, and also is used as
console for scripts (PowerShell).</p>
'@;

#Retrieve sample data
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

Write-Text -ParagraphStyle:'Header2' 'Work with data';

Write-Html -ParagraphStyle:'Text' @'
<b>SpreadCommander</b> allows to output not only into rich-text console
but also into <b>Spreadsheet</b> and <b>Data Grid</b>. This sample
shows use of <b>Spreadsheet</b>. Mose common cmdlet to output
into spreadsheet is <i>Out-SpreadTable</i>.
'@;

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
		
Write-Html -ParagraphStyle:'Text' @'
Review output on tab <b>Spreadsheet</b> in console.
<br><br>
Cmdlet <i>Out-SpreadTable</i> supports multiple options.
Particulary it is possible to calculate subtotals. This
requires -AsRange flag to not create a table and rows have to 
be correctly sorted. Sample sheet with subtotals is generated in sheets named 
<i>Energy use - Subtotal Regions</i> and <i>Energy use - Subtotal Income</i>.
See it in tab <b>Spreadsheet</b> in console.
'@;

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


Add-BookPageBreak;
Write-Text -ParagraphStyle:'Header2' 'Spreadsheet functions';

Write-Html -ParagraphStyle:'Text' @'
<b>SpreadCommander</b> supports most of functions available
in <i>Microsoft Excel</i>. In addition it supports following 
functions:
'@;

Write-Html -ParagraphStyle:'Text' @'
<br>
<i>Regular expressions</i>:

<ul>
	<li><p><b>REGEX.ISMATCH</b></p>
		<p align=justify>Indicates whether the regular expression finds a match 
			in a specified input string.</p>
	<li><p><b>REGEX.MATCH</b></p>
		<p align=justify>Searches the specified input string for the first occurrence 
			of the regular expression.</p>
	<li><p><b>REGEX.MATCHES</b></p>
		<p align=justify>Array function. Searches the specified input string 
			for all (up to size of the array) occurrences of a regular expression.</p>
	<li><p><b>REGEX.NAMEDMATCH</b></p>
		<p align=justify>Searches the specified input string for the first occurrence 
			of the regular expression and returns value of the named group.</p>
	<li><p><b>REGEX.NAMEDMATCHES</b></p>
		<p align=justify>Array function. Searches the specified input string 
			for all (up to size of the array) occurrences of a regular expression and
			returns value of the named group.</p>
	<li><p><b>REGEX.REPLACE</b></p>
		<p align=justify>Within a specified input string, replaces all strings 
			that match a regular expression pattern with a specified 
			replacement string.</p>
	<li><p><b>REGEX.SPLIT</b></p>
		<p align=justify>Splits the specified input string at the positions 
			defined by a regular expression pattern.</p>
</ul>
'@;

Write-Html -ParagraphStyle:'Text' @'
<br>
<i>String</i>:

<ul>
	<li><p><b>STRING.FORMAT</b></p>
		<p align=justify>Replaces one or more format items in a specified string 
			with the string representation of a specified object.</p>
</ul>
'@;

Write-Html -ParagraphStyle:'Text' @'
<br>
<i>Hash</i>:

<ul>
	<li><p><b>HASH.MD5</b></p>
		<p align=justify>Calculates MD5 hash of the string.</p>
	<li><p><b>HASH.SHA1</b></p>
		<p align=justify>Calculates SHA-1 hash of the string.</p>
	<li><p><b>HASH.SHA256</b></p>
		<p align=justify>Calculates SHA-256 hash of the string.</p>
	<li><p><b>HASH.SHA384</b></p>
		<p align=justify>Calculates SHA-384 hash of the string.</p>
	<li><p><b>HASH.SHA512</b></p>
		<p align=justify>Calculates SHA-512 hash of the string.</p>
</ul>
'@;

Write-Html -ParagraphStyle:'Text' @'
<br>
<i>Path</i>:

<ul>
	<li><p><b>PATH.CHANGEEXTENSION</b></p>
		<p align=justify>Changes the extension of a path string.</p>
	<li><p><b>PATH.COMBINE</b></p>
		<p align=justify>Combines multiple strings into a path.</p>
	<li><p><b>PATH.GETDIRECTORYNAME</b></p>
		<p align=justify>Returns the directory information for the specified path string.</p>
	<li><p><b>PATH.GETEXTENSION</b></p>
		<p align=justify>Returns the extension of the specified path string.</p>
	<li><p><b>PATH.GETFILENAME</b></p>
		<p align=justify>Returns the file name and extension of the specified path string.</p>
	<li><p><b>PATH.GETFILENAMEWITHOUTEXTENSION</b></p>
		<p align=justify>Returns the file name of the specified path string without the extension.</p>
	<li>
</ul>
'@;

Write-Html -ParagraphStyle:'Text' @'
<br>
<i>GUID</i>:

<ul>
	<li><p><b>NEWID</b></p>
		<p align=justify>Returns as new instance of GUID formatted using 
		    optional format specifier (N, D, B, P, X).</p>
</ul>
'@;


Add-BookPageBreak;
Write-Text -ParagraphStyle:'Header2' 'Spreadsheet charts';

Write-Html -ParagraphStyle:'Text' @'
<b>SpreadCommander</b> allows to insert chart sheets into workbooks.
This can be done using two cmdlets: <i>New-SimpleSpreadChart</i> and <i>New-SpreadChart</i>.
<i>New-SimpleSpreadChart</i> allows to insert multiple series with same type.
<i>New-SpreadChart</i> allows chart series with different types and provides more control
over other elements such as axes, title, legeng etc. Data for charts shall already
present in workbook. Charts can be copied into <b>Book</b> console using switch parameter
<i>-CopyToBook</i>.
<br>
In cmdlet <i>New-SpreadChart</i> series and axes are defined using objects with 
following properties:

<h3>Series</h3>
<ul>
	<li><p><b>Name</b></p>
		<p align=justify>Name of chart series. If skipped - column names will be using as series names.</p>
	<li><p><b>Arguments</b></p>
		<p align=justify>Data (column name) to plot as series arguments.</p>
	<li><p><b>Values</b></p>
		<p align=justify>Data (column names) to plot as series values. Multiple values will be using 
			for individual series. Bubble chart requires pairs of adjoined columns. 
			Stock chart requires set of adjoined columns.</p>
	<li><p><b>Type</b></p>
		<p align=justify>Chart types for series, when chart contain series with different types. 
			ChartType has to be set to the type compatible with types of all series.
			Possible values are: ColumnClustered, ColumnStacked, ColumnFullStacked, Column3DClustered, 
			Column3DStacked, Column3DFullStacked, Column3DStandard, Column3DClusteredCylinder, 
			Column3DStackedCylinder, Column3DFullStackedCylinder, Column3DStandardCylinder, 
			Column3DClusteredCone, Column3DStackedCone, Column3DFullStackedCone, Column3DStandardCone, 
			Column3DClusteredPyramid, Column3DStackedPyramid, Column3DFullStackedPyramid, 
			Column3DStandardPyramid, Line, LineStacked, LineFullStacked, LineMarker, 
			LineStackedMarker, LineFullStackedMarker, Line3D, Pie, Pie3D, PieExploded, Pie3DExploded, 
			PieOfPie, BarOfPie, BarClustered, BarStacked, BarFullStacked, Bar3DClustered, 
			Bar3DStacked, Bar3DFullStacked, Bar3DClusteredCylinder, Bar3DStackedCylinder, 
			Bar3DFullStackedCylinder, Bar3DClusteredCone, Bar3DStackedCone, Bar3DFullStackedCone, 
			Bar3DClusteredPyramid, Bar3DStackedPyramid, Bar3DFullStackedPyramid, Area, AreaStacked, 
			AreaFullStacked, Area3D, Area3DStacked, Area3DFullStacked, ScatterMarkers, 
			ScatterSmoothMarkers, ScatterSmooth, ScatterLine, ScatterLineMarkers, StockHighLowClose, 
			StockOpenHighLowClose, StockVolumeHighLowClose, StockVolumeOpenHighLowClose, 
			Surface, SurfaceWireframe, Surface3D, Surface3DWireframe, Doughnut, DoughnutExploded, 
			Bubble, Bubble3D, Radar, RadarMarkers, RadarFilled.</p>
	<li><p><b>Markers</b></p>
		<p align=justify>Shape of markers which can be painted at each data point in the series 
			on the line, scatter or radar chart and within the chart legend.
			Possible values are: Auto, None, Circle, Dash, Diamond, Dot, Picture, Plus, Square, 
			Star, Triangle, X.</p>
	<li><p><b>MarkerSize</b></p>
		<p align=justify>Size of the marker in points. Values from 2 to 72 are allowed, default is 7.</p>
	<li><p><b>AxisGroup</b></p>
		<p align=justify>Axis types for series. Default is primary axis. First axis group must be Primary.
			Possible values: Primary, Secondary.</p>
	<li><p><b>Explosion</b></p>
		<p align=justify>Explosion value for all slices in a pie or doughnut chart series.</p>
	<li><p><b>Shape</b></p>
		<p align=justify>Shape used to display data points in the 3D bar or column chart.
			Possible values are: Auto, Box, Cone, ConeToMax, Cylinder, Pyramid, PyramidToMax.</p>
	<li><p><b>Smooth</b></p>
		<p align=justify>Whether the curve smoothing is turned on for the line or scatter chart.</p>
	<li><p><b>FromIndex</b></p>
		<p align=justify>Start index for series. Index is 0-base. Negative values are allowed: 
			-1 is last value, -2 is value before last etc.</p>
	<li><p><b>ToIndex</b></p>
		<p align=justify>End index for series. Index is 0-base. Negative values are allowed: 
			-1 is last value, -2 is value before last etc.</p>
</ul>

<h3>Axes</h3>
<ul>
	<li><p><b>Title</b></p>
		<p align=justify>Title of axis.</p>
	<li><p><b>NumberFormat</b></p>
		<p align=justify>Format of axis labels.</p>
	<li><p><b>Font</b></p>
		<p align=justify>Axis font in form 'Tahoma, 8, Bold, Italic, Green'.</p>
	<li><p><b>MajorTickMarks</b></p>
		<p align=justify>Position of major tick marks on the axis.
			Possible values: Cross, Inside, None, Outside.</p>
	<li><p><b>MinorTickMarks</b></p>
		<p align=justify>Position of major tick marks on the axis.
			Possible values: Cross, Inside, None, Outside.</p>
	<li><p><b>ShowMajorGridLines</b></p>
		<p align=justify>Whether to draw major gridlines on the chart or no.</p>
	<li><p><b>ShowMinorGridLines</b></p>
		<p align=justify>Whether to draw minor gridlines on the chart or no.</p>
	<li><p><b>Position</b></p>
		<p align=justify>Position of axis on the chart.
			Possible values: Left, Top, Right, Bottom.</p>
	<li><p><b>BaseTimeUnit</b></p>
		<p align=justify>Base unit for the date axis.
			Possible values: Days, Months, Years, Auto.</p>
	<li><p><b>LabelAlignment</b></p>
		<p align=justify>Text alignment for the tick-mark labels on the category axis.
			Possible values: Center, Left, Rights.</p>
	<li><p><b>LogScaleBase</b></p>
		<p align=justify>Logarithmic base for the logarithmic axis.</p>
	<li><p><b>LogScale</b></p>
		<p align=justify>Whether the value axis should display its numerical values using 
			a logarithmic scale.</p>
	<li><p><b>Minimum</b></p>
		<p align=justify>Minimum value of the numerical or date axis.</p>
	<li><p><b>Maximum</b></p>
		<p align=justify>Maximum value of the numerical or date axis.</p>
	<li><p><b>Reversed</b></p>
		<p align=justify>Specifies that the axis must be reversed, so the axis starts 
			at the maximum value and ends at the minimum value.</p>
	<li><p><b>HideTickLabels</b></p>
		<p align=justify>Whether tick labels should be hidden.</p>
	<li><p><b>Visible</b></p>
		<p align=justify>Whether the axis should be displayed.</p>
	<li>
</ul>
'@;

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


Add-BookPageBreak;
Write-Text -ParagraphStyle:'Header2' 'Spreadsheet pivot tables';

Write-Html -ParagraphStyle:'Text' @'
<b>SpreadCommander</b> allows to pivot tables into workbooks.
This can be done using cmdlet <i>New-SpreadPivot</i>.
Pivot can be copied into <b>Book</b> console using switch parameter
<i>-CopyToBook</i>.
'@;

New-SpreadPivot -DataSheetName:'Energy use - Data' -DataTableName:'Energy use - Data' `
	-PivotSheetName:'Energy use - Pivot' -PivotTableName:'Energy use - Pivot' `
	-RowFields:'Region' -ColumnFields:'Income Group' -DataFields:'Value' `
	-Layout:Outline -MergeTitles -SummarizeValuesBy:Average -Style:Medium25 `
	-Formatting:'format with DataBar=Red, Gradient=true' -Replace -CopyToBook;
	

Add-BookPageBreak;
Write-Text -ParagraphStyle:'Header2' 'Spreadsheet templates';

Write-Html -ParagraphStyle:'Text' @'
<b>SpreadCommander</b> allows to use spreadsheet templates. Cmdlet <i>Out-SpreadTemplate</i>
accepts data from pipe and populate spreadsheet using pre-created template.
<br>
Spread template is a workbook with a single worksheet. It contains layout for the data fields.
<br>
Template also holds special defined names indicating whether data records should be merged 
into a worksheet.
<br>
To create template create a table in worksheet with same column names that will be in data source,
and populate this worksheet with sample data. Select cell(s) within this table and
in tab Operations on ribbon bar (group Table Tools) press button <i>Spreadsheet Template</i>.
<br>
<i>Spreadsheet Template Editor</i> contains ribbon tab <i>Mail Merge</i>. Select area that will
be used for data source row and press button <i>Detail</i>. It will create named range <i>DETAILRANGE</i>.
Drag fields from field list at right into this area and create wanted layout. If needed - create 
<i>Header</i> and <i>Footer</i> areas. To group data source rows by some field - press button 
<i>Source Fields</i> and add group fields. Then select area inside range <i>DETAILRANGE</i> and
create group's header and footer areas. Save template into a file.
<br>
Samples of use of cmdlet <i>Out-SpreadTemplate</i> are provided below. Produced sheet 
can be copied into <b>Book</b> using switch <i>-CopyToBook</i>.
'@;

$dataSet.Tables['Regions'] |
	Out-SpreadTable -SheetName:'Regions' `
		-TableName:'Regions' `
		-TableStyle:Medium25 -Replace;
		
$dataSet.Tables['Countries'] |
	Out-SpreadTable -SheetName:'Counrties' `
		-TableName:'Countries' `
		-TableStyle:Medium25 -Replace;
		
$dataSet.Tables['Series'] |
	Out-SpreadTable -SheetName:'Series' `
		-TableName:'Series' `
		-TableStyle:Medium25 -Replace;
		
$dataSet.Tables['Regions'] |
	Out-SpreadTemplate -SheetName:'Regions (template)' `
		-TemplateFileName:'~\Templates\Regions.xlsx' `
		-Replace -CopyToBook;
		
$dataSet.Tables['Countries'] |
	Out-SpreadTemplate -SheetName:'Countries (template)' `
		-TemplateFileName:'~\Templates\Countries.xlsx' `
		-Replace;
		
$dataSet.Tables['Series'] |
	Out-SpreadTemplate -SheetName:'Series (template)' `
		-TemplateFileName:'~\Templates\Series.xlsx' `
		-Replace;


$dataSet.Dispose();


Add-BookSection -ContinuePageNumbering -LinkHeaderToPrevious -LinkFooterToPrevious;

Write-Text -ParagraphStyle:'Header2' 'Spreadsheet cmdlets';

. $schost.MapPath('~\..\Common\CmdletHelp.ps1');

$cmdlets = [string[]]@(
	'Clear-Spreadsheet',
	'ConvertTo-Pivot',
	'ConvertTo-Unpivot',
	'Get-SpreadTable',
	'Merge-Spreadsheet',
	'New-SimpleSpreadChart',
	'New-SpreadChart',
	'New-SpreadPivot',
	'New-Spreadsheet',
	'Out-SpreadTable',
	'Out-SpreadTemplate',
	'Save-Spreadsheet',
	'Write-SpreadTable'
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
Save-Spreadsheet '~\ReadMe.xlsx' -Replace;