$schost = Get-SCHost;                              
$schost.Silent = $true;

Clear-Book;

Invoke-SCScript '~#\..\Common\InitBookStyles.ps1';

Set-BookSectionHeader '<b>Spread Commander</b> - <i>Examples: Map</i>' -Html;
Set-BookSectionFooter 'Page {PAGE} of {NUMPAGES}' -ExpandFields;

Write-Text -ParagraphStyle:'Header1' 'MAP';

Write-Html -ParagraphStyle:'Description' @'
<p><b>Map</b> is a tool that allows to visualize geographic data. <b>Map</b>
allows to output into a <b>Book</b> document or can be be saved in file for use
in other applications.</p>
'@;

Write-Html -ParagraphStyle:'Text' @'
<p align=justify>Creating new <b>Map</b> starts with cmdlet </i>New-Map</i>,
ends with cmdlet <i>Write-Map</i> or <i>Save-Map</i>. Multiple other cmdlets
(help on <b>Map</b> cmdlets is provided at end of this <b>Book</b>) allow
to customize maps, add multiple layers, titles etc.
Examples below demonstrate basic operations with <b>Map</b>.</p>
'@;

Write-Html -ParagraphStyle:'Text' @'
<p align=justify><b>Map</b> allows to build map using <i>Bing</i> and
<i>OpenStreet</i> providers, as well as loading .shp, and .kml files and
adding <i>Vector items</i> such as <i>Pushpin</i>, <i>Dot</i>, <i>Line</i>,
<i>Rectangle</i>, <i>Ellipse</i> etc.</p>
'@;

Write-Html -ParagraphStyle:'Description' @'
<p>Current script is accompanied with script file ReadMe.ps1 with 
source code.</p>
'@;


Add-BookPageBreak;
Write-Text -ParagraphStyle:'Header3' 'BASIC MAP';
Write-Html -ParagraphStyle:'Text' @'
<p align=justify>First map shows basic <b>Map</b> - <i>World Map</i> obtained
from <i>Bing</i> and <i>OpenStreet</i> providers with different zoom levels.</p>
'@;

New-Map -BackColor:White | 
	Add-MapLayerImage Bing Road | 
	Write-Map -CenterPoint:@(0,0) -Width:2000 -Height:2000 -ZoomLevel:0.7;
	
New-Map | 
	Add-MapLayerImage OpenStreet CycleMap | 
	Add-MapLayerImage OpenStreet PublicTransport | 
	Write-Map -CenterPoint:@(45,0) -Width:2000 -Height:2000 -ZoomLevel:8;
	
New-Map | 
	Add-MapLayerImage OpenStreet CycleMap | 
	Add-MapLayerImage OpenStreet PublicTransport | 
	Write-Map -CenterPoint:@(45,0) -Width:2000 -Height:2000 -ZoomLevel:12;


Add-BookPageBreak;
Write-Text -ParagraphStyle:'Header3' 'DRAW SHP FILE';
Write-Html -ParagraphStyle:'Text' @'
<p align=justify>Next sample shows how to draw <i>.shp</i> file and add <b>Map</b>'s elements -
<i>Colorizer</i>, <i>Legend</i>, <i>Overlay</i>.</p>
'@;
Write-Html -ParagraphStyle:'Description' @'
<p align=justify>Sample <i>.shp</i> file is from 
<a href="https://github.com/nvkelso/natural-earth-vector">Natural Eartch Project</a></p>
'@;
New-Map |
	Add-MapLayerVectorFile '~#\Data\ne_50m_admin_0_countries.shp' |
	Add-MapChoroplethColorizer 'GDP_MD_EST' @(0, 3000, 10000, 180000, 28000, 44000, 82000, 185000, 1000000, 2500000, 15000000) `
		@('#5F8B95', '#799689', '#A2A875', '#CEBB5F', '#F2CB4E', '#F1C149', '#E5A84D', '#D6864E', '#C56450', '#BA4D51') |
	Add-MapLegend ColorScale -Alignment:BottomCenter -Description:'Map legend description' -Header:'Map legend header' |
	Add-MapLegend ColorList -Alignment:MiddleLeft -Description:'Map legend description' -Header:'Map legend header' |
	Add-MapOverlay -Text:'GDP' -Alignment:TopCenter `
		-Font:'Arial,30,White' -Margin:20 -Padding:10 |
	Write-Map -CenterPoint:@(0, 0) -ZoomLevel:2 -Width:4000 -Height:3200 -Scale:0.5;
	
	
Add-BookPageBreak;
Write-Text -ParagraphStyle:'Header3' 'CUSTOM VECTOR ELEMENTS';
Write-Html -ParagraphStyle:'Text' @'
<p align=justify>Next sample shows how to place custom objects to <b>Map</b>.
First <b>Map</b> show how to add objects from <i>List</i> (different data sources
such as <i>DataTable</i> can be used); second <b>Map</b> adds items individually and
connects them with lines, it also shows how to use <b>Map</b>'s <i>Context</i> 
(variable <i>$map</i>).</p>
'@;
$mapItems = @(
	[PSCustomObject] @{ Latitude = 55.7496; Longitude = 37.6237;   Name = 'Moscow' },
	[PSCustomObject] @{ Latitude = 39.904;  Longitude = 116.4075;  Name = 'Beijing' },
	[PSCustomObject] @{ Latitude = 35.6895; Longitude = 139.6917;  Name = 'Tokyo' },
	[PSCustomObject] @{ Latitude = 37.7749; Longitude = -122.4194; Name = 'San-Francisco' },
	[PSCustomObject] @{ Latitude = 40.7144; Longitude = -74.006;   Name = 'New-York' },
	[PSCustomObject] @{ Latitude = 48.8566; Longitude = 2.3522;    Name = 'Paris' }
);

New-Map -BackColor:White | 
	Add-MapLayerImage Bing Road -CultureName:zh-Hant | 
	Add-MapLayerVectorData $mapItems -DefaultMapItemType:Pushpin `
		-LatitudeField:'Latitude' -LongitudeField:'Longitude' -TextField:'Name' |
	Write-Map -CenterPoint:@(0.0, 0.0) -Width:2000 -Height:1600 -ZoomLevel:1;

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
	Out-Null;

$map | Write-Map -CenterPoint:@(0.0, 0.0) -Width:2000 -Height:1600 -ZoomLevel:1;


Add-BookPageBreak;
Write-Text -ParagraphStyle:'Header3' 'SEARCH';
Write-Html -ParagraphStyle:'Text' @'
<p align=justify>Next sample shows how to search geo locations in <i>Bing</i> and
<i>OpenStreet</i> providers.</p>
'@;
New-Map -BackColor:White |
	Add-MapLayerImage Bing Road | 
	Add-MapLayerSearch Bing 'Moscow' |
	Write-Map -CenterPoint:@(0, 0) -ZoomLevel:1;
	
New-Map -BackColor:White |
	Add-MapLayerImage OpenStreet CycleMap | 
	Add-MapLayerSearch OpenStreet 'Moscow' |
	Write-Map -CenterPoint:@(0, 0) -ZoomLevel:1;
	

Add-BookPageBreak;
Write-Text -ParagraphStyle:'Header3' 'ROAD AND MINIMAP';
Write-Html -ParagraphStyle:'Text' @'
<p align=justify>Next sample shows how to draw a <i>Road</i> on <b>Map</b> and
how to add <i>MiniMap</i>.</p>
'@;
New-Map |
	Add-MapLayerImage Bing Road | 
	Add-MapLayerRoute Bing  @(@{Description='Marseille'; Location=@(43.2965, 5.3698)}, `
		@{Description='Bordeaux'; Location=@(44.8378, -0.5791)}, @{Description='Paris'; Location=@(48.8566, 2.3522)}) `
		-ShapeTitlesPattern:'{Description}' -StrokeColor:MidnightBlue -StrokeWidth:5 |
	Add-MapOverlay -Text:'ROAD Marseille - Bordeaux - Paris' -Alignment:TopCenter |
	Add-MapMiniMap -Alignment:TopLeft -CenterPoint:@(0,0) -ZoomLevel:0.3 |
	Add-MapLayerImage Bing Road -MiniMap |
	Write-Map -CenterPoint:@(46, 0) -ZoomLevel:5;


Add-BookPageBreak;
Write-Text -ParagraphStyle:'Header3' 'WEB MAP SERVICES';
Write-Html -ParagraphStyle:'Text' @'
<p align=justify><b>Map</b> allows to use <a href="https://en.wikipedia.org/wiki/Web_Map_Service">
WMS - Web Map Services</a>.</p>
'@;
New-Map -BackColor:White |
	Add-MapLayerImage WMS 'http://ows.mundialis.de/services/service' 'OSM-WMS' |
	Write-Map -CenterPoint:@(0, 0) -ZoomLevel:1;
	
New-Map -BackColor:White |
	Add-MapLayerImage WMS 'http://ows.mundialis.de/services/service' 'TOPO-OSM-WMS' |
	Write-Map -CenterPoint:@(0, 0) -ZoomLevel:1;
	
New-Map -BackColor:White |
	Add-MapLayerImage WMS 'http://ows.mundialis.de/services/service' 'SRTM30-Colored' |
	Write-Map -CenterPoint:@(0, 0) -ZoomLevel:0.7;

New-Map -BackColor:White |
	Add-MapLayerImage WMS 'http://ows.mundialis.de/services/service' 'SRTM30-Colored-Hillshade' |
	Write-Map -CenterPoint:@(0, 0) -ZoomLevel:0.7;


Add-BookPageBreak;
Write-Text -ParagraphStyle:'Header3' 'CARTESIAN COORDINATES';
Write-Html -ParagraphStyle:'Text' @'
<p align=justify><b>Map</b> allows to build block diagrams by adding items to the map
in Cartesian coordinates.</p>
'@;
New-Map -CoordinateSystem:Cartesian -BackColor:Black |
	Add-MapLayerVectorItems |
	Add-MapItem Rectangle @(0, 100) 60 60 -FillColor:Red  |
	Add-MapItem Ellipse @(50, 50) 60 60 -FillColor:Yellow  |
	Add-MapItem Line @(20, 20) @(80, 80) -StrokeColor:Blue -StrokeWidth:10 | 
	Write-Map -CenterPoint:@(50, 50) -ZoomLevel:1 -Width:2000 -Height:2000;


Add-BookSection -ContinuePageNumbering -LinkHeaderToPrevious -LinkFooterToPrevious;

Write-Text -ParagraphStyle:'Header2' 'Map cmdlets';

. $schost.MapPath('~#\..\Common\CmdletHelp.ps1');

$cmdlets = [string[]]@(
	'Add-MapChoroplethColorizer',
	'Add-MapClusterer',
	'Add-MapGraphColorizer',
	'Add-MapItem|Help for cmdlet <i>Add-MapItem</i> with parameter sets for different items (Line, Rectangle, Ellipse etc) is provided in later section of this <b>Book</b>',
	'Add-MapKeyColorizer',
	'Add-MapLayerImage|Help for cmdlet <i>Add-MapLayerImage</i> with parameter sets for different laters (Bing, OpenStreet, Heat, Wms) is provided in later section of this <b>Book</b>',
	'Add-MapLayerRoute|Help for cmdlet <i>Add-MapLayerRoute</i> with parameter sets for different providers (currently only Bing) is provided in later section of this <b>Book</b>',
	'Add-MapLayerSearch|Help for cmdlet <i>Add-MapLayerSearch</i> with parameter sets for different providers (Bing, OpenStreet) is provided in later section of this <b>Book</b>',
	'Add-MapLayerSql',
	'Add-MapLayerVectorData',
	'Add-MapLayerVectorFile',
	'Add-MapLayerVectorItems',
	'Add-MapLayerWkt',
	'Add-MapLegend',
	'Add-MapMiniMap',
	'Add-MapOverlay',
	'New-Map',
	'Save-Map',
	'Write-Map'
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

Write-Html -ParagraphStyle:'Header3' 'Cmdlet <i>Add-MapItem</i>: dynamic parameter sets';
Write-Html -ParagraphStyle:'Description' 'Table of Contents is at end of the <b>Book</b>.';

$firstParameterSet = $true;
@(
	'Dot',
	'Ellipse',
	'Line',
	'Polygon',
	'Polyline',
	'Rectangle',
	'Pushpin',
	'Custom',
	'Callout',
	'Bubble'
) |
%{
	if (-not $firstParameterSet) { Add-BookPageBreak; }
	
	$help = GenerateCmdletParametersHelp 'Add-MapItem' $_;
	
	Write-Html -ParagraphStyle:'Header4' "Parameters <i>$_</i>";
	Write-Html $help -ParagraphStyle:'CmdletHelp';
	
	$firstParameterSet = $false;
};


Add-BookPageBreak;

Write-Html -ParagraphStyle:'Header3' 'Cmdlet <i>Add-MapLayerImage</i>: dynamic parameter sets';
Write-Html -ParagraphStyle:'Description' 'Table of Contents is at end of the <b>Book</b>.';

$firstParameterSet = $true;
@(
	'Bing',
	'Heatmap',
	'OpenStreet',
	'WMS'
) |
%{
	if (-not $firstParameterSet) { Add-BookPageBreak; }
	
	$help = GenerateCmdletParametersHelp 'Add-MapLayerImage' $_;
	
	Write-Html -ParagraphStyle:'Header4' "Parameters <i>$_</i>";
	Write-Html $help -ParagraphStyle:'CmdletHelp';
	
	$firstParameterSet = $false;
};


Add-BookPageBreak;

Write-Html -ParagraphStyle:'Header3' 'Cmdlet <i>Add-MapLayerRoute</i>: dynamic parameter sets';
Write-Html -ParagraphStyle:'Description' 'Table of Contents is at end of the <b>Book</b>.';

$firstParameterSet = $true;
@(
	'Bing'
) |
%{
	if (-not $firstParameterSet) { Add-BookPageBreak; }
	
	$help = GenerateCmdletParametersHelp 'Add-MapLayerRoute' $_;
	
	Write-Html -ParagraphStyle:'Header4' "Parameters <i>$_</i>";
	Write-Html $help -ParagraphStyle:'CmdletHelp';
	
	$firstParameterSet = $false;
};


Add-BookPageBreak;

Write-Html -ParagraphStyle:'Header3' 'Cmdlet <i>Add-MapLayerSearch</i>: dynamic parameter sets';
Write-Html -ParagraphStyle:'Description' 'Table of Contents is at end of the <b>Book</b>.';

$firstParameterSet = $true;
@(
	'Bing',
	'OpenStreet'
) |
%{
	if (-not $firstParameterSet) { Add-BookPageBreak; }
	
	$help = GenerateCmdletParametersHelp 'Add-MapLayerSearch' $_;
	
	Write-Html -ParagraphStyle:'Header4' "Parameters <i>$_</i>";
	Write-Html $help -ParagraphStyle:'CmdletHelp';
	
	$firstParameterSet = $false;
};


Add-BookSection -ContinuePageNumbering -LinkHeaderToPrevious -LinkFooterToPrevious;
Write-Text -ParagraphStyle:'Header2' 'Table of Contents';
Add-BookTOC;

Save-Book '~#\ReadMe.docx' -Replace;
