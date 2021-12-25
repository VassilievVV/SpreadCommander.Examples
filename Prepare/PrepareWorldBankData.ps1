#Parameters
$indicators = @(
	'AG.LND.IRIG.AG.ZS',
	'AG.LND.FRST.ZS',
	'AG.LND.TOTL.K2',
	'AG.LND.AGRI.ZS',

	'NY.GDP.MKTP.KD',
	'NY.GDP.PCAP.KD',

	'NY.GDP.MKTP.CD',
	'NY.GDP.PCAP.CD',
	'NY.GDS.TOTL.CD',
	'NY.GDP.MKTP.PP.CD',
	'NY.GDP.PCAP.PP.CD',

	'EG.CFT.ACCS.ZS',
	'EG.ELC.LOSS.ZS',
	'EG.GDP.PUSE.KO.PP',
	'EG.USE.ELEC.KH.PC',
	'EG.USE.PCAP.KG.OE'
);

$connectionString = 'sqlite:~#\..\Data\WorldData.db';
$spreadTableStyle = 'Medium25';

#Common code
$schost = Get-SCHost;
Add-Type -AssemblyName System.IO.Compression.FileSystem;


#Import metadata
$dirData      = $schost.MapPath('~#\..\Data');
$fileMetadata = $schost.MapPath('~#\..\Data\WDI_Metadata.zip');

$fileCountries       = $schost.MapPath('~#\..\Data\WDICountry.csv');
$fileSeries          = $schost.MapPath('~#\..\Data\WDISeries.csv');
$fileCountriesSeries = $schost.MapPath('~#\..\Data\WDICountry-Series.csv');

if (Test-Path $fileCountries) 
	{ Remove-Item -LiteralPath:$fileCountries -Confirm:$false; }
if (Test-Path $fileSeries) 
	{ Remove-Item -LiteralPath:$fileSeries -Confirm:$false; }
if (Test-Path $fileCountriesSeries) 
	{ Remove-Item -LiteralPath:$fileCountriesSeries -Confirm:$false; }

[IO.Compression.ZipFile]::ExtractToDirectory($fileMetadata, $dirData);

Write-Text 'Importing Countries ...';
$csv = Import-Csv -Path:$fileCountries;
$csv | Out-SpreadTable -TableName:'Countries' -SheetName:'Countries' -TableStyle:$spreadTableStyle -Replace;
$csv | Export-TableToDatabase $connectionString 'Countries' -Replace `
		-PostScript:@'
create unique index IX_Countries_CountryCode on Countries ([Country Code]);
create unique index IX_Countries_ShortName on Countries ([Short Name]);
create unique index IX_Countries_TableName on Countries ([Table Name]);
create index IX_Countries_Region on Countries ([Region]);
create index IX_Countries_IncomeGroup on Countries ([Income Group]);
create index IX_Countries_WB2Code on Countries ([WB-2 code]);
create index IX_Countries_SNAPriceValuation on Countries ([SNA price valuation]);
create index IX_Countries_LendingCategory on Countries ([Lending category]);
create index IX_Countries_OtherGroups on Countries ([Other groups]);
create index IX_Countries_SystemOfNatinalAccounts on Countries ([System of National Accounts]);
create index IX_Countries_BalanceOfPaymentsManualInUse on Countries ([Balance of Payments Manual in use]);
create index IX_Countries_ExternalDebtReportingStatus on Countries ([External debt Reporting status]);
create index IX_Countries_SystemOfTrade on Countries ([System of trade]);
create index IX_Countries_GovernmentAccountingConcept on Countries ([Government Accounting concept]);
create index IX_Countries_IMFDataDisseminationStandard on Countries ([IMF data dissemination standard]);
create index IX_Countries_VitalRegistrationComplete on Countries ([Vital registration complete]);
'@;
Remove-Item -LiteralPath:$fileCountries -Confirm:$false;
	
Write-Text 'Importing Series ...';	
$csv = Import-Csv -Path:$fileSeries;
$csv | Out-SpreadTable -TableName:'Series' -SheetName:'Series' -TableStyle:$spreadTableStyle -Replace;
$csv | Export-TableToDatabase $connectionString 'Series' -Replace `
		-PostScript:@'
create unique index IX_Series_SeriesCode on Series ([Series Code]);
create index IX_Series_Topic on Series ([Topic]);
create unique index IX_Series_IndicatorName on Series ([Indicator Name]);
create index IX_Series_UnitOfMeasure on Series ([Unit of measure]);
create index IX_Series_Periodicity on Series ([Periodicity]);
create index IX_Series_BasePeriod on Series ([Base Period]);
create index IX_Series_AggregationMethod on Series ([Aggregation method]);
create index IX_Series_Source on Series ([Source]);
create index IX_Series_LicenseType on Series ([License Type]);
'@;
Remove-Item -LiteralPath:$fileSeries -Confirm:$false;

Write-Text 'Importing Countries-Series ...';
$csv = Import-Csv -Path:$fileCountriesSeries;
$csv | Out-SpreadTable -TableName:'CountriesSeries' -SheetName:'Countries-Series' -TableStyle:$spreadTableStyle -Replace;
$csv | Export-TableToDatabase $connectionString 'CountrySeries' -Replace `
		-PostScript:@'
create index IX_CountrySeries_CountryCode on CountrySeries ([CountryCode]);
create index IX_CountrySeries_SeriesCode on CountrySeries ([SeriesCode]);
'@;
Remove-Item -LiteralPath:$fileCountriesSeries -Confirm:$false;

$csv = $null;


#Import indicators
function DownloadIndicatorData([string]$indicator)
{
	Write-Text "Importing indicator: $indicator ...";
	
	$url     = "http://api.worldbank.org/v2/en/indicator/$($indicator)?downloadformat=xml";
	$zipName = $schost.MapPath("~#\..\Data\$indicator.zip");
	$xmlFile = $schost.MapPath("~#\..\Data\$indicator.xml");
	$filter  = "API_$($indicator)_*.xml";

	$wc = [Net.WebClient]::new();
	try
	{
		$wc.DownloadFile($url, $zipName);
	}
	finally
	{
		$wc.Dispose();
	}

	$zip = [IO.Compression.ZipFile]::OpenRead($zipName);
	try
	{
		$zip.Entries | 
		?{ $_.FullName -like $filter } |
		Select-Object -First:1 | 
		%{ [IO.Compression.ZipFileExtensions]::ExtractToFile($_, $xmlFile, $true); };
	}
	finally
	{
		$zip.Dispose();
	}

	$table = [Data.DataTable]::new();
	try
	{
		[void]$table.Columns.Add('Country Code', [string]);
		[void]$table.Columns.Add('Country or Area', [string]);
		[void]$table.Columns.Add('Year', [int]);
		[void]$table.Columns.Add('Value', [double]);

		Select-Xml -LiteralPath:$xmlFile -XPath:'/Root/data/record' | 
		%{ 
			$countryCode   = $_.Node.SelectSingleNode('field[@name="Country or Area"]/@key').Value;
			$countryOrArea = $_.Node.SelectSingleNode('field[@name="Country or Area"]').InnerText;
			$year          = [int]$_.Node.SelectSingleNode('field[@name="Year"]').InnerText;
			$value         = [double]$_.Node.SelectSingleNode('field[@name="Value"]').InnerText;
			
			if ($value -eq 0) { $value = $null; }
			
			[void]$table.Rows.Add($countryCode, $countryOrArea, $year, $value);
		};
		
		Remove-Item -LiteralPath:$xmlFile -Confirm:$false;

		#$table | Out-Data -TableName:$indicator -Replace;
		
		$table | 
			ConvertTo-Pivot @('Country Code', 'Country or Area') 'Year' 'Value' First |
			Out-SpreadTable -TableName:$indicator -SheetName:$indicator -TableStyle:$spreadTableStyle -Replace;

		#indicator name to use inside index names
		$indicator2 = $indicator.Replace('.', '_');
		#Do not save 'Country or Area' into SQLite database, 
		#if needed - it can be obtained through join with table Countries
		$table.Columns.Remove('Country or Area');

		$table |
			Export-TableToDatabase $connectionString $indicator -Replace `
				-PostScript:@"
create index [IX_$($indicator2)_CountryCode] on [$indicator] ([Country Code]);
create index [IX_$($indicator2)_Year] on [$indicator] ([Year]);
"@;
	}
	finally
	{
		$table.Dispose();
	}
}

$indicators | %{ DownloadIndicatorData($_) };

Save-Spreadsheet '~#\..\Data\WorldBank.xlsx' -Replace;