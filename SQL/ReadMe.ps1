$schost = Get-SCHost;                              
$schost.Silent = $true;

Clear-Book;
Clear-Data;

Invoke-SCScript '~\..\Common\InitBookStyles.ps1';

Set-BookSectionHeader '<b>Spread Commander</b> - <i>Examples: SQL</i>' -Html;
Set-BookSectionFooter 'Page {PAGE} of {NUMPAGES}' -ExpandFields;

Write-Text -ParagraphStyle:'Header1' 'SQL';

Write-Html -ParagraphStyle:'Description' @'
<p align=justify><b>SpreadCommander</b> includes a control to work with <b>SQL</b> queries.
<b>SpreadCommander</b> contains optimized connections for <i>MS SQL Server</i>,
<i>MySQL</i> and <i>SQLite</i> providers; it also allows to connect to 
<i>ODBC</i> and <i>OLEDB</i> providers.</p>
'@;

Write-Html -ParagraphStyle:'Text' @'
<p align=justify><b>SpreadCommander</b> provides <i>cmdlets</i> to simplify 
work with databases.</p>

<p align=justify>Connectuons (connection strings) can be stored 
with <i>cmdlet</i> <i>Set-DbConnenection</i> and retrieved by name 
with <i>cmdlet</i> <i>Get-DbConnection</i>. Although <b>SpreadCommander</b>
encrypts connection strings, it is an application with open source code 
so storing passwords this way is not completely reliable, when possible
it is recommended to use <i>Windows authentication</i>.</p>

<p align=justify><i>Cmdlet</i> <i>Get-DbConnection</i> allows to specify 
not only name of saved connection but also simplified connection string.
SQLite allows to specify file path under project directory with '~\'
(i.e. '~\Data\mydb.sqlite' is file mydb.sqlite in folder Data under
project's root directory. Supported custom connection strings are:</p>

<table>
	<tr>
		<td><b>sqlite:</b></td>
		<td align=justify><Prefix for <i>SQLite</i> database. First part of connection string
			is <i>Data Source</i>. For example 'sqlite:~\Data\MyDb.sqlite' opens
			connection to database located in file Data\MyDb.sqlite under 
			current project's directory.</td>
	</tr>
	<tr>
		<td><b>mssql:</b></td>
		<td align=justify>Prefix for <i>Microsoft SQL Server</i> database. First part of
			connection string is <i>Database</i>. Default server can be stored
			in settings; if User ID is provided - it will be used, if no - 
			Windows authentication will be used. Connection to default server
			allows to not specify it in connection string. For example
			'mssql:Northwind' opens connection to database Northwind located
			at default server (it has to be configured in settings) using
			Windows authentication; 'mssql:Northwind;Server=RemoteServer\ServerName;User ID=user;Password=pwd'
			connects to database Northwind at server specified in connection string.</td>
	</tr>
	<tr>
		<td><b>mysql:</b></td>
		<td align=justify>Prefix for <i>MySQL</i> database. First part of connection string is
			<i>Database</i>. Default server can be stored in settings. Connection to 
			default server allows to not specify it in connection string. For example
			'mysql:world' opens connection to database World located
			at default server (it has to be configured in settings); 
			'mysql:World;Server=127.0.0.1;User ID=user;Password=pwd'
			connects to database World at server specified in connection string.
		</td>
	</tr>
	<tr>
		<td><b>odbc:</b></td>
		<td align=justify>Prefix for <i>ODBC</i> database. Requires full 
			connection string.</td>
	</tr>
	<tr>
		<td><b>oldeb:</b></td>
		<td align=justify>Prefix for <i>OLEDB</i> database. Requires full 
			connection string.</td>
	</tr>
</table>
'@;

Write-Html -ParagraphStyle:'Description' @'
<p>Current script is accompanied with script file ReadMe.ps1 with 
source code.</p>
'@;

Add-BookPageBreak;
Write-Text -ParagraphStyle:'Header2' 'Extended SQL syntax';

Write-Html -ParagraphStyle:'Text' @'
<p align=justify><b>SpreadCommander</b> supports extended syntax for 
<i>SQL queries</i>. In addition to <i>SQL</i> commands that will be
sent to the <i>SQL Server</i> and executed, <b>SpreadCommander</b>
supports <i>client-side commands</i>.</p>

<p align=justify><i>Server-side commands</i> determine which data
will be returned from the <i>SQL Server</i>.</p>

<p align=justify><i>Client-side commands</i> determine how data
will be presented at client side. This includes table names, 
formatting, relations (master-detail) between tables. <i>Table</i>
here means table in a <i>DataSet</i> at client, not table in database.</p>

<p align=justify>Look at script <i>SampleData.sql</i> in this project
for an example for <i>client-side commands</i> in SQL script.
<i>Client-side commands</i> are replaced with spaces before 
sending command to <i>SQL Server</i> and will work even if 
<i>SQL Server</i> does not support comments.</p>

<p align=justify><i>Client-side commands</i> start with prefix <span style='color:green'>--#</span>
and continue to end of line, or embedded into block <span style='color:green'>/*#...#*/</span>
(where '...' is command text). In latter case commands are separated with semicolor.</p>

<p align=justify><i>Client-side commands</i> can be located anywhere in <i>script command</i>
(i.e. anywhere between GO lines). Table name assigned with command <i>TABLE</i> will be
applied to first, second etc. result set returned from <i>script command</i>, not necessarry
to result set that gramatically follows command <i>TABLE</i>.</p>
'@;

Write-Text -ParagraphStyle:'Header3' 'Command TABLE';
Write-Html -ParagraphStyle:'Text' @'
<p align=justify>First command is <i>TABLE</i>. It allows to specify table in 
at client side. This table name will be displayed as tab header in <i>Data grid</i> and
it is used in other commands (<i>FORMAT CONDITION</i>, <i>RELATION</i>) to refer
current table. Command <i>TABLE</i> accepts properties <i>GroupBy</i> and <i>OrderBy</i>,
these properties define how data will be shown in <b>Grid</b>.</p>

<p align=left style='color:green'>--#table Series with GroupBy="Topic Group", OrderBy="Topic Part desc";</p>
'@;

Write-Text -ParagraphStyle:'Header3' 'Command FORMAT CONDITION';
Write-Html -ParagraphStyle:'Text' @'
<p align=justify>Command <i>FORMAT CONDITION</i> allows to specify conditional formatting for data
returned from <i>SQL Server</i>.</p>

<p style='color:green'>--#format table Series column [Indicator Name] for "Contains([Indicator Name], 'Agricultural')" with BackColor=Green, ForeColor=White;</p>
<p style='color:green'>--#format table "Energy use - by Regions" column MaxValue with IconSet=Signs3;</p>

<p align=justify>To review supported format condition in <b>Grid</b> click on button "Formatting" on ribbon tab "Data"
when <b>Grid</b> is active, or right-click on <i>column header</i> and choose <i>Conditional Formatting/Manage Rules</i>. 
Displayed dialog allows to edit existing <i>format conditions</i> and add new ones.</p>

<p align=justify>Following <i>format conditions</i> are supported:</p>

<table>
	<tr>
		<td><b>Expression</b></td>
		<td><p align=justify>Applies format if row matches specified condition.</p>
			<p style='color:green'>--#format table Series column [Indicator Name] for "Contains([Indicator Name], 'Agricultural')" with BackColor=Green, ForeColor=White;</p></td>
	</tr>
	<tr>
		<td><b>DataBar</b></td>
		<td><p align=justify>Applies format using a <i>DataBar</i>. Bar length changes proprotionally to a cell value.</p>
			<p style='color:green'>--#format table [Energy use - by Income] column [AverageValue] with Rule=DataBar, BackColor='LightGreen', BackColor2='LightGray', Gradient=Horizontal</p></td>
	</tr>
	<tr>
		<td><b>IconSet</b></td>
		<td><p align=justify><i>IconSet</i> allows to classify values into ranges and display specific icon according to the range.</p>
			<p style='color:green'>--#format table "Energy use - by Regions" column MaxValue with IconSet=Signs3;</p></td>
	</tr>
	<tr>
		<td><b>ColorScale</b></td>
		<td><p align=justify>Allows to disploay data distribution and variation using a gradation of 2 or 3 colors.</p>
			<p style='color:green'>#format table [Energy use - by Income] column [AverageValue] with ColorScale='Red,LightGreen,White'</p></td>
	</tr>
	<tr>
		<td>
			<b>AboveAverage</b><br>
			<b>BelowAverage</b><br>
			<b>AboveOrEqualAverage</b><br>
			<b>BelowOrEqualAverage</b><br>
			<b>Unique</b><br>
			<b>Duplicate</b>
		</td>
		<td><p align=justify>Applies format if value corresponds specified condition.</p>
			<p style='color:green'>--#format table "Energy use" column Value with Rule=AboveAverage, BackColor=LightGreen</p></td>
	</tr>
	<tr>
		<td>
			<b>Top</b><br>
			<b>Bottom</b>
		</td>
		<td><p align=justify>Applies format if value belongs to top or bottom N values. 
			N can be a number or a percent.</p>
			<p style='color:green'>--#format table "Energy use" column Value with Rule=Top, Rank='10%', BackColor=LightGreen</p></td>
	</tr>
	<tr>
		<td><b>DateOccuring</b></td>
		<td><p align=justify>Applies a format if value refers to specific date or date interval relative to <i>today</i>.
			Day intervalus include Beyond, BeyondThisYear, Earlier, EarlierThisMonth, EarlierThisWeek, EarlierThisYear, 
			Empty, LastWeek, LaterThisMonth, LaterThisWeek, LaterThisYear, MonthAfter1, MonthAfter2, MonthAgo1, MonthAgo2, 
			MonthAgo3, MonthAgo4, MonthAgo5, MonthAgo6, NextWeek, PriorThisYear, SpecificDate, ThisMonth, ThisWeek, 
			Today, Tomorrow, User, Yesterday</p>
			<p style='color:green'>--#format table [Table1] column [Value] with DateOccuring=Today, BackColor=Yellow</p></td>
	</tr>
	<tr>
		<td>
			<b>Rule</b><br>
			<b>Comparison</b>
		</td>
		<td><p align=justify>Applies a format if value meets specified comparison condition. 
			Comparison types include Between, Equal, Expression, Greater, GreaterOrEqual, 
			Less, LessOrEqual, NotBetween, NotEqual.</p>
			<p style='color:green'>--#format table Regions column [Table Name] with Condition=Equal, Value1='OECD members', BackColor=Green, ApplyToRow=true</p></td>
	</tr>
</table>

<p align=justify>Most <i>format conditions</i> (except DataBar, IconSet and ColorScale) support properties that specify
<i>appearance</i> These properties include:</p>
<ul>
	<li>BackColor</li>
	<li>BackColor2 (to use with <i>Gradient</i></li>
	<li>BorderColor</li>
	<li>Font</li>
	<li>Gradient (requires <i>BackColor2</i>, one of following values: Horizontal, Vertical, ForwardDiagonal, BackwardDiagonal)</li>
</ul>
<br>
'@;

Write-Text -ParagraphStyle:'Header3' 'Command COMPUTED COLUMN';
Write-Html -ParagraphStyle:'Text' @'
<p align=justify><b>SpreadCommander</b> allows to add client-side computed columns to resulset of SQL queries.
To edit computed columns click on button "Computed columns" on ribbon tab "Data" when <b>Grid</b> is active.</p>

<p align=justify>Syntax is COMPUTED COLUMN [ColumnName] ColumnType IN [TableName] = 'Expression'.
ColumnType is one of: STRING (aliases - VARCHAR, CHAR, TEXT), INTEGER (alias INT), DECIMAL
(aliases NUMERIC, DOUBLE, FLOAT), BOOLEAN (alias BOOL).</p>

<p style='color:green'>--#computed column [Topic Part] string in Series = "REGEXMATCH([Topic], '(?<=:).*')"</p>
'@;

Write-Text -ParagraphStyle:'Header3' 'Command RELATION';
Write-Html -ParagraphStyle:'Text' @'
<p align=justify><b>SpreadCommander</b> executes <i>SQL queries</i> into <i>ADO.Net DataSet</i> and
allows to specify client-side <i>master-detail relations</i> between tables.</p>

<p align=justify>Syntax for <i>Relation</i> is: RELATION [RelationName] [ParentTable] ([ParentColumn(s)]) - 
[ChildTable] ([ChildColumn(s)]. Multiple column names are comma-separated.</p>

<p style='color:green'>--#relation Rel_Regions_EnergyUse [Regions] ([Short Name]) - [Energy use - by Regions] ([Region])</p>
'@;

Write-Text -ParagraphStyle:'Header3' 'Command CONNECTION';
Write-Html -ParagraphStyle:'Text' @'
<p align=justify><b>SpreadCommander</b> allows to specify different connections for different <i>SQL queries</i>
(part of script separated by line GO). Connection can be one defined in list of common connections 
(can be set with cmdlet <i>Set-DbConnection</i>) or connection string can be constructed for 
<i>Microsoft SQL Server</i>, <i>MySQL</i>, <i>SQLite</i>, <i>ODBC</i> or <i>OLEDB</i> providers.
Details see at beginning of this ReadMe file, in section <b>SQL</b>. If first query (part of script until first 
line GO) specifies <i>Connection</i> - it will be used by default to execute script, show execution plan and
other features.</p>

<p style='color:green'>--#connection "sqlite:~\..\Data\WorldData.db"</p>
'@;

Add-BookPageBreak;
Write-Text -ParagraphStyle:'Header2' 'Extended SQLite functions';

Write-Html -ParagraphStyle:'Text' @'
<p align=justify><b>SpreadCommander</b> adds extended set of functions to <i>SQLite</i>. These functions
allow basic data analysis in <i>SQL queries</i>.</p>
'@;

Write-Text -ParagraphStyle:'Header3' 'Scalar functions';
Write-Text -ParagraphStyle:'Header4' 'Hash';
Write-Html -ParagraphStyle:'Text' @'
<table>
	<tr>
		<td><b>MD5</b></td>
		<td>Computed <i>MD5</i> hash. Second optional parameter allows to specify encoding (Unicode, UTF8, ASCII etc).</td>
	</tr>
	<tr>
		<td><b>SHA1</b></td>
		<td>Computes <i>SHA1</i> hash. Second optional parameter allows to specify encoding (Unicode, UTF8, ASCII etc).</td>
	</tr>
	<tr>
		<td><b>SHA256</b></td>
		<td>Computes <i>SHA256</i> hash. Second optional parameter allows to specify encoding (Unicode, UTF8, ASCII etc).</td>
	</tr>
	<tr>
		<td><b>SHA384</b></td>
		<td>Computes <i>SHA384</i> hash. Second optional parameter allows to specify encoding (Unicode, UTF8, ASCII etc).</td>
	</tr>
	<tr>
		<td><b>SHA512</b></td>
		<td>Computes <i>SHA512</i> hash. Second optional parameter allows to specify encoding (Unicode, UTF8, ASCII etc).</td>
	</tr>
</table>
'@;

Write-Text -ParagraphStyle:'Header4' 'Math';
Write-Html -ParagraphStyle:'Text' @'
<table>
	<tr>
		<td><b>Acos</b></td>
		<td>Computes <i>acos</i> function.</td>
	</tr>
	<tr>
		<td><b>Asin</b></td>
		<td>Computes <i>asin</i> function.</td>
	</tr>
	<tr>
		<td><b>Atan</b></td>
		<td>Computes <i>atan</i> function.</td>
	</tr>
	<tr>
		<td><b>Ceiling</b></td>
		<td>Computes <i>Ceiling</i> function.</td>
	</tr>
	<tr>
		<td><b>Cos</b></td>
		<td>Computes <i>cos</i> function.</td>
	</tr>
	<tr>
		<td><b>Cosh</b></td>
		<td>Computes <i>cosh</i> function.</td>
	</tr>
	<tr>
		<td><b>Exp</b></td>
		<td>Computes <i>exp</i> function.</td>
	</tr>
	<tr>
		<td><b>Floor</b></td>
		<td>Computes <i>floor</i> function.</td>
	</tr>
	<tr>
		<td><b>IEEERemainder</b></td>
		<td>Computes <i>IEEEReminder</i> function.</td>
	</tr>
	<tr>
		<td><b>Log10</b></td>
		<td>Computes <i>log10</i> function.</td>
	</tr>
	<tr>
		<td><b>Log</b></td>
		<td>Computes <i>log</i> function.</td>
	</tr>
	<tr>
		<td><b>Pow</b></td>
		<td>Computes <i>power</i> function.</td>
	</tr>
	<tr>
		<td><b>Sign</b></td>
		<td>Computes <i>sign</i> function.</td>
	</tr>
	<tr>
		<td><b>Sin</b></td>
		<td>Computes <i>sin</i> function.</td>
	</tr>
	<tr>
		<td><b>Sinh</b></td>
		<td>Computes <i>sinh</i> function.</td>
	</tr>
	<tr>
		<td><b>Sqrt</b></td>
		<td>Computes <i>sqrt</i> function.</td>
	</tr>
	<tr>
		<td><b>Tan</b></td>
		<td>Computes <i>tan</i> function.</td>
	</tr>
	<tr>
		<td><b>Tanh</b></td>
		<td>Computes <i>tanh</i> function.</td>
	</tr>
	<tr>
		<td><b>Truncate</b></td>
		<td>Computes <i>truncate</i> function.</td>
	</tr>
</table>
'@;

Write-Text -ParagraphStyle:'Header4' 'Path';
Write-Html -ParagraphStyle:'Text' @'
<table>
	<tr>
		<td><b>PathChangeExtension</b></td>
		<td>Changes extension of provided filename.</td>
	</tr>
	<tr>
		<td><b>PathCombine</b></td>
		<td>Combines parts of the path.</td>
	</tr>
	<tr>
		<td><b>PathGetDirectoryName</b></td>
		<td>Returns directory name of the path.</td>
	</tr>
	<tr>
		<td><b>PathGetExtension</b></td>
		<td>Returns extension of the path.</td>
	</tr>
	<tr>
		<td><b>PathGetFileName</b></td>
		<td>Returns file name of the path.</td>
	</tr>
	<tr>
		<td><b>PathGetFileNameWithoutExtension</b></td>
		<td>Returns file name without extension of the path.</td>
	</tr>
</table>
'@;

Write-Text -ParagraphStyle:'Header4' 'Random';
Write-Html -ParagraphStyle:'Text' @'
<table>
	<tr>
		<td><b>RandNormal</b></td>
		<td>Returns random number with normal distribution.</td>
	</tr>
	<tr>
		<td><b>RandTriangular</b></td>
		<td>Returns random number with triangular distribution.</td>
	</tr>
	<tr>
		<td><b>RandUniform</b></td>
		<td>Returns random number with uniform distribution.</td>
	</tr>
</table>
'@;

Write-Text -ParagraphStyle:'Header4' 'Regular Expressions';
Write-Html -ParagraphStyle:'Text' @'
<table>
	<tr>
		<td><b>RegexIsMatch</b></td>
		<td>Returns 1 if string matches provided regular expression pattern and 0 otherwise.
			Parameters are input string and pattern.</td>
	</tr>
	<tr>
		<td><b>RegexMatch</b></td>
		<td>Searches the specified input string for the first occurrence of the regular expression.
			Parameters are input string, pattern and optional match number.</td>
	</tr>
	<tr>
		<td><b>RegexNamedMatch</b></td>
		<td>Searches the specified input string for the first occurrence of the regular expression
			and returns value of named group. Parameters are input string, pattern, 
			group name and optional match number.</td>
	</tr>
	<tr>
		<td><b>RegexReplace</b></td>
		<td>In a specified input string, replaces all strings that match a specified regular expression 
			with a specified replacement string. Parameters are input string, pattern and replacement.</td>
	</tr>
</table>
'@;

Write-Text -ParagraphStyle:'Header4' 'String';
Write-Html -ParagraphStyle:'Text' @'
<table>
	<tr>
		<td><b>StringFormat</b></td>
		<td>Formats string in .Net style.</td>
	</tr>
</table>
'@;

Write-Text -ParagraphStyle:'Header4' 'GUID';
Write-Html -ParagraphStyle:'Text' @'
<table>
	<tr>
		<td><b>NewID</b></td>
		<td>Returns new GUID. Format ("N", "D", "B, "P" or "X") can be specified as first parameter.
		    Default format is "D".</td>
	</tr>
</table>
'@;

Write-Text -ParagraphStyle:'Header3' 'Aggreagate functions';
Write-Html -ParagraphStyle:'Text' @'
<table>
	<tr>
		<td><b>AVG_GEOM</b></td>
		<td>Evaluates the geometric mean. </td>
	</tr>
	<tr>
		<td><b>AVG_HARMONIC</b></td>
		<td>Evaluates the harmonic mean.</td>
	</tr>
	<tr>
		<td><b>CORR</b></td>
		<td>Computes the Pearson Product-Moment Correlation coefficient.</td>
	</tr>
	<tr>
		<td><b>CORR_SPEARMAN</b></td>
		<td>Spearman correlation.</td>
	</tr>
	<tr>
		<td><b>COVAR</b></td>
		<td>Estimates the unbiased population covariance from the provided samples. 
			On a dataset of size N will use an N-1 normalizer (Bessel's correction).</td>
	</tr>
	<tr>
		<td><b>COVARP</b></td>
		<td>valuates the population covariance from the provided full populations. 
			On a dataset of size N will use an N normalizer and would thus be biased 
			if applied to a subset.</td>
	</tr>
	<tr>
		<td><b>STDEV</b></td>
		<td>Estimates the unbiased population standard deviation from the provided samples. 
			On a dataset of size N will use an N-1 normalizer (Bessel's correction). </td>
	</tr>
	<tr>
		<td><b>STDEVP</b></td>
		<td>Evaluates the variance from the provided full population. 
			On a dataset of size N will use an N normalizer and would thus be biased 
			if applied to a subset.</td>
	</tr>
	<tr>
		<td><b>VAR</b></td>
		<td>Estimates the unbiased population variance from the provided samples. 
			On a dataset of size N will use an N-1 normalizer (Bessel's correction).</td>
	</tr>
	<tr>
		<td><b>VARP</b></td>
		<td>Evaluates the variance from the provided full population. 
			On a dataset of size N will use an N normalizer and would thus be biased 
			if applied to a subset.</td>
	</tr>
	
	<tr>
		<td><b>InterQuantileRange</b></td>
		<td>Estimates the inter-quartile range from the provided samples. 
			Approximately median-unbiased regardless of the sample distribution.</td>
	</tr>
	<tr>
		<td><b>Kurtosis</b></td>
		<td>Estimates the unbiased population kurtosis from the provided samples. 
			Uses a normalizer (Bessel's correction; type 2).</td>
	</tr>
	<tr>
		<td><b>KurtosisP</b></td>
		<td>Evaluates the kurtosis from the full population. 
			Does not use a normalizer and would thus be biased 
			if applied to a subset (type 1).</td>
	</tr>
	<tr>
		<td><b>Median</b></td>
		<td>Estimates the sample median from the provided samples.</td>
	</tr>
	<tr>
		<td><b>Skewness</b></td>
		<td>Estimates the unbiased population skewness from the provided samples. 
			Uses a normalizer (Bessel's correction; type 2).</td>
	</tr>
	<tr>
		<td><b>SkewnessP</b></td>
		<td>Evaluates the skewness from the full population. 
			Does not use a normalizer and would thus be biased 
			if applied to a subset (type 1).</td>
	</tr>
	<tr>
		<td><b>Quantile</b></td>
		<td>Estimates the tau-th quantile from the provided samples. 
			The tau-th quantile is the data value where the cumulative 
			distribution function crosses tau. Approximately median-unbiased 
			regardless of the sample distribution.</td>
	</tr>
	<tr>
		<td><b>RMS</b></td>
		<td>Evaluates the root mean square (RMS) also known as quadratic mean.</td>
	</tr>
</table>
'@;

Write-Text -ParagraphStyle:'Header3' 'Collations';
Write-Html -ParagraphStyle:'Text' @'
<table>
	<tr>
		<td><b>Logical</b></td>
		<td>Logical collation. Strings '1', '2' .. '10', '20' are ordered according
			numeric values, not 1, 10, 2, 20. Case sensitive.</td>
	</tr>
	<tr>
		<td><b>LogicalCI</b></td>
		<td>Logical collation. Strings '1', '2' .. '10', '20' are ordered according
			numeric values, not 1, 10, 2, 20. Case insensitive.</td>
	</tr>
</table>
'@;


Add-BookSection -ContinuePageNumbering -LinkHeaderToPrevious -LinkFooterToPrevious;

Write-Text -ParagraphStyle:'Header2' 'SQL cmdlets';

Write-Html -ParagraphStyle:'Text' @'
This section contains help for <i>cmdlets</i> that allow to output data (not only
results of executing <i>SQL queries</i>) into <b>Data</b> tab.
'@;

. $schost.MapPath('~\..\Common\CmdletHelp.ps1');

$cmdlets = [string[]]@(
	'Out-Data',
	'Out-DataSet',
	'Remove-DataTable',
	'Clear-Data'
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
