$schost = Get-SCHost;                                
$schost.Silent = $true;

Clear-Book;

Invoke-SCScript '~#\..\Common\InitBookStyles.ps1';

Set-BookSectionHeader '<b>Spread Commander</b> - <i>Examples: Async</i>' -Html;
Set-BookSectionFooter 'Page {PAGE} of {NUMPAGES}' -ExpandFields;

Write-Text -ParagraphStyle:'Header1' 'ASYNC';

Write-Html -ParagraphStyle:'Description' @'
<p>This example shows how to use async commands in
<b>PowerShell</b> that supports <b>SpreadCommander</b>
environment, particularly <b>SpreadCommander</b> <i>cmdlets</i>.</p>
'@;

Write-Text -ParagraphStyle:'Header2' 'Sample code';

Write-Html -ParagraphStyle:'Text' @'
<p>Passing command parameters to async tasks:
'@;

@(
	[PSCustomObject]@{ X = 5;  Y = 2 },
	[PSCustomObject]@{ X = 10; Y = 3 },
	[PSCustomObject]@{ X = 15; Y = 4 }
) | Invoke-AsyncCommands -ScriptBlock:{
    param ($X, $Y);
	Write-Html "<b>X</b> * <b>Y</b> = $($X * $Y)";
	
	$t = [Threading.Thread]::CurrentThread;
	Write-Host "Thread ID: $($t.ManagedThreadId)";
} | Out-Null;

Write-Html -ParagraphStyle:'Text' @'
<p>Demonstrating tasks with different computation time:
'@;

1..10 |
    Select-Object -Property:@{ N = 'ID'; E = { $_ }} |
    Invoke-AsyncCommands -ThrottleLimit:10 {
        param ([int]$ID);
        [int]$interval = 10 - $ID;
        Start-Sleep -s:$interval;
        Write-Host "Task $ID completed.";
	} | Out-Null;

Write-Host 'Script completed.';

Write-Html -ParagraphStyle:'Description' @'
<br>
<i>Source code</i> of the script see in <i>ReadMe.ps1</i> .
'@;

Add-BookPageBreak;

Write-Html -ParagraphStyle:'Text' @'
<p align=justify>Cmdlets for async commands.</p>
'@;

Write-Html -ParagraphStyle:'Text' @'
<ul>
<li><h4>Invoke-AsyncCommands</h4>
<p align=justify>Invokes multiple commands in parallel runspaces.
Commands are executing in separate runspaces and local variables
cannot be used. Commands take parameters from pipeline, only
simple types (strings, numbers, dates) are allowed. Property
<i>CommonParameters</i> allows to pass common parameters
for each command, objects of any type are allowed. By default
cmdlet uses <i>RunspacePool</i> with count of <i>Runspaces</i>
equal to number of CPU cores but no more than 16, manually
<i>ThrottleLimit</i> (count of runspaces in runpool) can be set
in range between 1 and 256.</p></li>

<li><h4>New-SCRunspace</h4>
<p align=justify>Creates and initializes for <b>SpreadCommander</b>
PowerShell <i>Runspace</i>.</p></li>

<li><h4>New-SCRunspacePool</h4>
<p align=justify>Creates and initializes for <b>SpreadCommander</b>
PowerShell <i>RunspacePool<i>.</p></li>
'@;


Add-BookSection -ContinuePageNumbering -LinkHeaderToPrevious -LinkFooterToPrevious;

Write-Text -ParagraphStyle:'Header2' 'Book cmdlets';

. $schost.MapPath('~#\..\Common\CmdletHelp.ps1');

$cmdlets = [string[]]@(
	'Invoke-AsyncCommands',
	'New-SCRunspace',
	'New-SCRunspacePool'
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

Save-Book '~#\ReadMe.docx' -Replace;