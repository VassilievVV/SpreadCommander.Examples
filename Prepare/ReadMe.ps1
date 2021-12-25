$schost = Get-SCHost;
$schost.Silent = $true;

Clear-Book;

Set-BookSectionHeader '<b>Spread Commander</b> - <i>Examples: Prepare</i>' -Html;
Set-BookSectionFooter '<b>Spread Commander</b> - <i>Examples: Prepare</i>' -Html;

Invoke-SCScript '~#\..\Common\InitBookStyles.ps1';

Write-Text -ParagraphStyle:'Header1' 'PREPARE';

Write-Text -ParagraphStyle:'Description' @'
This project prepares data and templates for other example projects.
'@;

Save-Book '~#\ReadMe.docx' -Replace;
