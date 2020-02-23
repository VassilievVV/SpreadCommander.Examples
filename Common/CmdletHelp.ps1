<#
	Generates HTML-formatted help for cmdlet.
#>
function ConvertCmdletHelpToHtml([string]$help)
{
	$help = $help.Replace('<', '&lt;').Replace('>', '&gt;').
		Replace([Environment]::NewLine, '<br>' + [Environment]::NewLine);

	$help = $help -replace '(?m)(^\w[\w \t]+)', '<b><u>$1</u></b>';
	$help = $help -replace '(?m)^(?:[ \t]*(\-\w+(?:\s*\[?&lt;.*?&gt;\]?)?))\s*<br>\s*$', '<b>$1</b><br>';
	$help = $help -replace '(?m)^[ \t]*(&lt;CommonParameters&gt;)\s*<br>\s*$', '<b>$1</b><br>';

	$help = $help.Replace(' ', '&nbsp;');

	return $help;
}

function GenerateCmdletHelp([string]$cmdlet)
{
	$help = Get-Help -Name:"$cmdlet" -Full | Out-String -Width:10000;
	if ($help -eq $null) { $help = [string]::Empty }
	
	$help = ConvertCmdletHelpToHtml($help);
	return $help;
}

function GetTypeDescription([Type]$type)
{
	if ($type.IsGenericType) 
	{
		if ($type.GetGenericTypeDefinition().Name -eq 'Nullable`1')
		{
			return "$($type.GetGenericArguments()[0].Name)?";
		}
		else 
		{
			return $type.FullName;
		}
	} else 
	{
		return $type.Name;
	}
}

function GenerateCmdletParametersHelp([string]$cmdlet, [string[]]$parameters)
{
	$help = (Get-Command -Name:$cmdlet -ArgumentList:$parameters).ParameterSets | 
	%{
		"Parameter Set Name: $($_.Name) $(if ($_.IsDefault) {'(Default)'} else {''})";

		$_.Parameters |
		%{
@"
    -$($_.Name) <$(GetTypeDescription($_.ParameterType))>
        $($_.HelpMessage)
        
        Required?                    $($_.IsMandatory)
        Position?                    $(if ($_.Position -ge 0) {$_.Position} else {'Named'})
        Accept pipeline input?       $($_.ValueFromPipeline -or $_.ValueFromPipelineByPropertyName)
        Aliases                      $(if ($_.Aliases.Length -ge 0) {[string]::Join(', ', $_.Aliases)} else {'None'})
        Dynamic?                     $($_.IsDynamic)

"@;
		}
		"";
	} | Out-String -Width:10000;
	
	$help = ConvertCmdletHelpToHtml($help);
	return $help;
}