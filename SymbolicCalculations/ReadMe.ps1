using namespace MathNet.Symbolics;
using namespace System.Collections.Generic;

$schost           = Get-SCHost;
$schost.Silent = $true;

Clear-Book;

Invoke-SCScript '~\..\Common\InitBookStyles.ps1';

Set-BookSectionHeader '<b>Spread Commander</b> - <i>Examples: Symbolic Calculations</i>' -Html;
Set-BookSectionFooter 'Page {PAGE} of {NUMPAGES}' -ExpandFields;

Write-Text -ParagraphStyle:'Header1' 'Symbolic mathematics';

Write-Html -ParagraphStyle:'Description' @'
<p align=justify><b>SpreadCommander</b> includes <b>Math.Net.Numerics</b>
and <b>Math.Net.Symbolics</b> libraries and allows to perform
symbolic mathematic calculations.</p>
'@;

Write-Html -ParagraphStyle:'Text' @'
<p align=justify>Script below shows host to perform symbolic
calculations in <b>SpreadCommander</b>.</p>
'@;

$x = [SymbolicExpression]::Variable('x');
$y = [SymbolicExpression]::Variable('y');
$a = [SymbolicExpression]::Variable('a');
$b = [SymbolicExpression]::Variable('b');
$c = [SymbolicExpression]::Variable('c');
$d = [SymbolicExpression]::Variable('d');

Write-Text -ParagraphStyle:'Header4' 'Simple operations:';

$a + $a | Write-Text;
$a * $a | Write-Text;
2 + 1 / $x - 1 | Write-Text;
($a / $b / ($c * $a)) * ($c * $d / $a) / $d | Write-Text;

Write-Text -ParagraphStyle:'Header4' 'Output to different formats:';

(1 / ($a * $b)).ToString();
(1 / ($a * $b)).ToLaTeX();
(1 / ($a * $b)).ToLaTeX() | Write-Latex -FontSize:36;

Write-Text -ParagraphStyle:'Header4' 'Parsing:';

[SymbolicExpression]::Parse("1/(a*b)") | Write-Text; 
[SymbolicExpression]::Parse("1/(a*b)").ToLaTeX() | Write-LaTex -FontSize:24; 

Write-Text -ParagraphStyle:'Header4' 'Evaluating functions:';

$symbols = [Dictionary[string, FloatingPoint]]::new();
$symbols['a'] = 2.0;
$symbols['b'] = 3.0;

(1 / ($a * $b)).Evaluate($symbols).RealValue;

[Func[double, double, double]]$f = (1 / ($a * $b)).Compile('a', 'b');
$f.Invoke(2.0, 3.0) | Write-Text;

Write-Text -ParagraphStyle:'Header4' 'Output to LaTeX:';

$x.Cos().Pow(4).TrigonometricContract().ToLaTeX() | Write-LaTeX -FontSize:24;

Write-Text -ParagraphStyle:'Header4' 'Taylor expansion:';

function Taylor ([int]$k, 
	[SymbolicExpression]$symbol,
	[SymbolicExpression]$al, 
	[SymbolicExpression]$xl)
{
	[int]$factorial = 1;
	[SymbolicExpression]$accumulator = [SymbolicExpression]::Zero;
	[SymbolicExpression]$derivative  = $xl;
	
	0..($k-1) |
	%{
		[SymbolicExpression]$subs         = $derivative.Substitute($symbol, $al);
		[SymbolicExpression]$derivative   = $derivative.Differentiate($symbol);
		[SymbolicExpression]$accumulator += $subs / $factorial * (($symbol - $al).Pow($_));
		$factorial *= ($_ + 1);
	};
	
	return $accumulator.Expand();
};

$x    = [SymbolicExpression]::Variable('x');
$p1   = $x.Sin() + $x.Cos();
$zero = [SymbolicExpression]::Zero;

#Returns string "1 + x - x^2/2 - x^3/6"
$taylor = Taylor 4 $x $zero $p1;
$talyor | Write-Text;
$taylor.ToLaTeX() | Write-LaTeX -FontSize:24;


Add-BookSection -ContinuePageNumbering -LinkHeaderToPrevious -LinkFooterToPrevious;
Write-Text -ParagraphStyle:'Header2' 'Table of Contents';
Add-BookTOC;

Save-Book '~\ReadMe.docx' -Replace;