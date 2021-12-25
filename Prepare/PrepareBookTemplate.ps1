$book = New-Book;
try
{
	Invoke-SCScript -Book:$book '~#\..\Common\InitBookStyles.ps1';
	Save-Book -Book:$book '~#\..\Common\Template.docx' -Replace;
}
finally
{
	$book.Dispose();
}