<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet version="1.0" 
	xmlns:xsl="http://www.w3.org/1999/XSL/Transform" 
	xmlns:lx="http://www.landxml.org/schema/LandXML-1.2" 
	xmlns:msxsl="urn:schemas-microsoft-com:xslt"
	xmlns:landUtils="http://www.autodesk.com/land/civil/vcedit">

<msxsl:script language="JScript" implements-prefix="landUtils"><![CDATA[
function GetXmlHeader()
{
	var headStr;	
	headStr = "<?xml version=\"1.0\"?>";	
	return headStr;
}

function BeginWorkbook()
{
	var workStr;
	workStr = "<Workbook xmlns=\"urn:schemas-microsoft-com:office:spreadsheet\" xmlns:o=\"urn:schemas-microsoft-com:office:office\" " +
	"xmlns:x=\"urn:schemas-microsoft-com:office:excel\" xmlns:ss=\"urn:schemas-microsoft-com:office:spreadsheet\" " +
	"xmlns:html=\"http://www.w3.org/TR/REC-html40\"> ";
	return workStr;
}

function CloseWorkBook()
{
	return "</Workbook>"
}

function BeginWorkSheet(name)
{
	return "<Worksheet ss:Name=\"" + name + "\">            ";
}

function CloseWorkSheet()
{
	return "</Worksheet>"
}

function BeginTable(rowCount, colCount)
{
	var tabStr;
	//tabStr = "<Table ss:ExpandedColumnCount=\"" + colCount + "\" " +
	//"ss:ExpandedRowCount=\"" + rowCount + "\" x:FullColumns=\"1\"   x:FullRows=\"1\">";
	tabStr = "<Table>";
	
	return tabStr;
}

function CloseTable()
{
	return "</Table>";
}

function BeginRow()
{
	return "<Row>";
}

function CloseRow()
{
	return "</Row>";
}

function AddCell(type, contents, index)
{
	return "<Cell><Data ss:Type=\"" + type + "\">" + contents + "</Data></Cell>" ;
}

function AddFormula(formula)
{
	var cellStr = "<Cell ss:Formula=\"" + formula + "\">"
	var dataStr = "<Data ss:Type=\"Number\">0</Data>";
	return cellStr + dataStr + "</Cell>";
}
]]></msxsl:script>
</xsl:stylesheet>
