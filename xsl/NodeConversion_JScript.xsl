<?xml version="1.0" encoding="UTF-8"?>
<!-- (C) Copyright 2001 by Autodesk, Inc.  All rights reserved -->
<xsl:stylesheet version="1.0"
	xmlns:lx="http://www.landxml.org/schema/LandXML-1.2"
	xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
	xmlns:msxsl="urn:schemas-microsoft-com:xslt"
	xmlns:landUtils="http://www.autodesk.com/land/civil/vcedit"
	xmlns:lxml="urn:lx_utils">

<msxsl:script language="JScript" implements-prefix="landUtils"> 
<![CDATA[
function NodeToText(node)
{
	var str;
	if(node.length > 0)
	{
		return node.item(0).text;
	}
	else
	{
		return "" + node;
	}
}
function GetNodeType(node)
{
	return node;
}

]]></msxsl:script>
</xsl:stylesheet>
