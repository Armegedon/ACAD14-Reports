<?xml version="1.0" encoding="UTF-8"?>
<!-- (C) Copyright 2001 by Autodesk, Inc.  All rights reserved -->
<xsl:stylesheet 	xmlns:lx="http://www.landxml.org/schema/LandXML-1.2"
				xmlns:xsl="http://www.w3.org/1999/XSL/Transform" 
				xmlns:msxsl="urn:schemas-microsoft-com:xslt" 
				xmlns:landUtils="http://www.autodesk.com/land/civil/vcedit" 
				version="1.0">

<msxsl:script language="JScript" implements-prefix="landUtils"> 
<![CDATA[

function FormatCase(textStr, methodStr)
{
	var retStr = textStr;
	
	if(methodStr == "Uppercase")
	{
		retStr = textStr.toUpperCase();
	}
	else if(methodStr == "Lowercase")
	{
		retStr = textStr.toLowerCase();
	}
	else if(methodStr == "Sentence Case")
	{
		var sentStr = "";
		var i;
		for(i = 0; i< textStr.length; i++)
		{
			var c = textStr.charAt(i);
			if(i == 0)
			{
				var u = c.toUpperCase();
				sentStr = sentStr + u;
			}
			else
			{
				var u = c.toLowerCase();
				sentStr = sentStr + u;
			}
		}
		retStr = sentStr;
	}
	else if(methodStr == "Title Case")
	{
		var sentStr = "";
		var i;
		for(i = 0; i< textStr.length; i++)
		{
			if(i == 0)
			{
				var c = textStr.charAt(i);
				var u = c.toUpperCase();
				sentStr = sentStr + u;
			}
			else
			{
				var prevChar =  textStr.charAt(i - 1);
				if(prevChar == " ")
				{
					var c = textStr.charAt(i);
					var u = c.toUpperCase();
					sentStr = sentStr + u;
				}
				else
				{
					var c = textStr.charAt(i);
					var u = c.toLowerCase();
					sentStr = sentStr + u;
				}
			}
		}
		retStr = sentStr;
	}
	return retStr;
}


]]>
</msxsl:script>

</xsl:stylesheet>
