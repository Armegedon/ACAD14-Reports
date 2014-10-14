<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet 	xmlns:lx="http://www.landxml.org/schema/LandXML-1.2"
				xmlns:xsl="http://www.w3.org/1999/XSL/Transform" 
                		xmlns:msxsl="urn:schemas-microsoft-com:xslt"
                 		xmlns:landUtils="http://www.autodesk.com/land/civil/vcedit"
                		xmlns:lxml="urn:lx_utils"
                		version="1.0">
                		
<xsl:include href="Text_Formatting.xsl"/>
                		
<!-- Owner information parameters -->
<xsl:param name="Owner.Preparer.name"/>
<xsl:param name="Owner.Preparer.capitalization"/>
<xsl:param name="Owner.Company.name"/>
<xsl:param name="Owner.Company.capitalization"/>
<xsl:param name="Owner.Address1.name"/>
<xsl:param name="Owner.Address1.capitalization"/>
<xsl:param name="Owner.Address2.name"/>
<xsl:param name="Owner.Address2.capitalization"/>
<xsl:param name="Owner.City.name"/>
<xsl:param name="Owner.City.capitalization"/>
<xsl:param name="Owner.State.name"/>
<xsl:param name="Owner.State.capitalization"/>
<xsl:param name="Owner.Country.name"/>
<xsl:param name="Owner.Country.capitalization"/>
<xsl:param name="Owner.EMail.email"/>
<xsl:param name="Owner.URL.url"/>
<xsl:param name="Owner.Zip.number"/>
<xsl:param name="Owner.Phone.number"/>
<xsl:param name="Owner.Fax.number"/>

<!-- Client information parameters -->
<xsl:param name="Client.Contact.name"/>
<xsl:param name="Client.Contact.capitalization"/>
<xsl:param name="Client.Company.name"/>
<xsl:param name="Client.Company.capitalization"/>
<xsl:param name="Client.Address1.name"/>
<xsl:param name="Client.Address1.capitalization"/>
<xsl:param name="Client.Address2.name"/>
<xsl:param name="Client.Address2.capitalization"/>
<xsl:param name="Client.City.name"/>
<xsl:param name="Client.City.capitalization"/>
<xsl:param name="Client.State.name"/>
<xsl:param name="Client.State.capitalization"/>
<xsl:param name="Client.Country.name"/>
<xsl:param name="Client.Country.capitalization"/>
<xsl:param name="Client.EMail.email"/>
<xsl:param name="Client.URL.url"/>
<xsl:param name="Client.Zip.number"/>
<xsl:param name="Client.Phone.number"/>
<xsl:param name="Client.Fax.number"/>

<!-- This Report information -->
<xsl:param name="DateTime"/>

<!-- Auto Header example-->
<xsl:template name="AutoHeader">
	<xsl:param name="ReportTitle"/>
	<xsl:param name="ReportDesc"/>
	<center>
	<h1><xsl:value-of select="landUtils:FormatCase($Owner.Company.name, $Owner.Company.capitalization)"/></h1>
	<h2><xsl:value-of select="landUtils:FormatCase($Owner.Address1.name, $Owner.Address1.capitalization)"/></h2>
	<h2><xsl:value-of select="landUtils:FormatCase($Owner.Address2.name, $Owner.Address2.capitalization)"/></h2>
	<h2><xsl:value-of select="landUtils:FormatCase($Owner.City.name, $Owner.City.capitalization)"/>, 
	<xsl:text disable-output-escaping="yes"> </xsl:text><xsl:value-of select="landUtils:FormatCase($Owner.State.name, $Owner.State.capitalization)"/>
	<xsl:text disable-output-escaping="yes"> </xsl:text><xsl:value-of select="$Owner.Zip.number"/></h2>
	</center>
	<hr/>
	<table border="0" width="100%">
		<tr>
			<td><b><xsl:value-of select="$ReportTitle"/> </b> </td>
			<td align="right"><b>Client: </b><xsl:value-of select="landUtils:FormatCase($Client.Company.name, $Client.Company.capitalization)"/></td>  
		</tr>
		<tr>
			<td><b>Project Name: </b> <xsl:value-of select="$ReportDesc"/></td>
			<td align="right"><b>Project Description: </b><xsl:value-of select="//lx:Project/@desc"/></td>
		</tr>
		<tr>
			<td><b>Report Date: </b><xsl:value-of select="$DateTime"/></td>
			<td align="right"><b>Prepared by: </b><xsl:value-of select="landUtils:FormatCase($Owner.Preparer.name, $Owner.Preparer.capitalization)"/></td>
		</tr>
		
	</table>
	<hr/>
</xsl:template>
</xsl:stylesheet>
