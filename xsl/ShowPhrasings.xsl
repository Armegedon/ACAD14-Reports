<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform" xmlns:fo="http://www.w3.org/1999/XSL/Format">

<xsl:template match="LegalDescription">
	<h1><xsl:value-of select="@desc"/></h1>
	<h2>Legal Description Phrasings templates  </h2>
	<xsl:for-each select="./Style">
		<xsl:apply-templates select="."/>
	</xsl:for-each>
	<hr/>
</xsl:template>

<xsl:template match="Style">
	<h3>Style: <xsl:value-of select="@name"/></h3>
	<xsl:for-each select="./Geometry">
		<xsl:apply-templates select="."/>
	</xsl:for-each>
</xsl:template>

<xsl:template match="Geometry">
	<h4>Phrasing templates for: <xsl:value-of select="@type"/></h4>
	<table width="100%">
			<xsl:apply-templates select="./Bounds"/>
			<xsl:apply-templates select="Metes"/>
	</table>
	<p/>
</xsl:template>

<xsl:template match="Bounds">
			<tr>
				<th colspan="2" align="left">Bounds phrases</th>
			</tr>
			<tr>
				<th align="left" width="25%">Type</th>
				<th align="left">Phrase</th>
			</tr>
	<xsl:for-each select="./Bound">
	<tr>
		<td width="25%" valign="top"><xsl:value-of select="@type"/></td>
		<td valign="top"><xsl:value-of select="."/></td>
	</tr>
	</xsl:for-each>
</xsl:template>

<xsl:template match="Metes">
			<tr>
				<td height="10"></td>
			</tr>
			<tr>
				<th colspan="2" align="left">Metes phrases</th>
			</tr>
			<tr>
				<th align="left" width="25%">Type</th>
				<th align="left">Phrase</th>
			</tr>
	<xsl:for-each select="Mete">
	<tr>
		<td width="25%" valign="top"><xsl:value-of select="@type"/> to <xsl:value-of select="@connect"/></td>
		<td valign="top"><xsl:value-of select="."/></td>
	</tr>
	</xsl:for-each>
</xsl:template>

</xsl:stylesheet>
