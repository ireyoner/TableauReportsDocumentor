<?xml version="1.0" encoding="utf-8"?>
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
    xmlns:msxsl="urn:schemas-microsoft-com:xslt" exclude-result-prefixes="msxsl"
>
  <xsl:output method="xml" indent="yes" omit-xml-declaration="yes" />

  <xsl:template match="/workbook">
    <report xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
	    xsi:noNamespaceSchemaLocation="C:\Users\hp\Downloads\dok\dok\trc.xsd">
      <xsl:apply-templates select="datasources"/>
      <xsl:apply-templates select="worksheets"/>
    </report>
  </xsl:template>

  <xsl:template match="datasources" >
    <reportDataSources>
      <xsl:apply-templates select="datasource[@caption]"/>
    </reportDataSources>
  </xsl:template>

  <xsl:template match="datasource" >
    <reportDataSource>
      <xsl:attribute name="order">
        <xsl:value-of select="count(preceding::datasource)+1"/>      
      </xsl:attribute>
      <xsl:attribute name="visible">true</xsl:attribute>
      <caption>
        <xsl:value-of select="@caption"/>
      </caption>
      <tables>
        <xsl:apply-templates select="connection/metadata-records"/>
      </tables>
      <otherAttrigutes>
        <xsl:apply-templates select="column[@caption]"/>
      </otherAttrigutes>
    </reportDataSource>
  </xsl:template>

  <xsl:template match="column[@caption]" >
    <column>
      <xsl:attribute name="order">
        <xsl:value-of select="count(preceding::column)+1"/>
      </xsl:attribute>
      <xsl:attribute name="visible">true</xsl:attribute>
      <caption>
        <xsl:value-of select="@caption"/>
      </caption>
      <value>
        <xsl:value-of select="calculation/@formula"/>
      </value>
      <xsl:apply-templates select="local-name/text()"/>
    </column>
  </xsl:template>

  <xsl:template match="metadata-records" >
    <xsl:for-each select="metadata-record[@class='column' 
                           and not(parent-name = following::metadata-record[@class='column']/parent-name)]">
      <xsl:sort select=".//parent-name"/>
      <xsl:variable name="tableName" select=".//parent-name"/>
      <table>
        <xsl:attribute name="order">
            <xsl:value-of select="count(preceding::metadata-records//parent-name)+1"/>      
        </xsl:attribute>
        <xsl:attribute name="visible">true</xsl:attribute>
        <caption>
          <xsl:value-of select="$tableName"/>
        </caption>
        <columns>
          <xsl:apply-templates select="../metadata-record[@class='column' and parent-name/text() = $tableName]">
            <xsl:sort select="local-name"/>
          </xsl:apply-templates>
        </columns>
      </table>
    </xsl:for-each>
  </xsl:template>

  <xsl:template match="metadata-record" >
    <column>
      <xsl:attribute name="order">
        <xsl:value-of select="count(preceding::metadata-record)+1"/>      
      </xsl:attribute>
      <xsl:attribute name="visible">true</xsl:attribute>
      <xsl:apply-templates select="local-name/text()"/>
    </column>
  </xsl:template>


  <xsl:template match="worksheets" >
    <reportTabs>
      <xsl:apply-templates select="worksheet"/>
    </reportTabs>
  </xsl:template>

  <xsl:template match="worksheet" >
    <reportTab>
      <xsl:attribute name="order">
        <xsl:value-of select="count(preceding::worksheet)+1"/>
      </xsl:attribute>
      <xsl:attribute name="visible">true</xsl:attribute>
      <reportTabName>
        <xsl:value-of select="@name"/>
      </reportTabName>
      <reportTabTitle>
        <xsl:apply-templates select="layout-options/title"/>
      </reportTabTitle>
      <data>
        <xsl:apply-templates select="table/view/datasource-dependencies/column[@caption]"/>
      </data>
    </reportTab>
  </xsl:template>

  <xsl:template match="title" >
    <xsl:apply-templates select="formatted-text/run"/>
  </xsl:template>
  
  <xsl:template match="run" >
    <xsl:value-of select="."/>
  </xsl:template>

  <xsl:template match="@* | node()" priority="-10">
    <xsl:copy>
      <xsl:apply-templates select="@* | node()"/>
    </xsl:copy>
  </xsl:template>
</xsl:stylesheet>
