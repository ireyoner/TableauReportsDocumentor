<?xml version="1.0" encoding="utf-8"?>
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
    xmlns:msxsl="urn:schemas-microsoft-com:xslt" exclude-result-prefixes="msxsl"
>
  <xsl:output method="xml" indent="yes" omit-xml-declaration="yes" />

  <xsl:template match="/workbook">
    <report xsi:noNamespaceSchemaLocation="ImportValidator.xsd" 
      xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
      <xsl:apply-templates select="datasources"/>
      <xsl:apply-templates select="worksheets"/>
    </report>
  </xsl:template>

  <xsl:template match="datasources" >
    <section>
      <title>
        <xsl:value-of select="local-name(.)"/>
      </title>
      <xsl:apply-templates select="datasource"/>
    </section>
  </xsl:template>

  <xsl:template match="datasource" >
    <subsection>
      <title>
        <xsl:value-of select="@name"/>
      </title>
      <text>Caption: <xsl:value-of select="@caption"/></text>
      <table>
        <title>metadata-records</title>
        <xsl:apply-templates select="connection/metadata-records"/>
      </table>
      <xsl:apply-templates select="." mode="measures_table"/>
      <xsl:apply-templates select="." mode="dimensions_table"/>
    </subsection>
  </xsl:template>
 
  <xsl:template match="*" mode="measures_table">
    <table>
      <title>measures</title>
      <header>
        <cell>name</cell>
        <cell>caption</cell>
        <cell>formula</cell>
      </header>
      <xsl:for-each select="column[@role='measure']">
        <row>
          <cell>
            <xsl:value-of select="@name"/>
          </cell>
          <cell>
            <xsl:value-of select="@caption"/>
          </cell>
          <cell>
            <xsl:value-of select="calculation/@formula"/>
          </cell>
        </row> 
      </xsl:for-each>
    </table>
  </xsl:template>
  
  <xsl:template match="*" mode="dimensions_table">
    <table>
      <title>dimensions</title>
      <header>
        <cell>dimension</cell>
      </header>
      <xsl:for-each select="column[@role='dimension']">
        <row>
          <cell>
            <xsl:value-of select="@name"/>
          </cell>
        </row> 
      </xsl:for-each>
    </table>
  </xsl:template>
  
  <xsl:template match="*" mode="calculations_table">
    <table>
      <title>calculations</title>
      <header>
        <cell>name</cell>
        <cell>caption</cell>
        <cell>type</cell>
        <cell>role</cell>
        <cell>datatype</cell>
        <cell>formula</cell>
      </header>
      <xsl:for-each select="column[starts-with(@name, '[Calculation_')]">
        <row>
          <cell>
            <xsl:value-of select="@name"/>
          </cell>
          <cell>
            <xsl:value-of select="@caption"/>
          </cell>
          <cell>
            <xsl:value-of select="@type"/>
          </cell>
          <cell>
            <xsl:value-of select="@role"/>
          </cell>
          <cell>
            <xsl:value-of select="@datatype"/>
          </cell>
          <cell>
            <xsl:value-of select="calculation/@formula"/>
          </cell>
        </row> 
      </xsl:for-each>
    </table>
  </xsl:template>

  <xsl:template match="metadata-records" >
    <header>
      <cell>Table name</cell>
      <cell>Column name</cell>
      <cell>local-type</cell>
      <cell>aggregation</cell>
      <cell>contains-null</cell>
    </header>
    <xsl:for-each select="metadata-record[@class='column' 
                           and not(parent-name = following::metadata-record[@class='column']/parent-name)]">
      <xsl:sort select=".//parent-name"/>
      <xsl:variable name="tableName" select=".//parent-name"/>
      <xsl:for-each select="../metadata-record[@class='column' and parent-name/text() = $tableName]">
        <xsl:sort select="local-name"/>
        <row>
          <cell>
            <xsl:value-of select="$tableName"/>
          </cell>
          <cell>
            <xsl:value-of select="local-name/text()"/>
          </cell>
          <cell>
            <xsl:value-of select="local-type/text()"/>
          </cell>
          <cell>
            <xsl:value-of select="aggregation/text()"/>
          </cell>
          <cell>
            <xsl:value-of select="contains-null/text()"/>
          </cell>
        </row>
      </xsl:for-each>
    </xsl:for-each>
  </xsl:template>

  <xsl:template match="metadata-record" >
    <xsl:apply-templates select="local-name/text()"/>
  </xsl:template>


  <xsl:template match="worksheets" >
    <xsl:apply-templates select="worksheet"/>
  </xsl:template>

  <xsl:template match="worksheet" >
    <section>
      <title>
        <xsl:value-of select="@name"/>
      </title>
      <text>Title: <xsl:apply-templates select="layout-options/title"/>
      </text>
      <xsl:apply-templates select="table/view/datasource-dependencies" mode="calculations_table"/>
      <xsl:apply-templates select="table/view/datasource-dependencies" mode="measures_table"/>
      <xsl:apply-templates select="table/view/datasource-dependencies" mode="dimensions_table"/>
      <table>
        <title>quick-filter</title>
        <header>
          <cell>Filtr</cell>
        </header>
        <xsl:apply-templates select="table/style/style-rule[@element='quick-filter']/format"/>
      </table>
    </section>
  </xsl:template>

  <xsl:template match="format">
    <row>
      <cell>
        <xsl:value-of select="@value"/>
      </cell>
    </row>
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
