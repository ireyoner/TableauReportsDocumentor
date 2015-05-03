<?xml version="1.0" encoding="utf-8"?>
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
    xmlns:msxsl="urn:schemas-microsoft-com:xslt" exclude-result-prefixes="msxsl"
>
  <xsl:output method="xml" indent="yes" omit-xml-declaration="yes" />
  <!--<xsl:strip-space elements="*"/>-->

  <xsl:template match="/workbook">
    <report xsi:noNamespaceSchemaLocation="ImportValidator.xsd" 
      xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
      <title>
        <xsl:value-of select="local-name(.)"/>
      </title>
      <sections>
        <xsl:apply-templates select="datasources"/>
        <xsl:apply-templates select="worksheets"/>
      </sections>
    </report>
  </xsl:template>

  <xsl:template match="datasources" >
    <section visible="True">
      <title>
        <xsl:value-of select="local-name(.)"/>
      </title>
      <content/>
      <subsections>
        <xsl:apply-templates select="datasource"/>
      </subsections>
    </section>
  </xsl:template>

  <xsl:template match="datasource" >
    <subsection visible="True">
      <title>
        <xsl:value-of select="@name"/>
      </title>
      <content>
        <text visible="True">Caption: <xsl:value-of select="@caption"/></text>
        <xsl:apply-templates select="connection/metadata-records"/>
        <xsl:apply-templates select="." mode="measures_table"/>
        <xsl:apply-templates select="." mode="dimensions_table"/>
      </content>
    </subsection>
  </xsl:template>
 
  <xsl:template match="*" mode="measures_table">
    <table visible="True">
      <title>measures</title>
      <header>
        <hcell visible="True">name_tableau</hcell>
        <hcell visible="True">name</hcell>
        <hcell visible="True">formula</hcell>
      </header>
      <rows>
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
      </rows>
    </table>
  </xsl:template>
  
  <xsl:template match="*" mode="dimensions_table">
    <table visible="True">
      <title>dimensions</title>
      <header>
        <hcell visible="True">dimension</hcell>
      </header>
      <rows>
        <xsl:for-each select="column[@role='dimension']">
          <row>
            <cell>
              <xsl:value-of select="@name"/>
            </cell>
          </row> 
        </xsl:for-each>
      </rows>
    </table>
  </xsl:template>
  
  <xsl:template match="*" mode="calculations_table">
    <table visible="True">
      <title>calculations</title>
      <header>
        <hcell visible="True">name</hcell>
        <hcell visible="True">caption</hcell>
        <hcell visible="True">type</hcell>
        <hcell visible="True">role</hcell>
        <hcell visible="True">datatype</hcell>
        <hcell visible="True">formula</hcell>
      </header>
      <rows>
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
      </rows>
    </table>
  </xsl:template>

  <xsl:template match="metadata-records" >
    <table visible="True">
      <title>metadata-records</title>
      <header>
        <hcell visible="True">Table name</hcell>
        <hcell visible="True">Column name</hcell>
        <hcell visible="True">local-type</hcell>
        <hcell visible="True">aggregation</hcell>
        <hcell visible="True">contains-null</hcell>
      </header>
      <rows>
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
      </rows>
    </table>
  </xsl:template>

  <xsl:template match="*" mode="filter" >
    <table visible="True">
      <title>filter</title>
      <header>
        <hcell visible="True">column</hcell>
        <hcell visible="True">level</hcell>
        <hcell visible="True">member</hcell>
      </header>
      <rows>
      <xsl:for-each select="filter">
          <row>
            <cell>
              <xsl:value-of select="@column"/>
            </cell>
            <cell>
              <xsl:value-of select="groupfilter/@level"/>
            </cell>
            <cell>
              <xsl:value-of select="groupfilter/@member"/>
            </cell>
          </row>
      </xsl:for-each>
      </rows>
    </table>
  </xsl:template>

  <xsl:template match="worksheets" >
    <xsl:apply-templates select="worksheet"/>
  </xsl:template>

  <xsl:template match="worksheet" >
    <section visible="True">
      <title>
        <xsl:value-of select="@name"/>
      </title>
      <content>
        <text visible="True">Title: <xsl:apply-templates select="layout-options/title"/></text>
        <xsl:apply-templates select="table/view/datasource-dependencies" mode="calculations_table"/>
        <xsl:apply-templates select="table/view/datasource-dependencies" mode="measures_table"/>
        <xsl:apply-templates select="table/view/datasource-dependencies" mode="dimensions_table"/>
        <xsl:apply-templates select="table/style/style-rule[@element='quick-filter']"/>
        <xsl:apply-templates select="table/view" mode="filter"/>
      </content>
      <subsections/>
    </section>
  </xsl:template>

  <xsl:template match="style-rule[@element='quick-filter']">
    <table visible="True">
      <title>quick-filter</title>
      <header>
        <hcell visible="True">Filtr</hcell>
      </header>
      <rows>
        <xsl:for-each select="format">
          <row>
            <cell>
              <xsl:value-of select="@value"/>
            </cell>
          </row>
        </xsl:for-each>     
      </rows>
    </table>
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
