<?xml version="1.0" encoding="utf-8"?>
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform" xmlns:st="http://www.structualtext.co.jp/2011/xml"
                exclude-result-prefixes="st">
  <xsl:output method="xml" indent="yes" doctype-public="-//W3C//DTD XHTML 1.1//EN"
               doctype-system="http://www.w3.org/TR/xhtml11/DTD/xhtml11.dtd"/>

  <xsl:template match="/">
    <html xmlns="http://www.w3.org/1999/xhtml" xml:lang="ja">
      <head>
        <title></title>
      </head>
      <body>
        <xsl:apply-templates select="//st:paragraph|//st:chapter[./@title!='']|//st:mathematics|//st:image|//comment()" />
      </body>
    </html>
  </xsl:template>

  <xsl:template match="/*/st:chapter">
    <!-- 以下に"/"レベル以下のchapterに対する処理を示します。 -->
    <h1 xmlns="http://www.w3.org/1999/xhtml">
      <xsl:value-of select="@title"/>
    </h1>
    <xsl:apply-templates select="./chapter" />
  </xsl:template>

  <xsl:template match="/*/*/st:chapter">
    <!-- 以下に"//"レベル以下のchapterに対する処理を示します。 -->
    <h2 xmlns="http://www.w3.org/1999/xhtml">
      <xsl:value-of select="@title"/>
    </h2>
    <xsl:apply-templates select="./chapter" />
  </xsl:template>

  <xsl:template match="/*/*/*/st:chapter">
    <!-- 以下に"///"レベル以下のchapterに対する処理を示します。 -->
    <h3 xmlns="http://www.w3.org/1999/xhtml">
      <xsl:value-of select="@title"/>
    </h3>
    <xsl:apply-templates select="./chapter" />
  </xsl:template>

  <xsl:template match="/*/*/*/*/st:chapter">
    <!-- 以下に"////"レベル以下のchapterに対する処理を示します。 -->
    <h4 xmlns="http://www.w3.org/1999/xhtml">
      <xsl:value-of select="@title"/>
    </h4>
    <xsl:apply-templates select="./chapter" />
  </xsl:template>

  <xsl:template match="/*/*/*/*/*/st:chapter">
    <!-- 以下に"/////"レベル以下のchapterに対する処理を示します。 -->
    <h5 xmlns="http://www.w3.org/1999/xhtml">
      <xsl:value-of select="@title"/>
    </h5>
    <xsl:apply-templates select="./chapter" />
  </xsl:template>

  <xsl:template match="/*/*/*/*/*/*//st:chapter">
    <!-- 以下に"//////"レベル以下のchapterに対する処理を示します。 -->
    <h6 xmlns="http://www.w3.org/1999/xhtml">
      <xsl:value-of select="@title"/>
    </h6>
    <xsl:apply-templates select="./chapter" />
  </xsl:template>

  <xsl:template match="st:paragraph">
    <!-- 以下に平文に対応するparagraphに対する処理を示します。-->
    <p xmlns="http://www.w3.org/1999/xhtml">
      <xsl:value-of select="./text()"/>
    </p>
  </xsl:template>

  <xsl:template match="st:mathematics">
    <!-- 以下に"#Math:"に対応するmathematicsに対する処理を示します。 -->
    <p xmlns="http://www.w3.org/1999/xhtml" class="math">
      <xsl:value-of select="./st:expression[@type='text']/text()"/>
    </p>
  </xsl:template>

  <xsl:template match="comment()">
    <!-- 以下に"#"に対応するコメントに対する処理を示します。 -->
    <xsl:comment>
      <xsl:value-of select="."/>
    </xsl:comment>
  </xsl:template>


  <xsl:template match="st:image">
    <!-- 以下に"#Image:"に対応するimageに対する処理を示します。-->
    <xsl:element name="img" namespace="http://www.w3.org/1999/xhtml">
      <xsl:attribute name="alt">image</xsl:attribute>
      <xsl:attribute name="src">
        <xsl:value-of select="./@src"/>
      </xsl:attribute>
    </xsl:element>
  </xsl:template>

</xsl:stylesheet>