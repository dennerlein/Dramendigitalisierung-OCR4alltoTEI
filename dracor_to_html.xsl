<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
    xmlns:tei="http://www.tei-c.org/ns/1.0" exclude-result-prefixes="tei" version="3.0">
    
    <!-- Definiere die HTML-Ausgabe -->
    <xsl:output method="html" doctype-system="about:legacy-compat" encoding="UTF-8" indent="yes"/>
    
    <!-- Startseite mit HTML-Struktur -->
    <xsl:template match="/">
        <html>
            <head>
                <title>
                    <xsl:value-of
                        select="tei:TEI/tei:teiHeader/tei:fileDesc/tei:titleStmt/tei:title"/>
                </title>
                <style>
                    body {
                    font-family: Garamond, sans-serif;
                    line-height: 1.6;
                    padding: 20px;
                    background-color: #f4f4f4;
                    max-width: 500px;
                    justify-content: center;
                    }
                    h1,
                    h2,
                    h3,
                    h4 {
                    color: #333;
                    }
                    .scene-title {
                    font-weight: bold;
                    font-size: 1.2em;
                    color: #000;
                    margin-top: 20px;
                    }
                    .speech {
                    margin-bottom: 15px;
                    margin-top: 20px;
                    }
                    .speaker {
                    font-weight: bold;
                    }
                    .text {
                    margin-left: 20px;
                    } /* Hier wird der Sprechtext leicht eingerückt */
                    .stage-direction {
                    font-style: italic;
                    color: #555;
                    margin-left: 20px;
                    }
                    .roleDesc {
                    font-style: italic;
                    color: #555;
                    margin-left: 20px;
                    }
                    .role {
                    margin-bottom: 20px;
                    }
                    .cast-list h2 {
                    font-size: 1.4em;
                    font-weight: bold;
                    margin-bottom: 10px;
                    }
                    .cast-list ul {
                    list-style-type: none;
                    padding-left: 0;
                    }
                    </style>
            </head>
            <body>
                <!-- Titel anzeigen -->
                <h1>
                    <xsl:value-of
                        select="tei:TEI/tei:teiHeader/tei:fileDesc/tei:titleStmt/tei:title"/>
                </h1>
                
                <!-- Preface (Vorwort) -->
                <xsl:apply-templates select="tei:TEI/tei:text/tei:front"/>
                
                <!-- Figurenliste anzeigen -->
                <!--<xsl:apply-templates select="tei:TEI/tei:text/tei:front/tei:castList"/>-->
                
                <!-- Haupttext (Schauspiel) -->
                <xsl:apply-templates select="tei:TEI/tei:text/tei:body"/>
            </body>
        </html>
    </xsl:template>
    
    <!-- Figurenliste -->
    <xsl:template match="tei:castList">
        <div class="cast-list">
            <h2>
                <xsl:value-of select="tei:head"/>
            </h2>
            <ul>
                <xsl:for-each select="tei:castItem">
                    <li>
                        <xsl:value-of select="tei:role"/>
                        <xsl:text> </xsl:text>
                        <span style="font-style: italic; color: #555;">
                            <xsl:value-of select="tei:roleDesc"/>
                        </span>
                    </li>
                </xsl:for-each>
            </ul>
            
        </div>
    </xsl:template>
    
    <!-- Akte (acts) -->
    <xsl:template match="tei:div[@type = 'act']">
        <div class="act">
            <h2>
                <xsl:value-of select="tei:head" disable-output-escaping="yes"/>
            </h2>
            <xsl:apply-templates select="node()[not(self::tei:head)]"/>
        </div>
    </xsl:template>
    
    <!-- Szenen (scenes) -->
    <xsl:template match="tei:div[@type = 'scene']">
        <div class="scene">
            <p class="scene-title">
                <xsl:value-of select="tei:head" disable-output-escaping="yes"/>
            </p>
            <xsl:apply-templates select="node()[not(self::tei:head)]"/>
        </div>
    </xsl:template>
    
    <!-- Sprecher (speakers) -->
    <xsl:template match="tei:sp">
        <div class="speech">
            <!-- Sprecher -->
            <div class="speaker">
                <xsl:value-of select="tei:speaker"/>
            </div>
            <!-- Process alternating <p> and <stage> elements in order -->
            <div class="text">
                <xsl:apply-templates select="tei:p | tei:stage"/>
            </div>
        </div>
    </xsl:template>
    
    <!-- Template to handle <p> elements -->
    <xsl:template match="tei:p">
        <p>
            <xsl:value-of select="."/>
        </p>
    </xsl:template>
    
    <!-- Template to handle <l> elements -->
    <xsl:template match="tei:l">
            <xsl:value-of select="."/>
        
    </xsl:template>
    
    <!-- Template to handle <stage> elements -->
    <xsl:template match="tei:stage">
        <stage>
            <xsl:value-of select="."/>
        </stage>
    </xsl:template>
    
    <!-- Regieanweisungen (stage directions) -->
    <xsl:template match="tei:stage">
        <div class="stage-direction">
            <xsl:value-of select="."/>
        </div>
    </xsl:template>
    
    <!-- Allgemeine Absätze -->
    <xsl:template match="tei:p">
        <p>
            <xsl:apply-templates/>
        </p>
    </xsl:template>
    
    <xsl:template match="tei:l">
            <xsl:apply-templates/>
        <br/>
    </xsl:template>
    
    <!-- Text-Ausgabe, ohne wiederholtes Ausgeben von Text -->
    <xsl:template match="text()">
        <xsl:value-of select="normalize-space(.)"/>
    </xsl:template>
    
</xsl:stylesheet>
