package xlsx

import (
	"fmt"
	"regexp"
	"text/template"
	"time"
)

// Templates for the various XML content in XLST files
var (
	TemplateContentTypes          *template.Template
	TemplateRelationships         *template.Template
	TemplateWorkbook              *template.Template
	TemplateWorkbookRelationships *template.Template
	TemplateStyles                *template.Template
	TemplateStringLookups         *template.Template
	TemplateSheetStart            *template.Template
	TemplateApp                   *template.Template
	TemplateCore                  *template.Template
)

// Template function for integer addition. This is useful to convert between
// zero-based and one-based array offsets within templates
func plus(i int, n int) string {
	return fmt.Sprintf("%d", i+n)
}

// Template function for time formatting
func timeFormat(t time.Time) string {
	return t.Format(time.RFC3339)
}

func init() {
	re := regexp.MustCompile("\n[\t\n\f\r ]*")
	funcMap := template.FuncMap{"plus": plus, "timeFormat": timeFormat}

	TemplateContentTypes = template.Must(template.New("templateContentTypes").Funcs(funcMap).Parse(re.ReplaceAllLiteralString(templateContentTypes, "")))
	TemplateRelationships = template.Must(template.New("templateRelationships").Funcs(funcMap).Parse(re.ReplaceAllLiteralString(templateRelationships, "")))
	TemplateWorkbook = template.Must(template.New("templateWorkbook").Funcs(funcMap).Parse(re.ReplaceAllLiteralString(templateWorkbook, "")))
	TemplateWorkbookRelationships = template.Must(template.New("templateWorkbookRelationships").Funcs(funcMap).Parse(re.ReplaceAllLiteralString(templateWorkbookRelationships, "")))
	TemplateStyles = template.Must(template.New("templateStyles").Funcs(funcMap).Parse(re.ReplaceAllLiteralString(templateStyles, "")))
	TemplateStringLookups = template.Must(template.New("templateStringLookups").Funcs(funcMap).Parse(re.ReplaceAllLiteralString(templateStringLookups, "")))
	TemplateSheetStart = template.Must(template.New("templateSheetStart").Funcs(funcMap).Parse(re.ReplaceAllLiteralString(templateSheetStart, "")))
	TemplateApp = template.Must(template.New("templateApp").Funcs(funcMap).Parse(re.ReplaceAllLiteralString(templateApp, "")))
	TemplateCore = template.Must(template.New("templateCore").Funcs(funcMap).Parse(re.ReplaceAllLiteralString(templateCore, "")))
}

const templateContentTypes = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
  <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
      <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
      <Default Extension="xml" ContentType="application/xml"/>
      <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
      <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
      <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
      <Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>
      <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
      <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
  </Types>`

const templateRelationships = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
  <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
      <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
      <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
      <Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
  </Relationships>`

const templateWorkbook = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
  <workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
      <fileVersion appName="xl" lastEdited="5" lowestEdited="5" rupBuild="9303"/>
      <workbookPr defaultThemeVersion="124226"/>
      <bookViews>
          <workbookView xWindow="480" yWindow="60" windowWidth="18195" windowHeight="8505"/>
      </bookViews>
      <sheets>
          <sheet name="{{.Title}}" sheetId="1" r:id="rId1"/>
      </sheets>
      <calcPr calcId="145621"/>
  </workbook>`

const templateWorkbookRelationships = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
  <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
      <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
      <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>
      <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
  </Relationships>`

const templateStyles = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
  <styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x14ac" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac">
    <numFmts count="3">
      <numFmt numFmtId="43" formatCode="_-* #,##0.00_-;\-* #,##0.00_-;_-* &quot;-&quot;??_-;_-@_-"/>
      <numFmt numFmtId="164" formatCode="yyyy\-mm\-dd\ hh:mm"/>
      <numFmt numFmtId="165" formatCode="yyyy\-mm\-dd;@"/>
    </numFmts>
    <fonts count="2" x14ac:knownFonts="1">
      <font><sz val="11"/><color rgb="FF000000"/><name val="Calibri"/><family val="2"/><scheme val="minor"/></font>
      <font><sz val="11"/><color rgb="FF000000"/><name val="Arial Unicode MS"/></font>
    </fonts>
    <fills count="2">
      <fill>
        <patternFill patternType="none"/>
      </fill>
      <fill>
        <patternFill patternType="gray125"/>
      </fill>
    </fills>
    <borders count="1">
      <border>
        <left/>
        <right/>
        <top/>
        <bottom/>
        <diagonal/>
      </border>
    </borders>
    <cellStyleXfs count="1">
      <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
    </cellStyleXfs>
    <cellXfs count="3">
      <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
      <xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/>
      <xf numFmtId="164" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="0"/>
    </cellXfs>
    <cellStyles count="1">
      <cellStyle name="Normal" xfId="0" builtinId="0"/>
    </cellStyles>
    <dxfs count="0"/>
    <tableStyles count="0" defaultTableStyle="TableStyleMedium2" defaultPivotStyle="PivotStyleLight16"/>
    <extLst>
    </extLst>
  </styleSheet>`

const templateStringLookups = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="{{len .}}" uniqueCount="{{len .}}">
{{range .}}<si><t>{{.}}</t></si>{{end}}
</sst>`

const templateSheetStart = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
  <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x14ac" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac">
            <sheetViews>
        <sheetView workbookViewId="0"/>
      </sheetViews>
      <sheetFormatPr defaultRowHeight="15" x14ac:dyDescent="0.25"/>
        <cols>
          {{range $i, $e := .Cols}}
          <col min="{{plus $i 1}}" max="{{plus $i 1}}" width="{{$e.Width}}" customWidth="1" style="1"/>
          {{end}}
        </cols>
      <sheetData>`

const templateApp = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
  <Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
  <Application>None</Application>
  <DocSecurity>0</DocSecurity>
  <ScaleCrop>false</ScaleCrop>
  <HeadingPairs>
    <vt:vector size="2" baseType="variant">
      <vt:variant>
        <vt:lpstr>Worksheets</vt:lpstr>
      </vt:variant>
      <vt:variant>
        <vt:i4>1</vt:i4>
      </vt:variant>
    </vt:vector>
  </HeadingPairs>
  <TitlesOfParts>
    <vt:vector size="1" baseType="lpstr">
      <vt:lpstr>{{.Title}}</vt:lpstr>
    </vt:vector>
  </TitlesOfParts>
  <LinksUpToDate>false</LinksUpToDate>
  <SharedDoc>false</SharedDoc>
  <HyperlinksChanged>false</HyperlinksChanged>
</Properties>`

const templateCore = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
  <cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
    <dc:creator>{{.CreatedBy}}</dc:creator>
    <cp:lastModifiedBy>{{.ModifiedBy}}</cp:lastModifiedBy>
    <dcterms:created xsi:type="dcterms:W3CDTF">{{timeFormat .CreatedAt}}</dcterms:created>
    <dcterms:modified xsi:type="dcterms:W3CDTF">{{timeFormat .ModifiedAt}}</dcterms:modified>
  </cp:coreProperties>`
