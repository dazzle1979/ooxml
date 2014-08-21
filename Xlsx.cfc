component xlsx output="false" {
	variables.xlsxID = createUUID();
	variables.tmpDirPath = expandPath('\tmp');
	variables.xlsxDirPath = variables.tmpDirPath&"\"&variables.xlsxID;
	variables.zipPath = variables.tmpDirPath & "\" & variables.xlsxID & ".zip";
	variables.coreXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><dc:creator>Bas van der Graaf</dc:creator><cp:lastModifiedBy>Bas van der Graaf</cp:lastModifiedBy><dcterms:created xsi:type="dcterms:W3CDTF">2014-07-31T09:08:26Z</dcterms:created><dcterms:modified xsi:type="dcterms:W3CDTF">2014-07-31T09:09:32Z</dcterms:modified></cp:coreProperties>';
	variables.appXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"><Application>Microsoft Excel</Application><DocSecurity>0</DocSecurity><ScaleCrop>false</ScaleCrop><HeadingPairs><vt:vector size="2" baseType="variant"><vt:variant><vt:lpstr>Werkbladen</vt:lpstr></vt:variant><vt:variant><vt:i4>1</vt:i4></vt:variant></vt:vector></HeadingPairs><TitlesOfParts><vt:vector size="1" baseType="lpstr"><vt:lpstr>Blad1</vt:lpstr></vt:vector></TitlesOfParts><Company>Yoobi</Company><LinksUpToDate>false</LinksUpToDate><SharedDoc>false</SharedDoc><HyperlinksChanged>false</HyperlinksChanged><AppVersion>14.0300</AppVersion></Properties>';
	variables.relsXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>';
	variables.contentTypesXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/><Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/></Types>';
	variables.workbookRelsXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>';
	variables.workbookXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets></sheets></workbook>';

	public xlsx function init(struct config) {
		variables.config = arguments.config
		//DEV code
		//directoryDelete(variables.tmpDirPath,true);
		//directories regelen
		if(!directoryExists(variables.tmpDirPath)) directoryCreate(variables.tmpDirPath);
		if(!directoryExists(variables.xlsxDirPath)) directoryCreate(variables.xlsxDirPath);
		if(!directoryExists(variables.xlsxDirPath&"\docProps")) directoryCreate(variables.xlsxDirPath&"\docProps");
		if(!directoryExists(variables.xlsxDirPath&"\_rels")) directoryCreate(variables.xlsxDirPath&"\_rels");
		if(!directoryExists(variables.xlsxDirPath&"\xl")) directoryCreate(variables.xlsxDirPath&"\xl");
		if(!directoryExists(variables.xlsxDirPath&"\xl\_rels")) directoryCreate(variables.xlsxDirPath&"\xl\_rels");
		if(!directoryExists(variables.xlsxDirPath&"\xl\worksheets")) directoryCreate(variables.xlsxDirPath&"\xl\worksheets");

		this.createCore();
		this.createApp();
		this.createContentTypes();
		this.createStyles();
		this.createWorkbookRels();
		this.createRels();
		this.createWorkbook();

		return this;
	}

	public void function createSheet(string data, numeric sheetnumber) {
		arguments.data = ' <row> <c r="a1" s="1" t="inlineStr"><is><t>Project</t></is></c> <c r="b1" s="1" t="inlineStr"><is><t>Activiteit</t></is></c> <c r="c1" s="1" t="inlineStr"><is><t>Medewerker</t></is></c> <c r="d1" s="1" t="inlineStr"><is><t>Periode</t></is></c> <c r="e1" s="1" t="inlineStr"><is><t>Uren</t></is></c> <c r="f1" s="1" t="inlineStr"><is><t>Kosten</t></is></c> <c r="g1" s="1" t="inlineStr"><is><t>Aantal</t></is></c> <c r="h1" s="1" t="inlineStr"><is><t>Dag</t></is></c> <c r="i1" s="1" t="inlineStr"><is><t>Opmerkingen</t></is></c> </row> <row> <c r="a2" s="3" t="inlineStr"><is><t>AAH Alligator (11631)</t></is></c> <c r="b2" s="3" t="inlineStr"><is><t>Consultancy</t></is></c> <c r="c2" s="3" t="inlineStr"><is><t>Graaf, Bas van der</t></is></c> <c r="d2" s="3" t="d"><v>2014-08-04T00:00:00.000</v></c> <c r="e2" s="4" t="n"><v>1</v></c> <c r="f2" s="4" t="n"><v>0</v></c> <c r="g2" s="4" t="n"><v>0</v></c> <c r="h2" s="4" t="n"><v>0</v></c> <c r="i2" s="3" t="inlineStr"><is><t></t></is></c> </row> <row> <c r="a3" s="6" t="inlineStr"><is><t>AAH Alligator (11631)</t></is></c> <c r="b3" s="6" t="inlineStr"><is><t>Consultancy</t></is></c> <c r="c3" s="6" t="inlineStr"><is><t>Graaf, Bas van der</t></is></c> <c r="d3" s="6" t="d"><v>2014-08-08T00:00:00.000</v></c> <c r="e3" s="7" t="n"><v>1</v></c> <c r="f3" s="7" t="n"><v>0</v></c> <c r="g3" s="7" t="n"><v>0</v></c> <c r="h3" s="7" t="n"><v>0</v></c> <c r="i3" s="6" t="inlineStr"><is><t></t></is></c> </row> <row> <c r="a4" s="3" t="inlineStr"><is><t>Dell</t></is></c> <c r="b4" s="3" t="inlineStr"><is><t>Maatwerk</t></is></c> <c r="c4" s="3" t="inlineStr"><is><t>Graaf, Bas van der</t></is></c> <c r="d4" s="3" t="d"><v>2014-08-08T00:00:00.000</v></c> <c r="e4" s="4" t="n"><v>1.7500</v></c> <c r="f4" s="4" t="n"><v>0</v></c> <c r="g4" s="4" t="n"><v>0</v></c> <c r="h4" s="4" t="n"><v>0</v></c> <c r="i4" s="3" t="inlineStr"><is><t></t></is></c> </row> <row> <c r="a5" s="8"/> <c r="b5" s="8"/> <c r="c5" s="8"/> <c r="d5" s="8"/> <c r="e5" s="9" t="str"><f>SUM(e2:e4)</f><v>0</v></c> <c r="f5" s="9" t="str"><f>SUM(f2:f4)</f><v>0</v></c> <c r="g5" s="9" t="str"><f>SUM(g2:g4)</f><v>0</v></c> <c r="h5" s="9" t="str"><f>SUM(h2:h4)</f><v>0</v></c> <c r="i5" s="8"/> </row>';
		var sheetXml = xmlParse('<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>#arguments.data#</sheetData></worksheet>');
		
		file action="write" file="#variables.xlsxDirPath#\xl\worksheets\sheet#arguments.sheetnumber#.xml" output="#sheetXml#";
	}

	public void function send() {
		zip action="zip" source="#variables.xlsxDirPath#" file="#variables.zipPath#";
		
		var fileSize = getFileInfo(variables.zipPath).size;

		header name="content-disposition" value='attachment; filename="yoobi.xlsx"';
		header name="content-length" value="#fileSize#";
		content type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" file="#variables.zipPath#";

		file action="delete" file="#variables.zipPath#";
	}

	private void function createCore() {
		var coreXml = xmlParse(variables.coreXml);

		coreXml.XmlRoot["dc:creator"].xmlText = xmlFormat(config.creator);
		coreXml.XmlRoot["cp:lastModifiedBy"].xmlText =  xmlFormat(config.lastModifiedBy);
		coreXml.XmlRoot["dcterms:created"].xmlText = this.formatDate(config.created);
		coreXml.XmlRoot["dcterms:modified"].xmlText = this.formatDate(config.modified);

		file action="write" file="#variables.xlsxDirPath#\docProps\core.xml" output="#coreXml#";
	}

	private void function createApp() {
		var appXml = xmlParse(variables.appXml);
		//Set number of sheets
		appXml.XmlRoot.HeadingPairs["vt:vector"]["vt:variant"][2]["vt:i4"].xmlText = config.sheets.len();

		//Name of first sheet
		appXml.XmlRoot.TitlesOfParts["vt:vector"].xmlAttributes["size"] = config.sheets.len();
		appXml.XmlRoot.TitlesOfParts["vt:vector"]["vt:lpstr"][1].xmlText =  xmlFormat(config.sheets[1].name);
		//Add additional sheets and name them
		for(i=2;i<=config.sheets.len();i++) {
			arrayAppend(appXml.XmlRoot.TitlesOfParts["vt:vector"]["vt:lpstr"],xmlElemNew(appXml,"vt:lpstr"));
			appXml.XmlRoot.TitlesOfParts["vt:vector"]["vt:lpstr"][i].xmlText =  xmlFormat(config.sheets[i].name);
		}

		//Write xml
		file action="write" file="#variables.xlsxDirPath#\docProps\app.xml" output="#appXml#";
	}

	private void function createContentTypes() {
		var contentTypesXml = xmlParse(variables.contentTypesXml);
		var i = 0; var item=""; var el="";  var arr="";
		
		loop array="#config.sheets#" index="i" item="arr" {
			var el = xmlElemNew(contentTypesXml,"Override");
			el.xmlAttributes["ContentType"] = "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml";
			el.xmlAttributes["PartName"] = "/xl/worksheets/sheet#i#.xml";
			arrayAppend(contentTypesXml.XmlRoot.XmlChildren,el);
		}
		//Add Styles
		if(structKeyExists(variables.config,"styles")) {
			var el = xmlElemNew(contentTypesXml,"Override");
			el.xmlAttributes["ContentType"] = "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml";
			el.xmlAttributes["PartName"] = "/xl/styles.xml";
			arrayAppend(contentTypesXml.XmlRoot.XmlChildren,el);
		}

		//Write xml
		file action="write" file="#variables.xlsxDirPath#\[Content_Types].xml" output="#contentTypesXml#";
	}

	private void function createStyles() {
		if(structKeyExists(variables.config,"styles")) {
			//Write xml
			file action="write" file="#variables.xlsxDirPath#\xl\styles.xml" output="#variables.config.styles.getXml()#";
		}
		
	}

	private void function createWorkbook() {
		var workbookXml = xmlParse(variables.workbookXml);
		var i = 0; var item=""; var el="";  var arr="";

		loop array="#config.sheets#" index="i" item="arr" {
			var el = xmlElemNew(workbookXml,"sheet");
			el.xmlAttributes["name"] = xmlFormat(arr.name);
			el.xmlAttributes["sheetId"] = i;
			el.xmlAttributes["r:id"] = "rId#i#";
			arrayAppend(workbookXml.XmlRoot.sheets.XmlChildren,el);
		}

		//Write xml
		file action="write" file="#variables.xlsxDirPath#\xl\workbook.xml" output="#workbookXml#";
	}

	private void function createWorkbookRels() {
		var workbookRelsXml = xmlParse(variables.workbookRelsXml);
		var i = 0; var item=""; var el="";  var arr="";

		loop array="#config.sheets#" index="i" item="arr" {
			var el = xmlElemNew(workbookRelsXml,"Relationship");
			el.xmlAttributes["Type"] = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet";
			el.xmlAttributes["Target"] = "/xl/worksheets/sheet#i#.xml";
			el.xmlAttributes["Id"] = "rId#i#";
			arrayAppend(workbookRelsXml.XmlRoot.XmlChildren,el);
		}
		//Add Styles
		if(structKeyExists(variables.config,"styles")) {
			var el = xmlElemNew(workbookRelsXml,"Relationship");
			el.xmlAttributes["Type"] = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles";
			el.xmlAttributes["Target"] = "/xl/styles.xml";
			el.xmlAttributes["Id"] = "rId#config.sheets.len()+1#";
			arrayAppend(workbookRelsXml.XmlRoot.XmlChildren,el);
		}
		//Write xml
		file action="write" file="#variables.xlsxDirPath#\xl\_rels\workbook.xml.rels" output="#workbookRelsXml#";
	}

	private void function createRels() {
		var relsXml = xmlParse(variables.relsXml);
		//Write xml
		file action="write" file="#variables.xlsxDirPath#\_rels\.rels" output="#relsXml#";
	}

	private string function formatDate(date cfDate) {
		var dateString = dateFormat(arguments.cfDate,"YYYY-mm-dd");
		var timeString = timeFormat(arguments.cfDate,"HH:mm:ss");

		return dateString & "T" & timeString & "Z";
	}
}