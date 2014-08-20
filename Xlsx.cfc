component xlsx output="false" {
	variables.xlsxID = createUUID();
	variables.tmpDirPath = expandPath('\tmp');
	variables.xlsxDirPath = variables.tmpDirPath&"\"&variables.xlsxID;
	variables.coreXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><dc:creator>Bas van der Graaf</dc:creator><cp:lastModifiedBy>Bas van der Graaf</cp:lastModifiedBy><dcterms:created xsi:type="dcterms:W3CDTF">2014-07-31T09:08:26Z</dcterms:created><dcterms:modified xsi:type="dcterms:W3CDTF">2014-07-31T09:09:32Z</dcterms:modified></cp:coreProperties>';
	variables.appXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"><Application>Microsoft Excel</Application><DocSecurity>0</DocSecurity><ScaleCrop>false</ScaleCrop><HeadingPairs><vt:vector size="2" baseType="variant"><vt:variant><vt:lpstr>Werkbladen</vt:lpstr></vt:variant><vt:variant><vt:i4>1</vt:i4></vt:variant></vt:vector></HeadingPairs><TitlesOfParts><vt:vector size="1" baseType="lpstr"><vt:lpstr>Blad1</vt:lpstr></vt:vector></TitlesOfParts><Company>Yoobi</Company><LinksUpToDate>false</LinksUpToDate><SharedDoc>false</SharedDoc><HyperlinksChanged>false</HyperlinksChanged><AppVersion>14.0300</AppVersion></Properties>';

	public xlsx function init(struct config) {
		//DEV code
		directoryDelete(variables.tmpDirPath,true);
		//directories regelen
		if(!directoryExists(variables.tmpDirPath)) directoryCreate(variables.tmpDirPath);
		if(!directoryExists(variables.xlsxDirPath)) directoryCreate(variables.xlsxDirPath);
		if(!directoryExists(variables.xlsxDirPath&"\docProps")) directoryCreate(variables.xlsxDirPath&"\docProps");
		if(!directoryExists(variables.xlsxDirPath&"\_rels")) directoryCreate(variables.xlsxDirPath&"\_rels");
		if(!directoryExists(variables.xlsxDirPath&"\_xl")) directoryCreate(variables.xlsxDirPath&"\_xl");
		if(!directoryExists(variables.xlsxDirPath&"\_xl\_rels")) directoryCreate(variables.xlsxDirPath&"\_xl\_rels");
		if(!directoryExists(variables.xlsxDirPath&"\_xl\worksheets")) directoryCreate(variables.xlsxDirPath&"\_xl\worksheets");

		this.createCore(config);
		this.createApp(config);

		return this;
	}

	private void function createCore(struct config) {
		var coreXml = xmlParse(variables.coreXml);

		coreXml.XmlRoot["dc:creator"].xmlText = xmlFormat(config.creator);
		coreXml.XmlRoot["cp:lastModifiedBy"].xmlText =  xmlFormat(config.lastModifiedBy);
		coreXml.XmlRoot["dcterms:created"].xmlText = this.formatDate(config.created);
		coreXml.XmlRoot["dcterms:modified"].xmlText = this.formatDate(config.modified);

		file action="write" file="#variables.xlsxDirPath#\docProps\core.xml" output="#coreXml#";
	}

	private void function createApp(struct config) {
		var appXml = xmlParse(variables.appXml);
		
		writeDump(appXml.XmlRoot.HeadingPairs["vt:vector"]["vt:variant"][2]);

		file action="write" file="#variables.xlsxDirPath#\docProps\core.xml" output="#coreXml#";
	}

	private string function formatDate(date cfDate) {
		var dateString = dateFormat(arguments.cfDate,"yyy-mm-dd");
		var timeString = dateFormat(arguments.cfDate,"HH:mm:ss");

		return dateString & "T" & timeString & "Z";
	}
}