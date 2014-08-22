component xlsx output="false" {

	public function init(string data) {
		variables.stylesXML = data;
	}

	public function getXml() {
		return variables.stylesXML;
	}

}