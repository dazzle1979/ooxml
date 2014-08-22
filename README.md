ooxml
=====

Create Xlsx fast en easy (CFML/Railo)


Example usage without styles
=====
```coldfusion
<cfset Xlsx = createObject("component","Xlsx")>
<cfset Xlsx.init(	{	creator="Bas van der Graaf"
						,lastModifiedBy="Bas van der Graaf"
						,created=now()
						,modified=now()
						,sheets= [
							{ name="Sheet1" },
							{ name="Sheet2" }
						] 
					})>
<cfset Xlsx.createSheet('<row><c r="a1" t="inlineStr"><is><t>Bas</t></is></c> <c r="b1" t="n"><v>10</v></c></row>',1)>
<cfset Xlsx.createSheet('<row><c r="a1" t="inlineStr"><is><t>Bas</t></is></c> <c r="b1" t="n"><v>10</v></c></row>',2)>

<cfset Xlsx.send()>
```
Help
=====
For help mail me at bvdgraaf@gmail.com.
