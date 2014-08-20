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
<cfset Xlsx.createSheet("",1)>
<cfset Xlsx.createSheet("",2)>
<cfset Xlsx.send()>


