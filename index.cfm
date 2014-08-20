<cfset Xlsx = createObject("component","Xlsx")>
<cfset Xlsx.init(	{	creator="Bas van der Graaf"
						,lastModifiedBy="Bas van der Graaf"
						,created=now()
						,modified=now()
						,sheets= [
							{ name="Sheet 1" },
							{ name="Sheet 1" }
						] 
					})>
<cfabort />
<cffile action="read" file="ram:///filename.txt" variable="ramfile"/>

<cfdump var="#ramfile#" /><cfdump var="#GetVFSMetaData('ram')#"/><cfabort />

<cffile action="write" file="ram:///filename.txt" output="This file is an in-memory file"/>
<cfzip action="zip" source="ram:///filename.txt" file="test2.zip">

<cffile action="write" file="zip://test.zip!test.txt" output="test">
<cfdump var="#cgi#" />


