<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<!--#include file="../../ioRDFunctions.asp"-->
<%
	dim sResult1,sMarkedRQLData, sFileData, sKey, sOldKey
	Server.ScriptTimeout = 600
	Response.Expires=0
	response.flush()
	C34=Chr(34)
	sFileName=Request.ServerVariables("APPL_PHYSICAL_PATH") & "/PlugIns/ToolExecuteRQL/cache/executeRQL.txt"
	sRQLRequest=Request("sRQLRequest")
	sMarkedRQLData=Request("MarkedRQL")
	iScrollTop=Request("ScrollTop")
	sRQLResult=Request("sRQLResult")
	if sRQLRequest="" then
		sRQLRequest=objIO.GetFileData(sFileName)
	else
		objIO.PutFileData sFileName,sRQLRequest
	end if
	sKey=session("SessionKey")
	if sKey<>"" then
		sRQLRequest=Replace(sRQLRequest,"[!key!]",sKey)
	end if
	sLoginGuid=session("LoginGUID")
	if sLoginGuid<>"" then
		sRQLRequest=Replace(sRQLRequest,"[!guid_login!]",sLoginGuid)
	end if
	sServerGuid=session("EditorialServerGuid")
	if sServerGuid<>"" then
		sRQLRequest=Replace(sRQLRequest,"[!guid_editorialserver!]",sServerGuid)
	end if
	sUserGuid=session("UserGuid")
	if sUserGuid<>"" then
		sRQLRequest=Replace(sRQLRequest,"[!guid_user!]",sUserGuid)
	end if
	sProjectGuid=session("ProjectGuid")
	if sProjectGuid<>"" then
		sRQLRequest=Replace(sRQLRequest,"[!guid_project!]",sProjectGuid)
	end if
	sTemplateGuid=session("TemplateGuid")
	if sTemplateGuid<>"" then
		sRQLRequest=Replace(sRQLRequest,"[!guid_template!]",sTemplateGuid)
	end if
	sPageGuid=session("PageGuid")
	if sPageGuid<>"" then
		sRQLRequest=Replace(sRQLRequest,"[!guid_page!]",sPageGuid)
	end if
	sLinkGuid=session("LinkGuid")
	if sLinkGuid<>"" then
		sRQLRequest=Replace(sRQLRequest,"[!guid_link!]",sLinkGuid)
	end if
	'objIO.Host = Request("Host")
	if trim(sRQLRequest)<>"" then
		if sMarkedXmlData="" then
			sRQLResult=objIO.ServerExecuteXml(sRQLRequest,sError)
		else
			sRQLResult=objIO.ServerExecuteXml(sMarkedRQLData,sError)
		end if
		if sError<>"" then
			sRQLResult="Error:" & sError
		end if
	else
		sRQLResult="No Data!"
	end if	
%>
<html>
	<head>
		<title>Tool: Execute RQL</title>
		<meta http-equiv="expires" content="0" />
		<meta name="developer" content="Thomas Pollinger" />
		<meta name="version" content="1.0.1.2" />
		<link rel="stylesheet" type="text/css" href="css/executeRQL.css" />
		<script type="text/javascript" src="js/executeRQL.js"></script>
	</head>
	<body>
		<form id="ioInfo" name="ioInfo">
			<table>
				<tr>
					<td class="label" colspan="2">Project (<b><% =session("Project") %></b>)</td>
					<td class="label">ProjectGUID:</td><td><input type="text" onfocus="this.select()" value="<% =session("ProjectGuid") %>" /></td>
					<td class="label">SessionKey:</td><td><input type="text" onfocus="this.select()" value="<% =session("SessionKey") %>" /></td>
				</tr>
				<tr>
					<td class="label" colspan="2">PageName/ID (<b><% =session("PageHeadline") %></b>)/(<b><% =session("PageId") %></b>)</td>
					<td class="label"><i>Selected</i> PageGUID:</td><td><input type="text" onfocus="this.select()" value="<% =session("PageGuid") %>" /></td>
					<td class="label">LoginGUID:</td><td><input type="text" onfocus="this.select()" value="<% =session("LoginGuid") %>" /></td>
				</tr>
				<tr>
					<td class="label" colspan="2">Content Class (<b><% =session("TemplateTitle") %></b>)</td>
					<td class="label">TemplateGUID:</td><td><input type="text" onfocus="this.select()" value="<% =session("TemplateGuid") %>" /></td>
					<td class="label">UserGUID:</td><td><input type="text" onfocus="this.select()" value="<% =session("UserGuid") %>" /></td>
				</tr>
				<tr>
					<td class="label" colspan="2">LinkName (<b><% =session("LinkName") %></b>)</td>
					<td class="label">LinkGUID:</td><td><input type="text" onfocus="this.select()" value="<% =session("LinkGuid") %>" /></td>
					<td class="label">EditorialServerGUID:</td><td><input type="text" onfocus="this.select()" value="<% =session("EditorialServerGuid") %>" /></td>
				</tr>
			</table>
		</form>
		<form id="ioText" name="ioText" action="index.asp" method="post">
			<input type="button" name="btExecuteMarked" value="Execute marked RQL" onClick="ExecuteRQL()" />
			<input type="hidden" name="MarkedRQL" />
			<input type="hidden" name="ScrollTop" value="<% =iScrollTop %>" />
			<input type="hidden" name="Host" value="<% =Request("Host") %>" />
			<div>
				<b>Request:</b><br /><hr />
				<textarea class="sRQLRequest" name="sRQLRequest" rows="20" onMouseUp="SaveSelection()" onKeyUp="SaveSelection()"><% =sRQLRequest %></textarea>
				<hr />
				<br />
				<b>Result:</b><br /><hr />
				<textarea class="sRQLResult" rows="20" wrap="off"><% =Replace(sRQLResult,"><",">"+chr(13)+"<") %></textarea>
				<hr />
			</div>
		</form>
	</body>
</html>
