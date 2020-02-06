<%@ language=vbscript codepage=65001%>
<% Option Explicit 
'Response.Expires=0
'Response.ExpiresAbsolute=#Jan 01,1980 13:00#
'Response.AddHeader "Pragma", "No-Cache"
'Response.CacheControl = "no-cache"
%>
<!-- #INCLUDE FILE="IncContext.asp" -->
<!-- #INCLUDE FILE="IncErrMsg.asp" -->

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<meta http-equiv='Content-Type' content='text/html; charset=utf-8' />
<head>
	<title>Duplicate Publications</title>
	<link rel='stylesheet' type='text/css' href='Styles/InfoShare.css'/>
	<!--[if lte IE 8]>
	<link rel="stylesheet" type="text/css" href="Styles/InfoShare-IE8.css"/>
	<![endif]-->
</head>
<body bottommargin='0' topmargin='0' leftmargin='0'>

<%
	
	Dim sTitle
	Dim sContext
	Dim sMessage
	Dim sEventId
	dim sOldFolderId
	dim sNewFolderId
	dim sDestinationFolderName
	dim sSourceFolderName
	dim lTotal
	Dim lErrors
	
	Dim sCardIds
	
	Dim oEvent
    Dim oEventDataXML
	Dim oEventDataXMLRoot
	Dim oXMLElement
	
	Dim oLogObject
	'Dim sEventId
	Dim sEventName
	Dim sEventData	

	sContext = GetContext()
	
	sOldFolderId = Request("OldFolderId")
	sNewFolderId = Request("NewFolderId")
	sDestinationFolderName = Request("DestinationFolderName")
	sSourceFolderName = Request("SourceFolderName")
	lTotal = Request.Form("Total")
	
	sCardIds = Request.Form("CardIds")
	
	Set oEvent = CreateObject("ISAuthor.CEvent")
	Set oLogObject = oEvent.LogObject
    Set oEventDataXML = CreateObject("MSXML2.DOMDocument.6.0")
	
	sEventId = ""
	
	'1. Retrieve QueryString parameters	
    sEventName = "CLONEPUBLICATION"
	sTitle = "Clone Function"
	
	
	'2. Build eventdata structure
	set oEventDataXMLRoot = oEventDataXML.createElement("eventdata")
    set oEventDataXML.documentElement = oEventDataXMLRoot
	
	set oXMLElement = oEventDataXML.createElement("lngcardids")
    oXMLElement.nodetypedvalue = sCardIds
    oEventDataXMLRoot.appendChild oXMLElement
    set oXMLElement = nothing
	
	set oXMLElement = oEventDataXML.createElement("newfolderid")
    oXMLElement.nodetypedvalue = sNewFolderId
    oEventDataXMLRoot.appendChild oXMLElement
    set oXMLElement = nothing
	
	set oXMLElement = oEventDataXML.createElement("oldfolderid")
    oXMLElement.nodetypedvalue = sOldFolderId
    oEventDataXMLRoot.appendChild oXMLElement
    set oXMLElement = nothing
	
	sEventData = oEventDataXML.xml
	'3. Send Event and receive corresponding event id
    oEvent.SendEvent sContext, sEventName, sEventData, sEventId
    sMessage=oLogobject.resolvedContent(sContext)
    lErrors =oLogobject.countErrorsorwarnings
	
%>

    <script language='javascript'>
      document.location.href = 'ShowEventDetails.asp?EventId=<% =server.URLEncode(cstr(sEventId)) %>&Title=<% =server.URLEncode(sTitle) %>';
    </script>
<!--
<div>
	Request Details: <%=Reform.HtmlEncode(sOldFolderId)%> / <%=Reform.HtmlEncode(sNewFolderId)%> / <%=Reform.HtmlEncode(sDestinationFolderName)%> / <%=Reform.HtmlEncode(sSourceFolderName)%> / <%=Reform.HtmlEncode(sCardIds)%>
</div>
-->
<script language='javascript'>
	//document.location.href = 'ShowEventDetails.asp?EventId=<% =server.URLEncode(cstr(sEventId)) %>&Title=<% =server.URLEncode(sTitle) %>';
</script>

</body>
</html>