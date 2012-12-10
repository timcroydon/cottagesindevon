<html>
<%
	response.buffer=true
%>
<head>
<title>Contact Page</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>

<%

sch = "http://schemas.microsoft.com/cdo/configuration/"

strMailBody = "Cottages In Devon Enquiry:" & vbCrLf
	strMailBody = strMailBody & vbCrLf
	strMailBody = strMailBody & "Property " & vbTab & Request.Form("Property") & vbCrLf
	strMailBody = strMailBody & "Required " & vbTab & Request.Form("DateFrom") & " - " & Request.Form("DateTo") & vbCrLf
	strMailBody = strMailBody & "Alternative" & vbTab & Request.Form("AlternativeDateFrom") & " - " & Request.Form("AlternativeDateTo") & vbCrLf
	strMailBody = strMailBody & "Title " & vbTab & Request.Form("Title") & vbCrLf
	strMailBody = strMailBody & "Title-Other " & vbTab & Request.Form("TitleOther") & vbCrLf
	strMailBody = strMailBody & "Forename " & vbTab & Request.Form("Forename") & vbCrLf
	strMailBody = strMailBody & "Surname " & vbTab & Request.Form("surname") & vbCrLf
	strMailBody = strMailBody & "No/Street " & vbTab & Request.Form("Street") & vbCrLf
	strMailBody = strMailBody & "Town " & vbTab & vbTab & Request.Form("Town") & vbCrLf
	strMailBody = strMailBody & "County " & vbTab & Request.Form("County") & vbCrLf
	strMailBody = strMailBody & "Postcode " & vbTab & Request.Form("PostCode") & vbCrLf
	strMailBody = strMailBody & "Country " & vbTab & Request.Form("Country") & vbCrLf
	strMailBody = strMailBody & "Telephone " & vbTab & Request.Form("Telephone") & vbCrLf
	strMailBody = strMailBody & "Telephone (Work) " & vbTab & Request.Form("TelephoneWork") & vbCrLf
	strMailBody = strMailBody & "Email " & vbTab & Request.Form("Email") & vbCrLf
	strMailBody = strMailBody & "Party Total " & vbTab & Request.Form("TotalParty") & vbCrLf
	strMailBody = strMailBody & "No of Men " & vbTab & Request.Form("NoofMen") & vbCrLf
	strMailBody = strMailBody & "No of Women " & vbTab & Request.Form("NoofWomen") & vbCrLf
	strMailBody = strMailBody & "No of Children " & vbTab & Request.Form("NoofChildren") & vbCrLf
	strMailBody = strMailBody & "Childrens Ages " & vbTab & Request.Form("ChildrensAges") & vbCrLf
	strMailBody = strMailBody & "Description Pets " & vbTab & Request.Form("Pets") & vbCrLf
	strMailBody = strMailBody & "Linen Single " & vbTab & Request.Form("LinenSingle") & vbCrLf
	strMailBody = strMailBody & "Linen Double " & vbTab & Request.Form("LinenDouble") & vbCrLf
	strMailBody = strMailBody & "Cot Required " & vbTab & Request.Form("Cot") & vbCrLf
	strMailBody = strMailBody & "Comments " & vbTab & Request.Form("Comments") & vbCrLf

Set cdoConfig = Server.CreateObject("CDO.Configuration")
cdoConfig.Fields.Item(sch & "sendusing") = 2
cdoConfig.Fields.Item(sch & "smtpserver") = "mail.cottagesindevon.com"
cdoConfig.Fields.Item(sch & "smtpserverport") = 25
cdoConfig.fields.update
Set sendmail = Server.CreateObject("CDO.Message")
Set sendmail.Configuration = cdoConfig
Sendmail.From = "website@cottagesindevon.com"
Sendmail.To = "enquiry@cottagesindevon.com"
Sendmail.Subject = "Cottages In Devon Enquiry"
Sendmail.TextBody = strMailBody
Sendmail.Send

' free variables
Set sendmail = Nothing
Set cdoConfig = Nothing
	

	


	response.redirect "confirmation.htm"
%>
<body bgcolor="#FFFFFF">

</body>
</html>
