<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<title>Cottages In Devon - Availability Calendar</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<style>
  body{font-family:Verdana, Arial, Helvetica, sans-serif; color:#000000; background-color:#F0F0F0; margin:0px;}
  #cal{width:800px; height:300px; background-color:#F0F0F0; text-align:center;}
  #cal h1{font-size:16px; color:#000000;}
  #cal h2{font-size:14px; font-weight:bold; color:#000000; line-height:20px; height:20px;}
  #cal h2 a{font-size:13px; text-decoration:none; color:#000000;}
  #cal h2 a:hover{font-size:13px; text-decoration:underline;}
  #cal th{border-bottom:1px solid #999999; color:#000000;}
  #cal td.month{color:#666666; font-size:12px; font-weight:bold; width:80px; border-right:1px solid #999999; text-align:left;}
  #cal td.week{background-color:#EEEEEE;}
  td.green{background-color:#00FF00; border:1px solid #CCCCCC; color:#000000; text-align:center; font-size:11px;}
  td.red{background-color:#FF0000; border:1px solid #CCCCCC; color:#000000; text-align:center; font-size:11px;}
    
</style>
<script type="text/javascript">
  function highlight(element, part, color){			
	if(color == 'green'){
	  element.style.backgroundColor = '#B1FCB4'
	  if(document.getElementById(part) != null) document.getElementById(part).style.backgroundColor = '#B1FCB4'
	}else{
	  element.style.backgroundColor = '#FB9F9F'
	  if(document.getElementById(part) != null) document.getElementById(part).style.backgroundColor = '#FB9F9F'
	}
  }	  
  
  function normal(element, part, color){	
	if(color == 'green'){
	  element.style.backgroundColor = '#00FF00'
	  if(document.getElementById(part) != null) document.getElementById(part).style.backgroundColor = '#00FF00'
	}else{
	  element.style.backgroundColor = '#FF0000'
	  if(document.getElementById(part) != null) document.getElementById(part).style.backgroundColor = '#FF0000'
	}
  }
  
  function href(url){
    if(url != '#'){
	 window.opener.document.location.href = url;
	 window.opener.focus();
	 window.close();			
    }
  }	  
</script>
</head>  
<body>
<div id="cal">
<%
  Dim prop, yr
  prop = Request.QueryString("PID")  
  yr = Request.QueryString("yr")
  If yr = "" Then yr = Year(Now())
%>
<h1><% If prop = 1 Then Response.Write("Availability for 1 Stattens Cottages") 
       If prop = 2 Then Response.Write("Availability for Carpenters")
	   If prop <> 1 And prop <> 2 Then Response.Redirect("./default.htm") %></h1>
<h2>
<%
  If Cint(yr) > Cint(Year(Now())) Then 
    Response.Write("<a href=""calendar.asp?PID=" & prop & "&yr=" & yr-1 & """>Previous Year</a> &nbsp; ") 
  End If
  Response.Write(yr)
  Response.Write(" &nbsp; <a href=""calendar.asp?PID=" & prop & "&yr=" & yr+1 & """>Next Year</a>") 
%></h2>
<table cellspacing="0" cellpadding="2">
  <tr>
    <td>&nbsp;</td><th><img src="./img/1.jpg"></th><th><img src="./img/2.jpg"></th><th><img src="./img/3.jpg"></th><th><img src="./img/4.jpg"></th><th><img src="./img/5.jpg"></th><th><img src="./img/6.jpg"></th><th><img src="./img/7.jpg"></th><th><img src="./img/8.jpg"></th><th><img src="./img/9.jpg"></th><th><img src="./img/10.jpg"></th><th><img src="./img/11.jpg"></th>
    <th><img src="./img/12.jpg"></th><th><img src="./img/13.jpg"></th><th><img src="./img/14.jpg"></th><th><img src="./img/15.jpg"></th><th><img src="./img/16.jpg"></th><th><img src="./img/17.jpg"></th><th><img src="./img/18.jpg"></th><th><img src="./img/19.jpg"></th><th><img src="./img/20.jpg"></th><th><img src="./img/21.jpg"></th><th><img src="./img/22.jpg"></th>
    <th><img src="./img/23.jpg"></th><th><img src="./img/24.jpg"></th><th><img src="./img/25.jpg"></th><th><img src="./img/26.jpg"></th><th><img src="./img/27.jpg"></th><th><img src="./img/28.jpg"></th><th><img src="./img/29.jpg"></th><th><img src="./img/30.jpg"></th><th><img src="./img/31.jpg"></th>
  </tr>
<%
  Dim conn
  Set conn = Server.CreateObject("ADODB.connection")
  conn.Provider = "Microsoft.Jet.OLEDB.4.0"
  conn.Open(Server.MapPath("./cottagesindevon.mdb"))
    
  Dim i, j
  For i = 1 To 12
    Response.Write("  <tr><td class=""month"">" & MonthName(i) & "</td>" & vbcrlf)		
	Dim countdate, nextsat, cssclass, price, url, w, d 
	countdate = CDate("1/" & i & "/" & yr)
	j = Month(countdate)
	Do While j = i
	  nextsat = DateAdd("d", vbSaturday - Weekday(countdate), countdate)
	  If nextsat = countdate Then nextsat = DateAdd("d",7, countdate)	  
	  If Day(countdate) = 1 Then 
	    Set rs = conn.Execute("SELECT available, price, discount, FootnoteSymbol FROM Availability WHERE property = " & prop & " AND date = #" & Year(countdate) & "/" & month(countdate) & "/" & day(countdate) & "#")
		If rs.EOF Then
			d = DateAdd("d", vbSaturday - Weekday(countdate) - 7, countdate)
			rs.close
			Set rs = conn.Execute("SELECT available, price, discount, FootnoteSymbol FROM Availability WHERE property = " & prop & " AND date = #" & Year(d) & "/" & month(d) & "/" & day(d) & "#")
		End If
	  Else
	    Set rs = conn.Execute("SELECT available, price, discount, FootnoteSymbol FROM Availability WHERE property = " & prop & " AND date = #" & Year(countdate) & "/" & month(countdate) & "/" & day(countdate) & "#")
	  End If
	  If rs.EOF Then
	    cssclass = "green" 
		price = "POA"
		url = "./contact.htm" 
	  Else
	    If rs("available") = true Then
		  cssclass = "green"
		  url = "./contact.htm" 
		Else
		  cssclass = "red" 
		  url = "#" 
		End If
		If rs("Price") = 0 Then
		  price = "POA"
		Else
		  price = FormatNumber(rs("Price"),2)
		End If
		If rs("discount") = true Then  
		  discount = "*"
		Else
		  discount = ""
		End If
        If rs("FootnoteSymbol") <> "" Then  
		  discount = discount & rs("FootnoteSymbol")
		End If
		rs.close
	  End If
	  
	  If Day(countdate) = 1 Then
		Response.Write("    <td class=""week"" colspan=""" & DateDiff("d", countdate, nextsat)  & """>")
		Response.Write("<table width=""100%"" cellpadding=""2"" cellspacing=""0""><tr><td id=""f" & i & """ onClick=""href('" & url & "')"" onMouseOut=""normal(this, 'l"&i-1&"', '" & cssClass & "')"" onMouseOver=""highlight(this, 'l"&i-1&"', '" & cssClass & "')"" class=""" & cssclass & """>")
		If DateDiff("d", countdate, nextsat, countdate) < 4 Then
		  Response.Write("&nbsp;")
		  'Response.Write(countdate)
		Else
		  Response.Write("£" & price & discount)
		  'Response.Write(countdate)
		End If
		Response.Write("</td></tr></table>")  		  
		Response.Write("</td>" & vbcrlf)
	  Else
		If (day(countdate) + DateDiff("d", countdate, nextsat)) > Day(DateSerial(yr, i + 1, 0)) + 1 Then
		  Response.Write("    <td class=""week"" colspan=""" & DateDiff("d", countdate, DateSerial(yr, i + 1, 0)) + 1  & """>")
		  Response.Write("<table width=""100%"" cellpadding=""2"" cellspacing=""0""><tr><td id=""l" & i & """ onClick=""href('" & url & "')"" onMouseOut=""normal(this, 'f"&i+1&"', '" & cssClass & "')"" onMouseOver=""highlight(this, 'f"&i+1&"', '" & cssClass & "')"" class=""" & cssclass & """>")
		  If DateDiff("d", countdate, DateSerial(yr, i + 1, 0)) + 2 < 4 Then
		    Response.Write("&nbsp;")
		  Else
		    Response.Write("£" & price & discount)		
		  End If		
		  Response.Write("</td></tr></table>")
		  Response.Write("</td>" & vbcrlf)
		Else
		  Response.Write("    <td class=""week"" colspan=""" & DateDiff("d", countdate, nextsat) & """>")
		  Response.Write("<table width=""100%"" cellpadding=""2"" cellspacing=""0""><tr><td onClick=""href('" & url & "')"" onMouseOut=""normal(this, '', '" & cssClass & "')"" onMouseOver=""highlight(this, '', '" & cssClass & "')"" class=""" & cssclass & """>")
		  Response.Write("£" & price & discount)		  
		  Response.Write("</td></tr></table>")
		  Response.Write("</td>" & vbcrlf)
		End If
	  End If
	   
	  If DateDiff("d", countdate, DateAdd("d", vbSaturday - Weekday(countdate), countdate)) = 0 Then
	    countdate = DateAdd("d", 7, countdate)
      Else
	    countdate = DateAdd("d", vbSaturday - Weekday(countdate), countdate)
	  End If
	  j = Month(countdate)
	  Set rs = nothing
	  
	Loop
	Response.Write("  </tr>" & vbcrlf)
  Next
  conn.Close()
%>
</table>

<p><strong>Key</strong><br />
<font color="#00FF00">Available for Bookings</font><br><font color="#FF0000">Unavailable</font></p>
<%If prop = 1 Then 
	Response.Write("<p style='text-align:left; font-size:12px;'>* Discount available when only two people staying. Please contact us for details</p>")
  End If %>
</div>
</body>
</html>