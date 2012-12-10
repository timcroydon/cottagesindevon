
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "DTD/xhtml1-transitional.dtd">
<html>
<head>
<title>Holiday Cottages In Devon - Combe Martin, North Devon</title>
<meta name="keywords" content="Devon, Cottage, Devon cottage, holiday, rent, to let, Combe Martin, UK, GB, England, South West" />
<meta name="description" content="Cottages in Devon provide holiday rental accomodation in a North Devon cottage. Located in Combe Martin, our holiday home provides a great escape from your city or suburban life." />
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1"/>
<link href="../devon.css" rel="stylesheet" type="text/css" />

<style>
  #cal{width:780px; height:300px; background-color:#F0F0F0; text-align:center;}
  #cal h1{font-size:16px; color:#FF0000;}
  #cal h2{font-size:14px; font-weight:bold; color:#000000; line-height:20px; height:20px;}
  #cal h2 a{font-size:13px; text-decoration:none; color:#000000;}
  #cal h2 a:hover{font-size:13px; text-decoration:underline;}
  #cal th{border-bottom:1px solid #999999; color:#000000;}
  #cal td.month{color:#666666; font-size:12px; font-weight:bold; border-right:1px solid #999999; text-align:left;}
  #cal td.week{background-color:#EEEEEE;}
  td.green{background-color:#00FF00; border:1px solid #CCCCCC; color:#000000; text-align:center; font-size:11px;}
  td.red{background-color:#FF0000; border:1px solid #CCCCCC; color:#000000; text-align:center; font-size:11px;} 
  h1 a{color:#000000;}    
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
    if(url != '#') document.location.href = url
  }	  
</script>
</head><body>
<div id="container">
<div id="banner"><h1>COTTAGES IN DEVON</h1></div>
<div id="centercontent">
<h2>Availability Admin</h2>
<div id="cal">
<%
  Dim prop, yr
  prop = Request.QueryString("PID") 
  If prop = "" Then prop = 1 
  yr = Request.QueryString("yr")
  If yr = "" Then yr = Year(Now())
%>
<% If prop = 1 Then Response.Write("<h1>1 Stattens Cottages &nbsp; <a href=""./calendar.asp?PID=2&yr=" & yr & """>Carpenters</a></h1>") 
   If prop = 2 Then Response.Write("<h1><a href=""./calendar.asp?PID=1&yr=" & yr & """>1 Stattens Cottages</a> &nbsp; Carpenters</h1>")
%>
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
    <td>&nbsp;</td><th><img src="../img/1.jpg"></th><th><img src="../img/2.jpg"></th><th><img src="../img/3.jpg"></th><th><img src="../img/4.jpg"></th><th><img src="../img/5.jpg"></th><th><img src="../img/6.jpg"></th><th><img src="../img/7.jpg"></th><th><img src="../img/8.jpg"></th><th><img src="../img/9.jpg"></th><th><img src="../img/10.jpg"></th><th><img src="../img/11.jpg"></th>
    <th><img src="../img/12.jpg"></th><th><img src="../img/13.jpg"></th><th><img src="../img/14.jpg"></th><th><img src="../img/15.jpg"></th><th><img src="../img/16.jpg"></th><th><img src="../img/17.jpg"></th><th><img src="../img/18.jpg"></th><th><img src="../img/19.jpg"></th><th><img src="../img/20.jpg"></th><th><img src="../img/21.jpg"></th><th><img src="../img/22.jpg"></th>
    <th><img src="../img/23.jpg"></th><th><img src="../img/24.jpg"></th><th><img src="../img/25.jpg"></th><th><img src="../img/26.jpg"></th><th><img src="../img/27.jpg"></th><th><img src="../img/28.jpg"></th><th><img src="../img/29.jpg"></th><th><img src="../img/30.jpg"></th><th><img src="../img/31.jpg"></th>
  </tr>
<%
  Dim conn
  Set conn = Server.CreateObject("ADODB.connection")
  conn.Provider = "Microsoft.Jet.OLEDB.4.0"
  conn.Open(Server.MapPath("../cottagesindevon.mdb"))
    
  Dim i, j
  For i = 1 To 12
    Response.Write("  <tr><td class=""month"">" & Left(MonthName(i),1) & "</td>" & vbcrlf)		
	Dim countdate, nextsat, cssclass, price, d 
	countdate = CDate("1/" & i & "/" & yr)
	j = Month(countdate)
	Do While j = i
	  nextsat = DateAdd("d", vbSaturday - Weekday(countdate), countdate)
	  If nextsat = countdate Then nextsat = DateAdd("d",7, countdate)
	  If Day(countdate) = 1 Then 
	    Set rs = conn.Execute("SELECT available, price, discount FROM Availability WHERE property = " & prop & " AND date = #" & Year(countdate) & "/" & month(countdate) & "/" & day(countdate) & "#")
		If rs.EOF Then
			d = DateAdd("d", vbSaturday - Weekday(countdate) - 7, countdate)
			rs.close
			Set rs = conn.Execute("SELECT available, price, discount FROM Availability WHERE property = " & prop & " AND date = #" & Year(d) & "/" & month(d) & "/" & day(d) & "#")
		Else
		  d = countdate
		End If		
	  Else
	    d = countdate
		Set rs = conn.Execute("SELECT available, price, discount FROM Availability WHERE property = " & prop & " AND date = #" & Year(countdate) & "/" & month(countdate) & "/" & day(countdate) & "#")
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
		rs.close
	  End If
	  
	  If Day(countdate) = 1 Then
		Response.Write("    <td class=""week"" colspan=""" & DateDiff("d", countdate, nextsat)  & """>")
		Response.Write("<table width=""100%"" cellpadding=""2"" cellspacing=""0""><tr><td id=""f" & i & """ onClick=""href('./editweek.asp?PID=" & prop & "&yr=" & yr & "&price=" & price & "&week=" & d & "')"" onMouseOut=""normal(this, 'l"&i-1&"', '" & cssClass & "')"" onMouseOver=""highlight(this, 'l"&i-1&"', '" & cssClass & "')"" class=""" & cssclass & """>")
		If DateDiff("d", countdate, nextsat, countdate) < 4 Then
		  Response.Write("&nbsp;")
		Else
		  Response.Write("£" & price & discount)
		  'Response.Write(d)
		End If
		Response.Write("</td></tr></table>")  		  
		Response.Write("</td>" & vbcrlf)
	  Else
		If (day(countdate) + DateDiff("d", countdate, nextsat)) > Day(DateSerial(yr, i + 1, 0)) + 1 Then
		  Response.Write("    <td class=""week"" colspan=""" & DateDiff("d", countdate, DateSerial(yr, i + 1, 0)) + 1  & """>")
		  Response.Write("<table width=""100%"" cellpadding=""2"" cellspacing=""0""><tr><td id=""l" & i & """ onClick=""href('./editweek.asp?PID=" & prop & "&yr=" & yr & "&price=" & price & "&week=" & d & "')"" onMouseOut=""normal(this, 'f"&i+1&"', '" & cssClass & "')"" onMouseOver=""highlight(this, 'f"&i+1&"', '" & cssClass & "')"" class=""" & cssclass & """>")
		  If DateDiff("d", countdate, DateSerial(yr, i + 1, 0)) + 2 < 4 Then
		    Response.Write("&nbsp;")
		  Else
		    Response.Write("£" & price & discount)
		  End If		
		  Response.Write("</td></tr></table>")
		  Response.Write("</td>" & vbcrlf)
		Else
		  Response.Write("    <td class=""week"" colspan=""" & DateDiff("d", countdate, nextsat) & """>")
		  Response.Write("<table width=""100%"" cellpadding=""2"" cellspacing=""0""><tr><td onClick=""href('./editweek.asp?PID=" & prop & "&yr=" & yr & "&price=" & price & "&week=" & d & "')"" onMouseOut=""normal(this, '', '" & cssClass & "')"" onMouseOver=""highlight(this, '', '" & cssClass & "')"" class=""" & cssclass & """>")
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
<p>Click week to edit.</p>
</div>
<br />
<br />
<br />
<br />
<br />
<br />
<br />
</td>
</div>

<div id="footer"> </div>

</div>
</body>
</html>

