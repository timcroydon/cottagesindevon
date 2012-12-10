<%

  Dim conn, rs, d, available
  Set conn = Server.CreateObject("ADODB.connection")
  conn.Provider = "Microsoft.Jet.OLEDB.4.0"
  conn.Open(Server.MapPath("../cottagesindevon.mdb"))
  
  d = Request.Form("Week")
  If Not IsNumeric(Request.Form("Price")) then
    Response.Redirect("./editweek.asp?PID=" & Request.Form("property") & "&yr=" & Year(d) & "&price=" & Request.Form("Price") & "&week=" & d & "&err=1")
  End If
  available = "no"
  if Request.Form("Available") = "true" Then available = "yes"
  discount = "no"
  if Request.Form("discount") = "true" Then discount = "yes"
  
  Set rs = conn.Execute("SELECT * FROM Availability WHERE property = " & Request.Form("property") & " AND date = #" & Year(d) & "/" & month(d) & "/" & day(d) & "#")
  If rs.EOF Then
    conn.execute("INSERT INTO Availability ([Date],[Property],[Price],[Available],[Discount]) VALUES (#" & Year(d) & "/" & month(d) & "/" & day(d) & "#, " & Request.Form("property") & ", "&Request.Form("Price") & ", " & available & "," & discount & ")")
  Else
   conn.execute("UPDATE Availability SET Price = " & Request.Form("Price") & ", Available = " & available & ", Discount = " & discount & " WHERE Date = #" & Year(d) & "/" & month(d) & "/" & day(d) & "# AND Property = " & Request.Form("property"))
  End If
  rs.Close
  Response.Redirect("./calendar.asp?PID=" & Request.Form("property") & "&yr=" & Year(d))
%>