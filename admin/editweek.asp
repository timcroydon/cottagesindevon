
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "DTD/xhtml1-transitional.dtd">
<html>
<head>
<title>Holiday Cottages In Devon - Combe Martin, North Devon</title>
<meta name="keywords" content="Devon, Cottage, Devon cottage, holiday, rent, to let, Combe Martin, UK, GB, England, South West" />
<meta name="description" content="Cottages in Devon provide holiday rental accomodation in a North Devon cottage. Located in Combe Martin, our holiday home provides a great escape from your city or suburban life." />
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1"/>
<link href="../devon.css" rel="stylesheet" type="text/css" />
</head><body>
<div id="container">
<div id="banner"><h1>COTTAGES IN DEVON</h1></div>

<div id="centercontent">
<h2>Edit Week </h2>
<form action="editweeksubmit.asp" method="post" style="padding-left:40px;">
  <p>Property: <% If Request.QueryString("PID") = 1 Then Response.Write("1 Stattens Cottages") 
                  If Request.QueryString("PID") = 2 Then Response.Write("Carpenters")  %></p>
  <br>
  <p>Date: <%=Request.QueryString("week") %></p><br />
  <p>Price: £<input name="price" type="text" id="price" value="<%=Request.QueryString("Price") %>"><% If Request.QueryString("err") = 1 Then Response.Write("*must be numeric") %></p>
  <p>Available:
  <br>
  <br>
  <label><input type="radio" name="available" value="true" checked>Yes</label>
  <br>
  <label><input type="radio" name="available" value="false">No</label>
  </p>
  <p>Discount:
  <br>
  <br>
  <label><input type="radio" name="discount" value="true">Yes</label>
  <br>
  <label><input type="radio" name="discount" value="false" checked>No</label>
  </p>
  <br>
  <br>  
  
  <input type="hidden" name="week" value="<%=Request.QueryString("week") %>">
  <input type="hidden" name="property" value="<%=Request.QueryString("PID") %>">
  <input type="submit" name="Submit" value="Save">
  <input type="reset" name="Submit2" value="Clear">
  <input type="button" name="Submit3" value="Cancel" onclick="location.href='./calendar.asp?PID=<%=Request.QueryString("PID")%>&yr=<%=Request.QueryString("YR")%>'">
<br>
<br>
<br>

</form>

</div>

<div id="footer"> 
</div>

</div>
</body>
</html>
