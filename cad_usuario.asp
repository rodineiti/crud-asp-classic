<%
Response.Buffer  = true
Response.Expires = 0
Session.lcId     = 1033
%>
<!-- #include file="includes/conexao.asp" -->
<%

id   = Request.Form("id")
name = replace(trim(request.form("name")),"'","")
email = replace(trim(request.form("email")),"'","")

if (trim(id) = "") or (isnull(id)) then id = 0 end if
	
if cint(id) = 0 then
	
	strSQL = "INSERT INTO users (name, email) VALUES ('"&name&"','"&email&"');"
	conDB.execute(strSQL)

	response.redirect("index.asp?strStatus=INC")
	response.End
	
else

	strSQL = "UPDATE users SET name = '"&name&"', email = '"&email&"' WHERE id = " & id
	conDB.execute(strSQL)

	response.redirect("index.asp?strStatus=ALT")
	response.End
end if

%>