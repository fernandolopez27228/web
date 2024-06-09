<html>
<body>
<%@ LANGUAGE="VBSCRIPT" %>
<%
option explicit
Dim conn, nom, primer_apellido, segundo_apellido,correo, mensaje

NOM=Request.Form("nombre")
PRIMER_APELLIDO=Request.Form("primer_apellido")
SEGUNDO_APELLIDO=Request.Form("segundo_apellido")
correo=Request.Form("correo")
mensaje=Request.Form("mensaje")

Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\bryan y jesus\Documents\diseno TLAHUASHOP\COMPRAS.accdb.mdb"
conn.Execute "INSERT INTO COMPRAS(nombre,primer_apellido,segundo_apellido,correo,mensaje) VALUES ('" & nom & "','" & primer_apellido & "','" & segundo_apellido & "','" & correo & "','" & mensaje & "')"
conn.close
set conn=nothing
response.redirect("FORMULARIo.html")

%>
</body>
</html>