<html>
<body>
<%@ LANGUAGE="VBSCRIPT" %>
<%
option explicit
Dim conn,gor1,gor2,gor3,gor4,gor5,gor6,gor7,gor8,gor9,pt1,pt2,pt3,pt4,pt5,pt6,pt7,pt8,pt9,pt10,pt11,c1,c2,c3,c4,c5,c6,c7,c8,c9,c10,c11,c12,p1,p2,p3,p4,p5,s1,s2,s3,s4,s5,s6,s7,s8,s9
Dim id,cantidad,codigop,precio


id=Request.Form("id")
gor1=Request.Form("gor1")
gor2=Request.Form("gor2")
gor3=Request.Form("gor3")
gor4=Request.Form("gor4")
gor5=Request.Form("gor5")
gor6=Request.Form("gor6")
gor7=Request.Form("gor7")
gor8=Request.Form("gor8")
gor9=Request.Form("gor9")
pt1=Request.Form("pt1")
pt2=Request.Form("pt2")
pt3=Request.Form("pt3")
pt4=Request.Form("pt4")
pt5=Request.Form("pt5")
pt6=Request.Form("pt6")
pt7=Request.Form("pt7")
pt8=Request.Form("pt8")
pt9=Request.Form("pt9")
pt10=Request.Form("pt10")
pt11=Request.Form("pt11")
c1=Request.Form("c1")
c2=Request.Form("c2")
c3=Request.Form("c3")
c4=Request.Form("c4")
c5=Request.Form("c5")
c6=Request.Form("c6")
c7=Request.Form("c7")
c8=Request.Form("c8")
c9=Request.Form("c9")
c10=Request.Form("c10")
c11=Request.Form("c11")
c12=Request.Form("c12")
p1=Request.Form("p1")
p2=Request.Form("p2")
p3=Request.Form("p3")
p4=Request.Form("p4")
p5=Request.Form("p5")
s1=Request.Form("s1")
s2=Request.Form("s2")
s3=Request.Form("s3")
s4=Request.Form("s4")
s5=Request.Form("s5")
s6=Request.Form("s6")
s7=Request.Form("s7")
s8=Request.Form("s8")
s9=Request.Form("s9")
id=Request.Form("id")
cantidad=Request.Form("cantidad")
codigop=Request.Form("codigop")
precio=Request.Form("precio")


Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\bryan y jesus\Documents\diseno TLAHUASHOP\COMPRAS.accdb.mdb"


if gor1= 1 then
conn.execute "INSERT INTO PEDIDOS(id,cantidad,codigop,precio) values(1,"& CInt(cantidad) &",'gor1',400)"
end if

if gor2= 1 then
conn.execute "INSERT INTO PEDIDOS(id,cantidad,codigop,precio) values(1,"& CInt(cantidad) &",'gor2',350)"
end if

if gor3= 1 then
conn.execute "INSERT INTO PEDIDOS(id,cantidad,codigop,precio) values(1,"& CInt(cantidad) &",'gor3',450)"
end if

if gor4= 1 then
conn.execute "INSERT INTO PEDIDOS(id,cantidad,codigop,precio) values(1,"& CInt(cantidad) &",'gor4',300)"
end if

if gor5= 1 then
conn.execute "INSERT INTO PEDIDOS(id,cantidad,codigop,precio) values(1,"& CInt(cantidad) &",'gor5',450)"
end if

if gor6= 1 then
conn.execute "INSERT INTO PEDIDOS(id,cantidad,codigop,precio) values(1,"& CInt(cantidad) &",'gor6',450)"
end if

if gor7= 1 then
conn.execute "INSERT INTO PEDIDOS(id,cantidad,codigop,precio) values(1,"& CInt(cantidad) &",'gor7',400)"
end if

if gor8= 1 then
conn.execute "INSERT INTO PEDIDOS(id,cantidad,codigop,precio) values(1,"& CInt(cantidad) &",'gor8',400)"
end if

if gor9= 1 then
conn.execute "INSERT INTO PEDIDOS(id,cantidad,codigop,precio) values(1,"& CInt(cantidad) &",'gor9',2713)"
end if

if pt1= 1 then
conn.execute "INSERT INTO PEDIDOS(id,cantidad,codigop,precio) values(1,"& CInt(cantidad) &",'pt1',600)"
end if

if pt2= 1 then
conn.execute "INSERT INTO PEDIDOS(id,cantidad,codigop,precio) values(1,"& CInt(cantidad) &",'pt2',600)"
end if

if pt3= 1 then
conn.execute "INSERT INTO PEDIDOS(id,cantidad,codigop,precio) values(1,"& CInt(cantidad) &",'pt3',580)"
end if

if pt4= 1 then
conn.execute "INSERT INTO PEDIDOS(id,cantidad,codigop,precio) values(1,"& CInt(cantidad) &",'pt4',650)"
end if

if pt5= 1 then
conn.execute "INSERT INTO PEDIDOS(id,cantidad,codigop,precio) values(1,"& CInt(cantidad) &",'pt5',600)"
end if

if pt6= 1 then
conn.execute "INSERT INTO PEDIDOS(id,cantidad,codigop,precio) values(1,"& CInt(cantidad) &",'pt6',650)"
end if

if pt7= 1 then
conn.execute "INSERT INTO PEDIDOS(id,cantidad,codigop,precio) values(1,"& CInt(cantidad) &",'pt7',650)"
end if

if pt8= 1 then
conn.execute "INSERT INTO PEDIDOS(id,cantidad,codigop,precio) values(1,"& CInt(cantidad) &",'pt8',700)"
end if

if pt9= 1 then
conn.execute "INSERT INTO PEDIDOS(id,cantidad,codigop,precio) values(1,"& CInt(cantidad) &",'pt9',700)"
end if

if pt10= 1 then
conn.execute "INSERT INTO PEDIDOS(id,cantidad,codigop,precio) values(1,"& CInt(cantidad) &",'pt10',750)"
end if

if pt11= 1 then
conn.execute "INSERT INTO PEDIDOS(id,cantidad,codigop,precio) values(1,"& CInt(cantidad) &",'pt11',500)"
end if

if c1= 1 then
conn.execute "INSERT INTO PEDIDOS(id,cantidad,codigop,precio) values(1,"& CInt(cantidad) &",'c1',350)"
end if

if c2= 1 then
conn.execute "INSERT INTO PEDIDOS(id,cantidad,codigop,precio) values(1,"& CInt(cantidad) &",'c2',350)"
end if

if c3= 1 then
conn.execute "INSERT INTO PEDIDOS(id,cantidad,codigop,precio) values(1,"& CInt(cantidad) &",'c3',350)"
end if

if c4= 1 then
conn.execute "INSERT INTO PEDIDOS(id,cantidad,codigop,precio) values(1,"& CInt(cantidad) &",'c4',300)"
end if

if c5= 1 then
conn.execute "INSERT INTO PEDIDOS(id,cantidad,codigop,precio) values(1,"& CInt(cantidad) &",'c5',400)"
end if

if c6= 1 then
conn.execute "INSERT INTO PEDIDOS(id,cantidad,codigop,precio) values(1,"& CInt(cantidad) &",'c6',400)"
end if

if c7= 1 then
conn.execute "INSERT INTO PEDIDOS(id,cantidad,codigop,precio) values(1,"& CInt(cantidad) &",'c7',400)"
end if

if c8= 1 then
conn.execute "INSERT INTO PEDIDOS(id,cantidad,codigop,precio) values(1,"& CInt(cantidad) &",'c8',400)"
end if

if c9= 1 then
conn.execute "INSERT INTO PEDIDOS(id,cantidad,codigop,precio) values(1,"& CInt(cantidad) &",'c9',350)"
end if

if c10= 1 then
conn.execute "INSERT INTO PEDIDOS(id,cantidad,codigop,precio) values(1,"& CInt(cantidad) &",'c10',400)"
end if

if c11= 1 then
conn.execute "INSERT INTO PEDIDOS(id,cantidad,codigop,precio) values(1,"& CInt(cantidad) &",'c11',300)"
end if

if c12= 1 then
conn.execute "INSERT INTO PEDIDOS(id,cantidad,codigop,precio) values(1,"& CInt(cantidad) &",'c12',400)"
end if

if p1= 1 then
conn.execute "INSERT INTO PEDIDOS(id,cantidad,codigop,precio) values(1,"& CInt(cantidad) &",'p1',400)"
end if

if p2= 1 then
conn.execute "INSERT INTO PEDIDOS(id,cantidad,codigop,precio) values(1,"& CInt(cantidad) &",'p2',400)"
end if

if p3= 1 then
conn.execute "INSERT INTO PEDIDOS(id,cantidad,codigop,precio) values(1,"& CInt(cantidad) &",'p3',450)"
end if


if p4= 1 then
conn.execute "INSERT INTO PEDIDOS(id,cantidad,codigop,precio) values(1,"& CInt(cantidad) &",'p4',400)"
end if

if p5= 1 then
conn.execute "INSERT INTO PEDIDOS(id,cantidad,codigop,precio) values(1,"& CInt(cantidad) &",'p5',350)"
end if

if s1= 1 then
conn.execute "INSERT INTO PEDIDOS(id,cantidad,codigop,precio) values(1,"& CInt(cantidad) &",'s1',450)"
end if

if s2= 1 then
conn.execute "INSERT INTO PEDIDOS(id,cantidad,codigop,precio) values(1,"& CInt(cantidad) &",'s2',500)"
end if

if s3= 1 then
conn.execute "INSERT INTO PEDIDOS(id,cantidad,codigop,precio) values(1,"& CInt(cantidad) &",'s3',450)"
end if

if s4= 1 then
conn.execute "INSERT INTO PEDIDOS(id,cantidad,codigop,precio) values(1,"& CInt(cantidad) &",'s4',400)"
end if

if s6= 1 then
conn.execute "INSERT INTO PEDIDOS(id,cantidad,codigop,precio) values(1,"& CInt(cantidad) &",'s5',450)"
end if

if s7= 1 then
conn.execute "INSERT INTO PEDIDOS(id,cantidad,codigop,precio) values(1,"& CInt(cantidad) &",'s6',500)"
end if

if s8= 1 then
conn.execute "INSERT INTO PEDIDOS(id,cantidad,codigop,precio) values(1,"& CInt(cantidad) &",'s7',400)"
end if

if s9= 1 then
conn.execute "INSERT INTO PEDIDOS(id,cantidad,codigop,precio) values(1,"& CInt(cantidad) &",'s8',500)"
end if













conn.close
set conn=nothing
response.redirect("TLAHUASHOP.html")
%>
</body>
</html>