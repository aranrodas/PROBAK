<%@ Language=VBScript %>
<%Response.Expires=-1%>
<html>
<!--#include file="../seguridad.asp"-->
<!--#include file="../idioma.asp"-->
<!--#include file="menuIpScore.asp"-->
<head>
<link rel="stylesheet" type="text/css" href="../estilo.css">
<link rel="stylesheet" href="/jquery-ui/themes/ui-ekhi/jquery-ui-1.11.4.custom.css">

<script src="/jquery-ui/jquery-1.10.2.js"></script>
<script src="/jquery-ui/ui/jquery.ui.core.js"></script>
<script src="/jquery-ui/ui/jquery.ui.widget.js"></script>
<script src="/jquery-ui/ui/jquery.ui.dialog.js"></script>
<script src="/jquery-ui/ui/jquery.ui.button.js"></script>
<script src="/jquery-ui/ui/jquery.ui.position.js"></script>
<script src="/jquery-ui/js/jquery-ui-1.10.4.custom.min.js"></script>

<script type="text/javascript" src="../libraries/amcharts_3.19.4/amcharts/amcharts.js"></script>
<script type="text/javascript" src="../libraries/amcharts_3.19.4/amcharts/serial.js"></script>
<script type="text/javascript" src="../libraries/amcharts_3.19.4/amcharts/themes/light.js"></script>

<meta name="VI60_defaultClientScript" content="JavaScript">
<meta http-equiv="Content-Type" content="text/html; charset=ISO-8859-1" />
<meta charset="utf-8">

<style>
label.txt{ position: relative; display: inline-block; width: 100px; text-align:right;vertical-align: top;}
input.txt{ float:left; width: 35px; }

#Dialoganotacion{display:none;}
#DialogVerAnotacion{display:none;}
#Dialoginfo{display:none;}

.estructura td{
	width: 450px;
}

.colorAmarillo{
	background-color: #ffffb3;
}

.colorAzul{
	background-color: #66b3ff;
}

border-right:1px silver solid;
 border-bottom:1px silver solid;
 text-align:center

.bordeIzq{
	border-left:1px silver solid;
}

.bordeDer{
	border-right:1px white solid;
}

.nomLib{
	border-right:1px silver solid;
	border-bottom:1px silver solid;
	text-align: center
}

.valor{	
	border-top:1px silver solid;
	border-bottom:1px silver solid;
	text-align: center;
}

.valorTotal{
	border-right:1px silver solid;
	border-bottom:1px silver solid;
	text-align:right;
}
</style>
<script LANGUAGE="javascript">
	<!--
	function importe_adm(idic){
		thisForm.op.value="importe_adm";
		thisForm.valor1.value=idic;
		thisForm.submit();
	}
	
	function addanotacion(idic){
		$( "#Dialoganotacion" ).show();
			var Dialoganotacion = $('#Dialoganotacion').dialog({
				modal: true,
				autoOpen: true,
				title: 'ANOTACIONES',
				width: 480,
				height: 350,
				position: { my: 'center', at: 'center', of: window },
				show:{effect:"clip", duration:500},
				hide:{effect:"clip", duration:500},
				buttons: {
					'Guardar': function() {
						thisForm.op.value="addanotacion";
						thisForm.valor1.value=idic;
						thisForm.valor2.value=$("#addnota").val();
						thisForm.submit();
					},
					'Cancelar': function() {
						$(this).dialog( "close" );
						$( "#Dialoganotacion" ).hide();
					}
				},
				close: function() {
					$(this).dialog( "close" );
					$( "#Dialoganotacion" ).hide();
				}
			});	
	}
	
	function veranotacion(idic, anotacion){
		$("#nota").text(anotacion);
		
		$( "#DialogVerAnotacion" ).show();
			var DialogVerAnotacion = $('#DialogVerAnotacion').dialog({
				modal: true,
				autoOpen: true,
				title: 'ANOTACIONES',
				width: 480,
				height: 350,
				position: { my: 'center', at: 'center', of: window },
				show:{effect:"clip", duration:500},
				hide:{effect:"clip", duration:500},
				buttons: {
					'Guardar cambios': function() {
						thisForm.op.value="addanotacion";
						thisForm.valor1.value=idic;
						thisForm.valor2.value=$("#nota").val();
						thisForm.submit();
					},
					'Cancelar': function() {
						$(this).dialog( "close" );
						$( "#DialogVerAnotacion" ).hide();
					}
				},
				close: function() {
					$(this).dialog( "close" );
					$( "#DialogVerAnotacion" ).hide();
				}
			});	
	}
	
	function informacion(){
		$( "#Dialoginfo" ).show();
			var Dialoganotacion = $('#Dialoginfo').dialog({
				modal: true,
				autoOpen: true,
				title: 'INFORMACIÓN',
				width: 450,
				height: 200,
				position: { my: 'center', at: 'center', of: window },
				show:{effect:"clip", duration:500},
				hide:{effect:"clip", duration:500},
				buttons: {
					'Aceptar': function() {
						$(this).dialog( "close" );
						$( "#Dialoginfo" ).hide();
					}
				},
				close: function() {
					$(this).dialog( "close" );
					$( "#Dialoginfo" ).hide();
				}
			});	
	}
	//-->
</script>
</head>


<%
'Abrir Conexión
set conn=server.CreateObject("ADODB.Connection")
conn.Open Application("gespro_ConnectionString")
set rs=server.CreateObject("ADODB.Recordset")
set rs2=server.CreateObject("ADODB.Recordset")
set cmd=server.CreateObject("ADODB.Command")
cmd.ActiveConnection=conn
cmd.CommandType = 4

anio=request("selanio")
if anio="" then anio=request("anio")

function aDecPunto(num)
	aDecPunto = replace(replace(round(num, 2), ".", ""), ",", ".")
end function

if request("op")="importe_adm" then
	idic=request("valor1")
	'response.write idic
	strsql="select idic,nombre, dbo.ipscoreValorSDK(idic) as total_sdk from ipscoreIC where anio="&anio&" and idic="&idic 
	rs.open strsql,conn,adopenstatic
	'response.write " input: " &request("total_sdk"&rs("idic")) &" BD :"&rs("total_sdk")
	if replace(request("total_sdk"&rs("idic")),",",".")<>replace(rs("total_sdk"),",",".") then
		if request("total_sdk"&rs("idic"))="-1" then 
			conn.execute "update ipscoreIC set valor_adm= NULL where idic="&rs("idic")&"and anio="&anio
		else
			conn.execute "update ipscoreIC set valor_adm= "&replace(request("total_sdk"&rs("idic")),",",".")& " where idic="&rs("idic")&"and anio="&anio
		end if
	end if
	rs.close
end if

if request("op")="addanotacion" then
	idic=request("valor1")
	'response.write idic
	nota=request("valor2")
	conn.execute "update ipscoreIC set anotacion='"&nota&"' where idic="&idic&" and anio="&anio
end if

	' if request("op")="" then
	' cmd.CommandText=""
	' cmd.Parameters(1).Value=request("")
	' set rs=cmd.Execute
	' X=rs("resultado")
	' rs.close	
' end if

' if request("op")="" then	
	' cmd.CommandText=""
	' cmd.Parameters(1).Value=request("")
	' cmd.Execute
	' set cmd=nothing
' end if

%>

<body class="bodymargen"><FORM method=POST id=thisForm name=thisForm>
<input type="hidden" id="op" name="op">	
<input type="hidden" id="valor1" name="valor1">
<input type="hidden" id="valor2" name="valor2">
<input type="hidden" id="valor3" name="valor3">
<input type="hidden" id="anio" name="anio" value=<%=anio%>>
<%
if seguridad(0)>0 then	
%>
<ul class="breadcrumb">
		<li><a href="seleccionar.asp">Cambiar Año</a></li><li>|</li>
		<li><strong><%=anio%></strong></li>
</ul>
<br><br>
<%'Menu
response.write verMenu ("resultados", anio)%>
<br><br>
<table class=estructura cellpadding=0 cellspacing=5 width="100%">
	<tr>
		<td  width="33%" valign=top>
			<!--TABLA 1: FACTURACIÓN-->
			<table style="max-width:98%; min-width:98%" border=0 cellpadding=2 cellspacing=0>
				<tr><td colspan="3" bgcolor="  #ff6666" style="text-align:center"><font size=3><b>DATOS DE FACTURACIÓN</b></font></td></tr>
				<tr><th colspan= "3">Grupo 1 (Consolidado) </th></tr>
				<tr>
					<td class="valor bordeDer bordeIzq colorAzul"><b>Librería</b></td>
					<td class="valor bordeDer colorAzul"><b>Peso (%)</b></td>
					<td class="valor colorAzul"><b>Valor (€)</b></td>
				</tr>
				<%strsql="select sum(dbo.ipscoreValorSDK(idic)) total from ipscoreIC where anio="&anio&""
				rs.open strsql,conn,adopenstatic
				total_valor=rs("total")
				rs.close				
				strsql="select idic,nombre,dbo.ipscoreValorSDK(idic) as total_sdk from ipscoreIC where anio="&anio&" and grupo='Consolidado'"
				rs.open strsql,conn,adopenstatic
				total_Consolidado_valor=0
				total_Consolidado_peso=0
				while not rs.eof%>
				<tr>
					<td class="nomLib bordeIzq"><%=rs("nombre")%></td>
					<td class="valorTotal"> <%=numero((rs("total_sdk")/total_valor)*100)%>% </td>
					<td class="valorTotal"><%=numero(rs("total_sdk"))%> </td>
				</tr>
				<%total_Consolidado_valor=total_Consolidado_valor + rs("total_sdk")
				total_Consolidado_peso=total_Consolidado_peso + (rs("total_sdk")/total_valor)
				rs.movenext
				wend
				rs.close
				%>
				<tr>
					<td class="colorAmarillo nomLib" style="border-left:1px silver solid; "><b>TOTAL GRUPO 1</b></td>
					<td class="valorTotal colorAmarillo"><b><%=numero(total_Consolidado_peso*100)%>%</b></td>
					<td class="valorTotal colorAmarillo"><b><%=numero(total_Consolidado_valor)%></b></td>
				</tr>
				<tr><th colspan= "3">Grupo 2 (Emergente) </th></tr>
				<tr>
					<td class="valor bordeDer bordeIzq colorAzul"><b>Librería</b></td>
					<td class="valor bordeDer colorAzul"><b>Peso (%)</b></td>
					<td class="valor colorAzul"><b>Valor (€)</b></td>
				</tr>
				<%strsql="select idic,nombre,dbo.ipscoreValorSDK(idic) as total_sdk from ipscoreIC where anio="&anio&" and grupo='Emergente'"
				rs.open strsql,conn,adopenstatic
				total_Emergente_valor=0
				total_Emergente_peso=0
				while not rs.eof%>
				<tr>
					<td class="nomLib bordeIzq"><%=rs("nombre")%></td>
					<td class="valorTotal"><%=numero((rs("total_sdk")/total_valor)*100)%>%</td>
					<td class="valorTotal"> <%=numero(rs("total_sdk"))%> </td>
				</tr>
				<% total_Emergente_valor=total_Emergente_valor + rs("total_sdk")
				total_Emergente_peso=total_Emergente_peso + (rs("total_sdk")/total_valor)
				rs.movenext
				wend
				rs.close
				
				total_peso= total_Consolidado_peso + total_Emergente_peso
				%>
				<tr>
					<td class="colorAmarillo nomLib bordeIzq"><b>TOTAL GRUPO 2</b></td>
					<td class="valorTotal colorAmarillo"><b><%=numero(total_Emergente_peso*100)%>%</b></td>
					<td class="valorTotal colorAmarillo"><b><%=numero(total_Emergente_valor)%></b></td>
				</tr>
				<tr class=cajetintotales>
					<td class="nomLib bordeIzq"><b>TOTAL</b></td>
					<td class="valorTotal"><b><%=numero(total_peso*100)%>%</b></td>
					<td class="valorTotal"><b><%=numero(total_valor)%></td>
				</tr>
			</table>
		</td>
		<td width="33%" valign=top>
			<!--TABLA 2:DATOS ADMINISTRACIÓN-->
			<table style="max-width:98%; min-width:98%" border=0 cellpadding=2 cellspacing=0>
				<tr><td colspan="3" bgcolor=" #bb99ff" style="text-align:center"><font size=3><b>TOTAL APLICABLE AL IPSCORE</b></font></td></tr>
				<tr><th colspan= "3">Grupo 1 (Consolidado) </th></tr>
				<tr>
					<td class="valor bordeDer bordeIzq colorAzul"><b>Librería</b></td>
					<td class="valor bordeDer colorAzul"><b>Peso (%)</b></td>
					<td class="valor colorAzul"><b>Valor (€)</b>&nbsp;&nbsp;&nbsp;&nbsp; <img src="../images/information.png" width="12" height="12" onclick="informacion();"></td>
				</tr>
				<%strsql="select sum(isnull(valor_adm,dbo.ipscoreValorSDK(idic))) total from ipscoreIC where anio="&anio&""
				rs.open strsql,conn,adopenstatic
				total_adm=rs("total")
				rs.close
				
				strsql="select idic,nombre,anotacion, isnull(valor_adm,dbo.ipscoreValorSDK(idic)) as total_sdk from ipscoreIC where anio="&anio&" and grupo='Consolidado'"
				rs.open strsql,conn,adopenstatic
				
				total_Consolidado_valor=0
				total_Consolidado_peso=0
				while not rs.eof%>
				<tr>
					<td class="nomLib bordeIzq"><%=rs("nombre")%></td>
					<td class="valorTotal"> <%=numero((rs("total_sdk")/total_adm)*100)%>% </td>
					<td style="border-right:1px silver solid;border-bottom:1px silver solid; text-align:center">
						<input type=text id="total_sdk<%=rs("idic")%>" name="total_sdk<%=rs("idic")%>" value="<%=rs("total_sdk")%>" size=10 style="text-align:right" onkeyup="validanumero3(total_sdk<%=rs("idic")%>);">
					<%if isnull(rs("anotacion")) or rs("anotacion")="" then%>
						<img class="tooltip" title="Añadir anotación" src="../images/add-note-edit.png" id="Anotacion<%=rs("idic")%>" width="12" height="12" onclick="addanotacion(<%=rs("idic")%>);">&nbsp;
					<%else%>
						<img class="tooltip" title="Ver anotación" src="../images/note-edit.png" id="Anotacion<%=rs("idic")%>" width="12" height="12" onclick="veranotacion(<%=rs("idic")%>,'<%=rs("anotacion")%>');">&nbsp;
					<%end if%>
						
						<img class="tooltip" title="Guardar cambios" src="../images/save.gif" id="Guardar<%=rs("idic")%>" width="12" height="12" onclick="importe_adm(<%=rs("idic")%>);">
					</td>
				</tr>
				<%total_Consolidado_valor=total_Consolidado_valor + rs("total_sdk")	
				total_Consolidado_peso=total_Consolidado_peso + (rs("total_sdk")/total_adm)
				rs.movenext
				wend
				rs.close
				%>
				<tr>
					<td class="colorAmarillo nomLib bordeIzq"><b>TOTAL GRUPO 1</b></td>
					<td class="valorTotal colorAmarillo"><b><%=numero(total_Consolidado_peso*100)%>%</b></td>
					<td class="valorTotal colorAmarillo"><b><%=numero(total_Consolidado_valor)%></b></td>
				</tr>
				<tr><th colspan= "3">Grupo 2 (Emergente) </th></tr>
				<tr>
					<td class="valor bordeDer bordeIzq colorAzul"><b>Librería</b></td>
					<td class="valor bordeDer colorAzul"><b>Peso (%)</b></td>
					<td class="valor colorAzul"><b>Valor (€)</b> &nbsp;&nbsp;&nbsp;&nbsp;<img src="../images/information.png" width="12" height="12" onclick="informacion();"></td>
				</tr>
				<%strsql="select idic,nombre,anotacion, isnull(valor_adm,dbo.ipscoreValorSDK(idic)) as total_sdk from ipscoreIC where anio="&anio&" and grupo='Emergente'"
				rs.open strsql,conn,adopenstatic
				total_Emergente_valor=0
				total_Emergente_peso=0
				while not rs.eof%>
				<tr>
					<td class="nomLib bordeIzq"><%=rs("nombre")%></td>
					<td class="valorTotal"><%=numero((rs("total_sdk")/total_adm)*100)%>%</td>
					<td style="border-right:1px silver solid;border-bottom:1px silver solid;text-align: center">
						<input type=text id="total_sdk<%=rs("idic")%>" name="total_sdk<%=rs("idic")%>" value="<%=rs("total_sdk")%>" size=10 style="text-align:right" onkeyup="validanumero3(total_sdk<%=rs("idic")%>);"> 
					<%if isnull(rs("anotacion"))or rs("anotacion")="" then%>
						<img class="tooltip" title="Añadir anotación" src="../images/add-note-edit.png" id="Anotacion<%=rs("idic")%>" width="12" height="12" onclick="addanotacion(<%=rs("idic")%>);">&nbsp;
					<%else%>
						<img class="tooltip" title="Ver anotación" src="../images/note-edit.png" id="Anotacion<%=rs("idic")%>" width="12" height="12" onclick="veranotacion(<%=rs("idic")%>,'<%=rs("anotacion")%>');">&nbsp;
					<%end if%>
						<img class="tooltip" title="Guardar cambios" src="../images/save.gif" id="Guardar<%=rs("idic")%>" width="12" height="12" onclick="importe_adm(<%=rs("idic")%>);">
					</td>
				</tr>
				<% total_Emergente_valor=total_Emergente_valor + rs("total_sdk")
				total_Emergente_peso=total_Emergente_peso + (rs("total_sdk")/total_adm)
				rs.movenext
				wend
				rs.close
				total_peso= total_Consolidado_peso + total_Emergente_peso
				%>
				<tr>
					<td class="colorAmarillo nomLib bordeIzq"><b>TOTAL GRUPO 2</b></td>
					<td class="valorTotal colorAmarillo"><b><%=numero(total_Emergente_peso*100)%>%</b></td>
					<td class="valorTotal colorAmarillo"><b><%=numero(total_Emergente_valor)%></b></td>
				</tr>
				<tr class=cajetintotales>
					<td class="nomLib bordeIzq"><b>TOTAL</b></td>
					<td class="valorTotal"><b><%=numero(total_peso*100)%>%</b></td>
					<td class="valorTotal"><b><%=numero(total_adm)%></td>
				</tr>
				<tr>
					<td colspan="2" bgcolor="#ffb3b3" style="border-left:1px  silver solid;border-right:1px  #ff9999 solid; border-bottom:1px silver solid;text-align:center"><b>DIFERENCIA<b></td>
					<td bgcolor="#ffb3b3" class="valorTotal"><b><%=numero(total_valor - total_adm)%><b></td>
				</tr>
			</table>
		</td>
		<td width="33%" valign=top>
			<!--TABLA 3:DATOS ADMINISTRACIÓN + DATOS IPSCORE-->
			<table align="center" style="max-width:98%; min-width:98%" border=0 cellpadding=2 cellspacing=0 >
				<tr><td colspan="3" bgcolor="#ff9933" style="text-align:center"><font size=3><b>DATOS IPSCORE</b></font></td></tr>
				<tr><th colspan= "3">Grupo 1 (Consolidado) </th></tr>
				<tr>
					<td class="valor bordeDer bordeIzq colorAzul"><b>Librería</b></td>
					<td class="valor bordeDer colorAzul"><b>Peso (%)</b></td>
					<td class="valor colorAzul"><b>Valor (€)</b></td>
				</tr>
				<%'Una vez realizado el calculo del VAN, sacar datos desde la BD
				strsql="select van,grupo from ipscorePatentes where anio="&anio&" order by grupo"
				rs.open strsql,conn,adopenstatic
				if not rs.eof then
					consolidado_VAN=rs("van")
					'consolidado_VAN=replace(rs("van"),",",".")
					rs.movenext
					emergente_VAN=rs("van")
					'emergente_VAN=replace(rs("van"),",",".")
				end if
				rs.close
				total_ipscore= consolidado_VAN + emergente_VAN
				
				'Obtener datos de año anterior para grafica
				rs.open "SELECT min(anio) as anio FROM ipscoreIC", conn, adopenstatic
				if not rs.eof then
					min = rs("anio")
				else
					min = "2015"
				end if
				rs.close
				
				anterior =anio-1 	
				if anterior>=min then 
					total_Consolidado_valor2=0
					total_Emergente_valor2=0
					consolidado_VAN2 = 0
					emergente_VAN2 = 0
					strsql="select idic,nombre,anotacion, isnull(valor_adm,dbo.ipscoreValorSDK(idic)) as total_sdk from ipscoreIC where anio="&anterior&" and grupo='Consolidado'"
					rs.open strsql,conn,adopenstatic
					while not rs.eof 
						total_Consolidado_valor2=total_Consolidado_valor2 + rs("total_sdk")
						rs.movenext
					wend
					rs.close
					
					strsql="select idic,nombre,isnull(valor_adm,dbo.ipscoreValorSDK(idic)) as total_sdk from ipscoreIC where anio="&anterior&" and grupo='Emergente'"
					rs.open strsql,conn,adopenstatic
					while not rs.eof 
						total_Emergente_valor2=total_Emergente_valor2 + rs("total_sdk")
						rs.movenext
					wend
					rs.close
					
					strsql="select van,grupo from ipscorePatentes where anio="&anterior&" order by grupo"
					rs.open strsql,conn,adopenstatic
					if not rs.eof then
						consolidado_VAN2=rs("van")
						rs.movenext
						emergente_VAN2= rs("van")
					end if
					rs.close
				end if
				strsql="select idic,nombre,isnull(valor_adm,dbo.ipscoreValorSDK(idic)) as total_sdk from ipscoreIC where anio="&anio&" and grupo='Consolidado' order by nombre"
				rs.open strsql,conn,adopenstatic
				'total_Consolidado_valor=0
				'total_Consolidado_peso=0
				jsData1 = ""	'Variable para crear graficos (peso)
				jsData2 = ""	'Variable para crear graficos (valor)
				while not rs.eof
				%>
				<tr>
					<td style="border-right:1px silver solid;border-bottom:1px silver solid;border-left:1px silver solid;text-align: center"><%=rs("nombre")%></td>
					<td class="valorTotal"> <%=numero((rs("total_sdk")/total_Consolidado_valor)*100)%>% </td>
					
					<td style="border-right:1px silver solid;border-bottom:1px silver solid;text-align:right"><%=numero((rs("total_sdk")/total_Consolidado_valor)* consolidado_VAN) %></td>
				</tr>
				<%totalConPeso=totalConPeso + (rs("total_sdk")/total_Consolidado_valor)
				rs.movenext
				wend
				rs.close
				%>
				<tr>
					<td class="colorAmarillo nomLib bordeIzq"><b>TOTAL GRUPO 1</b></td>
					<td class="valorTotal colorAmarillo"><b><%=numero(totalConPeso*100)%>%</b></td>
					<td class="valorTotal colorAmarillo"><b><%=numero(consolidado_VAN)%></b></td>
				</tr>
				<tr><th colspan= "3">Grupo 2 (Emergente) </th></tr>
				<tr>
					<td class="valor bordeDer bordeIzq colorAzul"><b>Librería</b></td>
					<td class="valor bordeDer colorAzul"><b>Peso (%)</b></td>
					<td class="valor colorAzul"><b>Valor (€)</b></td>
				</tr>
				<%				
				strsql="select idic,nombre,isnull(valor_adm,dbo.ipscoreValorSDK(idic)) as total_sdk from ipscoreIC where anio="&anio&" and grupo='Emergente' order by nombre"
				rs.open strsql,conn,adopenstatic
				while not rs.eof
				%>
				<tr>
					<td class="nomLib bordeIzq"><%=rs("nombre")%></td>
					<td class="valorTotal"><%=numero((rs("total_sdk")/total_Emergente_valor)*100)%>%</td>
					<td class="valorTotal"><%=numero(emergente_VAN *rs("total_sdk")/total_Emergente_valor)%></td>
				</tr>
				<%
				totalEmgPeso=totalEmgPeso + (rs("total_sdk")/total_Emergente_valor)
				rs.movenext
				wend
				rs.close
				
				'VALORES Y PESOS DEL GRUPO 1 (para las graficas)
				strsql = "select idic,nombre,anio, isnull(valor_adm,dbo.ipscoreValorSDK(idic)) as total_sdk from ipscoreIC where (anio="&anio&" or anio='"&anterior&"') and grupo='Consolidado' order by nombre, anio"
				rs.open strsql, conn, adopenstatic
				libr = ""
				while not rs.eof
					if libr="" or libr<>ucase(rs("nombre")) then
						if libr<>ucase(rs("nombre")) and not libr="" then
							if not right(jsData2, 2)="}," then 
								jsData1 = jsData1 & ", ""valor2"":0},"
								jsData2 = jsData2 & ", ""valor2"":0},"
							end if
						end if
						jsData1 = jsData1 & "{"
						jsData1 = jsData1 & """libreria"":" & " """&rs("nombre")&""","
						jsData2 = jsData2 & "{"
						jsData2 = jsData2 & """libreria"":" & " """&rs("nombre")&""","
						if cint(rs("anio"))=cint(anterior) then
							jsData1 = jsData1 & """valor1"":" &aDecPunto(numero((rs("total_sdk")/total_Consolidado_valor2)*100))'&", "
							jsData2 = jsData2 & """valor1"":" &aDecPunto(numero((rs("total_sdk")/total_Consolidado_valor2)*consolidado_VAN2))'&", "
						else
							jsData1 = jsData1 & """valor1"":0, "
							jsData1 = jsData1 &	"""valor2"":" &aDecPunto(numero((rs("total_sdk")/total_Consolidado_valor)*100))&"},"

							jsData2 = jsData2 & """valor1"":0, "
							jsData2 = jsData2 &	"""valor2"":" &aDecPunto(numero((rs("total_sdk")/total_Consolidado_valor)* consolidado_VAN))&"},"
						end if
					else
						jsData1 = jsData1 & ", ""valor2"":" &aDecPunto(numero((rs("total_sdk")/total_Consolidado_valor)*100))&"},"
						jsData2 = jsData2 & ", ""valor2"":" &aDecPunto(numero((rs("total_sdk")/total_Consolidado_valor)* consolidado_VAN))&"},"
					end if
					libr = ucase(rs("nombre"))
					rs.movenext
				wend
				if not right(jsData2, 2)="}," then 
					jsData1 = jsData1 & ", ""valor2"":0},"
					jsData2 = jsData2 & ", ""valor2"":0},"
				end if
				rs.close
				
				'VALORES Y PESOS DEL GRUPO 2 (para las graficas) (misma estructura que arriba cambiando los calculos)
				strsql = "select idic,nombre,anio, isnull(valor_adm,dbo.ipscoreValorSDK(idic)) as total_sdk from ipscoreIC where (anio='"&anio&"' or anio='"&anterior&"') and grupo='Emergente' order by nombre, anio"
				rs.open strsql, conn, adopenstatic
				libr = ""
				while not rs.eof
					if libr="" or libr<>ucase(rs("nombre")) then
						if libr<>ucase(rs("nombre")) and not libr="" then
							if not right(jsData2, 2)="}," then 
								jsData1 = jsData1 & ", ""valor2"":0},"
								jsData2 = jsData2 & ", ""valor2"":0},"
							end if
						end if
						jsData1 = jsData1 & "{"
						jsData1 = jsData1 & """libreria"":" & " """&rs("nombre")&""","
						jsData2 = jsData2 & "{"
						jsData2 = jsData2 & """libreria"":" & " """&rs("nombre")&""","
						if cint(rs("anio"))=cint(anterior) then
							jsData1 = jsData1 & """valor1"":" &aDecPunto(numero((rs("total_sdk")/total_Emergente_valor2)*100))'&", "
							jsData2 = jsData2 & """valor1"":" &aDecPunto(numero(emergente_VAN2*rs("total_sdk")/total_Emergente_valor2))'&", "
						else
							jsData1 = jsData1 & """valor1"":0, "
							jsData1 = jsData1 &	"""valor2"":" &aDecPunto(numero((rs("total_sdk")/total_Emergente_valor)*100))&"},"

							jsData2 = jsData2 & """valor1"":0, "
							jsData2 = jsData2 &	"""valor2"":" &aDecPunto(numero(emergente_VAN*rs("total_sdk")/total_Emergente_valor))&"},"
						end if
					else
						jsData1 = jsData1 & ", ""valor2"":" &aDecPunto(numero((rs("total_sdk")/total_Emergente_valor)*100))&"},"
						jsData2 = jsData2 & ", ""valor2"":" &aDecPunto(numero(emergente_VAN*rs("total_sdk")/total_Emergente_valor))&"},"
					end if
					libr = ucase(rs("nombre"))
					rs.movenext
				wend
				if not right(jsData2, 2)="}," then 
					jsData1 = jsData1 & ", ""valor2"":0},"
					jsData2 = jsData2 & ", ""valor2"":0},"
				end if
				rs.close
				%>
				<tr>
					<td class="colorAmarillo nomLib bordeIzq"><b>TOTAL GRUPO 2</b></td>
					<td class="valorTotal colorAmarillo"><b><%=numero(totalEmgPeso*100)%>%</b></td>
					<td class="valorTotal colorAmarillo"><b><%=numero(emergente_VAN)%></b></td>
				</tr>
				<tr class=cajetintotales>
					<td class="nomLib bordeIzq bordeDer"><b>TOTAL</b></td>
					<td class="nomLib bordeDer"><b>TOTAL</b></td>
					<td class="valorTotal"><b><%=numero(total_ipscore)%></td>
				</tr>
			</table>
		</td>
	</tr>
</table>
	<script>
	
	var chart = AmCharts.makeChart( "chart", {
		"type": "serial",
		"theme": "light",
		"legend": {
			"enabled": true,
			"useGraphSettings": true
		},
		"dataProvider": [<%=jsData1%>],
		"synchronizeGrid":true,
		"valueAxes": [ {
			"axisColor": "#FF6600",
			"axisThickness": 2,
			"axisAlpha": 1,
			"position": "left"
		},],
		"graphs": [ {
			"valueAxis": "v1",
			"lineColor": "#80ccff",
			"bullet": "round",
			"bulletBorderThickness": 1,
			"hideBulletsCount": 30,
			"title": "<%=anterior%>",
			"valueField": "valor1",
			"fillAlphas": 0,
			"balloonText": "[[category]](<%=anterior%>): <b>[[value]]</b>"
		},{
			"valueAxis": "v2",
			"lineColor": "#006bb3",
			"bullet": "round",
			"bulletBorderThickness": 1,
			"hideBulletsCount": 30,
			"title": "<%=anio%>",
			"valueField": "valor2",
			"fillAlphas": 0,
			"balloonText": "[[category]](<%=anio%>): <b>[[value]]</b>"
		}],
		"categoryField": "libreria",
		"categoryAxis": {
			"axisColor": "#DADADA",
			"minorGridEnabled": true,
			"autoWrap": true
		},
		"export": {
			"enabled": true
		}
	} );
	
	var chart2 = AmCharts.makeChart("chart2", {
	"type": "serial",
    "theme": "light",
	"categoryField": "libreria",
	"dataProvider": [<%=jsData2%>],
	"startDuration": 1,
	"categoryAxis": {
		"gridPosition": "start",
		"position": "bottom",
		"autoWrap": true
	},
	"trendLines": [],
	"graphs": [
	{
		"balloonText": "[[category]](<%=anterior%>):[[value]]",
		"fillAlphas": 0.8,
		"id": "AmGraph-1",
		"lineAlpha": 0.2,
		"title": "<%=anterior%>",
		"type": "column",
		"valueField": "valor1"
	},
	{
		"balloonText": "[[category]](<%=anio%>):[[value]]",
		"fillAlphas": 0.8,
		"id": "AmGraph-2",
		"lineAlpha": 0.2,
		"title": "<%=anio%>",
		"type": "column",
		"valueField": "valor2"
	}],
	"guides": [],
	"valueAxes": [
		{
			"id": "ValueAxis-1",
			"axisAlpha": 0
		}
	],
	"allLabels": [],
	"balloon": {},
	"titles": [],
    "export": {
    	"enabled": true
     }
	
});
	</script>
<%

%>	

<table cellspacing=0 cellpadding=0 align="center" width="90%">
	<th style="font-size:18px;">Peso IPSCORE</th>
	<tr><td><div id="chart" style="width:100%;height:500px;font-size:11px;"></div></td>
	</tr>
</table>

<table cellspacing=0 cellpadding=0 align="center" width="90%">
	<th style="font-size:18px;">Valor IPSCORE</th>
	<tr><td><div id="chart2" style="width:100%;height:500px;font-size:11px;"></div></td>
	</tr>
</table>

<br><br>
<%

' note=request("value3")
%>
	<div id="Dialoganotacion"><br><textarea id="addnota" name="addnota"  cols=60 rows=10></textarea><br></div>
	<div id="DialogVerAnotacion"><br><textarea id="nota" name="nota"  cols=60 rows=10></textarea><br></div>
	<div id="Dialoginfo"><p>Puedes modificar los valores correspondientes a las librerías de software. Si quieres recurrir al valor calculado de una librería de software, escribe en el cuadrante correspondiente el valor -1.<br></p>
	<p>Para guardar cualquier cambio haga click en el disquete. </p>
	</div>
<%
else
	nopermisos()
end if 'seguridad
	
set rs=nothing
conn.Close
set conn=nothing
set cmd=nothing
%>
</FORM></body>
</html>

