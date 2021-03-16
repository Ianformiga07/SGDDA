<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<!--#include file ="lib/Conexao.asp"-->
<!--#include file="base.asp"-->
<%
dim modulos, faces
if request("Operacao") = 5 and request("Codigo") <> "" then
dim CodigoModulo
 CodigoModulo = request.form("Codigo")
call abreConexao
 SQL2 = "SELECT * FROM TB_Recipiente INNER JOIN TB_Modulos ON TB_Modulos.Codigo = TB_Recipiente.CodigoModulo INNER JOIN TB_Faces ON TB_Faces.CodigoFace = TB_Recipiente.CodigoFace WHERE Codigo = '"&request.form("Codigo")&"'"
 set rs1 = conn.Execute(SQL2)
 session("CodigoItensModuloFace") = rs1("CodigoItensModuloFace")
end if
'Session("modulo") = "MODULO 01"
'Session("faces") = "FACE 1"

'if request("Operacao") = 6 then
'Session("modulo") = request.form("modulo")
'Session("faces") = request.form("faces")
'Session("niveis") = request.form("niveis")
'Session("pasta") = request.form("pasta")
'response.write(session("CodigoItensModuloFace"))
'response.end
'response.redirect("CadDocumentos.asp?Codigo="&session("CodigoItensModuloFace"))

'end if

%>

<%

IF REQUEST("Operacao") = 1 THEN 'Cadastrar
call abreConexao
sql = "SELECT CodigoItensModuloFace FROM TB_Recipiente WHERE CodigoModulo = '"&request.Form("modulo")&"' and Codigoface = '"&request.Form("faces")&"'"
set rs_ItensModulo = conn.execute(sql)
sql = "INSERT INTO TB_Recipiente(CodigoItensModuloFace, Nivel, NumeroPasta) VALUES ('"&rs_ItensModulo("CodigoItensModuloFace")&"', '"&request.Form("niveis")&"', '"&request.Form("pasta")&"')"
call FechaConexao
END IF
%>

<style type="text/css">
.messagem_informacao {
    background: none repeat scroll 0 0 #00FFCC;
    border: 2px solid #00FF33;
    top:  -55px;
    left: 300px;	
}
</style>
<script src="https://cdn.jsdelivr.net/npm/sweetalert2@10.10.1/dist/sweetalert2.all.min.js"></script>
<script >

function validar()
{
var faces = frmCadRecipiente.faces.value;
var nivel = frmCadRecipiente.niveis.value;  
var pasta = frmCadRecipiente.pasta.value; 
  if(faces == ""){
      Swal.fire({
      title: "Ops!!!",
      text: "Selecione a Face!",
      icon: "error",
      button: "Ok!",
      });
      frmCadRecipiente.faces.focus()
      return false;
  } 
  if(nivel == ""){
      Swal.fire({
      title: "Ops!!!",
      text: "Selecione o Nível!",
      icon: "error",
      button: "Ok!",
      });
      frmCadRecipiente.niveis.focus()
      return false;
  } 
    if(pasta == ""){
      Swal.fire({
      title: "Ops!!!",
      text: "Selecione a Pasta!",
      icon: "error",
      button: "Ok!",
      });
      frmCadRecipiente.pasta.focus()
      return false;
  } 

	return true;
}	
function cadastrar()
{
	if(validar() == false)
	return false;
	
	document.frmCadRecipiente.Operacao.value = 1;
	document.frmCadRecipiente.action = "CadRecipiente.asp";
	document.frmCadRecipiente.submit();
}
function visualizar(codigo)
{
	document.frmCadRecipiente.Operacao.value = 2;
	document.frmCadRecipiente.CodigoVisualizar.value = codigo;
	document.frmCadRecipiente.action = "CadRecipiente.asp";
	document.frmCadRecipiente.submit();
}

function verificar_cadastro()
{
	var codigo = document.getElementById("modulo").value
	
	document.frmCadRecipiente.Operacao.value = 5;
	document.frmCadRecipiente.Codigo.value = codigo;
	document.frmCadRecipiente.action = "CadRecipiente.asp";
	document.frmCadRecipiente.submit();
}
function somente_numero(campo){   
var digits="0123456789"   
var campo_temp   
    for (var i=0;i<campo.value.length;i++){   
        campo_temp=campo.value.substring(i,i+1)   
        if (digits.indexOf(campo_temp)==-1){   
            campo.value = campo.value.substring(0,i);   
        }   
    }   
} 

</script>


      <!-- End Navbar -->
      <div class="content">
        <div class="row">
        
<div class="col-lg-11">
    <!-- Select2 -->
    <div class="card mb-6">
      <div class="card-header py-3 d-flex flex-row align-items-center justify-content-between">
        <h6 class="m-0 font-weight-bold text-primary">Cadastrar Recipiente</h6>
      </div>
      <div class="card-body">
       <form name="frmCadRecipiente" id="frmCadRecipiente" onSubmit="return validar();" method="post">
		<input type="hidden" name="Operacao" id="Operacao" />
		<input type="hidden" name="CodigoVisualizar" id="CodigoVisualizar" />
		<input type="hidden" name="Codigo" id="Codigo" />             
       <div class="form-group row g-3">
       <div class="col">
        <%
	call abreConexao
	sql = "SELECT Codigo, Modulo FROM TB_Modulos ORDER BY Modulo"
	 set rs = conn.execute(sql)
	
 		%>
            <label for="modulo" class="col-form-label col-form-label-sm" >Módulo <%=modulos%></label>
            <select class="select2-single form-control col-form-label-sm" name="modulo" id="modulo" onChange="verificar_cadastro();" required> 
                <option value="" selected>Módulo</option>
                <%do while not rs.eof%>
				<option value="<%=rs("Codigo")%>"<%if trim(rs("Codigo")) = trim(request.form("modulo")) then%>selected<%end if%>><%=rs("Modulo")%></option>
  			  <% rs.movenext
						loop
				call fechaConexao%>
			</select>
            </div>
             <div class="col">
            <label for="faces" class="col-form-label col-form-label-sm" >Faces <%=faces%></label>
            <select class="select2-single form-control col-form-label-sm" name="faces" id="faces" required> 
                <option value="" selected>Faces </option>
<%do while not rs1.eof%>
	<option value="<%=trim(rs1("CodigoFace"))%>" <%if trim(rs1("CodigoFace")) = trim(request.form("faces")) then%>Selected<%end if%>><%=trim(rs1("Face"))%></option>
    <% rs1.movenext
			loop
%>
</select>
            </div>
            </div>
      <div class="form-group row g-3">
        <div class="col">
            <label for="niveis" class="col-form-label col-form-label-sm" >Niveis</label>
            <select class="select2-single form-control col-form-label-sm" name="niveis" id="niveis" required> 
                <option value="" selected>Niveis</option>
                <option value="NIVEL 01" <%if niveis = "NIVEL 01" then%> selected <%end if%>> NÍVEL 01 </option>
 			    <option value="NIVEL 02" <%if niveis = "NIVEL 02" then%> selected <%end if%>> NÍVEL 02 </option>
  			    <option value="NIVEL 03" <%if niveis = "NIVEL 03" then%> selected <%end if%>> NÍVEL 03 </option>
    			<option value="NIVEL 04" <%if niveis = "NIVEL 04" then%> selected <%end if%>> NÍVEL 04 </option>
  			    <option value="NIVEL 05" <%if niveis = "NIVEL 05" then%> selected <%end if%>> NÍVEL 05 </option>
    			<option value="NIVEL 06" <%if niveis = "NIVEL 06" then%> selected <%end if%>> NÍVEL 06 </option>
   			    <option value="NIVEL 07" <%if niveis = "NIVEL 07" then%> selected <%end if%>> NÍVEL 07 </option>
			</select>
        </div>
        <div class="col">
             <label for="pasta" class="col-form-label col-form-label-sm" >N° da Pasta</label>
            <input type="text" name="pasta" id="pasta" size="13" class="form-control  col-form-sm" onKeyPress="somente_numero(pasta);" onBlur="somente_numero(pasta);" value="<%=NumAuto%>"/>
            </div>
      </div>
        <button class="btn btn-primary btn-icon-split" onClick="return cadastrar();">
          <span class="icon text-white-50">
          </span>
          <span class="text">Avançar</span>
        </button>
        <%if existe = 1 then%>
          <button class="btn btn-primary btn-icon-split" onClick="return cadastrar();">
          <span class="icon text-white-50">
          </span>
          <span class="text">Documentos</span>
        </button>
        <%end if%>
        </form>
      </div>
    </div>
  </div>
      </div>
  </div>
 