<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<!--#include file ="lib/Conexao.asp"-->
<!--#include file="base.asp"-->
<%
'response.write(session("CodigoItensModuloFace"))
'response.end

dim cont
IF REQUEST("Operacao") = 1 THEN
call abreConexao
sql = "INSERT INTO TB_CadDocumento(CodigoItensModuloFace, NumAuto, Serie, Observacao, Status) VALUES('"&clng(request.form("recipiente"))&"', '"&request.Form("txtNumAuto")&"', '"&request.Form("txtNumSerie")&"', '"&request.form("txtObs")&"', 1)"
conn.execute(sql)
call fechaConexao
'response.Redirect("cadDocumentos.asp?Resp=1")

ELSEIF REQUEST("Operacao") = 2 THEN ' visualizar
call abreConexao
sql = "SELECT NumAuto, Serie, Observacao, Status FROM TB_CadDocumento WHERE NumAuto = "&request.form("AutoVisualizar")&""
set rs = conn.execute(sql)
	if not rs.eof then 
	NumAuto = rs("NumAuto")
	Serie = rs("Serie")
	Observacao = rs("Observacao")
	statusUsuario = rs("Status")
	Existe = 1
	else
	NumAuto = request.form("AutoVisualizar")
	Serie = request.Form("txtNumSerie")
	Existe = 0
	end if 
call fechaConexao

ELSEIF REQUEST("Operacao") = 3 THEN 'ALTERAR
	call abreConexao
	sql = "UPDATE TB_CadDocumento SET Serie = '"&request.Form("txtNumSerie")&"', Observacao = '"&request.Form("txtObs")&"', Status = '"&request.Form("status")&"' WHERE NumAuto = '"&request.form("txtNumAuto")&"'"
	conn.execute(sql)
	call fechaConexao
	'Response.Redirect("cadDocumentos.asp?Resp=2")
END IF
%>
<script src="https://cdn.jsdelivr.net/npm/sweetalert2@10.10.1/dist/sweetalert2.all.min.js"></script>
<script src="javascript/Mascara.js"></script>
<script type="text/javascript">
var resp = "<%=request.form("operacao")%>"
mensagem(resp)
function mensagem(resp) {
if (resp == "1"){
     Swal.fire({
      title: "Ótimo!!!",
      text: "Documento Cadastrado com Sucesso!",
      icon: "success",
      button: "Ok!",
      });
      return false;
}
if (resp == "3"){
     Swal.fire({
      title: "Ótimo!!!",
      text: "Documento Alterado com Sucesso!",
      icon: "success",
      button: "Ok!",
      });
      return false;
}
}

function validar()
{
var NumAuto = frmCadDocumento.txtNumAuto.value;
var Serie = frmCadDocumento.txtNumSerie.value; 
var recipiente = frmCadDocumento.recipiente.value;
 if(recipiente == ""){
      Swal.fire({
      title: "Ops!!!",
      text: "Selecione o Recipiente!",
      icon: "error",
      button: "Ok!",
      });
      frmCadDocumento.recipiente.focus()
      return false;
  }  
  if(NumAuto == ""){
      Swal.fire({
      title: "Ops!!!",
      text: "Preencha o campo do Numero do Auto!",
      icon: "error",
      button: "Ok!",
      });
      frmCadDocumento.txtNumAuto.focus()
      return false;
  } 
  if(Serie == ""){
      Swal.fire({
      title: "Ops!!!",
      text: "Preencha o campo da Série!",
      icon: "error",
      button: "Ok!",
      });
      frmCadDocumento.txtNumSerie.focus()
      return false;
  } 

	return true;
}
function cadastrar()
{
	
	if(validar() == false)
	  return false;
	document.frmCadDocumento.Operacao.value = 1;
	document.frmCadDocumento.action = "CadDocumentos.asp";
	document.frmCadDocumento.submit();
}

function visualizar(auto)
{
	document.frmCadDocumento.Operacao.value = 2;
	document.frmCadDocumento.AutoVisualizar.value = auto;
	document.frmCadDocumento.action = "CadDocumentos.asp";
	document.frmCadDocumento.submit();
}
function verificar_cadastro()
{
	if( document.frmCadDocumento.txtNumAuto.value != "" && document.frmCadDocumento.txtNumSerie.value != "")
	{
	document.frmCadDocumento.Operacao.value = 2;
	document.frmCadDocumento.AutoVisualizar.value = document.frmCadDocumento.txtNumAuto.value;
	document.frmCadDocumento.action = "CadDocumentos.asp";
	document.frmCadDocumento.submit();
	}
}

function ConsultarModulo()
{
	document.frmCadDocumento.Operacao.value = 4;
	document.frmCadDocumento.action = "CadDocumentos.asp";
	document.frmCadDocumento.submit();
}
function alterar()
{
	if(validar() == false)
	return false;
	
	document.frmCadDocumento.Operacao.value = 3;
	document.frmCadDocumento.action = "CadDocumentos.asp";
	document.frmCadDocumento.submit();
}


function ApenasLetras(e, t) {
    try {
        if (window.event) {
            var charCode = window.event.keyCode;
        } else if (e) {
            var charCode = e.which;
        } else {
            return true;
        }
        if (
            (charCode > 64 && charCode < 91) || 
            (charCode > 96 && charCode < 123) ||
            (charCode > 191 && charCode <= 255) // letras com acentos
        ){
            return true;
        } else {
            return false;
        }
    } catch (err) {
        alert(err.Description);
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
        <h6 class="m-0 font-weight-bold text-primary">Cadastrar Documentos</h6>
      </div>
      <div class="col-7 col-md-8">
      <div class="card-body">
		<form name="frmCadDocumento" id="frmCadDocumento" method="post">
        <input type="hidden" id="Operacao" name="Operacao" />
        <input type="hidden" id="AutoVisualizar" name="AutoVisualizar" />
        <div class="form-group row g-3">
        <div class="col">
        <%
	call abreConexao
	sql = "SELECT TB_Recipiente.CodigoItensModuloFace, TB_Modulos.Modulo, TB_Faces.Face, TB_Nivel.Nivel, TB_Pastas.Pasta FROM TB_Recipiente INNER JOIN TB_Modulos ON TB_Recipiente.CodigoModulo = TB_Modulos.Codigo INNER JOIN TB_Faces ON TB_Recipiente.CodigoFace = TB_Faces.CodigoFace INNER JOIN TB_Nivel ON TB_Recipiente.ID_Nivel = TB_Nivel.ID_Nivel INNER JOIN TB_Pastas ON TB_Recipiente.ID_Pasta = TB_Pastas.ID"
	 set rs_ItensModulo = conn.execute(sql)
 		%>
         <label for="recipiente" class="col-form-label col-form-label-sm" >Recipientes</label>
        <select class="select2-single form-control col-form-label-sm" name="recipiente" id="recipiente" onChange="return ConsultarModulo();" > 
                <option value="" selected>Selecione</option>
                <%do while not rs_ItensModulo.eof%>
				<option value="<%=rs_ItensModulo("CodigoItensModuloFace")%>" <%if clng(request.form("recipiente")) = clng(rs_ItensModulo("CodigoItensModuloFace")) then%> selected <%end if%>><%=rs_ItensModulo("CodigoItensModuloFace")%> - <%=rs_ItensModulo("Modulo")%> - <%=rs_ItensModulo("Face")%> - <%=rs_ItensModulo("Nivel")%> - Pasta <%=rs_ItensModulo("CodigoItensModuloFace")%></option>
  			  <% rs_ItensModulo.movenext
				 loop
				call fechaConexao%>
			</select>
            </div>      
            </div>  
       <div class="form-group row g-3">
             <div class="col">
             <label for="txtNumAuto" class="col-form-label col-form-label-sm" >N° do Auto</label>
            <input type="text" name="txtNumAuto"  maxlength="6" id="txtNumAuto" size="13" class="form-control  col-form-sm" onKeyPress="somente_numero(txtNumAuto);" onBlur="verificar_cadastro(); somente_numero(txtNumAuto);" value="<%=NumAuto%>"/>
            </div>
             <div class="col">
            <label for="txtNumSerie" class="col-form-label col-form-label-sm" >Série</label>
            <input type="text" name="txtNumSerie" id="txtNumSerie" size="8" class="form-control  col-form-sm" onKeyPress="return ApenasLetras(event,this);" value="<%=Serie%>" onBlur="verificar_cadastro();"/>
            </div>
            </div>
      <div class="form-group row g-3">
                    <div class="col">
                        <label>Observação</label>
                        <textarea class="form-control textarea" id="txtObs" name="txtObs"><%=Observacao%></textarea>
                  </div>
      </div>
      <%if Existe = 1 then%>
     <div class="form-group row g-3">
		<div class="col">
            <label for="status" class="col-form-label col-form-label-sm" >Status</label>
            <select class="select2-single form-control col-form-label-sm" name="status" id="status" required> 
                <option value="" selected>Status</option>
                <option value="1" <%if statusUsuario = true then %> selected <%end if%>> Ativo </option>
 			    <option value="0" <%if statusUsuario = false then %> selected <%end if%>> Inativo </option>
			</select>
        </div>
        </div>
        <%end if%>
      
        <button class="btn btn-primary btn-icon-split" value="" onClick="return <%IF Existe = 1 THEN%>alterar();<%ELSE%>cadastrar();<%END IF%>" >
          <span class="icon text-white-50">
          </span>
          <span class="text"><%IF Existe = 1 THEN%>Alterar<%ELSE%>Cadastrar<%END IF%></span>
        </button>
        </form>
      </div>
    </div>
  </div>
      </div>
  </div>
  <%if request("Operacao") = 4 OR REQUEST("Operacao") = 2 or REQUEST("Operacao") = 1 or REQUEST("Operacao") = 3 then%>
  <%
  call abreConexao
	sql = "SELECT TB_Recipiente.CodigoItensModuloFace, TB_Modulos.Modulo, TB_Faces.Face, TB_Nivel.Nivel, TB_Pastas.Pasta, TB_CadDocumento.NumAuto, TB_CadDocumento.Serie, TB_CadDocumento.Status FROM TB_Recipiente INNER JOIN TB_Modulos ON TB_Recipiente.CodigoModulo = TB_Modulos.Codigo INNER JOIN TB_Faces ON TB_Recipiente.CodigoFace = TB_Faces.CodigoFace INNER JOIN TB_Nivel ON TB_Recipiente.ID_Nivel = TB_Nivel.ID_Nivel INNER JOIN TB_Pastas ON TB_Recipiente.ID_Pasta = TB_Pastas.ID INNER JOIN TB_CadDocumento ON TB_CadDocumento.CodigoItensModuloFace = TB_Recipiente.CodigoItensModuloFace WHERE TB_Recipiente.CodigoItensModuloFace = "&request("recipiente")&""
	set rs2 = conn.execute(sql)
  %>
  <div class="card">
              <div class="card-header">
                <h4 class="card-title">Tabela de Documentos</h4>
              </div>
              <div class="card-body">
                <div class="table-responsive">
                 <%if rs2.eof then%>
  <tr><td>Não Existe Nenhum Registro na base de Dados!</td></tr>
  <%else%>
                  <table class="table">
                    <thead class=" text-primary">
                      <th>
                        Módulo
                      </th>
                      <th>
                        Face
                      </th>
                      <th>
                        Nível
                      </th>
                      <th>
                        Número Pasta
                      </th>
                       <th>
                        Número Auto
                      </th>
                       <th>
                        Série
                      </th>
                      <th>
                        Status
                      </th>
                      <th>
                        Opções
                      </th>
                    </thead>
                    <%do while not rs2.eof
					cont =cont+1
					%>
                    <tbody>
                      <tr>
                        <td>
                          <%=rs2("Modulo")%>
                        </td>
                        <td>
                          <%=rs2("Face")%>
                        </td>
                        <td>
                          <%=rs2("Nivel")%>
                        </td>
                        <td>
                          <%=rs2("Pasta")%>
                        </td>
                        <td>
                          <%=rs2("NumAuto")%>
                        </td>
                        <td>
                          <%=rs2("Serie")%>
                        </td>
                        <td>
                        <%if rs2("Status") = true then%>
                          <font color="#009933"> ATIVO </font>
                          <%ELSE%>
  						  <font color="#FF0000"> INATIVO </font>
                          <%end if%>
                        </td>
                        <td>
                          &nbsp&nbsp <a href="#" onClick="visualizar(<%=rs2("NumAuto")%>)"><img src="Imagens\olho.png" width="30"/></a>
                        </td>
                      </tr>
                    </tbody>
                     <%
     rs2.movenext
	 loop
  %>
                  </table>
                  <br><p align="center">Total de Documentos cadastrados nessa Pasta: <%=cont%></p>
                  <%call fechaConexao
				  end if%>
                  <%end if%>
                </div>
              </div>
            </div>
  <!--   Core JS Files   -->
  <script src="../assets/js/core/jquery.min.js"></script>
  <script src="../assets/js/core/popper.min.js"></script>
  <script src="../assets/js/core/bootstrap.min.js"></script>
  <script src="../assets/js/plugins/perfect-scrollbar.jquery.min.js"></script>
  <!--  Google Maps Plugin    -->
  <script src="https://maps.googleapis.com/maps/api/js?key=YOUR_KEY_HERE"></script>
  <!-- Chart JS -->
  <script src="../assets/js/plugins/chartjs.min.js"></script>
  <!--  Notifications Plugin    -->
  <script src="../assets/js/plugins/bootstrap-notify.js"></script>
  <!-- Control Center for Now Ui Dashboard: parallax effects, scripts for the example pages etc -->
  <script src="../assets/js/paper-dashboard.min.js?v=2.0.1" type="text/javascript"></script><!-- Paper Dashboard DEMO methods, don't include it in your project! -->
  <script src="../assets/demo/demo.js"></script>
