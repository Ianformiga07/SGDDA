<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<!--#include file ="lib/Conexao.asp"-->
<!--#include file="base.asp"-->
<%
IF REQUEST("Operacao") = 2 THEN
	call abreConexao
	sql = "SELECT TB_Recipiente.CodigoItensModuloFace, TB_Modulos.Modulo, TB_Faces.Face, TB_Nivel.Nivel, TB_Pastas.Pasta, TB_CadDocumento.NumAuto, TB_CadDocumento.Serie, TB_CadDocumento.Observacao, TB_CadDocumento.Status FROM TB_Recipiente INNER JOIN TB_Modulos ON TB_Recipiente.CodigoModulo = TB_Modulos.Codigo INNER JOIN TB_Faces ON TB_Recipiente.CodigoFace = TB_Faces.CodigoFace INNER JOIN TB_Nivel ON TB_Recipiente.ID_Nivel = TB_Nivel.ID_Nivel INNER JOIN TB_Pastas ON TB_Recipiente.ID_Pasta = TB_Pastas.ID INNER JOIN TB_CadDocumento ON TB_CadDocumento.CodigoItensModuloFace = TB_Recipiente.CodigoItensModuloFace WHERE NumAuto = '"&request.form("txtAuto")&"' AND Serie = '"&request.form("txtSerie")&"';"
	
	'sql = "Select NumAuto, Serie FROM TB_CadDocumento WHERE NumAuto = '"&request.form("txtAuto")&"' AND Serie = '"&request.form("txtSerie")&"'"
	'response.write sql
	'response.end
	set rs = conn.execute(sql)
	if not rs.eof then
	NumAuto = rs("NumAuto")
	Serie = rs("Serie")
	statusUsuario = rs("Status")
	end if

END IF
%>
<script src="https://cdn.jsdelivr.net/npm/sweetalert2@10.10.1/dist/sweetalert2.all.min.js"></script>
<script src="javascript/Mascara.js"></script>
<script type="text/javascript">
/*var url_string = window.location.href;
'var url = new URL(url_string);
var resp = url.searchParams.post("Operacao");*/
var resp = "<%=request.form("Operacao")%>"
mensagem(resp)
function validar()
{
var NumAuto = frmBusca.txtAuto.value;
var Serie = frmBusca.txtSerie.value;  
  if(NumAuto == ""){
      Swal.fire({
      title: "Ops!!!",
      text: "Preencha o campo do Numero do Auto!",
      icon: "error",
      button: "Ok!",
      });
      frmBusca.txtAuto.focus()
      return false;
  } 
  if(Serie == ""){
      Swal.fire({
      title: "Ops!!!",
      text: "Preencha o campo da Série!",
      icon: "error",
      button: "Ok!",
      });
      frmBusca.txtSerie.focus()
      return false;
  } 

	return true;
}
function visualizar()
{
	if( document.frmBusca.txtAuto.value != "" && document.frmBusca.txtSerie.value != "")
	{
	document.frmBusca.Operacao.value = 2;
	document.frmBusca.action = "BuscarDocumentos.asp";
	document.frmBusca.submit();
	}
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
        <h6 class="m-0 font-weight-bold text-primary">Buscar Documentos</h6>
      </div>
      <div class="card-body">
		<form name="frmBusca" id="frmBusca" method="post" onSubmit="return validar()">
		<input type="hidden" id="Operacao" name="Operacao" />            
       <div class="form-group row g-3">
             <div class="col">
             <label for="txtAuto" class="col-form-label col-form-label-sm" >N° do Auto</label>
            <input type="text" name="txtAuto" id="txtAuto" size="13" class="form-control  col-form-sm" onKeyPress="somente_numero(txtAuto);" onBlur="somente_numero(txtAuto);" value="<%=(NumAuto)%>"/>
            </div>
         <div class="col">
            <label for="txtSerie" class="col-form-label col-form-label-sm" >Série</label>
            <input type="text" name="txtSerie" id="txtSerie" size="8" class="form-control  col-form-sm" onKeyPress="return ApenasLetras(event,this);" value="<%=(Serie)%>"/>
            </div>
            </div>
        <button class="btn btn-primary btn-icon-split" type="submit" onClick="visualizar();">
          <span class="icon text-white-100">
          </span>
          <span class="text">Pesquisar</span>
        </button>
 <%if request("Operacao") =  2 then%>

<div class="card">
              <div class="card-header">
                <h4 class="card-title">Tabela de Documentos</h4>
              </div>
              <div class="card-body">
                <div class="table-responsive">
                 <%if rs.eof then%>
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
                        Observações
                      </th>
                      <th>
                        Status
                      </th>
                    </thead>
                    <%do while not rs.eof%>
                    <tbody>
                      <tr>
                        <td>
                          <%=rs("Modulo")%>
                        </td>
                        <td>
                          <%=rs("Face")%>
                        </td>
                        <td>
                          <%=rs("Nivel")%>
                        </td>
                        <td>
                          <%=rs("Pasta")%>
                        </td>
                        <td>
                          <%=rs("NumAuto")%>
                        </td>
                        <td>
                          <%=rs("Serie")%>
                        </td>
                        <td>
                          <%=rs("Observacao")%>
                        </td>
                        <td>
                           <%if rs("Status") = true then%>
                          <font color="#009933"> ATIVO </font>
                          <%ELSE%>
  						  <font color="#FF0000"> INATIVO </font>
                          <%end if%>
                        </td>
                      </tr>
                    </tbody>
                     <%
     rs.movenext
	 loop
  %>
                  </table>
                  <%call fechaConexao
				  end if%>
                    <%end if%>
                </div>
              </div>
            </div>
        </form>
      </div>
    </div>
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
