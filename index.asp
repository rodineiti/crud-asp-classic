<%
Response.Buffer  = true
Response.Expires = 0
Session.lcId     = 1033
%>
<!-- #include file="includes/conexao.asp" -->
<%

strStatus = Request.Item("strStatus")
strMsg = ""

select case trim(ucase(strStatus))
  case "INC"
    strMsg = "Cadastro realizado com Sucesso"
  case "ALT"
    strMsg = "Cadastro atualizado com Sucesso"
  case "EXC"
    strMsg = "Cadastro deletado com Sucesso"
  case else
    strMsg = ""
end select

%>
<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="utf-8">
  <meta http-equiv="X-UA-Compatible" content="IE=edge">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>Tutorial CRUD ASP Clássico</title>
  <link href="css/bootstrap.min.css" rel="stylesheet">
  <style>
    body {
      padding-top: 50px;
    }
    .starter-template {
      padding: 40px 15px;
      text-align: center;
    }
  </style>
</head>
<body>

  <nav class="navbar navbar-inverse navbar-fixed-top">
    <div class="container">
      <div class="navbar-header">
        <button type="button" class="navbar-toggle collapsed" data-toggle="collapse" data-target="#navbar" aria-expanded="false" aria-controls="navbar">
          <span class="sr-only">Toggle navigation</span>
          <span class="icon-bar"></span>
          <span class="icon-bar"></span>
          <span class="icon-bar"></span>
        </button>
        <a class="navbar-brand" href="index.asp">CRUD ASP</a>
      </div>
    </div>
  </nav>

  <div class="container">
    
    <% if trim(strMsg) <> "" then %>
    <div class="alert alert-success">
      <a href="#" class="close" data-dismiss="alert" aria-label="close" title="close">×</a>
      <%=strMsg%>      
    </div>
    <% end if %>

    <div class="starter-template">
      <h1>Lista de Usuários</h1>
      <p align="left">
        <a href="frm_usuario.asp?id=0" class="btn btn-primary btn-cons" alt="Incluir Cadastro" title="Incluir Cadastro"><i class="glyphicon glyphicon-plus"></i> Adicionar</a>
      </p>

      <table class="table table-bordered"> 
        <thead>
          <tr>
            <th>#</th>
            <th>Nome</th>
            <th>E-mail</th>
            <th>Ação</th>
          </tr>
        </thead>
        <tbody>

          <%

          strSQL = "select * from users order by name asc;"

          set ObjRst = conDB.execute(strSQL)

          do while not ObjRst.EOF

            %>
            <tr>
              <td><%=ObjRst("id")%></td>
              <td><%=ObjRst("name")%></td>
              <td><%=ObjRst("email")%></td>
              <td>
                <a href="frm_usuario.asp?id=<%=ObjRst("id")%>" class="btn btn-success" alt="Editar Cadastro" title="Editar Cadastro"><i class="glyphicon glyphicon-pencil"></i></a>

                <a data-href="exc_usuario.asp?id=<%=ObjRst("id")%>" class="btn btn-danger" data-toggle="modal" data-target="#confirm-delete" alt="Excluir Cadastro" title="Excluir Cadastro"><i class="glyphicon glyphicon-remove"></i></a>
              </td>
            </tr>
            <%

            ObjRst.MoveNext()

          loop

          set ObjRst = Nothing

          %>

        </tbody>
      </table>
    </div>

    <!-- MODAL Exclusão-->
    <div class="modal fade stick-up" id="confirm-delete" tabindex="-1" role="dialog" aria-labelledby="myModalLabel" aria-hidden="true">
      <div class="modal-dialog">
        <div class="modal-content">
          <div class="modal-header clearfix text-left">
            <button type="button" class="close" data-dismiss="modal" aria-hidden="true"><i class="pg-close fs-14"></i>
            </button>
            <h5>Confirmação <span class="semi-bold">de Exclusão</span></h5>
          </div>
          <div class="modal-body">
            <!--<p class="debug-url"></p>-->
            <p>Confirmar a exclusão do usuário?</p>
          </div>                
          <div class="modal-footer">
            <button type="button" class="btn btn-default" data-dismiss="modal">Cancelar</button>
            <a class="btn btn-danger btn-deletar">Deletar</a>
          </div>
        </div>
      </div>
    </div>

  </div><!-- /.container -->

  <script src="js/jquery-3.1.1.min.js"></script>
  <script src="js/bootstrap.min.js"></script>
  <script>
    $(function()
    { 
      $('#confirm-delete').on('show.bs.modal', function(e) {
        $(this).find('.btn-deletar').attr('href', $(e.relatedTarget).data('href'));
        //$('.debug-url').html('Delete URL: <strong>' + $(this).find('.btn-deletar').attr('href') + '</strong>');
      });
    });
  </script>
</html>
<%

conDB.close()

set conDB = Nothing

%>
