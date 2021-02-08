<%
Response.Buffer  = true
Response.Expires = 0
Session.lcId     = 1033
%>
<!-- #include file="includes/conexao.asp" -->
<%
id   = Request.QueryString("id")

if (trim(id) = "") or (isnull(id)) then id = 0 end if

' Consiste o ID
if (cint(id) <> 0) then
			
	' Seleciona os dados do usuario
	strSQL = "SELECT * FROM users WHERE id = " & id
	
	' Executa a string sql.
	Set ObjRst = conDB.execute(strSQL)
			
	' Verifica se não é final de arquivo.	
	if not ObjRst.EOF then
				
		' Carrega as informações do usuario
		name = ObjRst("name")
		email     = ObjRst("email")

	end if
	
	set ObjRst = nothing
	
end if

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

    	<div class="starter">
    		<div class="row">
    			<div class="col-sm-6">
		    	    <h1>Formulário de Usuário</h1>
			        <div class="modal-body">
		    			<form role="form" id="db-form" name="db-form" method="post" action="cad_usuario.asp">
		    				<div class="form-group-attached">
		    					<div class="row">
	    							<div class="form-group">
	    								<label>Nome</label>
	    								<input type="text" name="name" id="name" class="form-control" placeholder="Informe o nome" value="<%=name%>" required>
	    							</div>
		    						
		    					</div>			
		    					<div class="row">
	    							<div class="form-group">
	    								<label>Email</label>
	    								<input type="email" name="email" id="email" class="form-control" placeholder="Informe o e-mail" value="<%=email%>" required>
	    							</div>
		    						
		    					</div>
		    				</div>
		    				<div class="row">
		    					<div class="form-group">
		    						<input type="hidden" name="id" id="id" value="<%=id%>">
		    						<button type="submit" class="btn btn-primary">Salvar</button>
		    						<a href="index.asp" class="btn btn-info">Voltar</a>
		    					</div>
		    				</div>
		    			</form>	    
		    		</div>
		    	</div>
		    </div>
    	</div>	  
    </div><!-- /.container -->

    <script src="js/jquery-3.1.1.min.js"></script>
    <script src="js/bootstrap.min.js"></script>

</html>
<%

conDB.close()

set conDB = Nothing

%>