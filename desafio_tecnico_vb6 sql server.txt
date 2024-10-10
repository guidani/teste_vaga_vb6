Desafio Técnico

Contexto da Empresa:
Você foi contratado pela XYZ Administradora de Cartões de Crédito para desenvolver uma série de funcionalidades em VB6(Aceitável VB.NET) e SQL Server, que visam melhorar o processo de gerenciamento de transações financeiras e relatórios. A empresa lida com um alto volume de transações diárias e precisa de uma solução eficiente para gerenciar essas informações.

Desafio:

1. CRUD em VB6(ou VB.NET):
     - Crie uma aplicação em VB6 para gerenciar um cadastro de transações de cartão de crédito.
     - O CRUD deve permitir:
     - Inserir uma nova transação, com os campos: Id_Transacao, Numero_Cartao, Valor_Transacao, Data_Transacao, Descricao.
     - Editar uma transação existente.
     - Excluir uma transação.
     - Consultar transações com filtros opcionais por Numero_Cartao, Data_Transacao e Valor_Transacao.
     - Utilize uma interface simples com caixas de texto, botões de ação e uma DataGrid para exibir os resultados das consultas.

2. Stored Procedure no SQL Server:
   - Crie uma stored procedure que calcule o total de transações para um dado período, agrupado por Numero_Cartao.
   - A procedure deve receber dois parâmetros: Data_Inicial e Data_Final.
   - Retorne o Numero_Cartao, o total de transações (Valor_Total) e a quantidade de transações (Quantidade_Transacoes) para cada cartão no período.

3. Function no SQL Server:
     - Crie uma função que receba um valor e retorne a categoria da transação:
     - Se Valor_Transacao for maior que 1000, a categoria é "Alta".
     - Se Valor_Transacao estiver entre 500 e 1000, a categoria é "Média".
     - Se Valor_Transacao for menor que 500, a categoria é "Baixa".
     - Utilize essa função em uma consulta para categorizar todas as transações existentes.

4. View no SQL Server:
     - Crie uma view que combine informações das tabelas Transacoes e Clientes.
     - A view deve retornar o Nome_Cliente, Numero_Cartao, Valor_Transacao, Data_Transacao e a categoria da transação utilizando a função criada anteriormente.

5. Exportação de Relatório em Excel:
     - Na aplicação VB6(ou VB.NET), adicione uma funcionalidade que permita exportar as transações do último mês para um arquivo Excel.
     - O relatório deve conter as colunas: Numero_Cartao, Valor_Transacao, Data_Transacao, Descricao, e Categoria.
     - O arquivo Excel deve ser salvo em um local específico, informado pelo usuário através de uma caixa de diálogo.

Requisitos de Avaliação:
- Qualidade do código.
- Eficiência das queries SQL.
- Facilidade de uso e clareza da interface VB6.
- Correta utilização dos recursos de Excel para geração de relatórios.
- Capacidade de manipular grandes volumes de dados de maneira eficiente.

Entregáveis:
- Código-fonte da aplicação VB6. (subir no github)
- Script SQL com a criação das tabelas, stored procedure, function e view. (Incluso script de inserção de dados de clientes, por mais que pra esses não seja necessário o CRUD, pedimos o script de inserção de dados.)
- Um exemplo de relatório gerado em Excel.