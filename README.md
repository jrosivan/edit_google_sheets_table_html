# edit_google_sheets_table_html

Editar uma Planilha Google com HTML + Bootstrap_Edit, com busca e gravação.
Utilizando bootstrap_edit em uma tabela google.
  * https://examples.bootstrap-table.com/#issues/137.html
  * https://bootstrap-table.com/docs/extensions/editable/

Com recursos de colunas fixas + scroll + Edição das células, em formato tabela com todos os dias do mes.

Busca os dados por:
 * EMPRESA: Empresas cadastradas na aba [Cadastros] da planilha - "D" na coluna DESATIVADO, desativa a mesma;
 * DEPARTAMENTOS: Abaixo das EMPRESAS, na aba [Cadastros] da planilha - "D" na coluna DESATIVADO, desativa o mesmo;
 * ANO: Cadastrado na aba [Ano] da planilha. (visando a possibilidade de fazer uma lista futuramente);
 * MES: Disponível apenas o mes atual e o anterior. (evitar alterar dados de mais de 02 meses);

Ao efetuar a busca dos dados selecionados, caso não exista, retornar a tabela vazia da aba [Modelo];

Os dados são gravados na aba [Dados], disponibilizando para geração de Gráficos ou ETL;


Link para o Google Sheets:
https://docs.google.com/spreadsheets/d/1jrLUBDutPBsUcJc68l--nY1pXHt877ZTaPEPRdrt0Rg/edit?usp=sharing

