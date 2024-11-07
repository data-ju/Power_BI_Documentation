**Projeto de Documentação Automática do Power BI**

Esse projeto tem o objetivo de gerar a documentação técnica de um relatório do Power BI.

Existem duas versões do código, uma utilizando apenas os arquivos JSON do relatório e outra também adicionando IA Generativa ao projeto. 
  Eles estão acessíveis pelas pastas "Python_Version" e "IA_Version".

**Para a versão que utiliza IA, é necessário ter acesso a Azure Open AI, e ter em mãos a chave de API e endpoint.**
  Se você participa de uma organização, você pode pedir essas informações para seu administrador de TI.
  Caso seja uma pessoa física, você pode criar uma nova conta da Azure (https://azure.microsoft.com/pt-br/) e utilizar os créditos disponíveis por 30 dias para testar. 
  Alternativamente, você também pode utilizar outros serviços gratuitos de IA Generativa e inclusive rodar localmente, para evitar vazamento de dados e manter a segurança de dados sensíveis.
  **Atenção: Cuide de seus dados e  pesquise sobre os modelos antes de realizar testes com dados sensíveis, para evitar retreinar modelos com dados privados.**

É necessário, antes de executar o código:

- Transformar seu arquivo .pbix em .pbit, para fazer isso, basta abrir seu arquivo no Power BI Desktop, ir em Arquivo > Exportar > Modelo do Power BI;
- Ter o Python baixado na sua máquina e, opcionalmente, o VSCode ou outro editor de código (utilizei o Jupyter Notebook);
- Instalar e importar as bibliotecas necessárias (baixe o arquivo requirements para poder importar quando rodar o código);
- Baixar o arquivo modelo em word ou utilizar o seu próprio (configurando o código conforme seu próprio modelo).
  
O resultado final será um arquivo Word com as informações de Páginas, Tabelas, Colunas, Medidas, Fontes e Relacionamentos de tabelas do projeto.
Caso utilize o código com IA, o resultado final será o mesmo, mas com as descrições adicionadas.

Espero que goste do projeto e que ele te auxilie em sua jornada de inteligência de dados.

Qualquer dúvida, pode me acionar que te auxiliarei.

**Atenção: Antes de executar o código, troque os caminhos das variáveis no arquivo config.py !**
