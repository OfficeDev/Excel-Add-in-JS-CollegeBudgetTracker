# Exemplo de Controlador de Orçamento Escolar do Suplemento do Painel de Tarefas para o Excel 2016

_Aplica-se ao: Excel 2016_

Esse suplemento do painel de tarefas mostra como criar um controlador de orçamento escolar usando as APIs JavaScript no Excel 2016. Há dois tipos: o editor de código e o Visual Studio.

![Exemplo de Controlador de Orçamento Escolar](../images/CollegeBudgetTracker_tracker.PNG)

## Experimente
### Versão do editor de código

A maneira mais fácil de implantar e testar o suplemento é copiar o manifesto em um compartilhamento de rede.

1.  Crie uma pasta em um compartilhamento de rede (por exemplo, \\\MyShare\CollegeBudgetTracker).  
2.  Copie o manifesto (CollegeBudgetTrackerManifest.xml) para um compartilhamento de rede (por exemplo, \\\MyShare\MyManifests).
3.  Adicione o local de compartilhamento que contém o manifesto como um catálogo de aplicativos confiáveis no Excel.

    a. Inicie o Excel e abra uma planilha em branco.  
    
    b. Escolha a guia **Arquivo** e escolha **Opções**.
    
    c. Escolha **Central de Confiabilidade** e, em seguida, escolha o botão **Configurações da Central de Confiabilidade**.
    
    d. Escolha **Catálogos de Suplementos Confiáveis**.
    
    e. Na caixa  **URL de Catálogo**, insira o caminho para o compartilhamento de rede que você criou na etapa 3 e escolha **Adicionar Catálogo**.
    
   f.  Marque a caixa de seleção **Mostrar no Menu** e escolha **OK**. Será exibida uma mensagem para informá-lo de que suas configurações serão aplicadas na próxima vez que você iniciar o Office. 
        
4.  Teste e execute o suplemento. 

    a. Na **guia Inserir** no Excel 2016, escolha **Meus Suplementos**. 
    
    b. Na caixa de diálogo **Suplementos do Office**, escolha **Pasta Compartilhada**.
    
    c. Clique no comando **Controlador de Orçamento Escolar** na guia Página Inicial. O suplemento abre um painel de tarefas e cria o controlador de orçamento escolar na planilha ativa, conforme mostrado neste diagrama. 
      
   ![Exemplo de Controlador de Orçamento Escolar](../images/CollegeBudgetTracker_tracker.PNG) 

    d. Adicione algumas despesas e o rendimento usando as guias **Adicionar despesas** e **Adicionar rendimento** e veja como os dados e os gráficos mudam dinamicamente.
    
      ![Amostra de Controlador de Orçamento Escolar](../images/CollegeBudgetTracker_taskpane1.PNG) 

Para usar o manifesto em seu próprio Suplemento, edite o elemento <SourceLocation> do arquivo de manifesto de modo que ele aponte para o arquivo HTML no local de compartilhamento do arquivo Home.html.
    
### Versão do Visual Studio
1.  Copie o projeto para uma pasta local e abra o Excel-Add-in-JS-CollegeBudgetTracker.sln no Visual Studio.
2.  Pressione F5 para criar e implantar o suplemento de exemplo. O Excel inicia e o suplemento abre em um painel de tarefas à direita da planilha em branco, conforme mostrado na figura a seguir. 
        
  ![Exemplo de Controlador de Orçamento Escolar](../images/CollegeBudgetTracker_tracker.PNG) 

3.  Adicione algumas despesas e o rendimento usando as guias **Adicionar despesas** e **Adicionar rendimento** e veja como os dados e os gráficos mudam dinamicamente.

  ![Exemplo de Controlador de Orçamento Escolar](../images/CollegeBudgetTracker_taskpane1.PNG) 


### Saiba mais

As APIs JavaScript para Excel têm muito mais a oferecer à medida que você desenvolve suplementos. Confira a seguir alguns dos recursos disponíveis. 

1.  [Visão geral da programação de Suplementos do Excel](https://github.com/OfficeDev/office-js-docs/blob/master/excel/excel-add-ins-programming-overview.md)
2.  [Explorador de trecho para Excel](http://officesnippetexplorer.azurewebsites.net/#/snippets/excel)
3.  [Exemplos de código de Suplementos do Excel](https://github.com/OfficeDev/office-js-docs/blob/master/excel/excel-add-ins-code-samples.md) 
4.  [Referência da API JavaScript para Suplementos do Excel](https://github.com/OfficeDev/office-js-docs/blob/master/excel/excel-add-ins-javascript-reference.md)
5.  [Crie seu primeiro Suplemento do Excel](https://github.com/OfficeDev/office-js-docs/blob/master/excel/build-your-first-excel-add-in.md)
