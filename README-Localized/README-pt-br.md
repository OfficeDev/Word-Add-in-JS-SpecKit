---
page_type: sample
products:
- office-word
- office-365
languages:
- javascript
extensions:
  contentType: samples
  technologies:
  - Add-ins
  createdDate: 3/24/2016 12:45:01 PM
---
# <a name="word-add-in-javascript-speckit"></a>Suplemento do Word JavaScript SpecKit

Saiba como você pode criar um suplemento que captura e insere texto clichê e como você pode implementar um processo de validação do documento simples.

## <a name="table-of-contents"></a>Sumário
* [Histórico de Alterações](#change-history)
* [Pré-requisitos](#prerequisites)
* [Configurar o projeto](#configure-the-project)
* [Executar o projeto](#run-the-project)
* [Compreender o código](#understand-the-code)
* [Perguntas e comentários](#questions-and-comments)
* [Recursos adicionais](#additional-resources)

## <a name="change-history"></a>Histórico de alterações

31 de março de 2016:
* Versão inicial do exemplo.

## <a name="prerequisites"></a>Pré-requisitos

* Word 2016 para Windows, build 16.0.6727.1000 ou superior.
* [Nó e npm](https://nodejs.org/en/)
* [Git Bash](https://git-scm.com/downloads) – você deve usar uma versão mais recente, já que as versões anteriores podem mostrar um erro ao gerar os certificados.

## <a name="configure-the-project"></a>Configurar o projeto

Execute os seguintes comandos do shell do Bash na raiz do projeto:

1. Clone este repositório em seu computador local.
2. ```npm install``` para instalar todas as dependências em package.json.
3. ```bash gen-cert.sh``` para criar os certificados necessários para executar este exemplo. Em seguida, no repositório em seu computador local, clique duas vezes em ca.crt e escolha **Instalar Certificado**. Escolha **Computador Local** e **Avançar** para continuar. Selecione **Colocar todos os certificados no armazenamento a seguir** e escolha **Procurar**.  Escolha **Autoridades de Certificação Raiz Confiáveis** e marque **OK**. Escolha **Avançar** e **Concluir** e o certificado de autoridade de certificação é adicionado ao armazenamento de certificados.
4. ```npm start``` para iniciar o serviço.

Nesse momento, você implantou esse suplemento de exemplo. Agora, você precisa informar ao Microsoft Word onde encontrar o suplemento.

1. Crie um compartilhamento de rede ou [compartilhe uma pasta para a rede](https://technet.microsoft.com/en-us/library/cc770880.aspx) e coloque o arquivo de manifesto [word-add-in-javascript-speckit-manifest.xml](word-add-in-javascript-speckit-manifest.xml) nele.
3. Inicie o Word e abra um documento.
4. Escolha a guia **Arquivo** e escolha **Opções**.
5. Escolha **Central de Confiabilidade**, e escolha o botão **Configurações da Central de Confiabilidade**.
6. Escolha **Catálogos de Suplementos Confiáveis**.
7. No campo **URL do Catálogo**, digite o caminho de rede para o compartilhamento de pasta que contém word-add-in-javascript-speckit-manifest.xml e escolha **Adicionar Catálogo**.
8. Selecione a caixa de seleção **Mostrar no Menu** e escolha **OK**.
9. Será exibida uma mensagem para informá-lo de que suas configurações serão aplicadas na próxima vez que você iniciar o Microsoft Office. Feche e reinicie o Word.

## <a name="run-the-project"></a>Executar o projeto

1. Abra um documento do Word.
2. Na guia **Inserir** no Word 2016, escolha **Meus Suplementos**.
3. Selecione a guia **Pasta compartilhada**.
4. Escolha **suplemento Word SpecKit** e **OK**.
5. Se os comandos de suplemento forem compatíveis com sua versão do Word, a interface do usuário informará que o suplemento foi carregado.

### <a name="ribbon-ui"></a>Faixa de Opções da Interface do Usuário
Na Faixa de Opções, você pode:
* Escolha a guia **suplemento SpecKit** para iniciar o suplemento na interface do usuário.
* Escolha **Inserir modelo de especificações** para iniciar o painel de tarefas e inserir um modelo de especificações no documento.
* Use os botões de validação na faixa de opções ou clique com botão direito do mouse no menu de contexto para validar o documento em uma lista de bloqueio de palavras.

 > Observação: O suplemento será carregado no painel de tarefas se os comandos de suplemento não forem compatíveis com sua versão do Word.

### <a name="task-pane-ui"></a>Interface do usuário do painel de tarefas
No painel de tarefas, você pode:
* Salve uma frase colocando o cursor na frase, nomeando-a no campo acima de **Adicionar frase ao texto clichê* no painel de tarefas e escolha **Adicionar frase ao texto clichê**. Você pode fazer o mesmo com parágrafos.
* Salvar frases e parágrafos também disponibilizará o texto clichê no menu suspenso **Inserir texto clichê**.
* Posicione o cursor do mouse no documento. Escolha um texto clichê do menu suspenso e o texto clichê será inserido no documento.
* Para alterar a propriedade *Author* do documento, altere o nome do autor e selecione o botão **Atualizar nome do autor**. Isso atualizará a propriedade do documento e o conteúdo de um controle de conteúdo associado.

## <a name="understand-the-code"></a>Compreender o código

Este exemplo usa o [conjunto de requisitos](http://dev.office.com/reference/add-ins/office-add-in-requirement-sets?product=word) 1.2 durante o período de visualização, mas exigirá o conjunto de requisitos 1.3 assim que ele estiver disponível.

### <a name="task-pane"></a>Task pane

A funcionalidade do painel de tarefas é configurada em sample.js que contém as seguintes funcionalidades:

* Configure a interface do usuário e manipuladores de eventos.
* Obtenha o modelo de especificações a partir de um serviço e insira-o no documento.
* Carregue uma lista de bloqueio que contenha palavras usadas para validar o documento. Essas palavras são consideradas palavras proibidas para este exemplo.
* Carregue um texto clichê padrão a partir de um serviço e guarde-o em um armazenamento local.
* Esqueleto de código para testar o código do arquivo de função. Convém desenvolver o código do comando de suplementos no painel de tarefas antes de movê-lo em um arquivo de função já que não é possível anexar um depurador ao arquivo de função.
* Carregue o nome padrão do autor a partir das propriedades do documento no painel de tarefas. Isso mostra como você pode acessar e alterar uma parte XML personalizada em um documento.
* Publique o texto clichê no serviço.

### <a name="document-validation-and-the-dialog-api"></a>Validação de documento e a API da caixa de diálogo

validation.js contém o código para validar o documento com base em uma lista de bloqueio de palavras. O método validateContentAgainstBlacklist() usa o novo método splitTextRanges para dividir o documento em intervalos com base em delimitadores. Os delimitadores nessa função identificam palavras no documento. Identificamos a interseção de palavras no documento e na lista de bloqueio e passamos esses resultados para o armazenamento local. Em seguida, usamos o método displayDialogAsync para abrir uma caixa de diálogo (dialog.html). Diálogo obtém os resultados da validação do armazenamento local e exibe os resultados.

### <a name="boilerplate-text-functionality"></a>Funcionalidade do texto clichê

boilerplate.js contém exemplos de como você pode salvar texto clichê no armazenamento local, atualizar uma lista suspensa do Fabric com entradas que correspondem ao texto clichê salvo e como inserir texto clichê selecionado em uma lista suspensa. Este arquivo contém exemplos de:
* splitTextRanges (novo no conjunto de requisitos WordApi 1.3) - essa API será substituída por split() em uma versão futura.
* compareLocationWith (novo no conjunto de requisitos WordApi 1.3)
* Atualize o menu suspenso do Fabric com as novas entradas.
* Inserir texto clichê no documento.

### <a name="custom-xml-binding-to-core-document-properties"></a>Associação personalizada de dados XML às propriedades principais do documento.

authorCustomXml.js contém métodos para obter e definir propriedades padrão do documento.

* Carregue a propriedade de autor no painel de tarefas quando o painel de tarefas for carregado. Observe se o documento também contém o valor da propriedade de autor. Isso ocorre porque o modelo contém um controle de conteúdo associado a essa propriedade do documento. Isso permite definir valores padrão no documento com base no conteúdo de uma parte XML personalizada.
* Atualize a propriedade de autor no painel de tarefas. Isso atualizará a propriedade do documento e o controle de conteúdo associado no documento.

### <a name="add-in-commands"></a>Comandos de suplemento

As declarações de comando do suplemento estão localizadas em word-add-in-javascript-speckit-manifest.xml. Este exemplo mostra como criar comandos de suplemento na faixa de opções e no menu de contexto.

## <a name="questions-and-comments"></a>Perguntas e comentários

Adoraríamos receber seus comentários sobre o exemplo do Word SpecKit. Você pode enviar seus comentários na seção *Problemas* deste repositório.

As perguntas sobre o desenvolvimento do Office 365 em geral devem ser postadas no [Stack Overflow](http://stackoverflow.com/questions/tagged/office-js+API). Não deixe de marcar as perguntas com [office-js] e [API].

## <a name="additional-resources"></a>Recursos adicionais

* [Documentação dos suplementos do Office](https://msdn.microsoft.com/en-us/library/office/jj220060.aspx)
* [Centro de Desenvolvimento do Office](http://dev.office.com/)
* [Exemplos de código e projetos iniciais de APIs do Office 365](http://msdn.microsoft.com/en-us/office/office365/howto/starter-projects-and-code-samples)

## <a name="copyright"></a>Copyright
Copyright (c) 2016 Microsoft Corporation. Todos os direitos reservados.



Este projeto adotou o [Código de Conduta de Software Livre da Microsoft](https://opensource.microsoft.com/codeofconduct/). Para saber mais, confira as [Perguntas frequentes sobre o Código de Conduta](https://opensource.microsoft.com/codeofconduct/faq/) ou contate [opencode@microsoft.com](mailto:opencode@microsoft.com) se tiver outras dúvidas ou comentários.
