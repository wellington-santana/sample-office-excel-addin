# Overview
Criar um suplemento para Office é muito útil nos momentos que temos que ter uma interação na aplicação da Microsoft e adicionar alguma regra de negócio antes do consumo. 

A facilidade de interagir via código com as funcionalidades nativas, agiliza e facilita a vida do desenvolvedor, além de salvar algumas horas no desenvolvimento de algum integrador. Qual é a abordagem mais simples? Criar uma API que consegue ler um documento Office e extraia as informações necessárias ou o próprio documento do Office enviar as informações?

Recentemente, diante da nessidade de um cliente, me fiz essa pergunta e acabei descobrindo o [Office Addin](https://docs.microsoft.com/pt-br/office/dev/add-ins/overview/office-add-ins)! Contudo, mesmo assim foi uma saga até conseguir finalmente realizar um deploy da solução (sim, isso foi um spoiler!).

# Limitações
Levanto em conta que o usuário final não é um desenvolvedor, é muito importante deixar claro que o resultado desse desenvolvimento será consumido facilmente e com manutenibilidade simples, apenas para usuários Office 365. As minhas dificuldades deixarei mais claras no decorrer do jornada, mas para simplificar posso dizer que simular o produto *vs* visualizar o produto final, são dois mundos distintos 😑.

# Setup
Todas as decrições abaixo, podêm ser visualizadas no próprio site com tutorial passo a passo clicando [aqui](https://docs.microsoft.com/pt-br/office/dev/add-ins/quickstarts/excel-quickstart-jquery?tabs=yeomangenerator).

Após instalar os pré-requisitos ([Node](https://nodejs.org/en/) e [Yeoman](https://yeoman.io/learning/)), vamos começar com o passo a passo:

1. Execute o comando na linha de comando:
> yo office

2. O setup será iniciado como um 'npm init'. Com uma bela surpresa, o projeto pode ser em Java Script, 🎉 Angular 🎉 ou React.
Basicamente, quando iniciamos o addin dentro do Office, é realizada uma chamada para o end-point que conterá nossa aplicação, que nada mais é que uma SPA.
Após finalizar o setup, tudo que é necessário para iniciar o desenvolvimento foi instalado automaticamente e com as divisões pastas e afins. No meu caso, usei Angular.

# Mão na massa

## Rodando o addin
Navegando pelo projeto, dêem uma olhada no 'package.json', haverá alguns comandos que usaremos para debug. Por exemplo, após realizar o setup inicial, execute o comando:
> npm run-script start:desktop

* O comando acima é muito bom para visualizar a versão desktop, mas não é a melhor abordagem para debug. Afinal, não temos console e a implementação de alerta é bem chata e nada agradável para verificar erros.

Uma janela do Node será inicializada, a compilação da aplicação será apresentada e automaticamente o seu aplicativo alvo do Office será inicializado com o seu addin. 

É possível visualizar no meu navegador? Sim! Mas não é pelo comando abaixo. ESSE COMANDO NÃO FUNCIONA.
> npm run-script start:web

A melhor maneira de desenvolver, na qual foi 100% utilizada após eu descobrir, é usar o Office 365 online. Estamos desenvolvendo uma SPA, desta maneira temos os mesmos recursos que estamos acontumados no desenvolvimento Web. Sim, estou falando do famoso F12.

Para publicar o seu addin é simples:

1. Abra a aplicação alvo no [Office 365](https://www.office.com/)
2. Inicie um documento
3. Na aba inserir, clique em suplementos. ![alt](/readme_img/1.png)
4. No seu projeto, há um arquivo chamado 'manifest.xml'. Informe esse arquivo para carregar o suplemento. ![alt](/readme_img/2.png)
5. Pronto, seu addin está totalmente vinculado ao projeto em desenvolvimento. Isso inclui um 'watch', todas as alterações realizadas em seu desktop, após compilação, serão atualizadas no addin carregado no navegador.
6. Todos os resursos de depuração também estão habilitadas, no meu caso usando o projeto em Angular, os arquivos '.ts' são carregados.

## Usando a biblioteca do Office
Em toda a tragetória de desenvolvimento, mantive o exemplo que o template oferece. Sendo assim, todas as funções eram assíncronas carregando o contexto no início da execução.
Com o Typescript, é muito tranquilo entender a bibilioteca quando fazemos alguma coisa errada, a descrição do erro é bem verbosa que nos permite solucionar o erro sem buscar a ajuda do StackOverflow.
Abaixo, exemplo comentado linha a linha de como interagir com as células selecionadas do Excel.

<code>
  
    try {
        //Função padrão para carregamento do contexto    
        Excel.run(async context => {
        
        //Capturando planilha que está selecionada
        let sheet = 
            context.workbook.worksheets.getActiveWorksheet();
        
        //Capturando as células que está selecionadas
        let range = 
            context.workbook.getSelectedRange().getUsedRange();

        /*
        * Carregando as variáveis que irei utilizar
        * Esse comando é obrigatório quando queremos capturar as 
        * informações das células
        */        
        range.load(["values", "columnIndex", "rowIndex", "address"]);

        /*
        * Com o comando sync, as informações do comando load
        * são carregadas e via promisse podemos capturar as 
        * informações das células
        */
        context.sync().then(() => {
            /*
            * Na variável range há os endereços das células
            * selecionadas. Desta maneira é possivel usar o 
            * restante da biblioteca para interação
            */
            let selecionList = 
                this.excelAddinService.mapSelection(range.address);

            this.importRangeSpecified(selecionList, sheet, range);
        });
      });
    } catch (error) {
        console.error(error);
        this.modal.open();
    }
</code>

# Surpresas
Tive algumas surpresas durante o desenvolvimento do addin, pretendo listar as mais chatas que me demandaram um certo tempo até encontrar alguma pista nos fóruns. Meu desenvolvimento foi em Angular, então quem realizará em Java Script ou React boa sorte 🤷‍♂️.

## Versão do Angular
Hoje, a versão que é instalada automaticamente é a 5.2.9. Entretanto, módulos que estamos acostumados a usar automaticamente em nosso dia a dia não são instaladas. Então, fique atendo nas dependências quando algo que você esteja acostumado a usar não funciona.

## One e Two-Way Data Biding
Gastei muitas horas tentando, mas não consegui fazer funcionar. Por favor, quem conseguir, me ensine esse milagre!
Para sair do outro lado, use as variavéis de template e às referenciem na ação.
> <input type="text" id="login" name="login" placeholder="Usuário" #usuario />

> <input type="text" id="password" name="login" placeholder="Senha" type="password" #senha />

> <input type="button" value="Acessar" (click)="login(usuario.value, senha.value)" />

## Injeção de serviços
A injeção de um serviço deve ser feita da maneira abaixo. Não consegui de outra maneira (no módulo continua como estamos acostumados).
>//x.component.ts
>
>     constructor(
>       @Inject(Router) router: Router) {
>       this.router = router;
>     }

>//x.module.ts
>
>     @NgModule({
>       providers: [MeuServico],
>       declarations: [...],
>       imports: [...],
>       bootstrap: [...]
>     })

## Template html e CSS
A maneira de declarar onde está o arquivo template e css não funciona como eu estava acostumado (ou estamos):
>     @Component({
>       selector: "login-component",
>       templateUrl: 'login.component.html',
>       styleUrls: ['login.compoent.css']
>     })

A maneira que utilizei é declarar o arquivo no html principal (taskpane.html):
>     <link href="taskpane.css" rel="stylesheet" type="text/css" />

Enquanto que o arquivo template sendo declarado via variável no próprio componente:
>     const template = require("./app.component.html");
> 
>     @Component({
>       selector: "app-home",
>       template: template,
>     })

## Debugging
A maneira mais simples de debugar é usar o comando console. Para usar, basta declarar globalmente no início do arquivo .ts como no exemplo:
> /* global console, Excel, require */

# Bora para Prod!
Esta é uma etapa que demandei um bom tempo pesquisando para finalizar. No início desde tutorial, comentei que desenvolvimento e produção são dois mundos distintos, e são!

Quando estamos desenvolvendo, conseguimos testar nosso addin no próprio Office que temos na máquina, no meu caso foi as versão Excel 2016. Após 'finalizar' o código, procurei várias maneiras de gerar o artefato de produção que poderia executar na minha própria máquina, falhei miseravelmente.

Temos três maneiras de publicar:

1. Compartilhar o código fonte ou o arquivo manifest.xml e quem precisar utilizar o recurso, realizar o mesmo processo de setup e execução. Nada viável.
2. Publicar o addin na Office Store
3. Publicar nos addins na store da organização

Passo a passo no [link](https://docs.microsoft.com/pt-br/office/dev/add-ins/publish/publish) sobre as publicações.

No meu caso, a aplicação é apenas para uso interno, desta maneira usei a 3ª opção. Abaixo, vou demonstrar como publiquei o addin em um servidor e o seu consumo.

A aplicação gerada é uma SPA, então realizei a publicação do fonte na Azure e modifiquei as configurações necessárias no arquivo 'manifest.xml'. Desta meneira, quem instalar o addin, consumirá aplicação que está na nuvem.

## Gerando o artefato
Para gerar o artefato é bem simples. No meu caso usando Angular, basta executar o comando abaixo e os arquivos serão gerados na pasta 'dist' na raiz do projeto.
> npm run-script build

Após o build, publique os artefatos para o servidor que receberá as requisições.

### Bug
No meu caso, tive que realizar uma modificação nas configurações do Webpack. Os atributos do Angular eram gerados em caixa baixa, então ao executar eram lançados erros. Insira o comando abaixo no arquivo 'webpack.config.js' para resolver.
>     {
>         test: /\.html$/,
>         exclude: /node_modules/,
>         use: {
>            loader: "html-loader", 
>            options:{
>              minimize: false
>            } 
>         },          
>     },

## Publicando o arquivo 'manifest.xml'
O arquivo está referenciando, provavelmente, o ambiente local. Substitua todos os apontamentos locais para URL do servidor que foi feita a publicação dos artefatos.

![alt](/readme_img/3.png)

Após realizar o ajuste, o arquivo 'manifest.xml' pode ser compartilhado para instalação manual pelo Office 365 online, publicar na Office Store, publicação na Store local ou no diretório compartilhado da empresa.

# Conclusões
O início de qualquer desenvolvimento, tecnologia ou coisas que fogem da nossa zona de conforto, sempre trazem grandes desafios, e este não foi diferente. Basta ter paciência e buscar informações nos fóruns para solucionar as dificuldades até a conclusão.

Espero realmente que essa descrição da minha jornada de desenvolvimento ajude os próximos que encontrarem dificuldades  e facilite o 'getting started'.

Fico a disposição para esclarecer dúvidas e discutir melhorias/ correções nos relatos acima.