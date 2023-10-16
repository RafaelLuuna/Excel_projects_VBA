# Excel_projects_VBA

Este repositório contém projetos antigos que fiz no excel usando o VBA. Cada projeto teve diversos desafios e muitos aprendizados, não vou entrar em detalhes de como os códigos funcionam más fique a vontade para explorar!

Com o intuito de facilitar a navegação, eu separei todos os scripts e formulários usados nos projetos em uma pasta chamada 'Script' dentro de cada projeto. Esses arquivos são apenas para consulta, não estão vinculados diretamente ao arquivo excel, ou seja, caso altere alguma coisa no código dessa pasta não será refletido na aplicação.

Para alterar o código da aplicação é preciso acessar o editor do visual basic dentro do próprio arquivo excel (você pode usar esse guia do blog ninjadoexcel caso tenha alguma dúvida: https://ninjadoexcel.com.br/como-abrir-o-visual-basic-no-excel/).


## Visão geral de cada projeto

### RPG GAME
Esse projeto trouxe grandes desafios, eu queria aproveitar a capacidade do formulário do VBA de carregar imagens, para montar um jogo estilo os pokémons antigos, onde o mundo era visto de cima para baixo.

Dentro desse jogo eu queria que fosse possível: 
* Desenhar o mundo sem ter que alterar muita coisa no código.
* Movimentar o player de maneira suave (ou seja, criar uma animação usando VBA :cold_sweat:).
* Houvesse objetos interativos como o baú ou a árvore por exemplo.
* Fosse possível entrar em interiores.
* Tenha um sistema de diálogo simples.

Para a minha alegria todos os requisitos acima foram cumpridos! :grin: 

Eu usei as próprias tabelas do excel como banco de dados, através delas eu consigo ler e escrever todas as informações que preciso, inclusive desenhar as layers do mundo, além de salvar informações como inventário, carteira, diálogos, etc...

Se você for análisar o código vai perceber que eu criei um sistema pensando no máximo de versátilidade possível. Eu quis facilitar ao máximo tarefas que são recorrentes no desenvolvimento do jogo como alterar sprites no mundo, cadastrar novos itens, cadastrar novos diálogos, etc... Por exemplo, eu posso alterar ou incluir falas de NPCs simplesmente mexendo nos valores da aba 'Script', além disso, na coluna 'H' eu posso usar um dos comandos abaixo para executar determinada ação de acordo com o roteiro.

Por exemplo, eu posso colocar 'OptionMode' na coluna 'H', e as próximas colunas definem quais vão ser as opções que o usuário pode selecionar, na próxima linha, eu posso colocar o comando 'OptionSelected' para executar alguma ação de acordo com a opção selecionada na linha anterior. A ação pode ser 'GoTo' para ir para outra linha do script, 'GaveItem' para incluír um item em um inventário ou 'UpdateWalllet' para alterar o valor da carteira de um personagem.

Outra coisa bem interessante é que a renderização do mundo não fica limiado ao número de sprites na tela. 

Por exemplo, essas constantes abaixo etão no módulo 'DATA.bas', elas definem o tamanho da grade de sprites que vai ser exibida, no caso, 17 sprites por 17 sprites.

```
Public Const xArraySize = 17
Public Const yArraySize = 17

```

O mundo do jogo tem 48 x 32, teoricamente não seria possível exibir o mundo todo na tela sem alterar essas constantes.

Para resolver esse problema, o código consegue identificar quando o player está fora da área de exibição (se ele chegou na posição 18 por exemplo) e então renderizar o próximo chunk do mundo baseado nessas constantes acima. 

Você pode alterar esses valores da maneira que quiser e perceberá que o jogo sempre se adapta ás dimensões que você definir.

Dessa forma o mundo não fica limitado apenas aos sprites que eu posso rederizar na tela, consigo criar cenas muito maiores.



Caso se interesse no projeto e queira saber mais, fique a vontade para entrar em contato através do discord: RafaelLuuna

### PS OS For Clothing

Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat. Duis aute irure dolor in reprehenderit in voluptate velit esse cillum dolore eu fugiat nulla pariatur. Excepteur sint occaecat cupidatat non proident, sunt in culpa qui officia deserunt mollit anim id est laborum.

### PDV ADEGA MARQUINHOS

Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat. Duis aute irure dolor in reprehenderit in voluptate velit esse cillum dolore eu fugiat nulla pariatur. Excepteur sint occaecat cupidatat non proident, sunt in culpa qui officia deserunt mollit anim id est laborum.

