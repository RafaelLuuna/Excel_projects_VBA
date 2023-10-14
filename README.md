# Excel_projects_VBA

Este repositório contém projetos antigos que fiz no excel usando o VBA. Cada projeto teve diversos desafios e muitos aprendizados, não vou entrar em detalhes de como os códigos funcionam más fique a vontade para explorar!

Com o intuito de facilitar a navegação, eu separei todos os scripts e formulários usados nos projetos em uma pasta chamada 'Script' dentro de cada projeto. Esses arquivos são apenas para consulta, não estão vinculados diretamente ao arquivo excel, ou seja, caso altere alguma coisa no código dessa pasta não será refletido na aplicação.

Para alterar o código da aplicação é preciso acessar o editor do visual basic dentro do próprio arquivo excel (você pode usar esse guia do blog ninjadoexcel caso tenha alguma dúvida: https://ninjadoexcel.com.br/como-abrir-o-visual-basic-no-excel/).


## Visão geral de cada projeto

### RPG GAME
Esse projeto trouxe grandes desafios, eu queria aproveitar a capacidade do formulário do VBA de carregar imagens, para montar um jogo estilo os pokémons antigos, onde o mundo era visto de cima para baixo.

Dentro desse jogo eu queria que fosse possível: 
* Desenhar o mundo da maneira que eu quisesse sem ter que alterar muita coisa no código.
* Movimentar o player de maneira suave (ou seja, criar uma animação usando VBA :cold_sweat:).
* Houvesse objetos interativos como o baú ou a árvore por exemplo.
* Fosse possível entrar em interiores.
* Tenha um sistema de diálogo simples.

Para a minha alegria todos os requisitos acima foram cumpridos! Se você for análisar o código vai perceber que eu criei um sistema pensando no máximo de versátilidade para alterar o tamanho da exibição, cadastrar novos itens, cadastrar novos diálogos, etc... Por exemplo, essas constantes abaixo definem o tamanho da grade de sprites que vai ser exibida, no caso 17 sprites por 17 sprites.

```
'Módulo -- DATA.bas
'Linha -- 47, 48

Public Const xArraySize = 17
Public Const yArraySize = 17

```

O mundo do jogo tem 48 x 32, o código consegue identificar quando o player está fora da área de exibição (ele chegou na posição 18 por exemplo) para renderizar o próximo chunk do mundo baseado nessas constantes acima. Você pode brincar com esses valores e perceberá que o jogo sempre se adapta ás dimensões que você definir.

Dessa forma o mundo não fica limitado apenas aos sprites que eu posso rederizar na tela, consigo criar cenas muito maiores.

Eu usei as próprias tabelas do excel como banco de dados, através delas eu consigo ler e escrever todas as informações que preciso, inclusive desenhar as layers do mundo, além de salvar informações como inventário, carteira, diálogos, etc...

Caso se interesse no projeto e queira saber mais, me chama no discord: RafaelLuuna



