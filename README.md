# Excel_projects_VBA

Este repositório contém projetos antigos que fiz no excel usando o VBA. Cada projeto teve diversos desafios e muitos aprendizados, não vou entrar em detalhes de como os códigos funcionam mas fique a vontade para explorar!

Com o intuito de facilitar a navegação, eu separei todos os scripts e formulários usados nos projetos em uma pasta chamada 'Script' dentro de cada projeto. Esses arquivos são apenas para consulta, não estão vinculados diretamente ao arquivo excel, ou seja, caso altere alguma coisa no código dessa pasta não será refletido na aplicação.

Para alterar o código da aplicação é preciso acessar o editor do visual basic dentro do próprio arquivo excel (você pode usar esse guia do blog ninjadoexcel caso tenha alguma dúvida: https://ninjadoexcel.com.br/como-abrir-o-visual-basic-no-excel/).


## Visão geral de cada projeto

### RPG GAME
Esse projeto trouxe grandes desafios, eu queria aproveitar a capacidade do formulário do VBA de carregar imagens, para montar um jogo estilo os pokémons antigos, onde o mundo era visto de cima para baixo e o mundo renderizado em blocos. A idéia premissa do jogo é ser um rougue-like onde você controla um lenhador que deve juntar madeira para vender e ganhar dinheiro para comprar comida e conseguir sobreviver ao inverno, e conforme fosse desmatando a floresta ficaria cada vez mais difícil achar madeira.

Infelizmente por falta de tempo não pude finalizar o jogo, mas consegui grandes avanços que valem a pena serem compartilhados:
* Sistema de inventário.
* Sistema de carteiras.
* Sistema de diálogos. (os itens do inventário e o saldo da carteira podem ser manipulados nesse sistema também)
* Animação do personagem principal para se movimentar no mundo. (no VBA isso traz um certo desafio, em JavaScript seria mais simples kk)
* Mundo aberto com a possibilidade de entrar em interiores.
* Interações com o mundo (baús e árvores).

Comentários sobre o desenvolvimento desse projeto:
Eu usei as próprias tabelas do excel como banco de dados, através delas eu consigo ler e escrever todas as informações que preciso, inclusive desenhar as layers do mundo, além de salvar informações do inventário, carteira, diálogos, etc...

Se você for análisar o código vai perceber que eu criei um sistema pensando no máximo de versátilidade possível. Eu quis facilitar ao máximo tarefas que são recorrentes no desenvolvimento do jogo como alterar sprites no mundo, cadastrar novos itens, cadastrar novos diálogos, etc... Por exemplo, eu posso alterar ou incluir falas de NPCs simplesmente alterando os valores da aba 'Script' do arquivo excel, além de conseguir executar comandos no meio das falas para realizar diversas interações, por exemplo:

Se colocar 'OptionMode' na coluna 'H' da aba scripts, as colunas 'I', 'J', 'K'... vão defir quais serão as opções que o usuário pode selecionar, na próxima linha do roteiro, eu posso colocar o comando 'OptionSelected' para ler a opção selecionada e executar alguma ação de acordo com a seleção do jogador. Essa ação pode ser 'GoTo' para ir para outra linha do script, 'GaveItem' para incluír um item em um inventário ou 'UpdateWalllet' para alterar o valor da carteira de um personagem.

Outra coisa bem interessante dess é que a renderização do mundo não fica limiado ao número de sprites na tela. 

Por exemplo, essas constantes abaixo etão no módulo 'DATA.bas', elas definem o tamanho da grade de sprites que vai ser exibida, no caso, 17 sprites por 17 sprites.

```
Public Const xArraySize = 17
Public Const yArraySize = 17

```

O mundo do jogo tem 48 x 32, teoricamente não seria possível exibir o mundo todo na tela sem alterar essas constantes.

Para resolver esse problema, o código consegue identificar quando o player está fora da área de exibição (se ele chegou na posição 18 por exemplo) e então renderizar o próximo chunk do mundo baseado nessas constantes acima. 

Você pode alterar esses valores da maneira que quiser e o jogo sempre se adapta ás dimensões que você definir.

Dessa forma o mundo não fica limitado apenas aos sprites que eu posso rederizar na tela, consigo criar cenas muito maiores.

Existem muitas outras funções interessantes no jogo que não vou conseguir detalhar o código em um único README, por exemplo o sistema de inventário que permite manipular os itens dentro do jogo ao precionar 'Enter'.


### PS OS For Clothing

Esse sistema foi desenvolvido para gerenciar uma pequena loja de roupas. Nele é possível cadastrar produtos, fornecedores, registrar vendas, e fazer um controle do fluxo de caixa.

O sistema permite cadastrar produtos de diversos tamanhos e fazer o gerenciamento do estoque, inclusive registrar pedidos de compras de cada fornecedor.

Ele conta com vários formulários automatizados por VBA que possuem formatações automáticas nos campos de input e faz o registro de todas as vendas.


### PDV ADEGA MARQUINHOS

Esse sistema de PDV permite 

