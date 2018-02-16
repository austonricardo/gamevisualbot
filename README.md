# gamevisualbot

Não é algo novo, encontrei uns código antigos que fiz no passado (2011) por diversão e resolvi compartilhá-los! 

Um bot que simula ações baseado em condições, em regra serve pra qualquer game mas na época quando usei era no L2, também pode ser usado para testar games :). Utiliza detecção nas mudanças de cores em regiões da tela, extrai informações e executa ações programadas a tempos regulares. Ele vincula ações a uma janela então é possível usar várias instancias vicnulados a telas diferentes. Lembro de ter testado com 4 contas e 4 char simultaneamente.

![Tela inicial visual game bot](https://github.com/austonricardo/gamevisualbot/blob/master/NovoAuto2.png "Tela inicial")

Feito em visual basic 6. Existem hoje várias versões portáveis desse editor/linguagem como esse:
[Artigo sobre VB6 Portavel](https://thementalclub.com/download-visual-basic-6-0-3038) cheque com seu antivirus antes de executar qualquer coisa executável da net.
 

Essa versão tem os recursos:
- Ação condicionada ao CP, HP e MP. Para utilizá-la, a parte da tela do game que mostra os atributos deve estar visível.
- Salvar e Abrir as configurações feitas.
- Alerta sonoro de captcha.

Como usar:
- Abra o game inicie a té a tela em que aparece os atributos do personagem.
- Clique o botão Janela e passe o mouse sobre a janela do jogo, ela será escolhida.
- Clique no botão "Full State", para capturar seus atributos, nesse momento os atributos devem estar todos cheios. 
- Se der tudo certo seus atributos aparecerao com 50%, senão será necessário calibrar as cores. Para fazer isso:
>-  clique no botão que tem a cor do atributo "Vermelho" (HP).
>-  Passe o Mouse sobre o "vermelho" do HP na figura obtida, e clique duplo para escolher.
>-  Repita para os demais atributos.
>-  Clique em "Full State" novamente e veja se estão todos com "XX%", onde XX é promixo de 50% se não estiver precisa ser calibrado.
>-  Informe os comando e tempo e inicie.

Dicas: Se vc precisou calibrar as cores é uma boa salvar o arquivo com a configurações pra carregar depois.
Você pode iniciar e acompanhar pra ver se os atributos estão ok, se tiver divergindo muito precisa calibrar pra corrigir.
