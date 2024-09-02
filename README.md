Descrição:

  Automação de cadastro em Python utilizando planilhas Excel. O projeto automatiza o processo de registro de dados, eliminando a necessidade de intervenção manual. Desenvolvido com pyautogui, openpyxl e pyperclip, é uma solução eficiente para economizar tempo em tarefas repetitivas.

Sobre o Projeto:

  Este projeto foi desenvolvido para automatizar o processo de cadastro de informações em um sistema, utilizando dados extraídos de uma planilha Excel. A automação é realizada através de Python, utilizando as bibliotecas pyautogui para controlar o cursor do mouse e teclado, openpyxl para manipulação da planilha Excel, e pyperclip para facilitar o trabalho com o clipboard.

Como Funciona:

  Planilha Excel: A automação se baseia em uma planilha Excel onde os dados de cadastro estão organizados.

  Coordenadas: O script utiliza coordenadas específicas para interagir com os campos do formulário de cadastro. É importante ajustar essas coordenadas no código conforme a posição dos campos na tela do seu sistema.

  Execução Automática: Após configurado, basta iniciar o script e deixar sua máquina executar o processo de cadastro automaticamente, sem interrupções.

Requisitos:

  Ambiente de Cadastro: Certifique-se de que o ambiente onde os cadastros serão realizados está configurado corretamente e sem interferências externas.

  Ajuste de Coordenadas: Antes de executar o script, ajuste as coordenadas no código para que correspondam aos campos de texto e botões do sistema de cadastro que você está utilizando.

Tecnologias Utilizadas

Python:

  pyautogui: Para automação de movimentos e cliques do mouse e teclado.
  
  openpyxl: Para leitura e manipulação de dados em planilhas Excel.
  
  pyperclip: Para copiar e colar dados do clipboard de forma eficiente.
  
Considerações:

Esse projeto é ideal para situações onde há a necessidade de realizar cadastros em massa de forma rápida e eficiente. No entanto, é fundamental que a máquina onde o script está rodando não seja usada para outras tarefas durante a execução da automação, a fim de evitar interrupções e erros.
