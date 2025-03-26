📜 Tradutor de Apresentações PowerPoint

Este projeto traduz automaticamente arquivos PowerPoint (.pptx) do inglês para o português (ou outro idioma suportado) usando o LibreTranslate.

---

📌 Funcionalidades

✅ Instalação automática de dependências
✅ Servidor de tradução integrado (LibreTranslate)
✅ Tradução preservando formatação do PowerPoint
✅ Interface para seleção de arquivos
✅ Detecção de teclas para interromper a tradução
✅ Barra de progresso para melhor visualização

---

🛠️ Instalação e Uso

🔹 1. Requisitos

- Python 3.7+
- Git

🔹 2. Clonar o Repositório

git clone https://github.com/seu-usuario/seu-repositorio.git
cd seu-repositorio

🔹 3. Executar o Script

python tradutor_ppt.py

Durante a execução, o script:

1. Instala automaticamente as dependências necessárias.
2. Configura e inicia o servidor LibreTranslate.
3. Abre uma janela para selecionar um arquivo PowerPoint.
4. Traduz os textos e salva um novo arquivo com a tradução aplicada.
5. Pressione 'q' para interromper a tradução a qualquer momento.

---

🖥️ Dependências

O script verifica e instala automaticamente os seguintes pacotes:

- requests - Para requisições HTTP
- python-pptx - Manipulação de arquivos PowerPoint
- tqdm - Barra de progresso
- argostranslate - Tradução offline
- keyboard - Captura de teclas
- pyautogui - Manipulação de janelas
- Flask, flask_cors - Servidor do LibreTranslate

Caso precise instalar manualmente:

pip install requests python-pptx tqdm argostranslate keyboard pyautogui Flask flask_cors

---

⚙️ Como Funciona?

1. Inicia o servidor LibreTranslate na porta 5000 (ou outra disponível).
2. Seleciona um arquivo .pptx para tradução.
3. Traduz o conteúdo slide a slide, mantendo formatação.
4. Salva um novo arquivo com o sufixo _traduzido_pt.pptx.

---

📝 Observações

- Caso o LibreTranslate não esteja instalado, o script baixa e configura automaticamente.
- Se a porta 5000 estiver em uso, ele tenta usar a porta 5001.
- Suporte a outros idiomas além do português (modificável no código).

---

📜 Licença

Este projeto está licenciado sob a MIT License. Veja o arquivo LICENSE para mais detalhes.

---

🤝 Contribuições

Sinta-se à vontade para abrir issues e pull requests para melhorias! 🚀
