# FSL Campo — Dashboard Operacional

Dashboard Flask para acompanhamento da operação FTTH/ICOMON, com visualização responsiva para celular e instalação como PWA.

## Recursos

- Indicadores de produção, reparos, pendências e execução em campo.
- Visões regionais Norte e Sul e filtros operacionais.
- Atalho móvel para colaboradores zerados.
- Interface mobile-first com tabelas apresentadas como cartões.
- Instalação no Android, iPhone e desktop como PWA.
- Tela offline para informar indisponibilidade de conexão.
- Coleta de relatórios do Salesforce por Selenium.

## Requisitos

- Python 3.11 ou superior.
- Google Chrome instalado.
- ChromeDriver compatível com a versão do Chrome.
- Acesso autorizado à organização Salesforce utilizada pelo projeto.

## Instalação local

```powershell
python -m venv .venv
.venv\Scripts\Activate.ps1
python -m pip install --upgrade pip
pip install -r requirements.txt
python app.py
```

Abra `http://localhost:5000`. Para carregar os relatórios, clique em **Iniciar Driver**, autentique-se no Salesforce na janela aberta e depois atualize a base.

## Instalação como aplicativo

Em `localhost`, o navegador permite testar a instalação diretamente. Em outro computador ou celular, publique a aplicação com HTTPS obrigatório.

- Android/Chrome: abra o menu e selecione **Instalar app** ou use o aviso exibido pelo dashboard.
- iPhone/Safari: toque em **Compartilhar** e depois em **Adicionar à Tela de Início**.
- Desktop/Chrome ou Edge: use o ícone de instalação na barra de endereço.

O service worker mantém apenas o shell visual e a tela offline. Os indicadores e endpoints `/api/` continuam dependendo da conexão com o servidor para evitar a exibição de dados operacionais antigos.

## Estrutura

```text
app.py                     Servidor Flask, Selenium e processamento dos dados
templates/index.html       Dashboard desktop/mobile
static/manifest.webmanifest Configuração da PWA
static/sw.js               Service worker
static/offline.html        Tela sem conexão
static/icon*.svg           Ícones instaláveis
metas_reparos.json         Metas persistidas por regional/coordenador
```

## Publicação do código no GitHub

Antes do primeiro envio, confirme que planilhas, áudios, downloads e o executável do ChromeDriver não serão versionados. Esses itens já estão contemplados no `.gitignore`.

```powershell
git init
git add .
git status
git commit -m "Publica dashboard FSL Campo com PWA"
git branch -M main
git remote add origin https://github.com/SEU-USUARIO/SEU-REPOSITORIO.git
git push -u origin main
```

O GitHub hospeda o código, mas o GitHub Pages não executa o backend Flask/Selenium. Para acesso remoto, use uma máquina ou servidor com Chrome disponível, mantenha o processo Flask em execução e coloque um proxy HTTPS na frente da aplicação.

## Segurança e dados

- Não publique planilhas operacionais ou arquivos baixados do Salesforce.
- Não armazene usuário, senha, cookies ou tokens no repositório.
- Restrinja o acesso ao dashboard quando ele for exposto fora da rede interna.
- Revise a URL e o ID do relatório Salesforce em `app.py` antes de compartilhar o repositório publicamente.

## Observação sobre produção

O servidor de desenvolvimento embutido do Flask não é indicado para exposição direta à internet. Em ambiente Windows interno, execute o app atrás de um servidor WSGI e de um proxy HTTPS compatível com Web/PWA, garantindo que a sessão gráfica necessária ao Selenium esteja disponível.
