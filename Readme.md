# SendMailsApp

SendMailsApp é uma aplicação de desktop desenvolvida em Python que permite o envio de e-mails em massa de forma simples e eficiente. Utilizando uma interface gráfica amigável construída com Tkinter, a aplicação permite que os usuários configurem o assunto e a mensagem do e-mail, anexem imagens e utilizem uma lista de destinatários a partir de um arquivo Excel.

## 📋 Recursos

- **Interface Intuitiva**: Fácil de usar com campos claros para assunto, mensagem, anexos e destinatários.
- **Envio de E-mails em Massa**: Envie e-mails para múltiplos destinatários de uma vez.
- **Anexar Imagens**: Adicione imagens diretamente no corpo do e-mail.
- **Leitura de Destinatários via Excel**: Importe a lista de destinatários a partir de um arquivo `.xlsx`.
- **Progresso Visual**: Barra de progresso que indica o status do envio dos e-mails.
- **Logs Detalhados**: Registre todas as operações e erros em um arquivo de log para referência futura.
- **Validação de E-mails**: Verifica se os endereços de e-mail possuem um formato válido antes de enviar.
- **Multithreading**: Envio de e-mails ocorre em uma thread separada para manter a interface responsiva.
- **Compatível com Windows e Outros Sistemas Operacionais**: Ajuste de ícone dependendo do sistema operacional.

## 🚀 Começando

### 📦 Pré-requisitos

Certifique-se de ter o Python 3 instalado em sua máquina. Você pode baixar o Python [aqui](https://www.python.org/downloads/).

### 🔧 Instalação

1. **Clone o repositório:**

```bash
git clone https://github.com/seu-usuario/SendMailsApp.git
cd SendMailsApp
```

2. **Crie e ative um ambiente virtual (opcional, mas recomendado):**

```bash
python -m venv venv
# No Windows
venv\Scripts\activate
# No macOS/Linux
source venv/bin/activate
```

3. **Instale as dependências:**

```bash
pip install -r requirements.txt
```

### 📂 Estrutura do Projeto

```bash
SendMailsApp/
├── icon.ico               # Ícone para Windows
├── icon.png               # Ícone para outros sistemas operacionais
├── LOG_SENDMAILS.log      # Arquivo de log (gerado automaticamente)
├── sendmail.py            # Código fonte da aplicação
├── requirements.txt       # Dependências do projeto
└── ...                    # Outros arquivos e diretórios
```

### 🛠️ Configuração

```bash
smtp_server = 'smtp.email.com'
smtp_port = 587
sender_email = 'email@email.com'
sender_password = 'senha'
```

**Substitua pelos detalhes do seu provedor de e-mail:**

- `smtp_server`: Endereço do servidor SMTP (por exemplo, `smtp.gmail.com` para Gmail).
- `smtp_port`: Porta do servidor SMTP (587 para TLS, 465 para SSL).
- `sender_email`: Seu endereço de e-mail.
- `sender_password`: Sua senha de e-mail ou senha de aplicativo.

> **Atenção:** Para contas que utilizam autenticação de dois fatores, pode ser necessário gerar uma senha de aplicativo específica.

### 🎨 Uso

1. **Inicie a aplicação:**
    
    Execute o script Python:
	```bash
	python sendmail.py
	```

1. **Preencha os campos:**
    
    - **Assunto**: Insira o assunto do e-mail.
    - **Mensagem**: Escreva a mensagem do e-mail. Você pode incluir texto e imagens.
    - **Imagem**: (Opcional) Clique em "IMAGEM" para selecionar uma imagem que será incorporada no e-mail.
    - **E-Mails**: Clique em "E-MAILS" para selecionar um arquivo Excel (`.xlsx`) contendo a lista de destinatários na coluna `Emails`.
    
1. **Enviar E-mails:**    

    - Clique no botão "ENVIAR" para iniciar o envio dos e-mails. A barra de progresso indicará o status do envio.
    - Caso ocorram erros durante o envio, uma janela de aviso exibirá os e-mails que não puderam ser enviados. 
1. **Limpar Campos:**
2. 
    - Clique em "LIMPAR" para limpar todos os campos preenchidos. 
3. **Sair:**
    
    - Clique em "SAIR" para fechar a aplicação.

## 🛡️ Segurança

Certifique-se de não expor suas credenciais de e-mail no código. Considere utilizar variáveis de ambiente ou arquivos de configuração seguros para armazenar informações sensíveis.