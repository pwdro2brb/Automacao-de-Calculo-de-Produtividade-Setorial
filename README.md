# Sistema de Automação de Produtividade

Este projeto automatiza a coleta de dados de diferentes sistemas corporativos e preenche um modelo de planilha de produtividade.

## O que o sistema faz

O script em [produtividade.py](produtividade.py) executa as seguintes etapas:

1. Acessa sistemas como Podio, Agilis, Bússola e MIR5.
2. Realiza login via autenticação Microsoft, incluindo suporte a MFA quando necessário.
3. Baixa relatórios e exportações em Excel.
4. Organiza e renomeia os arquivos gerados.
5. Processa os dados e preenche uma planilha de produtividade com informações de Agilis, Sedex, Lanctos, SAP/Miro e flags FSF.

## Requisitos

- Windows
- Python 3.10 ou superior
- Google Chrome instalado
- ChromeDriver compatível com a versão do Chrome
- Dependências Python:

```bash
pip install pandas openpyxl selenium pyautogui pillow pyperclip pywin32
```

## Configuração

Antes de executar, edite o arquivo [produtividade.py](produtividade.py) e ajuste:

- `EMAIL_USER`
- `SENHA_USER`
- Os caminhos de entrada/saída, se necessário
- Os nomes dos arquivos esperados, caso o ambiente tenha mudado

> Importante: mantenha suas credenciais em segurança e não as compartilhe.

## Como executar

Na pasta do projeto, execute:

```bash
python produtividade.py
```

## Observações

- O script depende de elementos visuais e fluxos das páginas web, então pode exigir ajustes se a interface mudar.
- Alguns passos podem aguardar confirmação manual, como login por MFA ou janelas de confirmação do sistema.
- Certifique-se de que o navegador e os downloads estejam liberados para o ambiente do usuário.

## Estrutura do projeto

- [produtividade.py](produtividade.py): script principal com a automação e o processamento dos dados.
- [README.md](README.md): documentação do projeto.
