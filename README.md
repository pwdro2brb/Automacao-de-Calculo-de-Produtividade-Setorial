📊 Automação de Cálculo de Produtividade Setorial
📝 Sobre o Projeto
Este projeto é uma automação em Python desenvolvida para calcular a produtividade da equipe, consolidando dados de quatro planilhas diferentes. O script elimina o trabalho manual de cruzamento de dados e geração de relatórios mensais.

⚙️ Funcionalidades
Extração Automática (Web Scraping): Acesso e download automático de relatórios dos sistemas Podio e Agilis.

Processamento de Dados: Integração das duas planilhas geradas automaticamente com duas planilhas de controle manual.

Geração de Relatórios: Criação de tabelas dinâmicas utilizando a biblioteca pandas, atribuindo os dados finais para cada colaborador do time na tabela principal.

Atualização Dinâmica: O código identifica e atualiza automaticamente o mês vigente para a geração do relatório.

🛠️ Tecnologias Utilizadas
Python

Pandas (Tratamento de dados e tabelas dinâmicas)

Selenium (Navegação nos sistemas Podio e Agilis)

openpyxl (Edição das planilhas)

🚀 Como executar
Você precisa já ter a planilha que é puxada do sistema Sap e do site mrv bússula, e deixa-las salvas na pasta dowloads no seu computador.

Daí é só rodar o código que ele faz o resto.
