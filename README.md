🤖 Sistema de Automação CMS V2.0
Sistema avançado de automação para processamento de dados em sistemas de gestão de clientes (CMS) com interface gráfica intuitiva

https://img.shields.io/badge/Python-3.8+-blue.svg
https://img.shields.io/badge/Automation-Desktop%2520GUI-orange.svg
https://img.shields.io/badge/version-2.0-green.svg
https://img.shields.io/badge/platform-Windows%2520%257C%2520Linux-lightgrey.svg
https://img.shields.io/badge/license-MIT-yellow.svg

📖 Sobre o Projeto
Sistema de automação completo desenvolvido em Python para processamento inteligente de grandes volumes de registros em plataformas CMS (Customer Management Systems). A ferramenta combina automação de interface gráfica com OCR inteligente para otimizar fluxos de trabalho complexos.

✨ Funcionalidades Principais
✅ Interface Gráfica Amigável - Controle completo através de GUI intuitiva

✅ Processamento em Lote - Automação de múltiplos registros sequenciais

✅ OCR Inteligente - Reconhecimento de texto em tempo real com Pytesseract

✅ Sistema de Coordenadas - Captura e gestão automatizada de posições na tela

✅ Detecção de Erros Avançada - Identificação visual de falhas no sistema

✅ Logs Detalhados - Sistema completo de registro e filtragem de atividades

✅ Processamento Condicional - Lógica adaptativa baseada no estado dos registros

🛠️ Tecnologias Utilizadas
Linguagem Principal
Python 3.8+ - Lógica principal de automação e interface

Bibliotecas de Automação
PyAutoGUI - Controle preciso de mouse e teclado

Keyboard - Detecção de eventos de teclado em tempo real

Pyperclip - Manipulação eficiente da área de transferência

Mouse - Captura avançada de eventos do mouse

Interface Gráfica
Tkinter - Interface gráfica nativa do Python

ttk - Componentes temáticos modernos

Processamento de Imagens & OCR
Pytesseract - Reconhecimento óptico de caracteres (OCR)

PIL (Pillow) - Processamento avançado de imagens e capturas de tela

Utilitários
JSON - Armazenamento e gestão de configurações

Time - Controle preciso de delays e sincronização

Re - Processamento de expressões regulares

OS - Gestão de arquivos e diretórios

Datetime - Timestamps e gestão temporal

Subprocess - Execução de processos externos

📦 Instalação
Pré-requisitos do Sistema
bash
# Instalar Tesseract OCR
# Windows: https://github.com/UB-Mannheim/tesseract/wiki
# Linux (Debian/Ubuntu): 
sudo apt update && sudo apt install tesseract-ocr

# Linux (RedHat/CentOS):
sudo yum install tesseract
Instalação das Dependências Python
bash
# Instalar pacotes necessários
pip install pyautogui pyperclip keyboard pillow pytesseract mouse

# Ou via requirements.txt
pip install -r requirements.txt
Configuração do Tesseract OCR
python
# No arquivo cms_grafica.py - Configuração automática incluída
pytesseract.pytesseract.tesseract_cmd = "/usr/bin/tesseract"  # Linux
# Para Windows, descomente e ajuste:
# pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'
🚀 Como Usar
Execução do Sistema
bash
# Executar aplicação principal
python cms_grafica.py

# Executar capturador de coordenadas
python coordenada.py
Fluxo de Trabalho Completo
Configuração Inicial

text
Executar coordenada.py → Capturar coordenadas da interface → Salvar automaticamente
Preparação de Dados

text
Carregar arquivo .txt com dados → Verificar prévia → Iniciar automação
Processamento Automatizado

text
Inserção de dados → Verificação de erros → Processamento OCR → 
Análise de estado → Ação condicional → Próximo registro
Estrutura de Arquivos
text
sistema-automacao-cms/
├── cms_grafica.py          # Aplicação principal com GUI
├── coordenada.py           # Capturador de coordenadas
├── coordenadas.json        # Configurações de coordenadas (gerado)
├── automacao_log.txt       # Logs de execução (gerado)
├── aprendizado.py          # Script de aprendizado (referência)
└── imagens/                # Diretório para imagens de referência
    ├── erro.png
    ├── erro1.png
    └── erro2.png
🔧 Funcionalidades Detalhadas
🎯 Aplicação Principal (cms_grafica.py)
Interface Gráfica:

Tela de carregamento com progresso

Controles intuitivos de automação

Visualização de logs em tempo real

Sistema de filtros avançado para logs

Recursos Principais:

Inicialização automática de coordenadas

Processamento em lote com progresso

Detecção visual de erros

Sistema de recuperação de falhas

Estatísticas detalhadas de execução

🖱️ Capturador de Coordenadas (coordenada.py)
Características:

Interface dedicada para captura de coordenadas

Detecção inteligente de cliques fora da janela

Progresso visual do processo de captura

Salvamento automático em JSON

Reinicialização e continuação seguras

Coordenadas Capturadas:

python
[
    "pf", "contador_btn", "ok_btn", "cliente", "seguinte", "processamento",
    "processamento_seguinte", "imprimir", "cancelar", "processamento_sim",
    "registado", "registado_seguinte", "registado_cmp1", "registado_cmp_potencia",
    "propriedade", "gis_x", "gis_y", "estado_instalacao", "num_luz", "quartos",
    "registro", "registado_sim", "anterior", "fim_rolagem"
]
🔍 Sistema OCR Avançado
python
# Captura e análise inteligente de texto
screenshot = pyautogui.screenshot(region=area)
texto_extraido = pytesseract.image_to_string(screenshot)

# Processamento de múltiplos estados
linhas = [linha.strip() for linha in texto_extraido.split('\n') if linha.strip()]
⚡ Estados de Processamento Detectados
"Registado" - Processamento completo e finalizado

"Em processamento" - Pendências que requerem ações específicas

"Introduzido" - Fluxo padrão de inserção inicial

"Suspeita de fraude" - Casos especiais com tratamento diferenciado

⚙️ Configuração
Arquivo de Coordenadas (coordenadas.json)
json
{
    "pf": [53, 250],
    "contador_btn": [366, 251],
    "ok_btn": [846, 475],
    "cliente": [490, 404],
    "seguinte": [512, 735],
    "processamento": [351, 356],
    "...": "..."
}
Configuração de Imagens de Erro
python
# Ajuste estos caminhos conforme seu ambiente
erro_img_path = "/home/zebito/Downloads/drive/cms/linux/erro.png"
erro_img_path1 = "/home/zebito/Downloads/drive/cms/linux/erro1.png" 
erro_img_path2 = "/home/zebito/Downloads/drive/cms/linux/erro2.png"
Formato do Arquivo de Dados
text
01317788162
54280992519
45135422488
78013515509
...
📊 Sistema de Logs e Monitoramento
Logs Detalhados
Timestamp preciso de todas as operações

Categorização por tipo de atividade

Filtragem avançada por data e categoria

Exportação automática para arquivo

Estatísticas em Tempo Real
python
print(f"Total processado: {total_processados}")
print(f"Introduzidos: {contador_introduzido}")
print(f"Em processamento: {contador_em_processamento}")
print(f"Registados: {contador_registados}")
print(f"Erros detectados: {numeros_errados}")
🚨 Gestão de Erros e Recuperação
Mecanismos de Detecção
Timeout de interface - Reinteligência automática

Imagens não encontradas - Fallbacks alternativos

Erros de OCR - Processamento condicional

Estados inesperados - Logs detalhados para debugging

Estratégias de Recuperação
Reintentos automáticos com delays progressivos

Múltiplos métodos de detecção de erro

Fallbacks para diferentes cenários

Interrupção segura com Ctrl+C

🔄 Fluxo de Processamento Completo
Fase 1: Inserção e Validação
text
Clique PF → Clique Contador → Colar Dado → Enter → Verificar Erro → Clique Cliente
Fase 2: Análise e Decisão
text
Captura Tela → OCR → Processar Texto → Identificar Estado → Selecionar Fluxo
Fase 3: Execução Condicional
text
Estado "Registado" → Fluxo Completo de Registro
Estado "Processamento" → Fluxo Simplificado de Acompanhamento
Estado "Introduzido" → Fluxo Padrão de Inserção
⚠️ Considerações Importantes
Requisitos de Sistema
Resolução de tela consistente durante a execução

Acesso administrativo para automação de interface

Ambiente estável sem interferências externas

Backup dos dados antes do processamento em lote

Limitações Conhecidas
Dependente da estabilidade da interface gráfica alvo

Requer calibração inicial das coordenadas

Sensível a mudanças no layout do CMS

Performance varia com a capacidade do sistema

🛡️ Boas Práticas Recomendadas
Antes da Execução
Execute backup completo dos dados

Teste em ambiente controlado com amostras pequenas

Verifique e calibre todas as coordenadas

Prepare plano de rollback para emergências

Durante a Execução
Monitore os logs constantemente

Mantenha o sistema estável sem interferências

Evite uso manual do computador durante o processamento

Tenha Ctrl+C preparado para interrupção rápida

🤝 Contribuindo para o Projeto
Como Contribuir
Reporte bugs através do sistema de Issues

Sugira melhorias no sistema de automação

Compartilhe configurações para diferentes ambientes

Documente casos de uso específicos

Áreas de Melhoria Futura
Interface de configuração gráfica mais avançada

Suporte a múltiplos layouts de CMS simultâneos

Sistema de templates para diferentes fluxos de trabalho

Relatórios em PDF automáticos com gráficos

API REST para integração com outros sistemas

📄 Licença
Distribuído sob licença MIT. Veja o arquivo LICENSE para mais informações.

👤 Autor
Zerdone Rocha

💼 LinkedIn: Zerdone Rocha

🐙 GitHub: HackMaster300

📈 Resultados e Benefícios

⏱️ Eficiência Operacional
Redução de até 90% no tempo de processamento manual

Processamento contínuo 24/7 sem intervenção humana

Escalabilidade para volumes ilimitados de dados

🎯 Precisão e Confiabilidade
Eliminação de erros humanos na digitação

Validação automática e consistente de dados

Processamento uniforme em todos os registros

📊 Controle e Transparência
Logs detalhados para auditoria e compliance

Métricas precisas de performance e eficiência

Detecção proativa de problemas e otimizações

<div align="center">
⚡ Automatize processos repetitivos e foque no que realmente importa!
https://img.shields.io/github/stars/HackMaster300/cms-automation?style=social
https://img.shields.io/github/forks/HackMaster300/cms-automation?style=social

🚀 Pronto para revolucionar seu fluxo de trabalho com automação inteligente!

</div>
