# Lembretes de ponto pelo WhatsApp

O programa lê os horários dos funcionários de um CSV e envia uma mensagem pelo
WhatsApp Web no horário correspondente a cada marcação. Ele deve permanecer
aberto durante o expediente e usa o fuso horário configurado no computador.

## Requisitos

- Python 3.9 ou superior.
- Google Chrome ou outro navegador padrão instalado.
- Uma sessão do WhatsApp Web autenticada no navegador padrão.

Instale a dependência, no PowerShell, dentro desta pasta:

```powershell
py -m pip install -r requirements.txt
```

Antes de iniciar o programa, abra `https://web.whatsapp.com` no navegador padrão
e autentique a sessão. A automação usa essa sessão para enviar as mensagens.
Mantenha o computador ligado, o navegador visível e sem minimizar, e não use o
mouse ou o teclado enquanto uma mensagem estiver sendo enviada: o programa
clica automaticamente no botão de envio do WhatsApp Web.

## Arquivo CSV

Por padrão, o arquivo deve se chamar `funcionarios.csv` e ficar nesta pasta.
O programa aceita arquivos CSV (UTF-8) ou planilhas Excel (`.xlsx`). Uma
planilha Excel também é reconhecida pelo conteúdo mesmo se tiver sido salva
com a extensão `.csv`; nesse caso, o programa lê a primeira planilha ativa.
Use seis colunas, nesta ordem: nome, telefone com código do país, entrada,
saída para almoço, volta do almoço e saída. No CSV, o separador pode ser
vírgula ou ponto e vírgula; arquivos exportados pelo Excel em português
normalmente usam ponto e vírgula.

Exemplo:

```csv
NOME DO FUNCIONARIO;TELEFONE DO FUNCIONARIO;ENTRADA;SAIDA ALMOÇO;VOLTA ALMOÇO;SAIDA
Maria Silva;+5511999999999;08:00;12:00;13:00;17:00
João Souza;+5521999999999;09:00;12:30;13:30;18:00
```

O cabeçalho é opcional, mas a ordem das colunas é obrigatória. Telefones devem
incluir o código do país (por exemplo, `+55` para o Brasil); se o `+` for
omitido, o programa o acrescenta. Horários usam o formato de 24 horas `HH:MM`.
Na planilha Excel, formate telefones como texto para preservar o `+` e zeros
iniciais.
Uma célula de horário vazia desativa somente aquele lembrete. Linhas em branco
são ignoradas.

## Uso

Valide o CSV sem enviar mensagens:

```powershell
py input.py --check
```

Inicie a automação:

```powershell
py input.py
```

Também é possível informar outros caminhos:

```powershell
py input.py --csv "C:\caminho\funcionarios.csv" --state-file "C:\caminho\envios.json"
```

O programa verifica os horários continuamente e envia uma mensagem diferente
para cada marcação: entrada, saída para almoço, volta do almoço e saída. Cada
mensagem inclui o nome e o horário. Os horários do CSV são carregados ao
iniciar; para aplicar alterações, encerre com `Ctrl+C` e inicie novamente.

O arquivo `estado_envios.json` registra as mensagens enviadas e evita que o
mesmo lembrete seja repetido no mesmo dia após reiniciar o programa. Mantenha
apenas uma instância em execução. Se o computador estiver desligado ou o
programa fechado no horário, o lembrete daquele horário não será enviado
retroativamente. O envio real pode atrasar enquanto o navegador abre ou envia
mensagens, especialmente quando vários lembretes vencem ao mesmo tempo.

Use a automação somente com autorização da empresa e dos funcionários,
respeitando as regras aplicáveis ao WhatsApp. O WhatsApp Web pode mudar e
interromper automações não oficiais; este programa não substitui uma integração
oficial do WhatsApp Business.
