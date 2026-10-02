# Central de Gestão de Obras — V3 Online

## Perfis
- Administrador: importa e publica RDO + Máquinas/Caminhões.
- Engenharia: somente consulta, filtros e BI. Não vê upload.

## Publicação
O projeto está pronto para hospedagem Streamlit. Configure dois segredos:
ADMIN_PASSWORD = senha do administrador
VIEWER_PASSWORD = senha compartilhada/definida para engenharia

## Importante sobre persistência
Os arquivos publicados ficam no armazenamento do servidor. Em hospedagens com filesystem efêmero,
um reinício/redeploy pode apagar os uploads. Para produção permanente, use um host com disco persistente
ou conecte banco/storage externo. O código já separa a camada de publicação da camada de visualização.


## V4 — Clima
Lê automaticamente a aba `Lançamento Diário` do RDO:
- dias ensolarados
- dias de chuva
- dias nublados
- dias praticáveis (Sim ou Parcial)
- dias impraticáveis (Não)
- horas paralisadas
A contagem é consolidada por Obra + Data para não duplicar um dia.
