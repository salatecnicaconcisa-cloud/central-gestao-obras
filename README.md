
# Central de Gestão de Obras

Primeira versão funcional do software para importar:
- RDO (.xlsx/.xlsm)
- Máquinas/Caminhões (.xlsx/.xlsm)

## Windows — iniciar
1. Instale Python 3.11 ou superior.
2. Dê dois cliques em `INICIAR_SISTEMA.bat`.
3. Na primeira execução, o Windows instalará as dependências.
4. O navegador abrirá o sistema.

## Uso
- Importe as duas planilhas no menu lateral.
- Escolha a obra e o período.
- O dashboard recalcula HH, HM, funcionários únicos, equipamentos únicos, parada, KM, combustível e HH/HM.
- Funcionários são identificados por Data + Obra + Nome.
- Equipamentos são identificados por Data + Obra + Prefixo.
- Os cards de quantidade contam nomes/prefixos únicos no período.

Observação: esta versão não grava os arquivos em servidor. O processamento ocorre durante a sessão do sistema.
