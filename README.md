# Automação de NF-e/XML para Prévia de ICMS e Controle de Comissões

Projeto demonstrativo de portfólio desenvolvido em Python para processar arquivos XML de NF-e e transformar dados fiscais e operacionais em uma planilha Excel.

## Objetivos

- consolidar uma prévia mensal do ICMS por filial;
- identificar itens de Diesel e ARLA;
- organizar nota, chave de acesso, data, produto, litros, valor, placa e motorista;
- separar abastecimentos de frota própria e terceiros;
- calcular uma comissão demonstrativa para motoristas terceiros;
- reduzir conferências manuais e melhorar a rastreabilidade das informações.

## Privacidade e segurança

Este repositório é uma **versão pública e anonimizada para portfólio**. Ele não contém código operacional proprietário, XMLs reais, CNPJs válidos, chaves de NF-e, placas, nomes de motoristas, valores, credenciais, caminhos internos ou regras específicas de qualquer empresa.

Todos os dados dos arquivos de exemplo são fictícios e servem exclusivamente para demonstração.

## Estrutura

```text
automacao-nfe-xml-icms-comissoes/
├── filiais_exemplo.json
├── motoristas_exemplo.json
├── nfe_demonstracao.xml
├── processar_notas_demo.py
├── requirements.txt
└── README.md
```

## Como executar

1. Instale o Python 3.11 ou superior.
2. Instale a dependência:

```bash
pip install -r requirements.txt
```

3. Copie XMLs **fictícios ou devidamente anonimizados** para uma pasta `entrada/` ao lado do programa.
4. Execute:

```bash
python processar_notas_demo.py
```

5. Consulte `saida/relatorio_nfe_demo.xlsx` (a pasta é criada automaticamente).

Para testar rapidamente, basta executar o programa: se `entrada/` estiver vazia, ele usa `nfe_demonstracao.xml` automaticamente.

## Saída

O arquivo Excel possui três abas:

- `PREVIA_ICMS`: total de ICMS por filial e mês;
- `ABASTECIMENTOS`: itens de Diesel e ARLA encontrados;
- `COMISSOES`: valores demonstrativos de comissão de terceiros.

## Tecnologias e competências

Python, XML, Excel, automação de processos, análise de dados, regras de negócio e organização de informações fiscais e logísticas.

## Aviso

Este projeto não substitui escrituração fiscal, apuração tributária oficial ou validação por profissional contábil. Os resultados representam apenas uma prévia gerencial baseada nos XMLs processados.
