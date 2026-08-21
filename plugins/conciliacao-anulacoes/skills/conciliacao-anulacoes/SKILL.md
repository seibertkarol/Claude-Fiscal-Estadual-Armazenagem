---
name: conciliacao-anulacoes
description: Executa a conciliação de ANULAÇÕES da SLC Agrícola — pareia cada nota de anulação (referência com sufixo -002 na Planilha6) com a nota de origem que ela está anulando, e pinta os dois lados de roxo. Use sempre que a usuária disser "faça a conciliação da anulação", "conciliar anulações", "parear as anulações", "achar o par das -002", ou qualquer variação. É um processo SEPARADO da conciliação de armazenagem (remessas x retornos) — não confundir nem misturar os dois. Sempre perguntar qual é o arquivo Excel e se existe relatório do ERP das anulações, nunca assumir.
---

# Conciliação de Anulações (-002) — SLC Agrícola

## O que este processo faz

Toda nota de **anulação** na Planilha6 aparece com o sufixo **`-002`** na coluna
Referência (ex: `000000990-002`) e tem valor negativo. Ela anula uma **nota de
origem** lançada antes, que aparece com sufixo `-001` e valor positivo.

Este processo encontra o par de cada anulação e pinta os dois lados de **ROXO
(`CC99FF`)**.

**A garantia matemática: filtrar a Planilha6 pela cor ROXA deve somar
exatamente R$0,00.** Se não somar, o pareamento está errado — ver
Troubleshooting.

> **Não confundir com a conciliação de armazenagem.** Aquela concilia remessas
> contra notas de retorno e usa as cores rosa/laranja/azul/amarelo/verde. Esta
> aqui só trata das anulações `-002` e usa exclusivamente o roxo. São processos
> independentes, com scripts e planilhas próprios. Se a usuária pedir "a
> conciliação" sem especificar, **pergunte qual das duas**.

## Arquivos envolvidos

| Arquivo | Função |
|---------|--------|
| `scripts/concilia_anulacoes.py` | Faz o pareamento e pinta (recebe `--arquivo` e opcionalmente `--relatorio`) |
| Excel da fazenda (varia) | Planilha com a aba Planilha6 — pode ser crua ou já conciliada |
| Relatório do ERP (varia, opcional) | Export do SAP com as notas de anulação e seus documentos de origem |

## Passo a passo

### 1. Perguntar os inputs — sempre

**Nunca assuma qual arquivo usar**, mesmo que pareça óbvio pelo contexto:

- Qual é o **arquivo Excel**? (pode ser o `EXPORT_*.xlsx` cru ou um
  `*_CONCILIADO.xlsx` — o script detecta sozinho e ajusta as colunas)
- Existe **relatório do ERP** das anulações? Se sim, qual é o arquivo?

O relatório é opcional, mas **muda muito o resultado** — ver a seção seguinte.
Se a usuária não souber se tem, explique o que é e ofereça rodar sem ele
primeiro para ela ver quantos casos ficam ambíguos.

### 2. Rodar

```bash
py -3.14 "CAMINHO_DA_SKILL/scripts/concilia_anulacoes.py" --arquivo="PLANILHA.xlsx" --relatorio="EXPORT_ERP.xlsx"
```

Sem o relatório do ERP:

```bash
py -3.14 "CAMINHO_DA_SKILL/scripts/concilia_anulacoes.py" --arquivo="PLANILHA.xlsx"
```

O script salva em um nome novo (`..._ANULACOES.xlsx`, ou `_v2`, `_v3`... se já
existir) — nunca sobrescreve nem trava por arquivo aberto no Excel.

### 3. Conferir e reportar

Leia o resumo do console e informe à usuária:

```
Anulacoes (-002) na planilha : 44
Pares confirmados            : 42   (84 linhas pintadas)
    via ERP         : 39
    via valor unico : 3
Pendentes de revisao manual  : 2
SALDO DAS LINHAS ROXAS       : R$ 0,00     <-- TEM que ser R$0
```

**Sempre confira o saldo.** Se não for R$0,00 o script avisa — nesse caso não
entregue o arquivo como pronto, investigue primeiro.

Diga o nome exato do arquivo gerado e explique que a aba **"Anulações"** traz
o detalhe de cada par (roxo) e cada pendência (vermelho), com os dois lados,
valores, datas e as linhas do Excel.

---

## Os dois métodos de pareamento

### 1. ERP — preferencial, determinístico

O relatório do SAP traz explicitamente qual documento cada anulação está
anulando:

| Coluna | Cabeçalho | Conteúdo |
|---|---|---|
| A | `Nº documento` | doc_sap da anulação |
| AF | `Número de nota fiscal eletrônica` | **NF da anulação** |
| BE | `Nº doc.original` | doc_sap da origem |
| BF | `Nº NOTA ORIGEM` | **NF da origem** |

O script usa **AF → BF**. As colunas são localizadas pelo nome do cabeçalho,
com as posições acima como fallback.

Detalhe importante: **no relatório do ERP a anulação aparece sem o `-002`**,
só o número da nota. Na Planilha6 ela tem o sufixo.

A coluna BF (`Nº NOTA ORIGEM`) é montada pela usuária traduzindo os doc_sap de
`Nº doc.original` para número de nota — porque na Planilha6 as origens estão
em formato de nota, não de doc_sap.

### 2. Valor único — fallback

Quando a anulação não está no relatório, procura uma linha com o valor
exatamente invertido. **Só aceita quando há UMA ÚNICA candidata e nenhuma
outra anulação disputa o mesmo valor.**

Essa restrição é essencial: valores pequenos se repetem muito (num caso real,
R$35,73 aparecia em **432 linhas**). Sem ela o script casaria por coincidência.
É melhor deixar ambíguo e revisar manualmente do que criar um par falso.

**Validação já feita:** num caso real com 44 anulações, 13 foram resolvidas
pelos dois métodos de forma independente e **os 13 apontaram para a mesma
origem, zero divergência**.

---

## As três regras que garantem o saldo zero

Estas regras existem porque cada uma delas já causou um saldo diferente de zero
num caso real:

1. **Parear por LINHA, nunca por número de nota.** Um mesmo número pode ter
   várias linhas na planilha e só uma é o par. Pintar todas as linhas daquele
   número quebra o saldo.

2. **Linhas já estornadas nunca entram como origem.** Se a linha tem "X" em
   `estornado` ou em `Documento de estorno`, ela já foi neutralizada por outra
   linha — o par delas soma zero sozinho. Usar uma dessas como origem deixa a
   anulação sem contrapartida real.

3. **Cada linha de origem só serve a UMA anulação.** Se duas anulações apontam
   para a mesma origem, a segunda vira PENDENTE. Caso real: duas anulações
   (NF 994 e NF 1006) apontavam para o mesmo doc_sap de origem no ERP, mas a
   Planilha6 tinha uma única linha daquela nota.

---

## Troubleshooting

**O saldo roxo não deu R$0,00:**
O script avisa no console. As causas prováveis, nesta ordem:
- Alguma linha foi pintada por número em vez de por linha específica
- Uma origem estornada entrou como par
- Duas anulações consumindo a mesma origem

Confira a aba "Anulações" — a coluna `Diferença` de cada par deve ser 0,00.

**Muitos casos "ambíguo — N candidatas com o mesmo valor":**
Normal quando se roda **sem** o relatório do ERP. Peça o relatório à usuária —
ele resolve quase tudo de forma determinística. Para ajudar a baixar, informe
o período das anulações pendentes (o script lista as NFs; a data está na aba).

**"origem não está na Planilha6":**
A anulação existe no ERP mas a nota de origem não está na planilha. Pode ser
conta, centro ou período diferente. Reporte para verificação manual.

**"anulação já estornada":**
A própria linha `-002` tem X de estorno. Ela já foi anulada por outra via e não
precisa de par.

**Colunas desalinhadas numa planilha nova:**
O script detecta as colunas pelo NOME do cabeçalho (Referência, Data de
lançamento, Valor em moeda da empresa, estornado, Documento de estorno), com
fallback para posições fixas. Se uma exportação usar nomes diferentes, ajuste
as keywords em `_find_col`.
