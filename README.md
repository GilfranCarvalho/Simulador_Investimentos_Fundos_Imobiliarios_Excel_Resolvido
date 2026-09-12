O projeto foi organizado com foco em **documentação técnica clara, rastreabilidade dos cálculos e facilidade de compartilhamento**, utilizando o GitHub como canal para armazenar a planilha, explicar sua lógica e permitir futuras evoluções.

## Objetivo do projeto

A ferramenta foi criada para responder, de forma simples e parametrizável, perguntas como:

- Quanto posso acumular investindo um valor fixo todos os meses?
- Qual seria o patrimônio estimado após 2, 5, 10, 20 ou 30 anos?
- Qual seria uma estimativa de dividendos mensais a partir do patrimônio acumulado?
- Como distribuir o aporte mensal entre diferentes tipos de FIIs de acordo com um perfil de investidor?

## Estrutura do arquivo

A planilha possui duas abas principais:

### `APP`

É a interface principal da ferramenta. Nela estão concentrados os parâmetros de entrada, os resultados da simulação, os cenários de longo prazo e a distribuição sugerida do aporte mensal.

Principais blocos:

| Bloco | Finalidade |
|---|---|
| **CONFIGURAÇÕES** | Define salário e rendimento utilizado para a estimativa de dividendos. |
| **INVESTIMENTO MENSAL** | Recebe aporte mensal, prazo e taxa mensal de rendimento e calcula o patrimônio projetado. |
| **Cenários** | Compara o patrimônio e os dividendos estimados em diferentes horizontes. |
| **PERFIL** | Define o perfil de alocação escolhido. |
| **VALOR A SER INVESTIDO POR MÊS** | Exibe o aporte utilizado na divisão da carteira. |
| **TIPO DE FII** | Apresenta o percentual sugerido e o valor financeiro por categoria. |

### `Planilha2`

Funciona como uma **tabela de apoio**, armazenando os percentuais sugeridos para cada combinação de perfil e tipo de FII.

Os perfis disponíveis são:

- Conservador
- Moderado
- Agressivo

Os tipos de FII considerados são:

- PAPEL
- TIJOLO
- HÍBRIDOS
- FOFs
- DESENVOLVIMENTO
- HOTELARIAS

Essa aba permite separar os parâmetros de negócio da interface principal e facilita a manutenção da ferramenta.

## Como a ferramenta foi construída

A lógica da planilha foi organizada em quatro etapas principais.

### 1. Definição das entradas

Na aba `APP`, o usuário informa ou ajusta os principais parâmetros da simulação:

- **Salário:** R$ 2.000,00 no cenário atual da planilha.
- **Rendimento da carteira para dividendos:** 0,6% ao mês.
- **Aporte mensal:** R$ 200,00.
- **Prazo:** 5 anos.
- **Taxa de rendimento mensal:** 1,079% ao mês.

A sugestão automática de investimento mensal é calculada como **30% do salário**:

```excel
=D12*30%
```

Com salário de R$ 2.000,00, a sugestão resulta em R$ 600,00.

> Observação: o valor efetivamente utilizado na simulação do cenário exibido é o aporte informado em `D17`, que atualmente está em R$ 200,00.

## 2. Cálculo do patrimônio acumulado

O patrimônio futuro é calculado utilizando a função financeira `FV` do Excel.

A fórmula principal da planilha é:

```excel
=FV(taxa_mensal,qtd_anos*12,aporte*-1)
```

Onde:

| Nome | Referência | Significado |
|---|---|---|
| `taxa_mensal` | `APP!D19` | Taxa de rendimento mensal da simulação. |
| `qtd_anos` | `APP!D18` | Quantidade de anos do investimento. |
| `aporte` | `APP!D17` | Valor investido mensalmente. |

O número de períodos é convertido de anos para meses por meio de:

```text
anos × 12
```

### Interpretação financeira

A fórmula considera uma sequência de aportes mensais e aplica uma taxa mensal de crescimento sobre o investimento ao longo do período. A função `FV` é apropriada para projetar o valor futuro de pagamentos periódicos com taxa constante.

No cenário atual da planilha:

- aporte mensal: **R$ 200,00**;
- prazo: **5 anos**;
- taxa mensal: **1,079%**;
- patrimônio projetado: aproximadamente **R$ 16.755,38**.

## 3. Cálculo de dividendos mensais

Depois de calcular o patrimônio, a ferramenta estima o valor mensal de dividendos multiplicando o patrimônio pela taxa definida em `rendimento_carteira`.

Fórmula utilizada:

```excel
=patrimonio*rendimento_carteira
```

Os nomes definidos do Excel são:

```text
patrimonio            = APP!$D$20
rendimento_carteira   = APP!$D$13
```

No cenário atual:

```text
Patrimônio × 0,6%
```

Com patrimônio de aproximadamente R$ 16.755,38, o resultado é aproximadamente:

```text
R$ 16.755,38 × 0,006 = R$ 100,53/mês
```

### Importante

Esse cálculo representa uma **estimativa matemática baseada em uma taxa fixa de 0,6% ao mês**. Ele não significa que os FIIs necessariamente pagarão esse valor em todos os meses, pois dividendos reais variam conforme os ativos, resultados dos fundos, vacância, contratos, mercado e outros fatores.

## 4. Construção dos cenários

A seção de cenários usa a mesma lógica de valor futuro para comparar diferentes horizontes:

- 2 anos
- 5 anos
- 10 anos
- 20 anos
- 30 anos

A fórmula do patrimônio para cada cenário é:

```excel
=FV($D$19,$A24*12,$D$17*-1)
```

O resultado do patrimônio de cada linha é então utilizado para estimar o dividendo correspondente:

```excel
=C24*rendimento_carteira
```

A estrutura permite alterar o aporte e a taxa uma única vez e observar automaticamente o impacto em todos os horizontes.

### Valores apresentados na planilha

| Horizonte | Patrimônio projetado | Dividendos mensais estimados |
|---:|---:|---:|
| 2 anos | R$ 5.445,53 | R$ 32,67 |
| 5 anos | R$ 16.755,38 | R$ 100,53 |
| 10 anos | R$ 48.656,84 | R$ 291,94 |
| 20 anos | R$ 225.039,68 | R$ 1.350,24 |
| 30 anos | R$ 864.433,93 | R$ 5.186,60 |

Esses valores são projeções baseadas exclusivamente nos parâmetros existentes na planilha e não constituem garantia de rentabilidade.

## Organização técnica das fórmulas

Para deixar a manutenção mais simples, a planilha utiliza **nomes definidos** no Excel em vez de depender apenas de referências de células.

Os nomes definidos identificados no arquivo são:

| Nome definido | Célula | Uso |
|---|---|---|
| `salario` | `APP!$D$12` | Salário informado. |
| `rendimento_carteira` | `APP!$D$13` | Taxa usada na estimativa dos dividendos. |
| `sugestao_investimento` | `APP!$D$14` | Sugestão de aporte equivalente a 30% do salário. |
| `aporte` | `APP!$D$17` | Aporte mensal da simulação. |
| `qtd_anos` | `APP!$D$18` | Prazo da simulação. |
| `taxa_mensal` | `APP!$D$19` | Taxa mensal usada no cálculo do valor futuro. |
| `patrimonio` | `APP!$D$20` | Patrimônio projetado. |

Essa abordagem melhora a legibilidade das fórmulas e aproxima a planilha de um modelo técnico reutilizável.

## Lógica de distribuição dos FIIs

A aba `Planilha2` armazena os percentuais sugeridos para cada perfil.

### Perfil Moderado

A configuração atualmente utilizada na aba `APP` é:

| Tipo de FII | Percentual | Valor sobre aporte de R$ 200 |
|---|---:|---:|
| PAPEL | 32% | R$ 64,00 |
| TIJOLO | 35% | R$ 70,00 |
| HÍBRIDOS | 8% | R$ 16,00 |
| FOFs | 5% | R$ 10,00 |
| DESENVOLVIMENTO | 10% | R$ 20,00 |
| HOTELARIAS | 10% | R$ 20,00 |
| **Total** | **100%** | **R$ 200,00** |

O percentual de cada categoria é recuperado por meio de `VLOOKUP` a partir de uma chave formada pela combinação entre perfil e tipo de FII.

Exemplo:

```excel
=VLOOKUP($C$32&"-"&B36,Planilha2!$A:$D,4,FALSE)
```

Depois de obter o percentual, o valor financeiro é calculado por:

```excel
=C36*$C$33
```

Ou seja:

```text
Valor da categoria = Percentual sugerido × Aporte mensal
```

O total da distribuição é conferido por:

```excel
=SUM(D36:D41)
```

Assim, o modelo consegue transformar automaticamente um percentual de alocação em valores monetários.

## Fluxo da solução

```text
Entrada de dados
      ↓
Salário + aporte + prazo + taxa mensal
      ↓
Cálculo do valor futuro (FV)
      ↓
Patrimônio projetado
      ↓
Estimativa de dividendos
      ↓
Cenários de 2 a 30 anos
      ↓
Definição do perfil de investidor
      ↓
Busca dos percentuais na tabela de apoio
      ↓
Distribuição do aporte por tipo de FII
```

## Tecnologias e recursos utilizados

- **Microsoft Excel** para construção do modelo financeiro.
- **Função `FV`** para cálculo do valor futuro de aportes periódicos.
- **Nomes definidos** para tornar as fórmulas mais legíveis e reutilizáveis.
- **`VLOOKUP`** para buscar percentuais conforme perfil e categoria.
- **`SUM`** para validação do total distribuído.

## Boas práticas de documentação técnica adotadas

A documentação registra:

- objetivo da solução;
- estrutura das planilhas;
- entradas e saídas;
- fórmulas financeiras utilizadas;
- lógica de relacionamento entre as abas;
- regras de distribuição da carteira;
- limitações e premissas do modelo;
- possibilidade de evolução futura.

Essa organização facilita a compreensão por outras pessoas e cria uma base para controle de versões e manutenção da ferramenta.

## Premissas e limitações

A simulação é um modelo matemático e deve ser interpretada como **projeção**, não como promessa de rentabilidade.

As principais premissas observadas no arquivo são:

- taxa mensal de rendimento constante para a projeção;
- aportes mensais constantes durante o período simulado;
- estimativa de dividendos baseada em uma taxa fixa de 0,6% ao mês;
- ausência, no modelo atual, de inflação, impostos, custos operacionais e mudanças na taxa de retorno;
- distribuição de FIIs baseada em percentuais definidos previamente por perfil.

Uma evolução futura do projeto pode incluir inflação, crescimento ou redução do aporte, taxas variáveis, reinvestimento explícito dos dividendos, cenários pessimista/base/otimista e comparação entre rentabilidade nominal e real.

# Metodologia de cálculo

## Patrimônio futuro

A projeção utiliza a função `FV` do Excel:

```excel
=FV(taxa_mensal,qtd_anos*12,aporte*-1)
```

A lógica é baseada em aportes mensais com taxa mensal constante.

## Dividendos estimados

A estimativa mensal de dividendos é calculada por:

```excel
=patrimonio*rendimento_carteira
```

Na configuração atual, `rendimento_carteira` corresponde a 0,6% ao mês.

## Alocação por perfil

Os percentuais de alocação são armazenados na aba `Planilha2`. A aba `APP` cria uma chave no formato:

```text
Perfil-TipoDeFII
```

A porcentagem correspondente é recuperada por `VLOOKUP` e multiplicada pelo aporte mensal.

