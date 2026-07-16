# MEMÓRIA DE CÁLCULO - FLUXO DE CAIXA MENSAL
## Loteamento Papagaios

---

### MÊS 01 - JANEIRO (Ano 1)

---

## 📋 PARÂMETROS GERAIS

| Parâmetro | Valor | Unidade |
|-----------|-------|---------|
| Preço de venda por lote | 80.000,00 | R$ |
| Total de lotes | 300 | un. |
| VGV Total | 24.000.000,00 | R$ |
| % lotes à vista/banco | 70,0% | % |
| % lotes financiados empreendedor | 30,0% | % |
| Prazo financiamento | 120 | meses |
| Percentual de sinal | 10,0% | % |
| Taxa de corretagem | 6,0% | % venda |
| Taxa administração | 3,0% | % VGV |
| Taxa marketing | 2,0% | % VGV |
| Taxa contingência | 3,0% | % custos |
| Taxa impostos (Lucro Presumido) | 6,5% | % receita |
| % Empreendedor | 65,0% | % |
| % Terrenista | 35,0% | % |
| Capital próprio investido | 1.000.000,00 | R$ |
| Taxa custo financeiro | 1,5% | a.m. |
| Total infraestrutura | 5.430.000,00 | R$ |

---

## 📊 DADOS DO MÊS 01

### Entradas Editáveis:
- **Lotes Vendidos:** 10 un.
- **Infraestrutura executada:** 200.000,00 R$
- **Marketing gasto:** 50.000,00 R$

---

## 🧮 CÁLCULOS DETALHADOS

### 1️⃣ RECEITAS

#### 1.1. Receita à Vista
```
Receita à Vista = Lotes Vendidos × Preço × % Vista
Receita à Vista = 10 × 80.000,00 × 70%
Receita à Vista = 560.000,00
```

#### 1.2. Cálculo da Parcela (financiamento empreendedor)
```
Valor financiado por lote = Preço × (1 - Sinal)
Valor financiado = 80.000,00 × (1 - 10%)
Valor financiado = 72.000,00

Parcela mensal = Valor financiado ÷ Prazo
Parcela mensal = 72.000,00 ÷ 120
Parcela mensal = 600,00 por lote
```

#### 1.3. Parcelas Recebidas no Mês 01
```
No mês 1, recebemos parcelas dos lotes vendidos em todos os meses anteriores
(mês atual - prazo) até o mês atual.

Como é o mês 1, só recebemos das vendas do próprio mês 1:
Parcelas Recebidas = Lotes vendidos mês 1 × Parcela
Parcelas Recebidas = 10 × 600,00
Parcelas Recebidas = 6.000,00
```

#### 1.4. Total Receita
```
Total Receita = Receita à Vista + Parcelas Recebidas
Total Receita = 560.000,00 + 6.000,00
Total Receita = 566.000,00
```

---

### 2️⃣ CUSTOS OPERACIONAIS

#### 2.1. Corretagem
```
Corretagem = Lotes Vendidos × Preço × % Corretagem
Corretagem = 10 × 80.000,00 × 6%
Corretagem = 48.000,00
```

#### 2.2. Administração Mensal
```
Admin Mensal = (VGV Total × % Admin) ÷ 96 meses
Admin Mensal = (24.000.000,00 × 3%) ÷ 96
Admin Mensal = 720.000,00 ÷ 96
Admin Mensal = 7.500,00
```

#### 2.3. Custos Diretos
```
Custos Diretos = Infraestrutura + Marketing + Corretagem + Admin
Custos Diretos = 200.000,00 + 50.000,00 + 48.000,00 + 7.500,00
Custos Diretos = 305.500,00
```

#### 2.4. Contingência
```
Contingência = Custos Diretos × % Contingência
Contingência = 305.500,00 × 3%
Contingência = 9.165,00
```

#### 2.5. Total Custos
```
Total Custos = Custos Diretos + Contingência
Total Custos = 305.500,00 + 9.165,00
Total Custos = 314.665,00
```

---

### 3️⃣ RESULTADO BRUTO

```
Resultado Bruto = Total Receita - Total Custos
Resultado Bruto = 566.000,00 - 314.665,00
Resultado Bruto = 251.335,00
```

---

### 4️⃣ IMPOSTOS (calculados ANTES do pagamento ao terrenista)

```
Impostos = Total Receita × % Impostos
Impostos = 566.000,00 × 6,5%
Impostos = 36.790,00
```

---

### 5️⃣ RESULTADO APÓS IMPOSTOS

```
Resultado após Impostos = Resultado Bruto - Impostos
Resultado após Impostos = 251.335,00 - 36.790,00
Resultado após Impostos = 214.545,00
```

---

### 6️⃣ PAGAMENTO AO TERRENISTA

```
Pagamento Terrenista = MAX(0, Resultado após Impostos) × % Terrenista
Pagamento Terrenista = MAX(0, 214.545,00) × 35%
Pagamento Terrenista = 214.545,00 × 0,35
Pagamento Terrenista = 75.090,75
```

---

### 7️⃣ RESULTADO DO EMPREENDEDOR

```
Resultado Empreendedor = Resultado após Impostos - Pagamento Terrenista
Resultado Empreendedor = 214.545,00 - 75.090,75
Resultado Empreendedor = 139.454,25
```

---

### 8️⃣ CUSTO FINANCEIRO

```
No mês 1, o custo financeiro é calculado sobre o capital investido:

Custo Financeiro = Capital × Taxa Financeiro
Custo Financeiro = 1.000.000,00 × 1,5%
Custo Financeiro = 15.000,00
```

**Nota:** A partir do mês 2, o custo financeiro será calculado sobre o saldo
acumulado negativo do mês anterior (se houver).

---

### 9️⃣ FLUXO LÍQUIDO DO MÊS

```
Fluxo Líquido = Resultado Empreendedor - Custo Financeiro
Fluxo Líquido = 139.454,25 - 15.000,00
Fluxo Líquido = 124.454,25
```

---

### 🔟 FLUXO ACUMULADO

```
No mês 1, descontamos o capital investido:

Fluxo Acumulado = Fluxo Líquido - Capital Investido
Fluxo Acumulado = 124.454,25 - 1.000.000,00
Fluxo Acumulado = -875.545,75
```

**Interpretação:** O projeto ainda está com saldo negativo de R$ 875.545,75
após o primeiro mês, pois o capital investido foi de R$ 1.000.000,00.

---

## 📈 RESUMO DO MÊS 01

| Item | Valor (R$) |
|------|------------|
| **RECEITAS** | |
| Receita à Vista | 560.000,00 |
| Parcelas Recebidas | 6.000,00 |
| **Total Receita** | **566.000,00** |
| | |
| **CUSTOS** | |
| Infraestrutura | 200.000,00 |
| Marketing | 50.000,00 |
| Corretagem | 48.000,00 |
| Administração | 7.500,00 |
| Contingência | 9.165,00 |
| **Total Custos** | **314.665,00** |
| | |
| **RESULTADO BRUTO** | **251.335,00** |
| (-) Impostos | 36.790,00 |
| **Resultado após Impostos** | **214.545,00** |
| (-) Pagamento Terrenista | 75.090,75 |
| **Resultado Empreendedor** | **139.454,25** |
| (-) Custo Financeiro | 15.000,00 |
| **FLUXO LÍQUIDO** | **124.454,25** |
| | |
| **FLUXO ACUMULADO** | **-875.545,75** |

---

## 🎯 INDICADORES DO MÊS

- **Margem Bruta:** 44,4% (251.335 / 566.000)
- **Margem após Impostos:** 37,9% (214.545 / 566.000)
- **Margem Líquida Empreendedor:** 24,6% (139.454 / 566.000)
- **Participação Terrenista:** 13,3% (75.091 / 566.000)
- **Lotes vendidos acumulado:** 10 un. (3,3% do total)
- **Receita acumulada:** R$ 566.000 (2,4% do VGV)

---

## 📝 OBSERVAÇÕES IMPORTANTES

1. **Ordem de cálculo:** Impostos são deduzidos ANTES do pagamento ao terrenista
2. **Base de cálculo dos impostos:** Sobre a receita total (não sobre o lucro)
3. **Base de cálculo do terrenista:** Sobre o resultado após impostos
4. **Custo financeiro:** No mês 1 incide sobre o capital investido. Nos próximos
   meses incidirá sobre o saldo acumulado negativo (se houver)
5. **Parcelas:** Cada lote vendido gera 120 parcelas mensais de R$ 600,00

---

**Data:** 16/07/2026
**Documento:** Memória de Cálculo - Mês 01
**Projeto:** Loteamento Papagaios
**Sistema:** FluxoCaixa_Interativo.html
