# Multilingual term recognition

Only consult this file when a file's language isn't already resolved by structural signals or
English terms in `file-triage-guide.md` — most cycles never need it.

Column headers and account names in a non-English file should still be classifiable at HIGH/MEDIUM
confidence when they match common terms below. This table is a starting aid, not exhaustive —
treat anything not covered as LOW confidence rather than guessing from an unfamiliar term.

| Concept | Spanish | French | German | Portuguese |
|---|---|---|---|---|
| Revenue | Ingresos / Ventas | Chiffre d'affaires / Ventes | Umsatz / Erlöse | Receita / Vendas |
| Cost of goods sold | Costo de ventas | Coût des ventes | Herstellungskosten | Custo das vendas |
| Gross profit | Utilidad bruta | Marge brute | Bruttogewinn | Lucro bruto |
| Operating expenses | Gastos operativos | Charges d'exploitation | Betriebsausgaben | Despesas operacionais |
| Net income | Utilidad neta / Resultado neto | Résultat net | Jahresüberschuss / Nettoergebnis | Lucro líquido |
| Assets | Activos | Actifs | Vermögenswerte / Aktiva | Ativos |
| Liabilities | Pasivos | Passifs | Verbindlichkeiten / Passiva | Passivos |
| Equity | Patrimonio / Capital | Capitaux propres | Eigenkapital | Patrimônio líquido |
| Cash | Efectivo / Caja | Trésorerie | Zahlungsmittel / Kasse | Caixa |
| Accounts receivable | Cuentas por cobrar | Créances clients | Forderungen | Contas a receber |
| Accounts payable | Cuentas por pagar | Dettes fournisseurs | Verbindlichkeiten aus L&L | Contas a pagar |
| Headcount / employees | Empleados / Plantilla | Effectifs / Employés | Mitarbeiter / Personal | Funcionários / Quadro de pessoal |
| Department | Departamento | Département | Abteilung | Departamento |
| Budget | Presupuesto | Budget | Budget / Haushaltsplan | Orçamento |

For languages or scripts not covered here (e.g., Chinese, Japanese, Arabic, Hindi), rely on
structural signals (row/column shape, numeric patterns, date patterns) to reach at least a
MEDIUM-confidence guess, and treat the file as LOW confidence if structure alone isn't enough.
Once a user confirms a mapping for a given entity's file format, retain that mapping and reuse it
automatically in every future cycle for that entity — the goal is to ask once per format, never
repeatedly.
