# 📊 Otimização do Processo de Cadastro de Clientes (BPM & Data Analytics)

> **Resumo do Impacto:** Reestruturação do fluxo de cadastro com redução estimada de **55% no Lead Time**, eliminação do retrabalho (**30% ➔ <3%**) e validação de dados na origem.

---

## 📌 Visão Geral do Problema

A operação apresentava alta percepção de lentidão no cadastro de novos clientes. A hipótese inicial indicava sobrecarga da equipe ou lentidão no ERP TOTVS Datasul. Através da análise dos dados de chamados e e-mails, foi realizado um diagnóstico para identificar a real causa raiz do gargalo.

---

## 🛠️ Ferramentas & Tecnologias

* **Python (Pandas/NumPy):** Saneamento da base, tratamento de outliers e análise estatística.
* **BPMN 2.0:** Mapeamento do fluxo de processos (As-Is vs. To-Be).
* **Power BI:** Modelagem e Ficha Técnica do KPI de governança (`% TRI`).
* **Figma & PowerPoint:** Hierarquia visual, UI/UX e Storytelling Executivo.

---

## 🔍 Diagnóstico Baseado em Dados

* **Touch Time vs. Wait Time:** A digitação no ERP leva apenas **14 minutos** (Touch Time). O tempo em fila representa **93% do Lead Time total** (3,32 horas).
* **Causa Raiz:** **30% dos chamados** chegavam com informações incompletas devido ao uso de fichas em Word sem validação obrigatória na entrada.

| Métrica | Cenário Atual (As-Is) | Cenário Proposto (To-Be) | Impacto |
| :--- | :---: | :---: | :---: |
| **Lead Time Médio** | 3,32 horas | ~1,5 horas | **-55%** |
| **Taxa de Retrabalho** | 30% | < 3% | **-90%** |
| **Gargalo Principal** | Fila e Troca de E-mails | Eliminado na Origem | **Ganho de Eficiência** |

---

## 💡 Solução & Plano de Ação

Substituição da ficha em Word por um **Formulário Digital com Validação em Tempo Real** (campos obrigatórios para CNPJ, CEP e Inscrição Estadual).

```text
Fase 1 (Semanas 1 e 2) ──► Desenho do Formulário + Travas de Validação (CNPJ/CEP)
Fase 2 (Semanas 3 e 4) ──► Go-Live, Treinamento Comercial e Desativação do Word
Fase 3 (Mês 2)         ──► Dashboard Power BI + Automação RPA para pré-cadastro
