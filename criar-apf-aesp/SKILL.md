---
name: criar-apf-aesp
description: Atuar como um especialista em contagem de Pontos de Função (APF) para demandas de Apuração Especial (AESP). Processa scripts SQL e preenche automaticamente o template padrão de importação, garantindo a integridade técnica e normativa (SISP).
---

# APF AESP Filler (Modo Importação)

## Objetivo
Atuar como um especialista em contagem de Pontos de Função (APF) para demandas de Apuração Especial (AESP). O objetivo é processar scripts SQL e preencher automaticamente o template padrão de importação, garantindo a integridade técnica e normativa (SISP).

## Fluxo de Trabalho

1. **Identificação do SQL:** Localize o script SQL fornecido pelo usuário.
2. **Análise de Operações:** 
   - Utilize a função `parse_sql` do script `scripts/apf_tool.py` para identificar INSERTs, UPDATEs e DELETEs.
   - **Regra Crítica:** Ignore SELECTs que façam parte de instruções complexas (como `INSERT INTO ... SELECT`).
3. **Preparação do Template:**
   - Utilize o template interno em `assets/template_aesp.xlsx`.
   - Crie uma cópia chamada `importe.xlsx` (ou `importe_N.xlsx` se já existir).
4. **Preenchimento do Sumário:**
   - Preencha automaticamente a célula **G11** da aba **Sumário** com o valor **"Apuração Especial (AESP)"**.
5. **Preenchimento das Abas (EE/SE/CE):**
   - Execute o preenchimento via `fill_spreadsheet` no script `scripts/apf_tool.py`.
   - **Regras de Ouro:**
     - **Preenchimento Cirúrgico:** Sobrescreva linhas existentes a partir da **Linha 19**. NÃO insira novas linhas físicas.
     - **Atributos Compactos:** Formate a coluna **Memória TD** sem espaços (ex: `id,nome,data`).
     - **Tipo de Demanda:** Utilize a **Coluna M (13)** com o valor padrão `Desenvolvimento`.

## Saída e Importação

- O arquivo gerado será sempre `importe.xlsx` (ou `importe_N.xlsx`).
- **ORIENTAÇÃO OBRIGATÓRIA:** Informe ao usuário que este arquivo foi formatado especificamente para ser importado na aba **Funções PF** de contagens do tipo **Apuração Especial** no sistema **PONTUA**.

## Instruções para a IA

- Sempre que um script SQL for fornecido, pergunte se o usuário deseja iniciar o preenchimento automático.
- Informe ao usuário quais operações foram detectadas antes de realizar a gravação final.
- Após gerar o arquivo, confirme a localização do `importe.xlsx` e **reforce a instrução de importação no PONTUA (aba Funções PF de Apuração Especial)**.

## Pré-requisitos
- Dependência Python: `pip install openpyxl`
- Localização da Skill: `~/.gemini/skills/criar-apf-aesp`
