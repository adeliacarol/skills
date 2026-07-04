# 🚀 Skill: Automação APF AESP (Especialista PONTUA)

Esta skill para o **Gemini CLI** automatiza a contagem de Pontos de Função (APF) a partir de scripts SQL, gerando arquivos prontos para importação no sistema **PONTUA**.

## 🛠️ O que ela faz
- **Análise SQL Inteligente:** Identifica automaticamente INSERTs, UPDATEs e DELETEs.
- **Parser Avançado:** Ignora SELECTs internos para evitar contagens duplicadas.
- **Preenchimento Cirúrgico:** Mantém o template original intacto, sobrescrevendo apenas as linhas necessárias.
- **Compatibilidade PONTUA:** Formata os atributos sem espaços (ex: `id,nome,data`) para garantir que o importador do sistema funcione corretamente.

## 📦 Instalação

1. **Pré-requisito:** Certifique-se de ter a biblioteca `openpyxl` instalada no seu Python:
   ```bash
   pip install openpyxl
   ```

2. **Adicionar Skill:** No seu terminal com o Gemini CLI ativo, execute:
   ```bash
   npx skills add adeliacarol/skills/criar-apf-aesp --skill criar-apf-aesp -y
   ```

## 🚀 Como usar

1.  No chat com o Gemini, ative a skill:
    > `activate_skill criar-apf-aesp`
2.  Envie o seu script SQL (ou referencie o arquivo):
    > `@meu_script_manutencao.sql`
3.  A IA analisará as operações e pedirá sua confirmação. Após o "sim", ela gerará o arquivo `importe.xlsx`.

## 📥 Importação no PONTUA

O arquivo gerado (`importe.xlsx`) foi formatado especificamente para ser utilizado da seguinte forma:

1.  Acesse o sistema **PONTUA**.
2.  Abra a contagem desejada (deve ser do tipo **Apuração Especial**).
3.  Vá até a aba **Funções PF**.
4.  Utilize a opção de **Importar** e selecione o arquivo gerado pela skill.

---
*Dúvidas ou sugestões? Entre em contato com a equipe de APF.*
