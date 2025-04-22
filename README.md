# 📦 Sistema de Conferência e Embalagem Automatizada em Excel + VBA

Este projeto tem como objetivo automatizar o processo logístico de **análise de pedidos, separação e embalagem**, utilizando Excel e VBA. Ele reduz erros manuais, melhora a eficiência operacional e garante rastreabilidade total do processo.

## 🔄 Etapas do Processo

### 1. 📥 Importação da Planilha de Necessidade
- Substitua o conteúdo da aba `Necessidade` com a planilha mais atualizada de pedidos de compra.
- Essa planilha contém os itens solicitados pelos clientes ou áreas internas.

### 2. 📊 Análise de Atendimento
- Execute a macro de análise na aba `Análise`.
- A macro irá cruzar os dados da aba `Necessidade` com o estoque atual.
- O resultado mostrará **quantos dos itens solicitados podem ser atendidos** com o que está disponível.
- Essa etapa é essencial para evitar promessas de entrega de itens que não temos.

### 3. 🧾 Separação de Itens
- Com base na análise, os itens disponíveis são listados na aba `Pedidos de Venda`.
- O operador utiliza essa aba para fazer a separação física dos produtos.

### 4. 📷 Verificação via Scanner
- Cada item separado é conferido com um **leitor de código de barras (scanner)**.
- O sistema preenche automaticamente a coluna `Scanner`:
  - `OK`: conferido com sucesso.
  - `Divergente`: conferido, mas com quantidade diferente.
  - `Não Scaneado`: não encontrado no estoque → **não entra no romaneio**.

### 5. 📦 Geração Automática de Romaneio
- Após a separação, execute a macro `GerarRomaneio`.
- A macro verifica os itens com status `OK` ou `Divergente` e os organiza em **caixas de até 16 jogos** cada.
- Se um SKU tiver mais de 16 jogos, o excesso é automaticamente dividido entre caixas.
- A aba `Romaneio` é gerada com as colunas:
  - Número da Caixa
  - Código SKU
  - Descrição
  - Quantidade de Jogos na Caixa
- A data e hora da geração são exibidas no topo da aba.

## ✅ Benefícios

- 🚫 Eliminação de erros manuais
- ⚡ Agilidade na separação e embalagem
- 📋 Rastreabilidade total do processo
- 🖨️ Romaneio pronto para impressão

## 🧠 Tecnologias Utilizadas

- Microsoft Excel
- VBA (Visual Basic for Applications)
- Scanner de código de barras


---

### ✍️ Autor
Desenvolvido por Gabriel Reis como solução prática para automatizar e organizar o processo de conferência e embalagem industrial.

---

