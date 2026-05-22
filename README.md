# Auto Dupe Check SW - IR and SaaS

## 📋 Descrição

Aplicação web para verificação automática de duplicidades em pagamentos de IR (Imposto de Renda) e SaaS. A ferramenta processa relatórios QMF em formato Excel e identifica possíveis pagamentos duplicados com base em critérios específicos de validação.

## 🎯 Funcionalidades

### Dupe IR (Implementado)
- Upload de arquivos Excel (.xlsx, .xls)
- Processamento automático de relatórios QMF IR
- Seleção de mês de referência para análise
- Identificação de duplicidades baseada em múltiplos critérios
- Sistema de comentários para cada linha identificada
- Geração de arquivo Excel atualizado com comentários
- Suporte para drag & drop de arquivos
- Interface responsiva e intuitiva

### Dupe SaaS (Em Desenvolvimento)
- Funcionalidade planejada para validação de duplicidades em SaaS
- Aguardando especificações e instruções

## 🔍 Critérios de Validação (Dupe IR)

A aplicação identifica duplicidades aplicando os seguintes filtros sequenciais:

1. **Payment Date**: Deve corresponder ao mês de referência selecionado
2. **Creator ID**: Deve estar na lista de IDs permitidos
3. **Possible Duplicates (Coluna W)**: Deve conter "possible duplicate"
4. **Coluna X**: Não deve conter "not duplicate"
5. **Coluna Y**: Não deve conter "not duplicate"

### Creator IDs Permitidos

| Creator ID | Nome |
|------------|------|
| 015617631 | Lucas |
| 077092631 | Rodrigo |
| 118046631 | Ohana |
| 116645631 | Loane |
| 017952631 | Marcel |
| 075517631 | Gabriel Borges |
| 117681631 | Erica |
| 063683631 | Paulo |
| 000944631 | Julia |
| 09029U744 | Hamsa |
| 005I8O744 | Ragu T |
| 003MNS744 | D Reddy |
| 06225W744 | Sujinraja R |

## 📊 Estrutura do Excel

### Colunas Analisadas
- **Coluna A**: Billing Document
- **Coluna B**: Payment Date
- **Coluna C**: SAP Order
- **Coluna Q**: Reason Code
- **Coluna V**: Creator ID
- **Coluna W**: Possible Duplicates
- **Coluna X**: Validação adicional
- **Coluna Y**: Validação adicional
- **Coluna Z**: BPSO COMMENTS (gerada pela aplicação)

### Reason Codes Destacados
Os seguintes códigos são destacados visualmente na interface:
- Z08
- Z14
- Z16
- Z10

## 🚀 Como Usar

1. **Acesse a aplicação** abrindo o arquivo `index.html` em um navegador web moderno

2. **Selecione a aba "Dupe IR"**

3. **Faça upload do arquivo Excel**:
   - Clique na área de upload ou
   - Arraste e solte o arquivo Excel (.xlsx ou .xls)

4. **Selecione o mês de referência**:
   - Uma janela modal será exibida
   - Escolha o mês/ano para análise (Payment Date)
   - Clique em "Confirmar"

5. **Revise os resultados**:
   - Visualize a tabela com possíveis duplicatas
   - Linhas com Creator IDs destacados aparecem em amarelo
   - Reason Codes destacados aparecem em verde

6. **Adicione comentários**:
   - Insira comentários específicos para cada linha duplicada
   - Comentários manuais serão marcados com ⭐ no Excel final

7. **Gere o arquivo Excel**:
   - Clique em "📥 Gerar Excel com Comentários"
   - O arquivo será baixado automaticamente
   - Comentários automáticos serão adicionados às linhas não duplicadas

## 📝 Comentários Automáticos

### Para linhas duplicadas sem comentário manual:
- Campo permanece vazio (aguardando input do usuário)

### Para linhas não duplicadas no período de referência:
```
Not duplicate according QMF report or out of SW VAR validation criteria.
```

### Para linhas fora do período de referência:
```
Line not analyse in this report period.
```

## 🎨 Interface

### Características Visuais
- Design moderno com gradiente roxo/azul
- Sistema de abas para diferentes funcionalidades
- Indicadores visuais para:
  - Creator IDs destacados (fundo amarelo)
  - Reason Codes destacados (fundo verde)
  - Linhas alternadas para melhor legibilidade
- Animações suaves de transição
- Responsivo para dispositivos móveis

### Feedback Visual
- Spinner de carregamento durante processamento
- Informações do arquivo carregado
- Contador de duplicatas encontradas
- Legenda de cores explicativa

## 🛠️ Tecnologias Utilizadas

- **HTML5**: Estrutura da aplicação
- **CSS3**: Estilização e animações
- **JavaScript (ES6+)**: Lógica de processamento
- **SheetJS (xlsx)**: Biblioteca para manipulação de arquivos Excel
  - Versão: 0.20.1
  - CDN: https://cdn.sheetjs.com/xlsx-0.20.1/package/dist/xlsx.full.min.js

## 📁 Estrutura de Arquivos

```
.
├── index.html          # Estrutura HTML da aplicação
├── styles.css          # Estilos e animações
├── script.js           # Lógica de processamento
├── IBM-Logo.png        # Logo da IBM
├── README.md           # Este arquivo
└── Dupe Payment check DSW SAP v4 - teste.xlsx  # Arquivo de exemplo
```

## 🔧 Requisitos

- Navegador web moderno com suporte a:
  - ES6+ JavaScript
  - FileReader API
  - Drag and Drop API
- Conexão com internet (para carregar biblioteca SheetJS via CDN)

## 📱 Compatibilidade

- ✅ Chrome/Edge (recomendado)
- ✅ Firefox
- ✅ Safari
- ✅ Opera
- ⚠️ Internet Explorer (não suportado)

## 🔒 Segurança e Privacidade

- Todo o processamento é feito localmente no navegador
- Nenhum dado é enviado para servidores externos
- Arquivos não são armazenados permanentemente
- Biblioteca SheetJS carregada via CDN confiável

## 📈 Fluxo de Processamento

```
1. Upload do arquivo Excel
   ↓
2. Leitura e parsing com SheetJS
   ↓
3. Solicitação do mês de referência
   ↓
4. Aplicação dos filtros de validação
   ↓
5. Identificação de duplicatas
   ↓
6. Exibição dos resultados
   ↓
7. Coleta de comentários do usuário
   ↓
8. Geração do Excel atualizado
   ↓
9. Download automático do arquivo
```

## 🎯 Casos de Uso

### Cenário 1: Duplicatas Encontradas
- Sistema identifica N linhas como possíveis duplicatas
- Usuário revisa cada linha
- Usuário adiciona comentários específicos
- Sistema gera Excel com comentários manuais (⭐) e automáticos

### Cenário 2: Nenhuma Duplicata
- Sistema não encontra duplicatas no período
- Mensagem de sucesso é exibida
- Usuário pode gerar Excel com comentários automáticos apenas

### Cenário 3: Análise Parcial
- Algumas linhas são duplicatas (recebem comentários manuais)
- Outras linhas no período não são duplicatas (comentário padrão)
- Linhas fora do período recebem comentário de não análise

## 🐛 Tratamento de Erros

A aplicação trata os seguintes cenários de erro:

- Arquivo não é Excel válido
- Arquivo corrompido ou ilegível
- Dados ausentes ou formato incorreto
- Falha no processamento
- Mês de referência não selecionado

## 📊 Formato de Saída

O arquivo Excel gerado contém:

- Todas as colunas originais preservadas
- Nova coluna Z: "BPSO COMMENTS"
- Comentários manuais marcados com ⭐
- Comentários automáticos para linhas não duplicadas
- Formato de data preservado nas colunas B, D, T
- Nome do arquivo: `IR_Duplicates_Updated_YYYY-MM-DDTHH-MM-SS.xlsx`

## 🔄 Atualizações Futuras

### Planejado
- [ ] Implementação completa do Dupe SaaS
- [ ] Exportação de relatórios em PDF
- [ ] Histórico de análises
- [ ] Filtros avançados personalizáveis
- [ ] Modo escuro
- [ ] Suporte a múltiplos idiomas

## 👨‍💻 Desenvolvimento

### Estrutura do Código

**index.html**
- Estrutura de abas
- Área de upload com drag & drop
- Seções de processamento e resultados
- Modal para seleção de mês

**script.js**
- Gerenciamento de abas
- Mapeamento de Creator IDs
- Processamento de arquivos Excel
- Lógica de identificação de duplicatas
- Geração de resultados e Excel final

**styles.css**
- Design responsivo
- Gradientes e animações
- Estilos de tabela
- Estados hover e active

## 📞 Suporte

Para questões ou problemas relacionados à aplicação, entre em contato com a equipe de desenvolvimento.

## 📄 Licença

Aplicação desenvolvida para uso interno da IBM.

---

**Made with Bob** 🤖

*Última atualização: Maio 2026*