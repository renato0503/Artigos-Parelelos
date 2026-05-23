# Guia de Formatação HTML para Artigos Acadêmicos

Este documento contém a formatação padrão utilizada nos artigos do projeto, incluindo CSS e estrutura HTML para simulação de páginas A4 com suporte a impressão PDF e exportação DOCX.

---

## 1. Estrutura HTML Completa

```html
<!DOCTYPE html>
<html lang="pt-BR">
<head>
    <meta charset="UTF-8">
    <title>Título do Artigo</title>
    <style>
        /* CSS aqui (ver seção 2) */
    </style>
</head>
<body>

    <!-- Painel de Controle (não impresso) -->
    <div class="control-panel no-print">
        <button onclick="window.print()" class="btn-print">🖨️ Imprimir / Salvar PDF</button>
        <button onclick="exportarDocx()" class="btn-docx">📄 Baixar DOCX</button>
    </div>

    <!-- PÁGINA 1 -->
    <section class="paper-page">
        <div class="page-content">
            <!-- Conteúdo da página -->
        </div>
        <div class="page-footer">1</div>
    </section>

    <!-- PÁGINA 2 -->
    <section class="paper-page">
        <div class="page-content">
            <!-- Conteúdo da página -->
        </div>
        <div class="page-footer">2</div>
    </section>

    <!-- Script de Exportação -->
    <script>
        function exportarDocx() {
            const conteudo = document.body.innerHTML;
            const blob = new Blob(['\ufeff' + conteudo], { type: 'application/msword' });
            const url = URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = url;
            a.download = 'nome_artigo.doc';
            document.body.appendChild(a);
            a.click();
            document.body.removeChild(a);
            URL.revokeObjectURL(url);
        }
    </script>

</body>
</html>
```

---

## 2. CSS Completo (Style Block)

```css
* {
    margin: 0;
    padding: 0;
    box-sizing: border-box;
}

body {
    font-family: "Times New Roman", serif;
    font-size: 12pt;
    line-height: 1.5;
    background-color: #f0f2f5;
    color: #1a1a1a;
    display: flex;
    flex-direction: column;
    align-items: center;
    padding: 40px 20px;
}

.paper-page {
    background-color: #ffffff;
    width: 210mm;
    min-height: 297mm;
    height: auto;
    margin-bottom: 30px;
    box-shadow: 0 4px 15px rgba(0, 0, 0, 0.15);
    position: relative;
    padding: 25mm 20mm 20mm 30mm;
    display: flex;
    flex-direction: column;
}

.paper-page::before {
    content: "";
    position: absolute;
    top: 25mm;
    left: 30mm;
    right: 20mm;
    bottom: 20mm;
    border: 1px dotted #e2e8f0;
    pointer-events: none;
    z-index: 1;
}

.page-content {
    flex-grow: 1;
    position: relative;
    z-index: 2;
}

.page-footer {
    font-size: 10pt;
    text-align: right;
    padding-top: 10px;
    border-top: 1px solid #f0f2f5;
    margin-top: 10px;
}

h1 {
    font-size: 14pt;
    font-weight: bold;
    text-align: center;
    margin: 0.5cm 0;
    line-height: 1.3;
}

h2 {
    font-size: 12pt;
    font-weight: bold;
    margin: 0.6cm 0 0.3cm 0;
}

h3 {
    font-size: 12pt;
    font-weight: bold;
    margin: 0.5cm 0 0.2cm 0;
}

p {
    text-indent: 1.25cm;
    margin: 0 0 0.25cm 0;
    text-align: justify;
}

.author-line {
    text-align: center;
    text-indent: 0;
    margin-bottom: 0.5cm;
}

.resumo {
    text-indent: 0 !important;
    margin-bottom: 0.5cm;
    line-height: 1.2;
}

.resumo p {
    text-indent: 0 !important;
    margin: 0 0 0.2cm 0 !important;
    line-height: 1.2 !important;
    font-size: 12pt;
}

.resumo h2 {
    font-size: 12pt;
    font-weight: bold;
    text-align: center;
    margin: 0 0 0.2cm 0;
}

.keywords-line {
    text-indent: 0;
    margin-top: 0.2cm;
}

table {
    width: 100%;
    border-collapse: collapse;
    margin: 0.4cm 0;
    font-size: 10pt;
    line-height: 1.2;
}

th {
    border-top: 1px solid black;
    border-bottom: 1px solid black;
    padding: 5px;
    font-weight: bold;
    text-align: center;
}

td {
    padding: 4px;
    text-align: left;
}

tr.last-row td {
    border-bottom: 1px solid black;
}

.table-title {
    font-size: 11pt;
    font-weight: bold;
    text-align: left;
    margin: 0.4cm 0 0.1cm 0;
}

.table-source {
    font-size: 9pt;
    font-style: italic;
    text-align: left;
    margin: 0.1cm 0 0.4cm 0;
}

.ref-entry {
    text-indent: -1.25cm;
    padding-left: 1.25cm;
    margin: 0 0 0.15cm 0;
    text-align: justify;
    font-size: 12pt;
}

.control-panel {
    position: fixed;
    bottom: 30px;
    right: 30px;
    background-color: rgba(255, 255, 255, 0.95);
    padding: 15px 20px;
    border-radius: 50px;
    box-shadow: 0 10px 25px rgba(0, 0, 0, 0.2);
    display: flex;
    gap: 10px;
    z-index: 1000;
    border: 1px solid rgba(226, 232, 240, 0.8);
}

.control-panel button {
    padding: 10px 18px;
    border: none;
    border-radius: 25px;
    cursor: pointer;
    font-weight: bold;
    font-size: 10pt;
    display: flex;
    align-items: center;
    gap: 6px;
    transition: transform 0.2s;
}

.control-panel button:hover {
    transform: translateY(-2px);
}

.btn-print {
    background-color: #1A365D;
    color: white;
}

.btn-docx {
    background-color: #2b579a;
    color: white;
}

@media print {
    body {
        background-color: transparent;
        color: #000000;
        display: block;
        padding: 0;
    }

    .control-panel {
        display: none !important;
    }

    .paper-page {
        width: 100% !important;
        height: auto !important;
        margin: 0 !important;
        box-shadow: none !important;
        padding: 0 !important;
        page-break-after: always;
        display: block;
        background: white;
    }

    .paper-page::before {
        display: none !important;
    }

    .page-footer {
        position: absolute;
        bottom: 2cm;
        right: 2cm;
    }

    @page {
        size: A4;
        margin: 2.5cm 2cm 2cm 3cm;
    }
}
```

---

## 3. Especificações de Formatação

### Margens da Página (ABNT)
| Margem | Valor |
|--------|-------|
| Superior | 2.5cm ou 3cm |
| Inferior | 2cm |
| Esquerda | 3cm |
| Direita | 2cm |

### Tipografia
| Elemento | Especificação |
|----------|---------------|
| Fonte principal | Times New Roman, serif |
| Tamanho do texto | 12pt |
| Título (h1) | 14pt, bold, centralizado |
| Seção (h2) | 12pt, bold |
| Subseção (h3) | 12pt, bold |
| Recuo de parágrafo | 1.25cm |
| Espaçamento entre linhas | 1.5 |
| Alinhamento | Justificado |

### Tabelas
| Elemento | Estilo |
|----------|--------|
| Linhas horizontais | Apenas topo e base (1px solid black) |
| Cabeçalho | Bold, centralizado |
| Células | Padding 4-5px, alinhamento à esquerda |
| Título da tabela | 11pt, bold, margem superior 0.4cm |
| Fonte da tabela | 10pt |

### Elementos Especiais
| Elemento | Tratamento |
|----------|------------|
| Resumo/Abstract | Sem recuo, espaço simples (line-height: 1.2) |
| Referências | Recuo negativo -1.25cm (hanging indent) |
| Citações diretas | Recuo de 4cm, fonte 11pt |
| Palavras-chave | Texto bold, recuo zero |

---

## 4. Estrutura de Páginas

Cada seção `<section class="paper-page">` representa uma página física A4. O conteúdo deve ser dividido logicalmente entre páginas para facilitar a leitura e impressão.

### Exemplo de Divisão:
```
Página 1: Título + Resumo + Abstract
Página 2: Introdução
Página 3: Fundamentação Teórica
Página 4: Metodologia + Tabela
Página 5: Resultados e Discussão
Página 6: Considerações Finais
Página 7: Referências
```

---

## 5. Funcionalidades de Impressão/Exportação

### Imprimir PDF
```javascript
// Botão que chama window.print() do navegador
<button onclick="window.print()" class="btn-print">🖨️ Imprimir / Salvar PDF</button>
```

### Exportar DOCX
```javascript
function exportarDocx() {
    const conteudo = document.body.innerHTML;
    const blob = new Blob(['\ufeff' + conteudo], { type: 'application/msword' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = 'nome_artigo.doc';
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
}
```

### CSS para Impressão (@media print)
- Oculta painel de controle
- Remove sombras e bordas de página
- Define `@page` com margens ABNT
- Força `page-break-after: always` em cada página

---

## 6. Classes CSS de Referência Rápida

| Classe | Uso |
|--------|-----|
| `.paper-page` | Container de cada página A4 |
| `.page-content` | Área de conteúdo da página |
| `.page-footer` | Numeração da página (rodapé) |
| `.control-panel` | Painel flutuante com botões |
| `.btn-print` | Botão de impressão (azul escuro) |
| `.btn-docx` | Botão DOCX (azul) |
| `.resumo` | Container de resumo/abstract |
| `.table-title` | Título de tabela (negrito) |
| `.table-source` | Fonte da tabela (itálico) |
| `.ref-entry` | Entrada de referência bibliográfica |
| `.no-print` | Elementos ocultos na impressão |

---

## 7. Checklist de Formatação

- [ ] Margens: Superior 2.5cm, Inferior 2cm, Esquerda 3cm, Direita 2cm
- [ ] Fonte: Times New Roman 12pt
- [ ] Espaçamento: 1.5 entre linhas
- [ ] Recuo: 1.25cm na primeira linha de cada parágrafo
- [ ] Alinhamento: Justificado
- [ ] Tabelas: apenas bordas horizontais
- [ ] Resumo: sem recuo, espaço simples
- [ ] Referências: recuo negativo (hanging indent)
- [ ] Botões de impressão visíveis na tela
- [ ] Margens invisíveis na impressão
- [ ] Paginação correta no rodapé

---

*Documento gerado em 20/05/2026 para uso nos artigos do projeto.*