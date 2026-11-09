# 📋 Gerador de Resumos em PDF (2 Colunas A4)

Uma ferramenta leve, rápida e **100% nativa para Windows** feita para estudantes e profissionais organizarem imagens de resumos, anotações de aula e fichas de estudo em PDFs diagramados no padrão A4.

---

## 📐 Como o PDF é montado

Cada página A4 é dividida automaticamente em **2 colunas e 4 quadrantes**, respeitando a ordem natural de leitura de fichamentos e notas:

```text
+------------------------------+------------------------------+
|                              |                              |
|   [1] Superior Esquerdo      |     [3] Superior Direito     |
|         (Início)             |                              |
|                              |                              |
+------------------------------+------------------------------+
|                              |                              |
|   [2] Inferior Esquerdo      |     [4] Inferior Direito     |
|                              |            (Fim)             |
|                              |                              |
+------------------------------+------------------------------+
```

* Se você colocar **4 imagens** ➔ Gera **1 página**.
* Se colocar **8, 12, 16 imagens** ➔ O script cria automaticamente as **páginas 1, 2, 3...** no mesmo documento.

---

## 🚀 Como Instalar em Qualquer Computador (Windows 10 ou 11)

Não é necessário instalar Python, Word, Node.js ou qualquer programa de terceiros. Funciona com o que todo Windows já tem de fábrica.

1. Baixe este repositório:
   * Clique no botão verde **Code** acima ➔ **Download ZIP**.
2. Extraia a pasta do arquivo ZIP em qualquer lugar do seu computador (por exemplo: na pasta `Documentos` ou `C:\Ferramentas\GeradorPDF`).
3. Abra a pasta extraída e dê **dois cliques** no arquivo:
   * **`Criar_Atalho_na_Area_de_Trabalho.bat`**
4. **Pronto!** O ícone **`Gerar PDF de Resumos`** aparecerá na sua Área de Trabalho com um ícone dedicado.

---

## 📖 Como Usar no Dia a Dia

Existem duas formas simples de usar:

### Modo 1: Arrastar e Soltar (Mais Rápido)
1. Coloque suas 4 imagens (ou mais) dentro de uma pasta.
2. **Arraste a pasta** (ou selecione as imagens e arraste) para cima do ícone **`Gerar PDF de Resumos`** na sua Área de Trabalho.

### Modo 2: Duplo Clique
1. Dê dois cliques no ícone **`Gerar PDF de Resumos`** na Área de Trabalho.
2. Selecione a pasta onde estão as fotos.

---

## 🖥️ A Janela de Prévia Visual

Assim que as fotos são carregadas, abre uma janela mostrando a folha A4 com as fotos já posicionadas:

* **Tudo certo?** Apenas aperte **Enter** (ou clique no botão azul *"Gerar PDF"*). Em 2 segundos o PDF final estará pronto na pasta!
* **Quer inverter a ordem de duas imagens?**
  1. Clique na imagem que quer mudar (ela fica com destaque azul).
  2. Clique na outra imagem com quem quer trocar.
  3. Elas invertem de lugar imediatamente na tela!
* **Quer ler uma imagem em tamanho grande?**
  * Dê **dois cliques** em cima de qualquer miniatura para abri-la em tela cheia no visualizador de fotos do Windows.
* **Múltiplas Páginas:**
  * Se o resumo tiver mais de 4 imagens, use os botões **`◀ Anterior`** e **`Próxima ▶`** no topo da janela para revisar cada página antes de gerar.

---

## 🛠️ Tecnologias Utilizadas
* **PowerShell 5.1 & .NET Windows Forms:** Interface gráfica leve, sem dependências externas.
* **Microsoft Edge Headless:** Renderização vetorial exata para o formato A4 com margens perfeitas e alta nitidez de imagem.

---

## 📄 Licença
Distribuído sob a licença MIT. Sinta-se livre para usar, modificar e compartilhar!
