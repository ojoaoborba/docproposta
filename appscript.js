function onOpen() {
  DocumentApp.getUi().createMenu('Funções')
    .addItem('Ativar preenchimento automático', 'activateAutoFill')
    .addToUi();
}

function activateAutoFill() {
  var doc = DocumentApp.getActiveDocument();
  var cursor = doc.getCursor();

  if (!cursor) {
    DocumentApp.getUi().alert('Por favor, posicione o cursor no final de um item.');
    return;
  }

  var element = cursor.getElement();
  if (!element) {
    DocumentApp.getUi().alert('Não foi possível localizar o texto. Certifique-se de que o cursor está em um item válido.');
    return;
  }

  // Tenta identificar o pai correto (parágrafo ou item de lista)
  var parent = element;
  while (parent) {
    if (parent.getType() === DocumentApp.ElementType.LIST_ITEM || parent.getType() === DocumentApp.ElementType.PARAGRAPH) {
      break;
    }
    parent = parent.getParent();
  }

  if (!parent) {
    DocumentApp.getUi().alert('Não foi possível encontrar um parágrafo ou item de lista válido.');
    return;
  }

  var text = parent.getText();

  // Incrementa o número no texto, se existir
  var regex = /\{Funcionalidade(\d+)\};$/;
  var match = text.match(regex);

  if (match) {
    var nextNumber = (parseInt(match[1]) + 1).toString().padStart(2, '0'); // Incrementa o número e mantém dois dígitos
    var newText = `{Funcionalidade${nextNumber}};`; // Novo texto para o próximo item

    // Verifica se é um item de lista
    if (parent.getType() === DocumentApp.ElementType.LIST_ITEM) {
      var listItem = parent.asListItem();
      var newItem = listItem.getParent().insertListItem(listItem.getParent().getChildIndex(listItem) + 1, newText);
      newItem.setGlyphType(listItem.getGlyphType()); // Copia o estilo do marcador
    } else if (parent.getType() === DocumentApp.ElementType.PARAGRAPH) {
      var body = doc.getBody();
      var index = body.getChildIndex(parent);
      body.insertParagraph(index + 1, newText);
    }
  } else {
    DocumentApp.getUi().alert('Formato incorreto. Certifique-se de que o texto segue o padrão: "{FuncionalidadeXX};".');
  }
}
