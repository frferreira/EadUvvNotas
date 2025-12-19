//Script para inserir o contéudo ministrado no fechamento da pauta.

// Caso queira usar a planilha de cronograma das turmas EAD, 
// 1) Insira o código a seguir na coluna E das linhas onde estão os topicos das aulas 
// = CONCAT("preencherConteudoMinistrado('"; SUBSTITUIR(ESQUERDA(SUBSTITUIR(B19;"Tópico 0";"");2);":";"");"', '"; DIREITA(B19;NÚM.CARACT(B19)-LOCALIZAR(":";B19));"');")
// A coluna E já montará a chamada da função javascript por conteúdo ministrado 

function preencherConteudoMinistrado(topico, conteudo) {
  topico = "Tópicos: " + topico;
  if (!topico || typeof topico !== 'string') throw new Error('Parâmetro topico obrigatório (string).');
  if (conteudo === undefined || conteudo === null) conteudo = '';

  // Normaliza texto: trim, lowercase, colapsa múltiplos espaços/newlines/tabs em um espaço
  function normalizeText(s) {
    return (s || '').toString()
      .replace(/\s+/g, ' ')   // converte múltiplos espaços/newlines/tabs em um único espaço
      .trim()
      .toLowerCase();
  }

  const target = normalizeText(topico);

  function triggerEvents(el) {
    ['focus', 'input', 'change', 'blur'].forEach(name => {
      try { el.dispatchEvent(new Event(name, { bubbles: true })); } catch (e) {}
    });
    try {
      const last = el.value;
      el.value = conteudo;
      const tracker = el._valueTracker;
      if (tracker && typeof tracker.setValue === 'function') tracker.setValue(last);
      el.dispatchEvent(new Event('input', { bubbles: true }));
    } catch (e) {}
  }

  // Elementos candidatos onde o título/tópico provavelmente aparece
  const selectors = ['h1','h2','h3','h4','h5','h6','label','td','th','span','div','p','a','li'];
  const nodeList = [];
  selectors.forEach(sel => nodeList.push(...Array.from(document.querySelectorAll(sel))));

  // Filtrar e ordenar por proximidade do topo (opcional)
  const candidates = nodeList
    .filter(n => n && n.textContent && normalizeText(n.textContent).includes(target))
    // reduzir falsos positivos: verificar correspondência ao token completo ou parcial com palavras
    .sort((a,b) => (a.compareDocumentPosition(b) & Node.DOCUMENT_POSITION_PRECEDING) ? 1 : -1);

  if (!candidates.length) {
    console.warn('Tópico não encontrado (após normalização):', topico);
    return false;
  }

  function localizarCampoNoContainer(container) {
    if (!container || container.nodeType !== 1) return null;
    const labelText = 'conteúdo ministrado';
    // 1) labels dentro do container
    const labels = Array.from(container.querySelectorAll('label'));
    for (const lab of labels) {
      if (normalizeText(lab.textContent).includes(labelText)) {
        const forAttr = lab.getAttribute('for');
        if (forAttr) {
          const el = container.querySelector(`#${CSS.escape(forAttr)}`) || document.getElementById(forAttr);
          if (el && (el.tagName === 'TEXTAREA' || el.tagName === 'INPUT' || el.isContentEditable)) return el;
        }
        const inner = lab.querySelector('textarea, input, [contenteditable="true"]');
        if (inner) return inner;
      }
    }
    // 2) procurar inputs/textarea com placeholder/aria-label/name/id correspondentes
    const candidatesFields = Array.from(container.querySelectorAll('textarea, input, [contenteditable="true"]'));
    for (const c of candidatesFields) {
      const combined = normalizeText(
        (c.getAttribute('placeholder') || '') + ' ' +
        (c.getAttribute('aria-label') || '') + ' ' +
        (c.getAttribute('name') || '') + ' ' +
        (c.getAttribute('id') || '') + ' ' +
        (c.getAttribute('title') || '')
      );
      if (combined.includes(labelText)) return c;
    }
    // 3) procurar texto "conteúdo ministrado" no container (text nodes), depois achar prox textarea/input
    const walker = document.createTreeWalker(container, NodeFilter.SHOW_TEXT, null, false);
    while (walker.nextNode()) {
      if (normalizeText(walker.currentNode.nodeValue).includes(labelText)) {
        const parent = walker.currentNode.parentElement;
        if (!parent) continue;
        const siblingInput = parent.querySelector('textarea, input, [contenteditable="true"]');
        if (siblingInput) return siblingInput;
        let next = parent.nextElementSibling;
        while (next) {
          const el = next.querySelector && next.querySelector('textarea, input, [contenteditable="true"]');
          if (el) return el;
          next = next.nextElementSibling;
        }
      }
    }
    // 4) fallback: se apenas 1 textarea no container, retornar
    const textareas = container.querySelectorAll ? container.querySelectorAll('textarea') : [];
    if (textareas.length === 1) return textareas[0];
    return null;
  }

  // Para cada candidato, suba até 3 níveis e tente preencher
  for (const node of candidates) {
    let container = node.nodeType === 1 ? node : node.parentElement;
    for (let level = 0; level <= 3; level++) {
      const campo = localizarCampoNoContainer(container);
      if (campo) {
        try {
          if (campo.isContentEditable) {
            campo.focus();
            campo.innerText = conteudo;
            triggerEvents(campo);
          } else if (campo.tagName === 'INPUT') {
            const type = (campo.getAttribute('type') || '').toLowerCase();
            if (['file','checkbox','radio','hidden','button','submit','reset'].includes(type)) {
              console.warn('Campo do tipo não suportado:', type);
              return false;
            }
            campo.focus();
            campo.value = conteudo;
            triggerEvents(campo);
          } else if (campo.tagName === 'TEXTAREA') {
            campo.focus();
            campo.value = conteudo;
            triggerEvents(campo);
          } else {
            campo.focus();
            try { campo.value = conteudo; } catch (e) { campo.innerText = conteudo; }
            triggerEvents(campo);
          }
          console.info('Campo "Conteúdo ministrado" preenchido para tópico:', topico);
          return true;
        } catch (err) {
          console.error('Erro ao preencher campo:', err);
          return false;
        }
      }
      container = container.parentElement;
      if (!container) break;
    }
  }

  console.warn('Não foi possível localizar o campo "Conteúdo ministrado" a partir do tópico fornecido.');
  return false;
}

//preencherConteudoMinistrado('1', ' Conceitos e definições da orientação a objetos');
//preencherConteudoMinistrado('2', ' : Introdução aos conceitos da linguagem C#');
//preencherConteudoMinistrado('3', ' Conceitos Avançados de Orientação a Objetos – Relacionamentos Horizontais');
//preencherConteudoMinistrado('4', '  Conceitos Avançados de Orientação a Objetos – Relacionamento de Herança');
//preencherConteudoMinistrado('5', ' Tratamento de exceções');
//preencherConteudoMinistrado('6', ' Acesso a Bancos de Dados');
//preencherConteudoMinistrado('7', '  Padrão de projeto');
//preencherConteudoMinistrado('8', ' Expressões Lambida e LINQ');
