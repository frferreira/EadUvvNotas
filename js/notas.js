function localizarNomeEInserirNota(nome, nota) {  
  // Normalize o nome para evitar problemas com case sensitivity.  
  const nomeNormalizado = nome.trim().toLowerCase();  

  // Obtenha todos os elementos que podem conter o nome.  
  const elementos = document.body.querySelectorAll('*');  
  let nomeEncontrado = false;  

  elementos.forEach((elemento) => {  
    // Verifique se o elemento possui texto visível e contém o nome que buscamos.  
    if (elemento.innerText && elemento.innerText.trim()) {  
      const texto = elemento.innerText.toLowerCase();  

      if (texto.includes(nomeNormalizado)) {  
        // Nome encontrado  
        console.log(`Nome "${nome}" encontrado no seguinte elemento:`, elemento);  
        nomeEncontrado = true;  

        // Procure o campo nota com a classe "quickgrade" dentro da mesma linha (ou seção relevante).  
        // Aqui assumindo que o campo de nota está associado ao nome na mesma linha (dentro de uma `<tr>` por exemplo).  
        const campoNota = elemento.closest('tr')?.querySelector('input.quickgrade');  

        if (campoNota) {  
          // Insira a nota no campo.  
          campoNota.value = nota;  
          console.log(`Nota "${nota}" inserida para "${nome}".`);  

          // Destaca os elementos para visualização (opcional).  
          elemento.style.backgroundColor = 'yellow';  
          campoNota.style.backgroundColor = 'lightgreen';  
        } else {  
          console.warn(`Campo de nota com class="quickgrade" não foi encontrado para o nome "${nome}".`);  
        }  
      }  
    }  
  });
}
