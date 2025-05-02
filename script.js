function script_string() {
  return "function localizarNomeEInserirNota(nome, nota) {    // Normalize o nome para evitar problemas com case sensitivity.  \n"+
  "  const nomeNormalizado = nome.trim().toLowerCase();\n"+
  "\n"+
  "  // Obtenha todos os elementos que podem conter o nome.  \n"+
  "  const elementos = document.body.querySelectorAll('*');\n"+
  "  let nomeEncontrado = false;\n"+
  "\n"+
  "  elementos.forEach((elemento) => {\n"+
  "    // Verifique se o elemento possui texto visível e contém o nome que buscamos.  \n"+
  "    if (elemento.innerText && elemento.innerText.trim()) {\n"+
  "      const texto = elemento.innerText.toLowerCase();\n"+
  "\n"+
  "      if (texto.includes(nomeNormalizado)) {\n"+
  "        // Nome encontrado  \n"+
  '        console.log(`Nome "${nome}" encontrado no seguinte elemento:`, elemento);\n'+
  "        nomeEncontrado = true;\n"+
  "\n"+
  '        // Procure o campo nota com a classe "quickgrade" dentro da mesma linha (ou seção relevante).  \n'+
  "        // Aqui assumindo que o campo de nota está associado ao nome na mesma linha (dentro de uma `<tr>` por exemplo).  \n"+
  "        const campoNota = elemento.closest('tr')?.querySelector('input.quickgrade');\n"+
  "\n"+
  "\n"+
  "        if (campoNota) {\n"+
  "          // Insira a nota no campo.  \n"+
  "          campoNota.value = nota;\n"+
  '          console.log(`Nota "${nota}" inserida para "${nome}".`);\n'+
  "\n"+
  "          // Destaca os elementos para visualização (opcional).  \n"+
  "          elemento.style.backgroundColor = 'yellow';\n"+
  "          campoNota.style.backgroundColor = 'lightgreen';\n"+
  "          return true;\n"+
  "        } else {\n"+
  "          const linhas = document.querySelectorAll('tr');\n"+
  "          for (let linha of linhas) {\n"+
  "            // Busca o nome na linha (ignorando maiúsculas/minúsculas e espaços extras)\n"+
  "            if (linha.innerText && linha.innerText.toLowerCase().trim().includes(nome.toLowerCase().trim())) {\n"+
  "              // Procura o campo input da AV1 (ajuste o seletor se necessário)\n"+
  "              // Supondo que há só 1 input de nota por linha e que é o de AV1:\n"+
  "              const inputAV1 = linha.querySelector('input');\n"+
  "              if (inputAV1) {\n"+
  "                inputAV1.value = nota;\n"+
  "                inputAV1.dispatchEvent(new Event('input', { bubbles: true }));\n"+
  "                inputAV1.dispatchEvent(new Event('change', { bubbles: true }));\n"+
  "                inputAV1.dispatchEvent(new Event('blur', { bubbles: true }));\n"+
  "                linha.style.backgroundColor = 'yellow';\n"+
  "                inputAV1.style.backgroundColor = 'lightgreen';\n"+
  "                return true; // Sucesso\n"+
  "              }\n"+
  "\n"+
  "            }\n"+
  '            console.warn(`Campo de nota com class="quickgrade" não foi encontrado para o nome "${nome}".`);\n'+
  "          }\n"+
  "        }\n"+
  "      }\n"+
  "    }\n"+
  "  });\n"+
  "  // Caso o nome não seja encontrado.  \n"+
  "  if (!nomeEncontrado) {\n"+
  '    console.log(`O nome "${nome}" não foi encontrado na página.`);\n'+
  "  }\n"+
  "}" + 
  "\n" +
  "\n"

}

function localizarNomeEInserirNota(nome, nota) {    // Normalize o nome para evitar problemas com case sensitivity.  
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
          return true;
        } else {
          const linhas = document.querySelectorAll('tr');
          for (let linha of linhas) {
            // Busca o nome na linha (ignorando maiúsculas/minúsculas e espaços extras)
            if (linha.innerText && linha.innerText.toLowerCase().trim().includes(nome.toLowerCase().trim())) {
              // Procura o campo input da AV1 (ajuste o seletor se necessário)
              // Supondo que há só 1 input de nota por linha e que é o de AV1:
              const inputAV1 = linha.querySelector('input');
              if (inputAV1) {
                inputAV1.value = nota;
                inputAV1.dispatchEvent(new Event('input', { bubbles: true }));
                inputAV1.dispatchEvent(new Event('change', { bubbles: true }));
                inputAV1.dispatchEvent(new Event('blur', { bubbles: true }));
                linhas.style.backgroundColor = 'yellow';
                inputAV1.style.backgroundColor = 'lightgreen';
                return true; // Sucesso
              }

            }
            console.warn(`Campo de nota com class="quickgrade" não foi encontrado para o nome "${nome}".`);
          }
        }
      }
    }
  });
  // Caso o nome não seja encontrado.  
  if (!nomeEncontrado) {
    console.log(`O nome "${nome}" não foi encontrado na página.`);
  }
}

async function abrirSiteNotas() {
  const url = document.getElementById("linkSiteNotas").value;
  if (url) {
    document.getElementById('iframeA').src = url;
  } else {
    alert("Por favor, permita pop-ups para este site.");
  }
}
async function processFiles() {
  const output = document.getElementById("output");
  output.value = "Processando arquivos...\n";
  const files = document.getElementById("fileInput").files;
  /*if (files.length === 0 || files.length > 5) {  */
  if (!files.length > 0) {
    alert("Selecione no minimo 1 arquivo XLSX.");
    return;
  }
  let combinedData = [];
 
  const status = document.getElementById("status");

  
  output.value = "Processando arquivos...\n";

  for (let i = 0; i < files.length; i++) {
    const file = files[i];
    const data = await readFile(file);
    if (data.length > 0) {
      combinedData = [...combinedData, ...data];
    }
  }
  try {

    // Insere fórmula na coluna H  
    formulas = combinedData.map((row, index) => {
      /*if (index === 0) {  
        row["H"] = "Fórmula (Coluna H)";  
      } else {  
        if (row["A"] && row["G"]) {  
          row["H"] = `=CONCAT("localizarNomeEInserirNota(";'"';${row["A"]};'", ";SUBSTITUIR(${row["G"]};",";".");")")`;  
        }  
      }  
      return row;  */
      /* está ignorando o primeiro aluno
       if (index === 0) {
        return null; // Ignoramos o cabeçalho  
      }*/
      const colA = row["Aluno"] !== undefined ? row["Aluno"] : ""; // Conteúdo da coluna A  
      const colG = row["Resultado"] !== undefined ? row["Resultado"].toString() : ""; // Conteúdo da coluna G  
      return `localizarNomeEInserirNota("${colA}", ${colG.replace(",", ".")})`;
    }).filter(Boolean); // Remove valores nulos  ;  
    // Exibe as fórmulas no elemento <textarea>  
    //output.value = script_string().join(formulas.join("\n"));  
    output.value = script_string() + formulas.join("\n");
    output.focus();
    output.select();
    // copia o conteudo do output para a area de transferencia
    document.execCommand("copy");
    //output.value = script_string() + formulas.join("\n");


    //output.value = formulas.join("\n");
  } catch (error) {
    console.error("Erro ao processar arquivo:", error);
    output.value = "Erro ao processar arquivo. Verifique se ele está no formato correto.";
  }
  status.innerHTML = "Arquivos processados com sucesso! e Script copiados para memoria !";
}


function readFile(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (event) => {
      const workbook = XLSX.read(event.target.result, { type: "binary" });
      const firstSheet = workbook.SheetNames[0];
      const data = XLSX.utils.sheet_to_json(workbook.Sheets[firstSheet], { defval: "" });
      resolve(data);
    };
    reader.onerror = (error) => reject(error);
    reader.readAsBinaryString(file);
  });
}

function s2ab(s) {
  const buf = new ArrayBuffer(s.length);
  const view = new Uint8Array(buf);
  for (let i = 0; i < s.length; i++) {
    view[i] = s.charCodeAt(i) & 0xff;
  }
  return buf;
}

async function processXLSX(fileInput) {
  //const fileInput = document.getElementById("xlsxInput");  
  const output = document.getElementById("output");

  /*if (!fileInput.files.length) {  
    alert("Por favor, selecione um arquivo XLSX.");  
    return;  
  } */

  const file = fileInput.files[0];

  try {
    // Lê o arquivo XLSX  
    const data = await readXLSX(file);

    // Gera as fórmulas para cada linha do arquivo  
    const formulas = data.map((row, index) => {
      // Pula a primeira linha, caso ela seja o cabeçalho  
      if (index === 0) {
        return null; // Ignoramos o cabeçalho  
      }

      const colA = row["A"] !== undefined ? row["A"] : ""; // Conteúdo da coluna A  
      const colG = row["G"] !== undefined ? row["G"] : ""; // Conteúdo da coluna G  
      return `localizarNomeEInserirNota("${colA}", ${colG.replace(",", ".")})`;
    }).filter(Boolean); // Remove valores nulos  

    // Exibe as fórmulas no elemento <textarea>  
    output.value = formulas.join("\n");
  } catch (error) {
    console.error("Erro ao processar arquivo:", error);
    output.value = "Erro ao processar arquivo. Verifique se ele está no formato correto.";
  }
}

// Função para ler o arquivo XLSX e converter para JSON  
function readXLSX(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (event) => {
      try {
        const workbook = XLSX.read(event.target.result, { type: "binary" });
        const firstSheet = workbook.SheetNames[0];
        const data = XLSX.utils.sheet_to_json(workbook.Sheets[firstSheet], { defval: "" });
        resolve(data);
      } catch (error) {
        reject(error);
      }
    };
    reader.onerror = (error) => reject(error);
    reader.readAsBinaryString(file);
  });
}  