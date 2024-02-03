const express = require('express');
const multer = require('multer');
const xlsx = require('xlsx');
const path = require('path');

const app = express();
const PORT = 54;

// configurar o multer para lidar com uploads
const armazenamento = multer.memoryStorage();
const upload = multer({ storage: armazenamento });

app.use(express.static(path.join(__dirname, 'public')));

app.post('/upload', upload.single('excelFile'), (req, res) => {
  try {
    // extrai o buffer do arquivo excel enviado na req
    const excelBuffer = req.file.buffer;

    // lê o arquivo excel usando a biblioteca xlsx
    const leituraExcel = xlsx.read(excelBuffer, { type: 'buffer' });
    const sheet_name_list = leituraExcel.SheetNames;
    const worksheet = leituraExcel.Sheets[sheet_name_list[0]];

    // converte a primeira planilha em um formato JSON
    const conversorPlanilha = xlsx.utils.sheet_to_json(worksheet, {
      raw: false,
      defval: '',
      header: 1,
      dateNF: 'yyyy-mm-dd'
    });

    let html_table = '<table class="table"><thead class="thead-dark"><tr>';

    // adiciona cabeçalhos
    Object.keys(conversorPlanilha[0]).forEach(column => {
      html_table += `<th scope="col">${column}</th>`;
    });

    html_table += '</tr></thead><tbody>';

    // adiciona linhas
    conversorPlanilha.forEach(row => {
      html_table += '<tr>';
      Object.values(row).forEach(celula => {
        // verifica se a celula eh uma data e converte
        if (celula instanceof Date) {
          html_table += `<td>${celula.toLocaleDateString()}</td>`;
        } else {
          html_table += `<td>${celula}</td>`;
        }
      });
      html_table += '</tr>';
    });

    html_table += '</tbody></table>';

    res.send(html_table);
  } catch (error) {
    res.status(500).send(`Erro no processamento do arquivo Excel: ${error.message}`);
  }
});

app.listen(PORT, () => {
  console.log(`Servidor rodando em http://localhost:${PORT}/`);
});
