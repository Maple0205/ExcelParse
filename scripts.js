
var page_number = 0;
var content = [];
var allFiles = [];
var sheets = [];

const input_excel = document.getElementById("input_excel");
const textContent = document.getElementById("text_content");
const parse_button = document.getElementById("parse_button");
const reset_button = document.getElementById("reset_button");
const merge_button = document.getElementById("merge_button");
const drop_zone = document.getElementById("drop_zone");
const prev_page = document.getElementById("prev_page");
const next_page = document.getElementById("next_page");
const page = document.getElementById("page");

prev_page.addEventListener('click', handlePrevPage, false);
next_page.addEventListener('click', handleNextPage, false);
reset_button.addEventListener('click', handleReset, false);
merge_button.addEventListener('click', handleMerge, false);

function handleReset() {
  var dt = new DataTransfer();
  content = [];
  input_excel.files = dt.files;
  textContent.innerHTML = content;
  page.innerHTML = null;
  allFiles = [];
  sheets = [];
  drop_zone.innerHTML = "Drag and drop files here";
}

const page_function =(page_number)=> {
  textContent.innerHTML = content[page_number];
  page.innerHTML = `${page_number+1} | ${content.length}`;
}

function handlePrevPage() {
  if (page_number>=1) {
    page_number--;
    page_function(page_number);
  }
}

function handleNextPage() {
  if (page_number<content.length-1) {
    page_number++;
    page_function(page_number);
  }
}

drop_zone.addEventListener('dragover', function(e) {
  e.preventDefault();
});

drop_zone.addEventListener(
  "drop",
  (e) => {
    e.preventDefault();
    if (input_excel) {
      input_excel.files = e.dataTransfer.files;
      input_excel.dispatchEvent(new Event('change'));
    }
  },
  false,
);

input_excel.addEventListener("change", handleFiles, false);
function handleFiles() {
  const fileList = this.files;
  const file = fileList[0];
  if (file) {
    const type = file.name.split('.')[1];
    const size = calculateSize(file.size);
    const content = `
      <div class="file-info">File type: ${type}</div>
      <div class="file-info">File name: ${file.name}</div>
      <div class="file-info">File size: ${size}</div>
    `;
    drop_zone.innerHTML = content;
  } else {
    drop_zone.innerHTML = `<div>You haven't upload any data</div>`
  }
}

function calculateSize(numberOfBytes) {
  const units = [
    "B",
    "KB",
    "MB",
    "GB",
    "TB",
    "PB",
    "EB",
    "ZB",
    "YB",
  ];
  const exponent = Math.min(
    Math.floor(Math.log(numberOfBytes) / Math.log(1024)),
    units.length - 1,
  );
  const approx = numberOfBytes / 1024 ** exponent;
  const output =
    exponent === 0
      ? `${numberOfBytes} bytes`
      : `${approx.toFixed(3)} ${
          units[exponent]
        } (${numberOfBytes} bytes)`;
  return output;
}

['dragenter', 'dragover'].forEach(eventName => {
  drop_zone.addEventListener(eventName, highlight, false);
});

['dragleave', 'drop'].forEach(eventName => {
  drop_zone.addEventListener(eventName, unhighlight, false);
});

function highlight() {
  drop_zone.classList.add('highlight');
}

function unhighlight() {
  drop_zone.classList.remove('highlight');
}

parse_button.addEventListener('click', handleParse, false)

function handleParse() {
  const file = input_excel.files[0];
  if (file) {
    const supportedTypes = ['xlsx', 'xls'];
    const fileType = file.name.split('.').pop().toLowerCase();

    if (!supportedTypes.includes(fileType)) {
      alert("Unsupported file type. Please upload an Excel file.");
      return;
    }

    const reader = new FileReader();
    reader.onload = (e) => {
      const data = new Uint8Array(e.target.result);
      try {
        const workbook = XLSX.read(data, {type: 'array'});
        let i = 0;
        const processSheet = () => {
          if (i < workbook.SheetNames.length) {
            const sheetName = workbook.SheetNames[i];
            const worksheet = workbook.Sheets[sheetName];
            if (worksheet["!ref"]) {
              const htmlstr = XLSX.utils.sheet_to_html(worksheet, {editable: false});
              content.push(htmlstr);
            }
            i++;
            percentage = i/workbook.SheetNames.length;
            progressBarFunc(percentage);  // Update progress
            setTimeout(processSheet, 100); // Wait for 500 ms before processing the next sheet
          } else {
            // Once all sheets are processed
            if (content.length > 0) {
              page_function(page_number);
              sheets.push(workbook.SheetNames.length);
              allFiles.push(file);
            } else {
              alert("No data found in the Excel file.");
            }
          }
        };
        processSheet(); // Start processing
      } catch (error) {
        alert("Failed to parse the file: " + error.message);
      }
    };
    reader.onerror = (e) => {
      alert("Failed to read the file: " + reader.error);
    };
    reader.readAsArrayBuffer(file);
  } else {
    alert("Please select a file to parse.");
  }
}


function handleMerge() {
  const fileList = allFiles;
  if (!fileList.length) {
    alert("No files selected.");
    return;
  }

  let workbookOut = XLSX.utils.book_new();
  let sheets_number = sheets.reduce((acc, val) => acc + val, 0);
  let processedSheets = 0;

  function processFile(file, index) {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = (e) => {
        const data = new Uint8Array(e.target.result);
        const workbookIn = XLSX.read(data, { type: 'array' });
        let i = 0;
        const processSheet = () => {
          if (i < workbookIn.SheetNames.length) {
            const sheetName = workbookIn.SheetNames[i];
            const worksheet = workbookIn.Sheets[sheetName];
            XLSX.utils.book_append_sheet(workbookOut, worksheet, `${sheetName}_${index}`);
            processedSheets++;
            let percentage = processedSheets / sheets_number;
            progressBarFunc(percentage);  // Update progress
            i++;
            setTimeout(processSheet, 100);
          } else {
            resolve();
          }
        };
        processSheet();
      };
      reader.onerror = (err) => reject(err);
      reader.readAsArrayBuffer(file);
    });
  }

  async function readFile(index) {
    if (index >= fileList.length) {
      XLSX.writeFile(workbookOut, "merged_output.xlsx");
      return;
    }
    await processFile(fileList[index], index);
    readFile(index + 1);
  }
  readFile(0);
}


function progressBarFunc(percentage) {
  const progressContainer = document.getElementById('progressContainer');
  const progressBar = document.getElementById('progressBar');
  const roundPercentage = Math.round(percentage * 100);

  if (roundPercentage<100) {
    removeHidden(progressContainer);
    progressBar.style.width = roundPercentage + '%';
    progressBar.innerHTML = roundPercentage + '%';
  } else {
    progressBar.style.width = '100%';
    progressBar.innerHTML = '100%';
    setTimeout(() => {
      addHidden(progressContainer);
      alert("Parse Success!");
    }, 100);
  }
}

function removeHidden(element) {
  element.classList.remove('hidden');
}

function addHidden(element) {
  element.classList.add('hidden');
}




//switch tab
function showPage(pageId) {
  document.querySelectorAll('.page').forEach(page => {
      page.classList.add('inactive');
  });
  document.getElementById(pageId).classList.remove('inactive');
}

//convert
const { jsPDF } = window.jspdf;

const imageInput = document.getElementById('imageInput');
const imagePreview = document.getElementById('imagePreview');
const convertToPdfBtn = document.getElementById('convertToPdfBtn');
const clearImageBtn = document.getElementById('clearImageBtn');

imageInput.addEventListener('change', function(event) {
  const file = event.target.files[0];
  if (file) {
      const reader = new FileReader();
      reader.onload = function(e) {
          imagePreview.src = e.target.result;
          imagePreview.classList.remove('hidden');
          convertToPdfBtn.disabled = false;
          clearImageBtn.disabled = false;
      };
      reader.readAsDataURL(file);
  }
});

convertToPdfBtn.addEventListener('click', function() {
  const img = imagePreview;
  const pdf = new jsPDF();
  const imgWidth = img.naturalWidth;
  const imgHeight = img.naturalHeight;
  const pdfWidth = pdf.internal.pageSize.getWidth();
  const pdfHeight = (imgHeight * pdfWidth) / imgWidth;

  pdf.addImage(img.src, 'JPEG', 0, 0, pdfWidth, pdfHeight);
  pdf.save('converted.pdf');
});

clearImageBtn.addEventListener('click', function() {
  imageInput.value = '';
  imagePreview.src = '#';
  imagePreview.classList.add('hidden');
  convertToPdfBtn.disabled = true;
  clearImageBtn.disabled = true;
});

const pdfInput = document.getElementById('pdfInput');
const convertToImageBtn = document.getElementById('convertToImageBtn');
const clearPdfBtn = document.getElementById('clearPdfBtn');
const pdfPreview = document.getElementById('pdfPreview');

pdfInput.addEventListener('change', function(event) {
  const file = event.target.files[0];
  if (file) {
      convertToImageBtn.disabled = false;
      clearPdfBtn.disabled = false;
  }
});

convertToImageBtn.addEventListener('click', function() {
  const file = pdfInput.files[0];
  if (file) {
      const fileReader = new FileReader();
      fileReader.onload = function() {
          const typedarray = new Uint8Array(this.result);
          renderPdfToImages(typedarray);
      };
      fileReader.readAsArrayBuffer(file);
  }
});

clearPdfBtn.addEventListener('click', function() {
  pdfInput.value = '';
  pdfPreview.innerHTML = '';
  convertToImageBtn.disabled = true;
  clearPdfBtn.disabled = true;
});

function renderPdfToImages(data) {
  pdfPreview.innerHTML = '';
  pdfjsLib.getDocument(data).promise.then(pdf => {
      for (let pageNum = 1; pageNum <= pdf.numPages; pageNum++) {
          pdf.getPage(pageNum).then(page => {
              const scale = 1.5;
              const viewport = page.getViewport({ scale });
              const canvas = document.createElement('canvas');
              const context = canvas.getContext('2d');
              canvas.height = viewport.height;
              canvas.width = viewport.width;

              const renderContext = {
                  canvasContext: context,
                  viewport: viewport
              };
              page.render(renderContext).promise.then(() => {
                  const img = document.createElement('img');
                  img.src = canvas.toDataURL();
                  img.classList.add('preview-img');
                  pdfPreview.appendChild(img);
              });
          });
      }
  });
}