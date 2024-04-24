
var page_number = 0;
var content = [];

const input_excel = document.getElementById("input_excel");
const textContent = document.getElementById("text_content");
const parse_button = document.getElementById("parse_button");
const reset_button = document.getElementById("reset_button");
const drop_zone = document.getElementById("drop_zone");
const prev_page = document.getElementById("prev_page");
const next_page = document.getElementById("next_page");
const page = document.getElementById("page");

prev_page.addEventListener('click', handlePrevPage, false);
next_page.addEventListener('click', handleNextPage, false);
reset_button.addEventListener('click', handleReset, false);

function handleReset() {
  var dt = new DataTransfer();
  input_excel.files = dt.files;
  textContent.innerHTML = null;
  page.innerHTML = null;
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

function highlight(e) {
  drop_zone.classList.add('highlight');
}

function unhighlight(e) {
  drop_zone.classList.remove('highlight');
}

parse_button.addEventListener('click', handleParse, false)

function handleParse() {
  const file = input_excel.files[0];
  if(file) {
    const reader = new FileReader();
    reader.onload = (e) => {
      const data = new Uint8Array(e.target.result);
      const workbook = XLSX.read(data, {type: 'array'});
      workbook.SheetNames.forEach((sheetName)=>{
        const worksheet = workbook.Sheets[sheetName];
        // const json = XLSX.utils.sheet_to_json(worksheet);
        const htmlstr = XLSX.utils.sheet_to_html(worksheet, {editable: false});
        content.push(htmlstr);
      })
      page_number = 0;
      page_function(page_number);
    };
    reader.readAsArrayBuffer(file);
  }
}
