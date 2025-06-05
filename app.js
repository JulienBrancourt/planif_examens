document.getElementById("fileInput").addEventListener("change", function (e) {
  const file = e.target.files[0];
  if (!file) return;

  const reader = new FileReader();
  reader.onload = function (e) {
    const decoder = new TextDecoder("iso-8859-1");
    const text = decoder.decode(e.target.result);
    parseCSV(text);
  };
  reader.readAsArrayBuffer(file); // Important pour bien gérer l'encodage
});

function parseCSV(content) {
  const rows = content.trim().split(/\r?\n/);
  const planning = {};
  const jours = ["lundi", "mardi", "mercredi", "jeudi", "vendredi"];

  rows.slice(1).forEach(line => {
    const cols = line.split(/,(?=(?:(?:[^"]*"){2})*[^"]*$)/).map(s => s.replace(/^"|"$/g, '').trim());

    const [day, start, end, , category, moduleFull] = cols;

    if (category !== "Examen") return;

    const jour = day.toLowerCase();
    if (!jours.includes(jour)) return;

    const moduleName = moduleFull.includes(" - ") ? moduleFull.split(" - ")[1] : moduleFull;
    const time = `${start} - ${end}`;

    if (!planning[jour]) planning[jour] = [];
    planning[jour].push({ start, time, module: moduleName });
  });

  displayPlanning(planning, jours);
}

function displayPlanning(planning, jours) {
  const container = document.getElementById("planning");
  container.innerHTML = "";

  jours.forEach(jour => {
    if (planning[jour]) {
      planning[jour].sort((a, b) => a.start.localeCompare(b.start));

      const section = document.createElement("section");
      const title = document.createElement("h3");
      title.textContent = jour;
      section.appendChild(title);

      const tableWrapper = document.createElement("div");
      tableWrapper.style.overflowX = "auto";

      const table = document.createElement("table");
      tableWrapper.appendChild(table);

      const thead = document.createElement("thead");
      const tbody = document.createElement("tbody");

      const headRow = document.createElement("tr");
      planning[jour].forEach(item => {
        const td = document.createElement("td");
        td.textContent = item.time;
        headRow.appendChild(td);
      });
      thead.appendChild(headRow);

      const dataRow = document.createElement("tr");
      planning[jour].forEach(item => {
        const td = document.createElement("td");
        td.textContent = item.module;
        dataRow.appendChild(td);
      });
      tbody.appendChild(dataRow);

      table.appendChild(thead);
      table.appendChild(tbody);

      section.appendChild(tableWrapper);

      // ✅ Bouton d'export Excel
      const exportBtn = document.createElement("button");
      exportBtn.textContent = "Exporter vers Excel";
      exportBtn.style.marginTop = "15px";
      exportBtn.style.padding = "8px 12px";
      exportBtn.style.border = "none";
      exportBtn.style.backgroundColor = "#3498db";
      exportBtn.style.color = "white";
      exportBtn.style.borderRadius = "6px";
      exportBtn.style.cursor = "pointer";

      exportBtn.addEventListener("click", () => {
        exportTableToExcel(table, `Planning_${jour}`);
      });

      section.appendChild(exportBtn);
      container.appendChild(section);
    }
  });
}

// ✅ Fonction d'export vers Excel
function exportTableToExcel(table, filename = "table") {
  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.table_to_sheet(table);
  XLSX.utils.book_append_sheet(wb, ws, "Planning");
  XLSX.writeFile(wb, `${filename}.xlsx`);
}

