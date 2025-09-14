document.addEventListener("DOMContentLoaded", () => {
  const MONTHS = ["Январь","Февраль","Март","Апрель","Май","Июнь","Июль","Август","Сентябрь","Октябрь","Ноябрь","Декабрь"];
  const DOW    = ["Пн","Вт","Ср","Чт","Пт","Сб","Вс"]; // lundi→dimanche
  const today  = new Date();

  const daysInMonth = (y, m) => new Date(y, m + 1, 0).getDate();
  const mondayIndex = (jsDay) => (jsDay + 6) % 7; // JS: 0=Sun..6=Sat → 0=Mon..6=Sun

  function buildMonthTable(year, monthIndex) {
    const wrapper = document.createElement("div");
    wrapper.className = "month";

    const title = document.createElement("h3");
    title.textContent = MONTHS[monthIndex];
    wrapper.appendChild(title);

    const table = document.createElement("table");
    const thead = document.createElement("thead");
    const trh = document.createElement("tr");
    DOW.forEach(lbl => {
      const th = document.createElement("th");
      th.textContent = lbl;
      trh.appendChild(th);
    });
    thead.appendChild(trh);
    table.appendChild(thead);

    const tbody = document.createElement("tbody");

    const first = new Date(year, monthIndex, 1);
    const offset = mondayIndex(first.getDay());
    const dim = daysInMonth(year, monthIndex);
    const prevDim = daysInMonth(year, monthIndex - 1);

    let day = 1, next = 1;

    for (let w = 0; w < 6; w++) {
      const tr = document.createElement("tr");
      for (let d = 0; d < 7; d++) {
        const td = document.createElement("td");
        const idx = w * 7 + d;
        let display, isCurr = true, cellDate;

        if (idx < offset) {
          display = prevDim - (offset - 1 - idx);
          isCurr = false;
          cellDate = new Date(year, monthIndex - 1, display);
        } else if (day <= dim) {
          display = day++;
          cellDate = new Date(year, monthIndex, display);
        } else {
          display = next++;
          isCurr = false;
          cellDate = new Date(year, monthIndex + 1, display);
        }

        const num = document.createElement("div");
        num.className = "num" + (isCurr ? "" : " muted");
        num.textContent = display;
        td.appendChild(num);

        const isToday =
          isCurr &&
          cellDate.getFullYear() === today.getFullYear() &&
          cellDate.getMonth() === today.getMonth() &&
          cellDate.getDate() === today.getDate();
        if (isToday) td.classList.add("today");

        tr.appendChild(td);
      }
      tbody.appendChild(tr);

      // supprime la 6e ligne si elle ne contient pas de jours du mois courant
      if (w >= 4) {
        const hasCurrent = Array.from(tr.children).some(td => !td.firstChild.className.includes("muted"));
        if (!hasCurrent) { tbody.removeChild(tr); break; }
      }
    }

    table.appendChild(tbody);
    wrapper.appendChild(table);
    return wrapper;
  }

  function renderYear(year, containerId) {
    const grid = document.getElementById(containerId);
    for (let m = 0; m < 12; m++) grid.appendChild(buildMonthTable(year, m));
  }

  renderYear(2025, "grid-2025");
  renderYear(2026, "grid-2026");

  // -------- Export Word (.docx) --------
  document.getElementById("downloadWord").addEventListener("click", async () => {
    if (!window.docx) {
      alert("La librairie docx n'est pas chargée.");
      return;
    }
    const { Document, Packer, Paragraph, Table, TableRow, TableCell, WidthType, HeadingLevel } = window.docx;

    const doc = new Document({ sections: [] });

    [2025, 2026].forEach((year) => {
      const children = [ new Paragraph({ text: `${year} год`, heading: HeadingLevel.HEADING_1 }) ];

      for (let m = 0; m < 12; m++) {
        children.push(new Paragraph({ text: MONTHS[m], heading: HeadingLevel.HEADING_2 }));

        // Ligne d'en-têtes
        const headerRow = new TableRow({
          children: DOW.map(lbl =>
            new TableCell({
              width: { size: 100/7, type: WidthType.PERCENTAGE },
              children: [ new Paragraph(lbl) ]
            })
          )
        });

        const rows = [headerRow];

        // Calcul du mois pour Word : on laisse vide les cases hors mois
        const first = new Date(year, m, 1);
        const offset = mondayIndex(first.getDay());
        const dim = daysInMonth(year, m);
        let day = 1;

        for (let w = 0; w < 6; w++) {
          const cells = [];
          for (let d = 0; d < 7; d++) {
            const idx = w * 7 + d;
            const txt = (idx >= offset && day <= dim) ? String(day++) : "";
            cells.push(new TableCell({ children: [ new Paragraph(txt) ] }));
          }
          rows.push(new TableRow({ children: cells }));
          if (day > dim) break;
        }

        children.push(new Table({ rows }));
      }

      doc.addSection({ children });
    });

    const blob = await Packer.toBlob(doc);
    const a = document.createElement("a");
    a.href = URL.createObjectURL(blob);
    a.download = "Календарь-2025-2026.docx";
    document.body.appendChild(a);
    a.click();
    a.remove();
    setTimeout(() => URL.revokeObjectURL(a.href), 1000);
  });
});