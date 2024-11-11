Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
  }
});

async function generateTable() {
  const tableNumber = parseInt(document.getElementById("tableNumber").value);

  if (isNaN(tableNumber) || tableNumber < 1) {
      alert("Please enter a valid number greater than 0.");
      return;
  }

  await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      sheet.getRange("A1").values = [[`Multiplication Table for ${tableNumber}`]];

      const headers = [["Number", `x ${tableNumber}`]];
      sheet.getRange("A2:B2").values = headers;

      for (let row = 1; row <= 10; row++) {
          const rowData = [[row, row * tableNumber]];
          sheet.getRange(`A${row + 2}:B${row + 2}`).values = rowData;
      }

      sheet.getUsedRange().format.autofitColumns();
      sheet.getUsedRange().format.autofitRows();
      await context.sync();
  })
}

function openTab(evt, tabName) {
  document.getElementById("welcomeImage").classList.add("hidden");

  const tabContents = document.getElementsByClassName("tabcontent");
  for (let i = 0; i < tabContents.length; i++) {
      tabContents[i].classList.remove("active");
  }

  const tabLinks = document.getElementsByClassName("tablinks");
  for (let i = 0; i < tabLinks.length; i++) {
      tabLinks[i].classList.remove("active");
  }

  document.getElementById(tabName).classList.add("active");
  evt.currentTarget.classList.add("active");
}


