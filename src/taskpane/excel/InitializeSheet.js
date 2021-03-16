import excelColumnName from "excel-column-name";

const titleArr = [
  {
    values: "Base Case",
    fillColor: "#4472C4",
    fontColor: "white"
  },
  {
    values: "Bull Case",
    fillColor: "#70AD47",
    fontColor: "white"
  },
  {
    values: "Bear Case",
    fillColor: "#ED7D31",
    fontColor: "white"
  }
];

export const getPosition = async (arr, callback) => {
  try {
    await Excel.run(async context => {
      const sheet = context.workbook.worksheets.getFirst();

      const eRanges = sheet.findAll("E", {
        completeMatch: true,
        matchCase: false
      });
      const aRanges = sheet.findAll("A", {
        completeMatch: true,
        matchCase: false
      });
      const totalRevenuesRange = sheet.findAll("Total revenues", {
        completeMatch: true,
        matchCase: false
      });
      const EBITRange = sheet.findAll("Pre Exceptional EBIT", {
        completeMatch: true,
        matchCase: false
      });
      const EPSRange = sheet.findAll("Pre Exceptional EPS", {
        completeMatch: true,
        matchCase: false
      });

      eRanges.load("address, values");
      aRanges.load("address, values");
      totalRevenuesRange.load("address, values");
      EBITRange.load("address, values");
      EPSRange.load("address, values");
      sheet.load("name");

      await context.sync();
      let rangeObj = {
        year: {
          e: eRanges.address.replace(/(\d)+/g, n => {
            return --n;
          }),
          a: aRanges.address.replace(/(\d)+/g, n => {
            return --n;
          })
        },
        totalRevenues: {
          e: eRanges.address.replace(/(\d)+/g, totalRevenuesRange.address.replace(/^\D+/g, "")),
          a: aRanges.address.replace(/(\d)+/g, totalRevenuesRange.address.replace(/^\D+/g, "")),
          title: totalRevenuesRange.address
        },
        EBIT: {
          e: eRanges.address.replace(/(\d)+/g, EBITRange.address.replace(/^\D+/g, "")),
          a: aRanges.address.replace(/(\d)+/g, EBITRange.address.replace(/^\D+/g, "")),
          title: EBITRange.address
        },
        EPS: {
          e: eRanges.address.replace(/(\d)+/g, EPSRange.address.replace(/^\D+/g, "")),
          a: aRanges.address.replace(/(\d)+/g, EPSRange.address.replace(/^\D+/g, "")),
          title: EPSRange.address
        }
      };

      for (let i = 0, l = arr.length; i < l; i++) {
        rangeObj[`customize${i}`] = {
          e: eRanges.address.replace(/(\d)+/g, arr[i]),
          a: aRanges.address.replace(/(\d)+/g, arr[i]),
          title: EPSRange.address.replace(/(\d)+/g, arr[i])
        };
      }

      if (callback) {
        callback(rangeObj);
      }
    });
  } catch (e) {}
};

export const getYearValues = async (yearPosStr, callback) => {
  if (!yearPosStr) {
    return;
  }
  await Excel.run(async context => {
    const sheet = context.workbook.worksheets.getFirst();
    // const range = sheet.getRange(yearPosStr);
    const arr = yearPosStr.split(",");
    const rangeArr = [];
    const valuesArr = [];
    for (let i = 0, l = arr.length; i < l; i++) {
      const range = sheet.getRange(arr[i]);
      range.load("values");
      rangeArr.push(range);
    }
    // range.load("values");
    await context.sync();

    for (let i = 0, l = rangeArr.length; i < l; i++) {
      valuesArr.push(rangeArr[i].values);
    }
    callback(valuesArr);
  });
};

export const drawTable = async yearValueArr => {
  if (!yearValueArr || !yearValueArr.length) {
    return;
  }
  const len = yearValueArr.length;
  await Excel.run(async context => {
    const sheets = context.workbook.worksheets;
    let sheet = sheets.getLast();

    sheet.load("name");
    // const firstSheet = context.workbook.worksheets.getFirst();
    // firstSheet.getRange(yearValueArr);
    await context.sync();

    if (sheet.name == "Scenario Analysis") {
      sheet.delete();
    }
    sheet = sheets.add("Scenario Analysis");
    sheet.showGridlines = false;
    sheet.freezePanes.freezeAt(sheet.getRange("A1:B3"));
    sheet.activate();

    for (let i = 0, l = titleArr.length; i < l; i++) {
      const title = titleArr[i];
      sheet
        .getRange(`${excelColumnName.intToExcelCol(3 * i + len)}2:${excelColumnName.intToExcelCol(3 * i + len + 2)}2`)
        .merge();

      for (let ii = 0, ll = yearValueArr.length; ii < ll; ii++) {
        const range = sheet.getRange(`${excelColumnName.intToExcelCol(3 * i + len + ii)}3`);
        range.values = yearValueArr[ii];
        range.format.font.bold = true;
      }

      const _range = sheet.getRange(`${excelColumnName.intToExcelCol(3 * i + len)}2`);
      _range.values = title.values;
      _range.format.fill.color = title.fillColor;
      _range.format.font.color = title.fontColor;
    }
    sheet
      .getRange(`B3:${excelColumnName.intToExcelCol(len * titleArr.length + 2)}3`)
      .format.borders.getItem("EdgeBottom").style = "Continuous";

    sheet.getRange(`B1`).format.columnWidth = 120;
  });
};

export const linkTableData = async matrix => {
  if (!matrix || !matrix.length) {
    return;
  }
  await Excel.run(async context => {
    const sheets = context.workbook.worksheets;
    let sheet = sheets.getLast();
    for (let i = 0, l = matrix.length; i < l; i++) {
      const title = matrix[i][0];
      const titleRange = sheet.getRange(`B${i + 4}`);
      titleRange.formulas = [[`=${title}`]];
      titleRange.format.font.bold = true;
      for (let ii = 1, ll = matrix[i].length; ii < ll; ii++) {
        const col = matrix[i][ii];
        {
          const range = sheet.getRange(`${excelColumnName.intToExcelCol(ii + 2)}${i + 4}`);
          range.formulas = [[`=${col}`]];
          range.numberFormat = "0.00";
        }
        {
          const range = sheet.getRange(`${excelColumnName.intToExcelCol(ii + 5)}${i + 4}`);
          range.formulas = [[`=${col}`]];
          range.numberFormat = "0.00";
        }
        {
          const range = sheet.getRange(`${excelColumnName.intToExcelCol(ii + 8)}${i + 4}`);
          range.formulas = [[`=${col}`]];
          range.numberFormat = "0.00";
        }
        // await context.sync();
      }
    }
    await context.sync();
  });
};
