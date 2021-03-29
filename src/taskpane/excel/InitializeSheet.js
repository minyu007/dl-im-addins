import excelColumnName from "excel-column-name";

const titleArr = [
  {
    values: "Base Case",
    fillColor: "#2A4979",
    fontColor: "white"
  },
  {
    values: "Bull Case",
    fillColor: "#4EAD5B",
    fontColor: "white"
  },
  {
    values: "Bear Case",
    fillColor: "#B02318",
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

export const getDriver = async (row, callback) => {
  try {
    await Excel.run(async context => {
      const sheet = context.workbook.worksheets.getFirst();

      const estimateRanges = sheet.findAll("E", {
        completeMatch: true,
        matchCase: false
      });

      const totalRevenuesRange = sheet.findAll("Total revenues", {
        completeMatch: true,
        matchCase: false
      });

      estimateRanges.load("address, values");
      totalRevenuesRange.load("address, values");

      await context.sync();
      const address = estimateRanges.address.replace(/(\d)+/g, row);
      const addressArr = address.split(",");
      const arr = addressArr.map(v => {
        const range = sheet.getRange(v);
        range.load("address, values,valueTypes, formulas, format, numberFormat, formulasLocal, formulasR1C1");
        return range;
      });
      const titleAddress = totalRevenuesRange.address.replace(/(\d)+/g, row);
      const titleRange = sheet.getRange(titleAddress);
      titleRange.load("address, values,valueTypes, formulas, format, numberFormat, formulasLocal, formulasR1C1");

      await context.sync();
      if (callback) {
        callback(
          arr.map(v => {
            return {
              address: v.address,
              values: v.values,
              format: v.format,
              valueTypes: v.valueTypes,
              numberFormat: v.numberFormat,
              formulasLocal: v.formulasLocal,
              formulasR1C1: v.formulasR1C1
            };
          }),
          {
            address: titleRange.address,
            values: titleRange.values,
            format: titleRange.format,
            valueTypes: titleRange.valueTypes,
            numberFormat: titleRange.numberFormat,
            formulasLocal: titleRange.formulasLocal,
            formulasR1C1: titleRange.formulasR1C1
          }
        );
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
    const arr = yearPosStr.split(",");
    const rangeArr = [];
    const valuesArr = [];
    for (let i = 0, l = arr.length; i < l; i++) {
      const range = sheet.getRange(arr[i]);
      range.load("values");
      rangeArr.push(range);
    }
    await context.sync();

    for (let i = 0, l = rangeArr.length; i < l; i++) {
      valuesArr.push(rangeArr[i].values);
    }
    callback(valuesArr);
  });
};

export const drawTableFixedMatrix = async (aYear, eYear, matrix) => {
  if (!aYear || !eYear) {
    return;
  }
  const _aYear = aYear.split(",");
  const _eYear = eYear.split(",");
  const aYearLen = _aYear.length;
  const eYearLen = _eYear.length;
  const matrixLen = matrix.length;
  await Excel.run(async context => {
    const sheets = context.workbook.worksheets;
    let sheet = sheets.getLast();
    sheet.load("name");
    await context.sync();
    if (sheet.name == "Scenario Analysis") {
      sheet.delete();
    }
    sheet = sheets.add("Scenario Analysis");
    sheet.showGridlines = false;
    sheet.freezePanes.freezeAt(sheet.getRange("A1:A2"));
    sheet.activate();
    const eYearRangeArr = [];
    const caseRangeArr = [];
    const titleRangeArr = [];

    for (let i = 0, l = titleArr.length; i < l; i++) {
      const title = titleArr[i];
      sheet
        .getRange(
          `${excelColumnName.intToExcelCol(i * eYearLen + 2 + aYearLen)}1:${excelColumnName.intToExcelCol(
            i * eYearLen + 1 + aYearLen + eYearLen
          )}1`
        )
        .merge();
      const caseArr = [];
      for (let ii = 0, ll = _eYear.length; ii < ll; ii++) {
        const range = sheet.getRange(`${excelColumnName.intToExcelCol(i * eYearLen + 2 + aYearLen + ii)}2`);
        range.formulas = [[`=${_eYear[ii]}`]];
        range.format.font.bold = true;
        range.load("address");
        caseArr.push(range);
      }
      const _range = sheet.getRange(`${excelColumnName.intToExcelCol(i * eYearLen + 2 + aYearLen)}1`);
      _range.values = title.values;
      _range.format.fill.color = title.fillColor;
      _range.format.font.color = title.fontColor;
      _range.load("address");
      caseRangeArr.push(_range);
      eYearRangeArr.push(caseArr);
    }

    sheet.getRange(`A1`).format.columnWidth = 120;
    const range = sheet.getRange(`B1:${excelColumnName.intToExcelCol(aYearLen + 1)}1`);
    range.merge();
    range.values = "Actual";
    range.format.fill.color = "#808080";
    range.format.font.color = "white";

    const aYearRangeArr = [];
    for (let i = 0, l = _aYear.length; i < l; i++) {
      const range = sheet.getRange(`${excelColumnName.intToExcelCol(i + 2)}2`);
      range.formulas = [[`=${_aYear[i]}`]];
      range.format.font.bold = true;
      range.load("address");
      aYearRangeArr.push(range);
    }
    sheet.getRange(`B:${excelColumnName.intToExcelCol(aYearLen)}`).group("ByColumns");
    sheet.getRange(`B:${excelColumnName.intToExcelCol(aYearLen)}`).hideGroupDetails("ByColumns");

    const aYearValueRangeArr = [];
    for (let i = 0, l = matrixLen; i < l; i++) {
      const title = matrix[i][0];
      const titleRange = sheet.getRange(`A${i + 3}`);
      titleRange.formulas = [[`=${title}`]];
      titleRange.format.font.bold = true;
      titleRange.load("address");
      titleRangeArr.push(titleRange);
      // for (let ii = 1, ll = matrix[i].length; ii < ll; ii++) {}
      const arr = [];
      for (let ii = 1, ll = matrix[i].length; ii < ll; ii++) {
        const col = matrix[i][ii];
        const range = sheet.getRange(`${excelColumnName.intToExcelCol(ii + 1)}${i + 3}`);
        range.formulas = [[`=${col}`]];
        range.numberFormat = "0.00";
        range.load("address");
        arr.push(range);
      }
      aYearValueRangeArr.push(arr);
    }
    sheet.getRange(`A${matrixLen + 5}`).values = "Current Drivers:";
    sheet.getRange(`A${matrixLen + 5}`).format.font.bold = true;

    await context.sync();
    for (let iii = 0, lll = titleRangeArr.length; iii < lll; iii++) {
      for (let i = 0, l = aYearRangeArr.length; i < l; i++) {
        const range = sheet.getRange(
          `${excelColumnName.intToExcelCol(2 + aYearRangeArr.length + i)}${matrixLen +
            60 +
            (caseRangeArr.length + 2) * iii}`
        );
        range.formulas = `=${aYearRangeArr[i].address}`;
        range.format.font.bold = true;
      }
      for (let i = 0, l = eYearRangeArr.length; i < l; i++) {
        const eYearRange = eYearRangeArr[i];
        if (i == 0) {
          for (let ii = 0, ll = eYearRange.length; ii < ll; ii++) {
            const range = sheet.getRange(
              `${excelColumnName.intToExcelCol(2 + aYearLen * 2 + ii)}${matrixLen +
                60 +
                (caseRangeArr.length + 2) * iii}`
            );
            range.formulas = `=${eYearRange[ii].address}`;
            range.format.font.bold = true;
          }
        }
      }
    }

    for (let i = 0, l = titleRangeArr.length; i < l; i++) {
      const range = sheet.getRange(
        `${excelColumnName.intToExcelCol(aYearLen + 1)}${matrixLen + 61 + (caseRangeArr.length + 2) * i}`
      );
      range.formulas = `=${titleRangeArr[i].address}`;
      range.format.font.bold = true;
      for (let ii = 0, ll = caseRangeArr.length; ii < ll; ii++) {
        const range = sheet.getRange(
          `${excelColumnName.intToExcelCol(aYearLen + 1)}${matrixLen + 61 + (caseRangeArr.length + 2) * i + ii + 1}`
        );
        range.formulas = `=${caseRangeArr[ii].address}`;
        range.format.font.bold = true;
      }
    }

    for (let i = 0, l = aYearValueRangeArr.length; i < l; i++) {
      const arr = aYearValueRangeArr[i];
      for (let ii = 0, ll = arr.length; ii < ll; ii++) {
        const range = sheet.getRange(
          `${excelColumnName.intToExcelCol(ii + 2 + aYearLen)}${matrixLen + 61 + (caseRangeArr.length + 2) * i}`
        );
        range.formulas = `=${arr[ii].address}`;
        range.numberFormat = "0.00";
      }
      for (let ii = 0, ll = caseRangeArr.length; ii < ll; ii++) {
        const range = sheet.getRange(
          `${excelColumnName.intToExcelCol(arr.length + 1 + aYearLen)}${matrixLen +
            61 +
            (caseRangeArr.length + 2) * i +
            ii +
            1}`
        );
        range.formulas = `=${arr[arr.length - 1].address}`;
        range.numberFormat = "0.00";
        for (let iii = 0, lll = eYearRangeArr[ii].length; iii < lll; iii++) {
          const range = sheet.getRange(
            `${excelColumnName.intToExcelCol(arr.length + 2 + iii + aYearLen)}${matrixLen +
              62 +
              (caseRangeArr.length + 2) * i +
              ii}`
          );
          range.formulas = `=${eYearRangeArr[ii][iii].address.replace(/(\d)+/g, n => parseInt(n) + 1 + i)}`;
          range.numberFormat = "0.00";
        }
      }
    }
    await context.sync();
    sheet.getRange("A1").select();
  });
};

export const drawChart = async () => {
  await Excel.run(async context => {
    const sheets = context.workbook.worksheets;
    const sheet = sheets.getItem("Scenario Analysis");
    const dataRange = sheet.getRange("G63:P67");
    const chart = sheet.charts.add("XYScatterSmooth", dataRange, "auto");
    chart.setPosition("H12", "L25");
    chart.title.text = "Total Revenue";
    chart.legend.position = "bottom";
    chart.legend.format.fill.setSolidColor("white");

    let series = chart.series;
    let series0 = series.getItemAt(0);
    let series1 = series.getItemAt(1);
    let series2 = series.getItemAt(2);
    let series3 = series.getItemAt(3);
    series1.markerStyle = "Dot";
    series2.markerStyle = "Dot";
    series3.markerStyle = "Dot";
    series1.markerForegroundColor = "#2A4979";
    series2.markerForegroundColor = "#4EAD5B";
    series3.markerForegroundColor = "#B02318";

    const dataRange2 = sheet.getRange("G68:P72");
    const chart2 = sheet.charts.add("Line", dataRange2, "auto");
    chart2.setPosition("M12", "Q25");
    chart2.title.text = "Pre Exceptional EBIT";
    chart2.legend.position = "bottom";
    chart2.legend.format.fill.setSolidColor("white");

    const dataRange3 = sheet.getRange("G73:P77");
    const chart3 = sheet.charts.add("Line", dataRange3, "auto");
    chart3.setPosition("H26", "L38");
    chart3.title.text = "Pre Exceptional EPS";
    chart3.legend.position = "bottom";
    chart3.legend.format.fill.setSolidColor("white");

    await context.sync();
  });
};

export const updateDriverNames = async (len, driver) => {
  await Excel.run(async context => {
    const sheets = context.workbook.worksheets;
    let sheet = sheets.getItem("Scenario Analysis");
    // sheet.getRange("A10").values = 1;
    for (let i = 0, l = driver.length; i < l; i++) {
      const range = sheet.getRange(`A${len + 6 + i}`);
      if (driver[i].driverName) {
        range.values = (driver[i].driverName + " - row( " + driver[i].row + ")").trim();
      } else {
        range.values = "";
      }
    }
    await context.sync();
  });
};

export const setDriverIntoFirstSheet = async (address, values, callback) => {
  await Excel.run(async context => {
    const sheets = context.workbook.worksheets;
    let firstSheet = sheets.getFirst();
    firstSheet.getRange(address).values = values;
    await context.sync();
    await callback();
    await context.sync();
  });
};

export const resetDrivers = async ranges => {
  // return new Promise(function(resolve) {
  await Excel.run(async function(context) {
    const sheets = context.workbook.worksheets;
    let firstSheet = sheets.getFirst();
    for (let i = 0, l = ranges.length; i < l; i++) {
      const range = ranges[i];
      if (range.formulasLocal) {
        firstSheet.getRange(range.address).formulasLocal = range.formulasLocal;
      } else {
        firstSheet.getRange(range.address).values = range.values;
      }
    }
    await context.sync();
    // return context.sync().then(function() {
    //   resolve();
    // });
    // }).catch(errorHandlerFunction);
  });
  // await Excel.run(async context => {
  //   await callback();
  //   await context.sync();
  // });
};

// export const drawChartTable = async () => {
//   await Excel.run(async context => {
//     const sheets = context.workbook.worksheets;
//     let sheet = sheets.getItem("Scenario Analysis");

//     await context.sync();
//   });
// };

export const updateTableBaseCase = async (pos, matrix) => {
  if (!matrix || !matrix.length || !pos) {
    return;
  }
  await Excel.run(async context => {
    const sheets = context.workbook.worksheets;
    let firstSheet = sheets.getFirst();
    let sheet = sheets.getItem("Scenario Analysis");
    const _matrix = matrix.map(v => {
      return v.map(vv => {
        const range = firstSheet.getRange(vv);
        range.load("values");
        return range;
      });
    });
    await context.sync();
    for (let i = 0, l = _matrix.length; i < l; i++) {
      for (let ii = 0, ll = _matrix[i].length; ii < ll; ii++) {
        const cell = _matrix[i][ii];
        const range = sheet.getRange(`${excelColumnName.intToExcelCol(pos.startCol + ii)}${i + 3}`);
        range.values = cell.values;
        range.numberFormat = "0.00";
      }
    }
    await context.sync();
  });
};

export const updateTableBullCase = async (pos, matrix) => {
  if (!matrix || !matrix.length || !pos) {
    return;
  }
  await Excel.run(async context => {
    const sheets = context.workbook.worksheets;
    let firstSheet = sheets.getFirst();
    let sheet = sheets.getItem("Scenario Analysis");
    const _matrix = matrix.map(v => {
      return v.map(vv => {
        const range = firstSheet.getRange(vv);
        range.load("values");
        return range;
      });
    });
    await context.sync();
    for (let i = 0, l = _matrix.length; i < l; i++) {
      for (let ii = 0, ll = _matrix[i].length; ii < ll; ii++) {
        const cell = _matrix[i][ii];
        const range = sheet.getRange(`${excelColumnName.intToExcelCol(pos.startCol + ii)}${i + 3}`);
        range.values = cell.values;
        range.numberFormat = "0.00";
      }
    }
    await context.sync();
  });
};

export const updateTableBearCase = async (pos, matrix) => {
  if (!matrix || !matrix.length || !pos) {
    return;
  }
  await Excel.run(async context => {
    const sheets = context.workbook.worksheets;
    let firstSheet = sheets.getFirst();
    let sheet = sheets.getItem("Scenario Analysis");
    const _matrix = matrix.map(v => {
      return v.map(vv => {
        const range = firstSheet.getRange(vv);
        range.load("values");
        return range;
      });
    });
    await context.sync();
    for (let i = 0, l = _matrix.length; i < l; i++) {
      for (let ii = 0, ll = _matrix[i].length; ii < ll; ii++) {
        const cell = _matrix[i][ii];
        const range = sheet.getRange(`${excelColumnName.intToExcelCol(pos.startCol + ii)}${i + 3}`);
        range.values = cell.values;
        range.numberFormat = "0.00";
      }
    }
    await context.sync();
  });
};

export const activeCaseByPos = async (toPos, fromPos) => {
  await Excel.run(async context => {
    const sheets = context.workbook.worksheets;
    let sheet = sheets.getItem("Scenario Analysis");
    if (fromPos) {
      const range = sheet.getRange(
        `${excelColumnName.intToExcelCol(fromPos.startCol)}${fromPos.startRow - 1}:${excelColumnName.intToExcelCol(
          fromPos.endCol
        )}${fromPos.endRow}`
      );
      range.format.fill.color = "white";
    }
    const range = sheet.getRange(
      `${excelColumnName.intToExcelCol(toPos.startCol)}${toPos.startRow - 1}:${excelColumnName.intToExcelCol(
        toPos.endCol
      )}${toPos.endRow}`
    );
    range.format.fill.color = "#FFFED1";
    await context.sync();
  });
};

export const duplicateSheet = async worksheetName => {
  await Excel.run(async context => {
    var worksheet = context.workbook.worksheets.getFirst();
    var range = worksheet.getUsedRange();
    range.load("values", "address");
    var newWorksheet = context.workbook.worksheets.add("Backup");
    await context.sync();
    var newAddress = range.address.substring(range.address.indexof("!") + 1);
    newWorksheet.getRange(newAddress).values = range.values;
  });
};

// export const inactiveAllCase = async (pos1, pos2, pos3) => {
//   await Excel.run(async context => {
//     const sheets = context.workbook.worksheets;
//     let sheet = sheets.getItem("Scenario Analysis");

//     await context.sync();
//   });
// };
