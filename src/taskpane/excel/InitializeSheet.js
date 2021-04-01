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
      let totalRevenuesRange = sheet.findAllOrNullObject("Total revenues", {
        completeMatch: true,
        matchCase: false
      });

      let EBITRange = sheet.findAllOrNullObject("Pre Exceptional EBIT", {
        completeMatch: true,
        matchCase: false
      });

      let EPSRange = sheet.findAllOrNullObject("Pre Exceptional EPS", {
        completeMatch: true,
        matchCase: false
      });

      await context.sync();
      if (!totalRevenuesRange.isNullObject) {
        totalRevenuesRange.load("address, values");
      }
      if (!EBITRange.isNullObject) {
        EBITRange.load("address, values");
      }
      if (!EPSRange.isNullObject) {
        EPSRange.load("address, values");
      }

      eRanges.load("address, values");
      aRanges.load("address, values");

      sheet.load("name");

      await context.sync();

      let str1 = eRanges.address.replace(/(\d)+/g, n => {
        return --n;
      });
      let str2 = aRanges.address.replace(/(\d)+/g, n => {
        return --n;
      });

      let rangeE = str1.split(",");
      let rangeA = str2.split(",");

      rangeE = rangeE.map(v => {
        const range = sheet.getRange(v);
        range.load("address, values");
        return range;
      });
      rangeA = rangeA.map(v => {
        const range = sheet.getRange(v);
        range.load("address, values");
        return range;
      });

      await context.sync();

      const arr1 = [];
      const arr2 = [];
      for (let i = 0, l = rangeE.length; i < l; i++) {
        const values = `${rangeE[i].values}`;
        const address = rangeE[i].address;
        if (values.length == 4 && /^[0-9]+[0-9]*[0-9]*$/.test(values) && parseInt(values) >= 2015) {
          arr1.push(address);
        }
      }
      for (let i = 0, l = rangeA.length; i < l; i++) {
        const values = `${rangeA[i].values}`;
        const address = rangeA[i].address;
        if (values.length == 4 && /^[0-9]+[0-9]*[0-9]*$/.test(values) && parseInt(values) >= 2015) {
          arr2.push(address);
        }
      }
      const strE = arr1.join(",");
      const strA = arr2.join(",");

      let rangeObj = {
        year: {
          e: strE,
          a: strA
        }
      };
      if (!totalRevenuesRange.isNullObject) {
        rangeObj["totalRevenues"] = {
          e: strE.replace(/(\d)+/g, totalRevenuesRange.address.replace(/^\D+/g, "")),
          a: strA.replace(/(\d)+/g, totalRevenuesRange.address.replace(/^\D+/g, "")),
          title: totalRevenuesRange.address
        };
      }
      if (!EBITRange.isNullObject) {
        rangeObj["EBIT"] = {
          e: strE.replace(/(\d)+/g, EBITRange.address.replace(/^\D+/g, "")),
          a: strA.replace(/(\d)+/g, EBITRange.address.replace(/^\D+/g, "")),
          title: EBITRange.address
        };
      }

      if (!EPSRange.isNullObject) {
        rangeObj["EPSRange"] = {
          e: strE.replace(/(\d)+/g, EPSRange.address.replace(/^\D+/g, "")),
          a: strA.replace(/(\d)+/g, EPSRange.address.replace(/^\D+/g, "")),
          title: EPSRange.address
        };
      }

      for (let i = 0, l = arr.length; i < l; i++) {
        rangeObj[`customize${i}`] = {
          e: strE.replace(/(\d)+/g, arr[i]),
          a: strA.replace(/(\d)+/g, arr[i]),
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
    const firstSheet = context.workbook.worksheets.getFirst();
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
        range.format.fill.color = title.fillColor;
        range.format.font.color = title.fontColor;
        range.format.horizontalAlignment = "Center";
        range.load("address");
        caseArr.push(range);
      }
      const _range = sheet.getRange(`${excelColumnName.intToExcelCol(i * eYearLen + 2 + aYearLen)}1`);
      _range.values = title.values;
      _range.format.font.bold = true;
      _range.format.fill.color = title.fillColor;
      _range.format.font.color = title.fontColor;
      _range.format.horizontalAlignment = "Center";
      _range.load("address,numberFormat");
      caseRangeArr.push(_range);
      eYearRangeArr.push(caseArr);
    }

    sheet.getRange(`A1`).format.columnWidth = 120;
    const range = sheet.getRange(`B1:${excelColumnName.intToExcelCol(aYearLen + 1)}1`);
    range.merge();
    range.values = "Actual";
    range.format.font.bold = true;
    range.format.fill.color = "#808080";
    range.format.font.color = "white";
    range.format.horizontalAlignment = "Center";

    const aYearRangeArr = [];
    for (let i = 0, l = _aYear.length; i < l; i++) {
      const range = sheet.getRange(`${excelColumnName.intToExcelCol(i + 2)}2`);
      range.formulas = [[`=${_aYear[i]}`]];
      range.format.font.bold = true;
      range.format.fill.color = "#808080";
      range.format.font.color = "white";
      range.format.horizontalAlignment = "Center";
      range.load("address,numberFormat");
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
      titleRange.load("address, numberFormat");
      titleRangeArr.push(titleRange);
      // for (let ii = 1, ll = matrix[i].length; ii < ll; ii++) {}
      const arr = [];
      for (let ii = 1, ll = matrix[i].length; ii < ll; ii++) {
        const col = matrix[i][ii];
        const range = sheet.getRange(`${excelColumnName.intToExcelCol(ii + 1)}${i + 3}`);
        const _range = firstSheet.getRange(col);
        _range.load("numberFormat");
        await context.sync();
        range.formulas = [[`=${col}`]];
        range.numberFormat = _range.numberFormat;
        range.load("address", "numberFormat");
        arr.push(range);
      }
      aYearValueRangeArr.push(arr);
    }

    const driverRange = sheet.getRange(`A${matrixLen + 3}`);
    driverRange.values = "Drivers";
    driverRange.format.font.bold = true;
    driverRange.format.font.color = "blue";

    await context.sync();
    for (let iii = 0, lll = titleRangeArr.length; iii < lll; iii++) {
      for (let i = 0, l = aYearRangeArr.length; i < l; i++) {
        const range = sheet.getRange(
          `${excelColumnName.intToExcelCol(2 + aYearRangeArr.length + i)}${matrixLen +
            100 +
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
                100 +
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
        `${excelColumnName.intToExcelCol(aYearLen + 1)}${matrixLen + 101 + (caseRangeArr.length + 2) * i}`
      );
      range.formulas = `=${titleRangeArr[i].address}`;
      range.format.font.bold = true;
      for (let ii = 0, ll = caseRangeArr.length; ii < ll; ii++) {
        const range = sheet.getRange(
          `${excelColumnName.intToExcelCol(aYearLen + 1)}${matrixLen + 101 + (caseRangeArr.length + 2) * i + ii + 1}`
        );
        range.formulas = `=${caseRangeArr[ii].address}`;
        range.format.font.bold = true;
      }
    }

    for (let i = 0, l = aYearValueRangeArr.length; i < l; i++) {
      const arr = aYearValueRangeArr[i];
      for (let ii = 0, ll = arr.length; ii < ll; ii++) {
        const range = sheet.getRange(
          `${excelColumnName.intToExcelCol(ii + 2 + aYearLen)}${matrixLen + 101 + (caseRangeArr.length + 2) * i}`
        );
        range.formulas = `=${arr[ii].address}`;
        range.numberFormat = arr[ii].numberFormat;
      }
      for (let ii = 0, ll = caseRangeArr.length; ii < ll; ii++) {
        const range = sheet.getRange(
          `${excelColumnName.intToExcelCol(arr.length + 1 + aYearLen)}${matrixLen +
            101 +
            (caseRangeArr.length + 2) * i +
            ii +
            1}`
        );
        range.formulas = `=${arr[arr.length - 1].address}`;
        range.numberFormat = arr[arr.length - 1].numberFormat;
        for (let iii = 0, lll = eYearRangeArr[ii].length; iii < lll; iii++) {
          const range = sheet.getRange(
            `${excelColumnName.intToExcelCol(arr.length + 2 + iii + aYearLen)}${matrixLen +
              102 +
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

export const updateDriverNames = async (len1, len2, driver) => {
  await Excel.run(async context => {
    const sheets = context.workbook.worksheets;
    let sheet = sheets.getItem("Scenario Analysis");
    const range1 = sheet.getRange(`A${len1 + 4}:${excelColumnName.intToExcelCol(12 + len2 + len1)}${len1 + 10}`);
    // range1.format.fill.color = "white";
    range1.clear();

    for (let i = 0, l = driver.length; i < l; i++) {
      const range = sheet.getRange(`A${len1 + 4 + i}`);
      range.values = (driver[i].driverName + " - row( " + driver[i].row + ")").trim();
      range.format.font.bold = true;

      // for (let ii = 0, ll = driver[i].cols.length; ii < ll; ii++) {
      //   const col = driver[i].cols[ii];
      //   const _range = sheet.getRange(`${excelColumnName.intToExcelCol(ii + 2 + len2)}${len1 + 4 + i}`);
      //   _range.formulas = `=${col.address}`;
      //   _range.numberFormat = col.numberFormat;
      // }
    }
    await context.sync();
  });
};

export const updateDriverData = async (pos, driver) => {
  if (!driver || !driver.length || !pos) {
    return;
  }
  await Excel.run(async context => {
    const sheets = context.workbook.worksheets;
    let firstSheet = sheets.getFirst();
    let sheet = sheets.getItem("Scenario Analysis");
    const _driver = driver.map(v => {
      return v.map(vv => {
        const range = firstSheet.getRange(vv.address);
        range.load("values, numberFormat");
        return range;
      });
    });
    await context.sync();
    for (let i = 0, l = _driver.length; i < l; i++) {
      for (let ii = 0, ll = _driver[i].length; ii < ll; ii++) {
        const cell = _driver[i][ii];
        const range = sheet.getRange(`${excelColumnName.intToExcelCol(pos.startCol + ii)}${i + 7}`);
        range.values = cell.values;
        range.numberFormat = cell.numberFormat;
      }
    }
    await context.sync();
  });
};

export const setDriverIntoFirstSheet = async (address, values, callback) => {
  return new Promise(function(resolve) {
    Excel.run(function(context) {
      const sheets = context.workbook.worksheets;
      let firstSheet = sheets.getFirst();
      firstSheet.getRange(address).values = values;
      return context.sync().then(function() {
        resolve();
      });
    });
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
  });
};

export const updateTable = async (pos, matrix) => {
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
        range.load("values, numberFormat");
        return range;
      });
    });
    await context.sync();
    for (let i = 0, l = _matrix.length; i < l; i++) {
      for (let ii = 0, ll = _matrix[i].length; ii < ll; ii++) {
        const cell = _matrix[i][ii];
        const range = sheet.getRange(`${excelColumnName.intToExcelCol(pos.startCol + ii)}${i + 3}`);
        range.values = cell.values;
        range.numberFormat = cell.numberFormat;
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
        `${excelColumnName.intToExcelCol(fromPos.startCol)}${fromPos.startRow}:${excelColumnName.intToExcelCol(
          fromPos.endCol
        )}${fromPos.endRow}`
      );

      range.format.fill.color = "white";
    }
    const range = sheet.getRange(
      `${excelColumnName.intToExcelCol(toPos.startCol)}${toPos.startRow}:${excelColumnName.intToExcelCol(
        toPos.endCol
      )}${toPos.endRow}`
    );
    range.format.fill.color = "yellow";
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

export const drawChart = async (aYear, eYear, matrix) => {
  await Excel.run(async context => {
    const _aYear = aYear.split(",");
    const _eYear = eYear.split(",");
    const aYearLen = _aYear.length;
    const eYearLen = _eYear.length;
    const matrixLen = matrix.length;

    const sheets = context.workbook.worksheets;
    const sheet = sheets.getItem("Scenario Analysis");

    for (let i = 0, l = matrixLen; i < l; i++) {
      const titleRange = sheet.getRange(`A${i + 3}`);
      const col1Range = sheet.getRange(`B${i + 3}`);
      const b2Range = sheet.getRange(`B2`);
      const bMaxRange = sheet.getRange(`${excelColumnName.intToExcelCol(2 + aYearLen + aYearLen)}2`);
      titleRange.load("values");
      col1Range.load("values");
      b2Range.load("values");
      bMaxRange.load("values");
      await context.sync();

      // const dataRange = sheet.getRange("G63:P67");
      const dataRange = sheet.getRange(
        `${excelColumnName.intToExcelCol(1 + aYearLen)}${matrixLen +
          100 +
          i * (2 + titleArr.length)}:${excelColumnName.intToExcelCol(1 + aYearLen + aYearLen + eYearLen)}${matrixLen +
          100 +
          1 +
          titleArr.length +
          i * (2 + titleArr.length)}`
      );
      const chart = sheet.charts.add("XYScatterSmooth", dataRange, "auto");
      chart.setPosition(
        `${excelColumnName.intToExcelCol(aYearLen + 1 + (i % 2 == 0 ? 1 : 10))}${matrixLen +
          10 +
          i * 11 +
          (i % 2 == 0 ? 0 : -11)}`,
        `${excelColumnName.intToExcelCol(aYearLen + 1 + (i % 2 == 0 ? 9 : 18))}${matrixLen +
          10 +
          11 +
          i * 11 +
          (i % 2 == 0 ? 10 : -1)}`
      );

      chart.title.text = titleRange.values + "";

      chart.legend.position = "bottom";
      chart.legend.format.fill.setSolidColor("white");
      const series0 = chart.series.getItemAt(0);
      const series1 = chart.series.getItemAt(1);
      const series2 = chart.series.getItemAt(2);
      const series3 = chart.series.getItemAt(3);

      series0.format.line.color = "#808080";
      series1.format.line.color = "#2A4979";
      series2.format.line.color = "#4EAD5B";
      series3.format.line.color = "#B02318";

      // series0.markerStyle = "Dot";
      // series1.markerStyle = "Dot";
      // series2.markerStyle = "Dot";
      // series3.markerStyle = "Dot";
      series0.markerBackgroundColor = "white";
      series1.markerBackgroundColor = "white";
      series2.markerBackgroundColor = "white";
      series3.markerBackgroundColor = "white";

      series0.markerForegroundColor = "#808080";
      series1.markerForegroundColor = "#2A4979";
      series2.markerForegroundColor = "#4EAD5B";
      series3.markerForegroundColor = "#B02318";

      series1.format.line.lineStyle = "Dash";
      series2.format.line.lineStyle = "Dash";
      series3.format.line.lineStyle = "Dash";

      const valueAxis = chart.axes.valueAxis;
      const categoryAxis = chart.axes.categoryAxis;

      valueAxis.minimum = parseFloat(col1Range.values);
      categoryAxis.minimum = parseFloat(b2Range.values);
      categoryAxis.maximum = parseFloat(bMaxRange.values);
      await context.sync();
    }
  });
};
