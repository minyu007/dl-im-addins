import React, { useState, useEffect, useBoolean } from "react";
import _ from "lodash";
import {
  Dialog,
  DialogType,
  DialogFooter,
  PrimaryButton,
  TextField,
  Checkbox,
  Separator,
  DetailsList,
  DetailsListLayoutMode,
  Icon,
  CommandBarButton,
  MessageBarType,
  MessageBar,
  MessageBarButton,
  ActionButton,
  SelectionMode,
  SpinButton,
  Stack,
  Label,
  SearchBox,
  ChoiceGroup,
  Spinner,
  SpinnerSize
} from "office-ui-fabric-react";
// import { mergeStyles, mergeStyleSets } from "office-ui-fabric-react/lib/Styling";
/* global Button, console, Excel, Header, HeroList, HeroListItem, Progress */

import {
  getYearValues,
  drawTableFixedMatrix,
  getPosition,
  updateTableBaseCase,
  updateTableBullCase,
  updateTableBearCase,
  activeCaseByPos,
  getDriver,
  duplicateSheet,
  updateDriverNames,
  setDriverIntoFirstSheet,
  resetDrivers,
  drawChart
} from "../excel";

let tableFixedMatrix = [];
let tableDynamicMatrix = [];
let locationObjects = {};
let cacheArr = [];

let currentType = "base";

const getBasePos = () => {
  try {
    const aLen = locationObjects.year.a.split(",").length;
    const eLen = locationObjects.year.e.split(",").length;
    return Object.assign(
      {},
      {
        startCol: aLen + 2,
        endCol: aLen + 1 + eLen,
        startRow: 3,
        endRow: 2 + tableDynamicMatrix.length
      }
    );
  } catch (e) {
    return null;
  }
};
const getBullPos = () => {
  try {
    const aLen = locationObjects.year.a.split(",").length;
    const eLen = locationObjects.year.e.split(",").length;
    return Object.assign(
      {},
      {
        startCol: aLen + 2 + eLen,
        endCol: aLen + 1 + eLen * 2,
        startRow: 3,
        endRow: 2 + tableDynamicMatrix.length
      }
    );
  } catch (e) {
    return null;
  }
};
const getBearPos = () => {
  try {
    const aLen = locationObjects.year.a.split(",").length;
    const eLen = locationObjects.year.e.split(",").length;
    return Object.assign(
      {},
      {
        startCol: aLen + 2 + eLen * 2,
        endCol: aLen + 1 + eLen * 3,
        startRow: 3,
        endRow: 2 + tableDynamicMatrix.length
      }
    );
  } catch (e) {
    return null;
  }
};

const getValueByFormat = range => {
  if (range.numberFormat == "0.0%") {
    const formattedValue = Math.floor(range.values * 1000) / 1000;
    const arr = `${formattedValue}`.split(".");
    let step = 0.01;
    if (arr.length > 1) {
      let len = arr[1].length;
      if (len == 1) {
        step = 0.1;
      }
      if (len == 2) {
        step = 0.01;
      }
      if (len == 3) {
        step = 0.001;
      }
    }
    return { ...range, formattedValue: formattedValue, step: step };
  } else if (range.numberFormat == "0%") {
    const formattedValue = Math.floor(range.values * 100) / 100;
    const arr = `${formattedValue}`.split(".");
    let step = 0.01;
    if (arr.length > 1) {
      let len = arr[1].length;
      if (len == 1) {
        step = 0.1;
      }
      if (len == 2) {
        step = 0.01;
      }
    }
    return { ...range, formattedValue: formattedValue, step: step };
  } else if (range.valueTypes == "Double") {
    let values = range.values;
    let step = 1;
    if (values >= 0) {
      if (values > 10) {
        values = Math.round(values);
      } else {
        values = Math.floor(values * 10) / 10;
        step = 0.1;
      }
    } else {
      values = Math.abs(values);
      if (values > 10) {
        values = -Math.round(values);
      } else {
        values = -Math.floor(values * 10) / 10;
        step = 0.1;
      }
    }
    return { ...range, formattedValue: values, step: step };
  } else {
    return range;
  }
};

const isValidDrive = ranges => {
  let isValid = true;
  for (let i = 0, l = ranges.length; i < l; i++) {
    const range = ranges[i];
    if (range.valueTypes == "Empty" || range.valueTypes == "Error") {
      isValid = false;
    }
  }
  return isValid;
};

const App = () => {
  const options = [
    { key: "base", text: "Base", value: "base", iconProps: { iconName: "LineChart", styles: { color: "blue" } } },
    { key: "bull", text: "Bull", value: "bull", iconProps: { iconName: "Market" } },
    { key: "bear", text: "Bear", value: "bear", iconProps: { iconName: "MarketDown" } }
  ];

  const getColumns = () => {
    if (yearArr && yearArr.length) {
      const arr = [];
      arr.push({
        key: "column1",
        name: "Driver",
        fieldName: "driver",
        minWidth: 90,
        maxWidth: 90,
        isResizable: true,
        // eslint-disable-next-line react/display-name
        onRender: item => {
          return (
            <SearchBox
              placeholder="Row No."
              value={item.row ? item.row : ""}
              underlined={true}
              onClear={v => handleDriverRowClear(v, item)}
              onSearch={v => handleDriverRowSearch(v, item)}
            />
          );
        }
      });
      for (let i = 0, l = yearArr.length; i < l; i++) {
        arr.push({
          key: `column${i + 3}`,
          name: yearArr[i],
          fieldName: `y${yearArr[i]}`,
          minWidth: 90,
          maxWidth: 90,
          isResizable: true,
          // eslint-disable-next-line react/display-name
          onRender: item => (
            <SpinButton
              defaultValue={0}
              value={item[`y${i}`] ? item[`y${i}`] : 0}
              min={item.min}
              max={item.max}
              onIncrement={v => increment(v, item[`step${i}`], item[`address${i}`])}
              onDecrement={v => decrement(v, item[`step${i}`], item[`address${i}`])}
              step={item[`step${i}`] ? item[`step${i}`] : 1}
              disabled={item.disable}
            />
          )
        });
      }
      return arr;
    }
    return null;
  };

  const [testValue, setTestValue] = useState([]);
  const [test1Value, setTest1Value] = useState("");
  const [yearArr, setYearArr] = useState([]);
  const [customize1Value, setCustomize1Value] = useState("");
  const [customize2Value, setCustomize2Value] = useState("");
  const [isHideDialog, hideDialog] = useState(false);
  const [selfCustomize1Checked, setSelfCustomize1] = useState(false);
  const [selfCustomize2Checked, setSelfCustomize2] = useState(false);
  const [loading, setLoading] = useState(false);
  const [items, setItems] = useState([
    {
      key: 1,
      driver: 1,
      disable: true,
      driverName: "",
      min: -1000000,
      max: 1000000
    },
    {
      key: 2,
      driver: 2,
      disable: true,
      driverName: "",
      min: -1000000,
      max: 1000000
    },
    {
      key: 3,
      driver: 3,
      disable: true,
      driverName: "",
      min: -1000000,
      max: 1000000
    }
  ]);

  const handleInit = async () => {
    setLoading(true);
    hideDialog(true);
    const customizeArr = [];
    if (customize1Value) {
      customizeArr.push(customize1Value);
    }
    if (customize2Value) {
      customizeArr.push(customize2Value);
    }
    await getPosition(customizeArr, async payload => {
      locationObjects = Object.assign({}, payload);
      for (let key in payload) {
        if (key != "year") {
          const elem = payload[key];
          const arr1 = [elem.title];
          const arr2 = [];
          const estimateArr = elem.e.split(",");
          const actualArr = elem.a.split(",");
          for (let i = 0, l = actualArr.length; i < l; i++) {
            arr1.push(actualArr[i]);
          }
          for (let i = 0, l = estimateArr.length; i < l; i++) {
            arr2.push(estimateArr[i]);
          }
          tableFixedMatrix.push(arr1);
          tableDynamicMatrix.push(arr2);
        }
      }

      locationObjects = { ...payload };
      await getYearValues(payload.year.e, async yearValueArr => {
        setYearArr(yearValueArr);
        await drawTableFixedMatrix(payload.year.a, payload.year.e, tableFixedMatrix);
        await updateTableBaseCase(getBasePos(), tableDynamicMatrix);
        await updateTableBullCase(getBullPos(), tableDynamicMatrix);
        await updateTableBearCase(getBearPos(), tableDynamicMatrix);

        await activeCaseByPos(getBasePos());
        await drawChart();
        setLoading(false);
      });
    });
  };

  const handleSelfCustomize1 = () => {
    if (selfCustomize1Checked) {
      setCustomize1Value("");
    }
    setSelfCustomize1(!selfCustomize1Checked);
  };

  const handleSelfCustomize2 = () => {
    if (selfCustomize2Checked) {
      setCustomize2Value("");
    }
    setSelfCustomize2(!selfCustomize2Checked);
  };

  const handleDriverRowSearch = async (value, item) => {
    if (!value) {
      return;
    }
    const arr = [...items];
    const _item = arr[item.key - 1];
    await getDriver(value, async (driverArr, title) => {
      if (driverArr && driverArr.length && isValidDrive(driverArr)) {
        const index = cacheArr.findIndex(v => v.key == item.key);
        // setTest1Value(index);
        if (index != -1) {
          cacheArr.splice(index, 1);
        }

        cacheArr.push({
          key: item.key,
          row: value,
          title,
          cols: driverArr.map(v => ({
            ...v,
            formattedValue: getValueByFormat(v).formattedValue,
            step: getValueByFormat(v).step
          }))
        });

        const obj = cacheArr.find(v => v.row == value);
        _item.disable = false;
        _item.row = obj.row;
        _item.driverName = obj.title.values;
        for (let i = 0, l = obj.cols.length; i < l; i++) {
          _item[`y${i}`] = obj.cols[i].formattedValue;
          _item[`step${i}`] = obj.cols[i].step;
          _item[`address${i}`] = obj.cols[i].address;
        }
        setItems(arr);
        setTestValue(cacheArr);
        // setTest1Value({ value: cacheArr.length, address: 111 });
        await updateDriverNames(
          tableFixedMatrix.length,
          locationObjects.year.a.split(",").length,
          cacheArr.map(v => ({ driverName: v.title.values, row: v.row, cols: v.cols }))
        );
      }
    });
  };

  const handleDriverRowClear = async (e, item) => {
    const arr = [...items];
    const index = cacheArr.findIndex(v => v.row == item.row);
    // setTest1Value(index);
    if (index != -1) {
      arr[item.key - 1].disable = true;
      arr[item.key - 1].row = "";
      arr[item.key - 1].driverName = "";
      for (let i = 0, l = yearArr.length; i < l; i++) {
        arr[item.key - 1][`y${i}`] = "";
        arr[item.key - 1][`step${i}`] = "";
        arr[item.key - 1][`address${i}`] = "";
      }
      setItems(arr);

      await resetDrivers(cacheArr[index].cols);
      cacheArr.splice(index, 1);
      await updateDriverNames(
        tableFixedMatrix.length,
        locationObjects.year.a.split(",").length,
        cacheArr.map(v => ({ driverName: v.title.values, row: v.row, cols: v.cols }))
      );
      await setCase();
    }
  };

  // const resetDriver = () => {};

  const changeCaseType = async e => {
    const { value } = e.currentTarget;
    const basePos = getBasePos();
    const bullPos = getBullPos();
    const bearPos = getBearPos();
    let fromPos = null;
    if (currentType == "base") {
      fromPos = basePos;
    } else if (currentType == "bull") {
      fromPos = bullPos;
    } else {
      fromPos = bearPos;
    }
    if (value == "base" && currentType != "base") {
      await activeCaseByPos(basePos, fromPos);
      currentType = "base";
    }
    if (value == "bull" && currentType != "bull") {
      await activeCaseByPos(bullPos, fromPos);
      currentType = "bull";
    }
    if (value == "bear" && currentType != "bear") {
      await activeCaseByPos(bearPos, fromPos);
      currentType = "bear";
    }
  };

  const increment = (value, step, address) => {
    const newValue = Math.round((value + step) * 1e12) / 1e12;
    setDriver(address, newValue);
    return newValue;
  };

  const decrement = (value, step, address) => {
    const newValue = Math.round((value - step) * 1e12) / 1e12;
    setDriver(address, newValue);
    return newValue;
  };

  const setDriver = _.debounce(async (address, newValue) => {
    await setDriverIntoFirstSheet(address, newValue);
    await setCase();
  }, 1000);

  const setCase = async () => {
    if (currentType == "base") {
      await updateTableBaseCase(getBasePos(), tableDynamicMatrix);
    } else if (currentType == "bull") {
      await updateTableBullCase(getBullPos(), tableDynamicMatrix);
    } else {
      await updateTableBearCase(getBearPos(), tableDynamicMatrix);
    }
  };

  // const copy1 = async () => {
  //   setTest1Value({ value: 1, address: 2 });
  //   await duplicateSheet();
  // };
  return (
    <div className="App">
      <Dialog
        hidden={isHideDialog}
        onDismiss={handleInit}
        dialogContentProps={{
          type: DialogType.largeHeader,
          title: "Initialization",
          subText: "Please select"
        }}
        modalProps={{
          isBlocking: false,
          styles: { main: { maxWidth: 450 } }
        }}
      >
        <Stack tokens={{ childrenGap: 10 }}>
          <Checkbox label="Total revenues" disabled defaultChecked />
          <Checkbox label="Pre Exceptional EBIT" defaultChecked />
          <Checkbox label="Pre Exceptional EPS" defaultChecked />
          <Checkbox
            onChange={() => {
              handleSelfCustomize1();
            }}
            onRenderLabel={() => (
              <TextField
                label="Self-customize 1:"
                underlined
                placeholder="Row NO."
                value={customize1Value}
                disabled={!selfCustomize1Checked}
                onChange={e => {
                  setCustomize1Value(e.target.value);
                }}
              />
            )}
          />
          <Checkbox
            onChange={() => {
              handleSelfCustomize2();
            }}
            onRenderLabel={() => (
              <TextField
                label="Self-customize 2"
                value={customize2Value}
                disabled={!selfCustomize2Checked}
                underlined
                onChange={e => {
                  setCustomize2Value(e.target.value);
                }}
                placeholder="Row NO."
              />
            )}
          />
          {/* <Checkbox label="Disabled checked checkbox" /> */}
        </Stack>
        <DialogFooter>
          <PrimaryButton onClick={handleInit} text="Go" />
        </DialogFooter>
      </Dialog>
      <Separator alignContent="center">
        <Label>DaLian IM Scenario Analysis Tool</Label>
      </Separator>
      {loading ? (
        <Stack
          tokens={{
            padding: "m 20px"
          }}
        >
          <Spinner label="Please Waiting ..." />
        </Stack>
      ) : (
        <>
          <Stack
            tokens={{
              padding: "m 20px"
            }}
          >
            <Stack horizontal>
              <ChoiceGroup
                label="Please select a case"
                onChange={changeCaseType}
                defaultSelectedKey="base"
                options={options}
              />
            </Stack>
            {/* <Stack horizontal>
              {
                <ul>
                  {testValue &&
                    testValue.map((v, i) => {
                      return (
                        <li key={i}>
                          {v.row}
                          {v.values}
                        </li>
                      );
                    })}
                </ul>
              }
            </Stack> */}
            <Stack horizontal>{test1Value && <span>{test1Value}</span>}</Stack>
            {/*  <Stack horizontal>
              <MessageBar
                messageBarType={MessageBarType.warning}
                isMultiline={false}
                onDismiss={() => {}}
                dismissButtonAriaLabel="Close"
              >
                Warning MessageBar content.
              </MessageBar>
            </Stack> */}

            <Stack horizontal>
              {yearArr && yearArr.length && (
                <DetailsList
                  items={items}
                  columns={getColumns()}
                  setKey="set"
                  compact={true}
                  layoutMode={DetailsListLayoutMode.justified}
                  selectionPreservedOnEmptyClick={true}
                  selectionMode={SelectionMode.none}
                />
              )}
            </Stack>
            <Stack horizontal>
              <PrimaryButton>New Driver</PrimaryButton>
            </Stack>
          </Stack>
        </>
      )}
    </div>
  );
};

export default App;
