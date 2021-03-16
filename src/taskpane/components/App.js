import React, { useState, useEffect, useBoolean } from "react";
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
  ActionButton,
  SelectionMode,
  SpinButton,
  Stack,
  Label,
  SearchBox,
  ChoiceGroup,
  Spinner,
  SpinnerSize,
} from "office-ui-fabric-react";
// import { mergeStyles, mergeStyleSets } from "office-ui-fabric-react/lib/Styling";
/* global Button, console, Excel, Header, HeroList, HeroListItem, Progress */

import { getYearValues, drawTable, getPosition, linkTableData } from "../excel";

const years = [2020, 2021, 2022];

let expectMatrix = [];

let locationObjects = {};

const App = () => {
  const options = [
    { key: "base", text: "Base", iconProps: { iconName: "LineChart", styles: { color: "blue" } } },
    { key: "bull", text: "Bull", iconProps: { iconName: "Market" } },
    { key: "bear", text: "Bear", iconProps: { iconName: "MarketDown" } },
  ];
  const columns = [
    {
      key: "column1",
      name: "Driver",
      fieldName: "driver",
      minWidth: 100,
      maxWidth: 100,
      isResizable: true,
      // eslint-disable-next-line react/display-name
      onRender: (item) => {
        console.log(item);
        return (
          <SearchBox
            placeholder="Row No."
            underlined={true}
            onClear={(e) => handleDriverRowClear(e, item)}
            onSearch={(e) => handleDriverRowSearch(e, item)}
          />
        );
      },
    },
    {
      key: "column2",
      name: "2020",
      fieldName: "y2020",
      minWidth: 80,
      maxWidth: 80,
      isResizable: true,
      // eslint-disable-next-line react/display-name
      onRender: (item) => <SpinButton defaultValue={item.y2020} min={0} max={10} step={0.1} disabled={item.disable} />,
    },
    {
      key: "column3",
      name: "2021",
      fieldName: "y2021",
      minWidth: 80,
      maxWidth: 80,
      isResizable: true,
      // eslint-disable-next-line react/display-name
      onRender: (item) => <SpinButton defaultValue={item.y2021} min={0} max={10} step={0.1} disabled={item.disable} />,
    },
    {
      key: "column4",
      name: "2022",
      fieldName: "y2022",
      minWidth: 80,
      maxWidth: 80,
      isResizable: true,
      // eslint-disable-next-line react/display-name
      onRender: (item) => <SpinButton defaultValue={item.y2022} min={0} max={10} step={0.1} disabled={item.disable} />,
    },
  ].map((v, i) => {
    if (i >= 1) {
      return { ...v, name: years[i - 1] };
    } else {
      return v;
    }
  });

  // const [expectMatrix, setExpectMatrix] = useState([]);
  const [locate, setLocate] = useState(null);
  const [yearArr, setYearArr] = useState([]);

  const [customize1Value, setCustomize1Value] = useState("");
  const [customize2Value, setCustomize2Value] = useState("");
  const [isHideDialog, hideDialog] = useState(false);
  const [selfCustomize1Checked, setSelfCustomize1] = useState(false);
  const [selfCustomize2Checked, setSelfCustomize2] = useState(false);
  const [loading, setLoading] = useState(false);
  const [items, SetItems] = useState([
    {
      key: 1,
      driver: 1,
      y2020: 0,
      y2021: 0,
      y2022: 0,
      disable: true,
    },
    {
      key: 2,
      driver: 2,
      y2020: 0,
      y2021: 0,
      y2022: 0,
      disable: true,
    },
    {
      key: 3,
      driver: 3,
      y2020: 0,
      y2021: 0,
      y2022: 0,
      disable: true,
    },
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
    await getPosition(customizeArr, async (payload) => {
      setLocate(payload);

      for (let key in payload) {
        if (key != "year") {
          const elem = payload[key];
          const arr = [elem.title];
          const expectArr = elem.e.split(",");

          for (let i = 0, l = expectArr.length; i < l; i++) {
            arr.push(expectArr[i]);
          }
          expectMatrix.push(arr);
        }
      }
      // setExpectMatrix(matrix);

      locationObjects = { ...payload };
      await getYearValues(payload.year.e, async (yearValueArr) => {
        setYearArr(yearValueArr);
        await drawTable(yearValueArr);
        setLoading(false);
        await linkTableData(expectMatrix);
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

  const handleDriverRowSearch = (e, item) => {
    const arr = [...items];
    arr[item.key - 1].disable = false;
    SetItems(arr);
  };

  const handleDriverRowClear = (e, item) => {
    const arr = [...items];
    arr[item.key - 1].disable = true;
    SetItems(arr);
  };

  return (
    <div className="App">
      <Dialog
        hidden={isHideDialog}
        onDismiss={handleInit}
        dialogContentProps={{
          type: DialogType.largeHeader,
          title: "Initialization",
          subText: "Please select",
        }}
        modalProps={{
          isBlocking: false,
          styles: { main: { maxWidth: 450 } },
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
                onChange={(e) => {
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
                onChange={(e) => {
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
            padding: "m 20px",
          }}
        >
          <Spinner label="Please Waiting ..." />
        </Stack>
      ) : (
        <>
          <Stack
            tokens={{
              padding: "m 20px",
            }}
          >
            <Stack horizontal>
              <ChoiceGroup label="Please select a case" defaultSelectedKey="base" options={options} />
            </Stack>
            <Stack horizontal>
              {/* <ul>
                {expectMatrix &&
                  expectMatrix.map((v, i) => {
                    return v.map((vv, ii) => {
                      return <li key={i + ii}>{vv}</li>;
                    });
                  })}
              </ul> */}
            </Stack>
            <Stack horizontal>
              <DetailsList
                items={items}
                columns={columns}
                setKey="set"
                compact={true}
                layoutMode={DetailsListLayoutMode.justified}
                selectionPreservedOnEmptyClick={true}
                selectionMode={SelectionMode.none}
              />
            </Stack>
            <Stack horizontal>
              <ActionButton iconProps={{ iconName: "Add" }}>New Driver</ActionButton>
            </Stack>
          </Stack>
        </>
      )}
    </div>
  );
};

export default App;
