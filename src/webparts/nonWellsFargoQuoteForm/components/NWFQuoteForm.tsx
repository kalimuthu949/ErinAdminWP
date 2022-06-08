import * as React from "react";
import styles from "./NonWellsFargoQuoteForm.module.scss";
import { Fragment, useCallback } from "react";
import { useState, useEffect, useRef } from "react";
import { Icon, IIconProps } from "@fluentui/react/lib/Icon";
import {
  ActionButton,
  DefaultButton,
  PrimaryButton,
} from "@fluentui/react/lib/Button";
import { TextField, MaskedTextField } from "@fluentui/react/lib/TextField";
import {
  DatePicker,
  DayOfWeek,
  Dropdown,
  IDropdownOption,
  mergeStyles,
  defaultDatePickerStrings,
  Checkbox,
  ThemeProvider,
  Label,
  ChoiceGroup,
  IChoiceGroupOption,
  Fabric,
} from "@fluentui/react";
import { IStyleSet, ILabelStyles, Pivot, PivotItem } from "@fluentui/react";

import { loadTheme, createTheme, Theme } from "@fluentui/react";
import * as strings from "AdminDashboardWebPartStrings";

let formID = 0;
const paramsString = window.location.href.split("?")[1].toLowerCase();
const searchParams = new URLSearchParams(paramsString);
searchParams.has("formid") ? (formID = Number(searchParams.get("formid"))) : "";

const addIcon: IIconProps = { iconName: "Add" };

const choiceGroupStyles = {
  flexContainer: {
    display: "flex",
    label: {
      marginRight: "1rem",
    },
  },
};
let arrMilestones = [
  {
    id: 10000,
    title: "Milestone-1",
    description: "",
    startDate: new Date().toLocaleString(),
    endDate: new Date().toLocaleString(),
    amount: "",
  },
];
let arrnwParts = [
  {
    isSelected: false,
    PartNo: "",
    PartName: "",
    PartDescription: "",
    ListPrice: 0,
    itemFor: "",
    NetPrice: 0,
    Note: "",
    PartDescriptionSort: "",
    id: 0,
    quantity:0,
    MultiplierApply:false
  },
];

const myTheme = createTheme({
  palette: {
    themePrimary: "#004fa2",
    themeLighterAlt: "#f1f6fb",
    themeLighter: "#cadcf0",
    themeLight: "#9fc0e3",
    themeTertiary: "#508ac8",
    themeSecondary: "#155fae",
    themeDarkAlt: "#004793",
    themeDark: "#003c7c",
    themeDarker: "#002c5b",
    neutralLighterAlt: "#faf9f8",
    neutralLighter: "#f3f2f1",
    neutralLight: "#edebe9",
    neutralQuaternaryAlt: "#e1dfdd",
    neutralQuaternary: "#d0d0d0",
    neutralTertiaryAlt: "#c8c6c4",
    neutralTertiary: "#a19f9d",
    neutralSecondary: "#605e5c",
    neutralPrimaryAlt: "#3b3a39",
    neutralPrimary: "#323130",
    neutralDark: "#201f1e",
    black: "#000000",
    white: "#ffffff",
  },
});
let objValues = {
  ProjectNo: "",
  Date: new Date(),
  ConsultantName: "",
  ConsultantCity: "",
  ConsultantContactNo: "",
  ConsultantPinCode: "",
  ConsultantAddress: "",
  ClientName: "",
  ClientCity: "",
  ClientContactNo: "",
  ClientPinCode: "",
  ClientAddress: "",
  SentVai: "",
  ProjectDescription: "",
  TypesOfProposal: "",
  Multiplier: "1",
  ProposedBy: "",
  ProposedName: "",
  ProposedTitle: "",
  ProposedDate: new Date(),
  AcceptedBy: "",
  AcceptedByName: "",
  AcceptedByDate: new Date(),
  AcceptedByTitle: "",
  StatementOfWork: "",
  Services: "",
  AcceptedForClient: "",
};
let objSelectedServices = {
  JENEsysEDGE: [],
  ONYXX: [],
  ONYXXLX: [],
  Niagara4: [],
  HardwareAccessories: [],
  JENEsysEngineeringTools: [],
  JENEsysEnclosures: [],
  Renewals: [],
  JENEsysThermostatsPeripherals: [],
  BACnetControllers: [],
  Distech: [],
  Veris: [],
  Belimo: [],
  TemperatureRHCO2Sensors: [],
  PowerMeters: [],
  DifferentialPressureTransmittersSwitches: [],
  Relays: [],
  CurrentSensorsTransmitters: [],
  PowerSupplies: [],
  Transformers: [],
  LynxspringUniversity: [],
  TAPA: [],
  DGLux: [],
  SkyFoundry: [],
  TridiumAnalytics: [],
};
let arrPartsCount = {
  JENEsysEDGE: 0,
  ONYXX: 0,
  ONYXXLX: 0,
  Niagara4: 0,
  HardwareAccessories: 0,
  JENEsysEngineeringTools: 0,
  JENEsysEnclosures: 0,
  Renewals: 0,
  JENEsysThermostatsPeripherals: 0,
  BACnetControllers: 0,
  Distech: 0,
  Veris: 0,
  Belimo: 0,
  TemperatureRHCO2Sensors: 0,
  PowerMeters: 0,
  DifferentialPressureTransmittersSwitches: 0,
  Relays: 0,
  CurrentSensorsTransmitters: 0,
  PowerSupplies: 0,
  Transformers: 0,
  LynxspringUniversity: 0,
  TAPA: 0,
  DGLux: 0,
  SkyFoundry: 0,
  TridiumAnalytics: 0,
};

const halfWidthInput = {
  root: { width: 300, margin: "0 1rem 0.5rem 0" },
};
let arrSentViaOptions = [];
let arrTypesOfProposal = [];

const NWFQuoteForm = (props) => {
  loadTheme(myTheme);
  const [firstDayOfWeek, setFirstDayOfWeek] = useState(DayOfWeek.Sunday);
  const [selectedKey, setSelectedKey] = useState(1);
  const [milestones, setMilestones] = useState(arrMilestones);
  const [fetchTable, setFetchTable] = useState(true);
  const [partsDetails, setPartsDetails] = useState(arrnwParts);
  const [fetchPartsTable, setFetchPartsTable] = useState(true);
  const [objToPost, setObjToPost] = useState(objValues);
  const [renderObjValue, setRenderObjValue] = useState(true);
  const [selectedServices, setSelectedServices] = useState(objSelectedServices);
  const [fetchSelectedServices, setFetchSelectedServices] = useState(true);
  const [sentViaOptions, setSentViaOptions] = useState(arrSentViaOptions);
  const [typesOfProposalOptions, setTypesOfProposalOptions] =
    useState(arrTypesOfProposal);
  const [partsCount, setPartsCount] = useState(arrPartsCount);
  const [dateValue, setDateValue] = useState<Date | undefined>();
  const [isFullDescription, setIsFullDescription] = useState(false);

  useEffect(() => {
    props.spcontext.web.lists
      .getByTitle("GeneralQuoteRequestList")
      .fields.filter("EntityPropertyName eq 'SentVia'")
      .get()
      .then((SentVia) => {
        SentVia[0].Choices.forEach((option) => {
          arrSentViaOptions.push({
            key: option,
            text: option,
          });
        });
      });
    setSentViaOptions(arrSentViaOptions);
    props.spcontext.web.lists
      .getByTitle("GeneralQuoteRequestList")
      .fields.filter("EntityPropertyName eq 'TypesOfProposal'")
      .get()
      .then((types) => {
        types[0].Choices.forEach((option) => {
          arrTypesOfProposal.push({
            key: option,
            text: option,
          });
        });
        console.log(arrTypesOfProposal);

        setTypesOfProposalOptions(arrTypesOfProposal);
      });
    props.spcontext.web.lists
      .getByTitle("GeneralParts")
      .items.top(5000)
      .get()
      .then((partList: any) => {
        arrnwParts = partList.map((partItem, i) => {
          return {
            isSelected: false,
            PartNo: partItem.Title,
            PartName: partItem.PartName,
            PartDescription: partItem.PartDescription,
            PartDescriptionSort: partItem.PartDescription.slice(0, 100) + "...",
            ListPrice: partItem.ListPrice,
            itemFor: partItem.itemFor,
            NetPrice:partItem.MultiplierApply? +partItem.ListPrice*0:+partItem.ListPrice,
            Note: "",
            id: i,
            quantity:partItem.quantity?partItem.quantity:0,
            MultiplierApply:partItem.MultiplierApply
          };
        });
        console.log(arrnwParts);
        setPartsDetails(arrnwParts);
      })
      .then(async () => {
        props.spcontext.web.lists
          .getByTitle("GeneralQuoteRequestList")
          .items.select("*")
          .getById(formID)
          .get()
          .then((qrItem) => {
            console.log(qrItem);
            objValues.ProjectNo = qrItem.ProjectNo;
            setRenderObjValue(true);
          });
      });
  }, []);
  useEffect(() => {
    if (fetchTable) {
      setMilestones([...arrMilestones]);
      setFetchTable(false);
    }
  }, [fetchTable]);
  useEffect(() => {
    if (fetchPartsTable) {
      setPartsDetails([...arrnwParts]);
      setFetchPartsTable(false);
    }
  }, [fetchPartsTable]);
  useEffect(() => {
    if (fetchSelectedServices) {
      setSelectedServices(objSelectedServices);
      setFetchSelectedServices(false);
    }
  }, [fetchSelectedServices]);
  useEffect(() => {
    if (renderObjValue) {
      setObjToPost(objValues);
      setRenderObjValue(false);
    }
  }, [renderObjValue]);
  const addMilestoneHandler = () => {
    let newId =
      arrMilestones.length > 0
        ? arrMilestones[arrMilestones.length - 1].id + 1
        : 100000;
    let title =
      arrMilestones.length > 0
        ? `Milestone-${
            +arrMilestones[arrMilestones.length - 1].title.split("-")[1] + 1
          }`
        : "Milestone-1";
    arrMilestones.push({
      id: newId,
      title: title,
      description: "",
      startDate: new Date().toLocaleString(),
      endDate: new Date().toLocaleString(),
      amount: "",
    });
    setFetchTable(true);
  };
  const multiplierHander = (multiplier) => {
    arrnwParts = arrnwParts.map((arrPart) => {
      if (multiplier > 0) {
        return {
          isSelected: arrPart.isSelected,
          PartNo: arrPart.PartNo,
          PartName: arrPart.PartName,
          PartDescription: arrPart.PartDescription,
          PartDescriptionSort: arrPart.PartDescription.slice(0, 100) + "...",
          ListPrice: arrPart.ListPrice,
          itemFor: arrPart.itemFor,
          NetPrice: arrPart.MultiplierApply?+arrPart.ListPrice * multiplier*+arrPart.quantity:+arrPart.ListPrice * +arrPart.quantity,
          Note: "",
          id: arrPart.id,
          quantity:arrPart.quantity?arrPart.quantity:0,
          MultiplierApply:arrPart.MultiplierApply
        };
      } else {
        return {
          isSelected: arrPart.isSelected,
          PartNo: arrPart.PartNo,
          PartName: arrPart.PartName,
          PartDescription: arrPart.PartDescription,
          PartDescriptionSort: arrPart.PartDescription.slice(0, 100) + "...",
          ListPrice: arrPart.ListPrice,
          itemFor: arrPart.itemFor,
          NetPrice: arrPart.MultiplierApply?+arrPart.ListPrice * multiplier*+arrPart.quantity:+arrPart.ListPrice * +arrPart.quantity,
          Note: "",
          id: arrPart.id,
          quantity:arrPart.quantity?arrPart.quantity:0,
          MultiplierApply:arrPart.MultiplierApply
        };
      }
    });
  };
  const tableCheckboxHandler = (id, part) => {
    part.isSelected
      ? (arrnwParts.filter((arrpart) => arrpart.id == id)[0].isSelected = false)
      : (arrnwParts.filter((arrpart) => arrpart.id == id)[0].isSelected = true);
    console.log(arrnwParts.filter((arrpart) => arrpart.id == id)[0]);
    setFetchTable(true);
    console.log(part.itemFor);
    console.log(part.isSelected);
    part.isSelected
      ? arrPartsCount[part.itemFor]++
      : arrPartsCount[part.itemFor]--;
    setPartsCount(arrPartsCount);
    if (part.itemFor == "JENEsysEDGE") {
      objSelectedServices.JENEsysEDGE = [];
      objSelectedServices.JENEsysEDGE = arrnwParts
        .filter(
          (arrpart) =>
            arrpart.isSelected == true && arrpart.itemFor == "JENEsysEDGE"
        )
        .map((part) => part.PartNo);

      setFetchSelectedServices(true);
    } else if (part.itemFor == "ONYXX") {
      objSelectedServices.ONYXX = [];
      objSelectedServices.ONYXX = arrnwParts
        .filter(
          (arrpart) => arrpart.isSelected == true && arrpart.itemFor == "ONYXX"
        )
        .map((part) => part.PartNo);
      setFetchSelectedServices(true);
    } else if (part.itemFor == "ONYXXLX") {
      objSelectedServices.ONYXXLX = [];
      objSelectedServices.ONYXXLX = arrnwParts
        .filter(
          (arrpart) =>
            arrpart.isSelected == true && arrpart.itemFor == "ONYXXLX"
        )
        .map((part) => part.PartNo);
      setFetchSelectedServices(true);
    } else if (part.itemFor == "Niagara4") {
      objSelectedServices.Niagara4 = [];
      objSelectedServices.Niagara4 = arrnwParts
        .filter(
          (arrpart) =>
            arrpart.isSelected == true && arrpart.itemFor == "Niagara4"
        )
        .map((part) => part.PartNo);

      setFetchSelectedServices(true);
    } else if (part.itemFor == "HardwareAccessories") {
      objSelectedServices.HardwareAccessories = [];
      objSelectedServices.HardwareAccessories = arrnwParts
        .filter(
          (arrpart) =>
            arrpart.isSelected == true &&
            arrpart.itemFor == "HardwareAccessories"
        )
        .map((part) => part.PartNo);
      setFetchSelectedServices(true);
    } else if (part.itemFor == "JENEsysEngineeringTools") {
      objSelectedServices.JENEsysEngineeringTools = [];
      objSelectedServices.JENEsysEngineeringTools = arrnwParts
        .filter(
          (arrpart) =>
            arrpart.isSelected == true &&
            arrpart.itemFor == "JENEsysEngineeringTools"
        )
        .map((part) => part.PartNo);
      setFetchSelectedServices(true);
    } else if (part.itemFor == "JENEsysEnclosures") {
      objSelectedServices.JENEsysEnclosures = [];
      objSelectedServices.JENEsysEnclosures = arrnwParts
        .filter(
          (arrpart) =>
            arrpart.isSelected == true && arrpart.itemFor == "JENEsysEnclosures"
        )
        .map((part) => part.PartNo);
      setFetchSelectedServices(true);
    } else if (part.itemFor == "Renewals") {
      objSelectedServices.Renewals = [];
      objSelectedServices.Renewals = arrnwParts
        .filter(
          (arrpart) =>
            arrpart.isSelected == true && arrpart.itemFor == "Renewals"
        )
        .map((part) => part.PartNo);
      setFetchSelectedServices(true);
    } else if (part.itemFor == "JENEsysThermostatsPeripherals") {
      objSelectedServices.JENEsysThermostatsPeripherals = [];
      objSelectedServices.JENEsysThermostatsPeripherals = arrnwParts
        .filter(
          (arrpart) =>
            arrpart.isSelected == true &&
            arrpart.itemFor == "JENEsysThermostatsPeripherals"
        )
        .map((part) => part.PartNo);
      setFetchSelectedServices(true);
    } else if (part.itemFor == "BACnetControllers") {
      objSelectedServices.BACnetControllers = [];
      objSelectedServices.BACnetControllers = arrnwParts
        .filter(
          (arrpart) =>
            arrpart.isSelected == true && arrpart.itemFor == "BACnetControllers"
        )
        .map((part) => part.PartNo);
      setFetchSelectedServices(true);
    } else if (part.itemFor == "Distech") {
      objSelectedServices.Distech = [];
      objSelectedServices.Distech = arrnwParts
        .filter(
          (arrpart) =>
            arrpart.isSelected == true && arrpart.itemFor == "Distech"
        )
        .map((part) => part.PartNo);
      setFetchSelectedServices(true);
    } else if (part.itemFor == "Veris") {
      objSelectedServices.Veris = [];
      objSelectedServices.Veris = arrnwParts
        .filter(
          (arrpart) => arrpart.isSelected == true && arrpart.itemFor == "Veris"
        )
        .map((part) => part.PartNo);

      setFetchSelectedServices(true);
    } else if (part.itemFor == "Belimo") {
      objSelectedServices.Belimo = [];
      objSelectedServices.Belimo = arrnwParts
        .filter(
          (arrpart) => arrpart.isSelected == true && arrpart.itemFor == "Belimo"
        )
        .map((part) => part.PartNo);
      setFetchSelectedServices(true);
    } else if (part.itemFor == "TemperatureRHCO2Sensors") {
      objSelectedServices.TemperatureRHCO2Sensors = [];
      objSelectedServices.TemperatureRHCO2Sensors = arrnwParts
        .filter(
          (arrpart) =>
            arrpart.isSelected == true &&
            arrpart.itemFor == "TemperatureRHCO2Sensors"
        )
        .map((part) => part.PartNo);
      setFetchSelectedServices(true);
    } else if (part.itemFor == "PowerMeters") {
      objSelectedServices.PowerMeters.push(part.PartNo);
      setFetchSelectedServices(true);
    } else if (part.itemFor == "DifferentialPressureTransmittersSwitches") {
      objSelectedServices.DifferentialPressureTransmittersSwitches.push(
        part.PartNo
      );
      setFetchSelectedServices(true);
    } else if (part.itemFor == "Relays") {
      objSelectedServices.Relays.push(part.PartNo);
      setFetchSelectedServices(true);
    } else if (part.itemFor == "CurrentSensorsTransmitters") {
      objSelectedServices.CurrentSensorsTransmitters.push(part.PartNo);
      setFetchSelectedServices(true);
    } else if (part.itemFor == "PowerSupplies") {
      objSelectedServices.PowerSupplies.push(part.PartNo);
      setFetchSelectedServices(true);
    } else if (part.itemFor == "Transformers") {
      objSelectedServices.Transformers.push(part.PartNo);
      setFetchSelectedServices(true);
    } else if (part.itemFor == "LynxspringUniversity") {
      objSelectedServices.LynxspringUniversity.push(part.PartNo);
      setFetchSelectedServices(true);
    } else if (part.itemFor == "TAPA") {
      objSelectedServices.TAPA.push(part.PartNo);
      setFetchSelectedServices(true);
    } else if (part.itemFor == "DGLux") {
      objSelectedServices.DGLux.push(part.PartNo);
      setFetchSelectedServices(true);
    } else if (part.itemFor == "SkyFoundry") {
      objSelectedServices.SkyFoundry.push(part.PartNo);
      setFetchSelectedServices(true);
    } else if (part.itemFor == "TridiumAnalytics") {
      objSelectedServices.TridiumAnalytics.push(part.PartNo);
      setFetchSelectedServices(true);
    }
  };
  const listPriceChangeHandler = (id, value) => {
    console.log(id, value);
    let updateItem = arrnwParts.filter((part) => part.id == id)[0];
    updateItem.ListPrice = value;
    +objValues.Multiplier>0 && updateItem.MultiplierApply&&updateItem.quantity!=null&&updateItem.quantity!=undefined&&!isNaN(updateItem.quantity)?(updateItem.NetPrice = updateItem.quantity * value * +objValues.Multiplier)
      : updateItem.quantity!=null&&updateItem.quantity!=undefined&&!isNaN(updateItem.quantity)?(updateItem.NetPrice = value*updateItem.quantity):updateItem.NetPrice=0;
    setFetchTable(true);
  };

  const quantityChangeHandler = (id, value) => {
    console.log(id, value);
    let updateItem = arrnwParts.filter((part) => part.id == id)[0];
    updateItem.quantity = value;
    +objValues.Multiplier>0 && updateItem.MultiplierApply&&updateItem.quantity!=null&&updateItem.quantity!=undefined&&!isNaN(updateItem.quantity)?(updateItem.NetPrice = value * updateItem.ListPrice * +objValues.Multiplier)
      : updateItem.quantity!=null&&updateItem.quantity!=undefined&&!isNaN(updateItem.quantity)?(updateItem.NetPrice = value*updateItem.ListPrice):updateItem.NetPrice=0;
    setFetchTable(true);
  };

  const submitBtnHandler = () => {
    props.spcontext.web.lists
      .getByTitle("GeneralQuoteRequestList")
      .items.getById(formID)
      .update({
        AcceptedBy: objValues.AcceptedBy,
        AcceptedByDate: objValues.AcceptedByDate,
        AcceptedByName: objValues.AcceptedByName,
        AcceptedByTitle: objValues.AcceptedByTitle,
        ClientAddress: objValues.ClientAddress,
        ClientCity: objValues.ClientCity,
        ClientContactNo: objValues.ClientContactNo,
        ClientName: objValues.ClientName,
        ClientPinCode: objValues.ClientPinCode,
        ConsultantAddress: objValues.ConsultantAddress,
        ConsultantCity: objValues.ConsultantCity,
        ConsultantContactNo: objValues.ConsultantContactNo,
        ConsultantName: objValues.ConsultantName,
        ConsultantPinCode: objValues.ConsultantPinCode,
        Date: objValues.Date,
        Milestones:
          arrMilestones.length == 1
            ? arrMilestones[0].description == "" &&
              arrMilestones[0].amount == ""
              ? ""
              : JSON.stringify(arrMilestones)
            : arrMilestones.length > 0
            ? JSON.stringify(arrMilestones)
            : "",
        Multiplier: objValues.Multiplier,
        ProjectDescription: objValues.ProjectDescription,
        ProposedBy: objValues.ProposedBy,
        ProposedDate: objValues.ProposedDate,
        ProposedName: objValues.ProposedName,
        ProposedTitle: objValues.ProposedTitle,
        Services: objValues.Services,
        StatementOfWork: objValues.StatementOfWork,
        TypesOfProposal: objValues.TypesOfProposal,
        SentVia: objValues.SentVai,
        AcceptedForClient: objValues.AcceptedForClient,
        Status: "Order in production",
        ProposedServicesFees: JSON.stringify(
          arrnwParts.filter((part) => part.isSelected)
        ),
      })
      .then((data) => {
        history.back();
      })
      .catch((error) => console.log(error));
  };
  const onNextClick = () => {
    selectedKey < 26 ? setSelectedKey(selectedKey + 1) : "";
  };
  const onPrevClick = () => {
    selectedKey > 1 ? setSelectedKey(selectedKey - 1) : "";
  };

  return (
    <div style={{ backgroundColor: "#F2F2F2", padding: "1rem 2rem" }}>
      <div className={styles.formHeader}>
        <Icon
          iconName="NavigateBack"
          styles={{
            root: {
              fontSize: 30,
              fontWeight: 600,
              color: myTheme.palette.themePrimary,
              marginRight: "1rem",
              cursor: "pointer",
            },
          }}
          onClick={() => {
            history.back();
          }}
        />
        <h2
          style={{
            textAlign: "center",
            color: myTheme.palette.themePrimary,
            width: "100%",
          }}
        >
          Proposal of Services
        </h2>
      </div>
      <div className={`${styles.projectDetails} ${styles.section}`}>
        <TextField
          label="Project No"
          styles={halfWidthInput}
          value={objValues.ProjectNo}
          disabled={true}
        />
        <DatePicker
          styles={halfWidthInput}
          label="Date"
          firstDayOfWeek={firstDayOfWeek}
          placeholder="Select a date..."
          ariaLabel="Select a date"
          formatDate={(date: Date): string => {
            return (
              date.getMonth() +
              1 +
              "/" +
              date.getDate() +
              "/" +
              date.getFullYear()
            );
          }}
          onSelectDate={(date) => {
            setDateValue as (date?: Date) => void;
            objValues.Date = date;
          }}
        />
      </div>
      {/* Section */}
      <div className={styles.section}>
        {/* Consultant Section */}
        <h3 style={{ color: myTheme.palette.themePrimary }}>
          Form (Consultant)
        </h3>
        <div className={styles.consultantClient}>
          <div>
            <TextField
              label="Name"
              styles={halfWidthInput}
              value={objValues.ConsultantName}
              onChange={(e) => {
                objValues.ConsultantName = e.target["value"];
                setRenderObjValue(true);
              }}
            />
            <TextField
              label="Contact No"
              styles={halfWidthInput}
              value={objValues.ConsultantContactNo}
              onChange={(e) => {
                objValues.ConsultantContactNo = e.target["value"];
                setRenderObjValue(true);
              }}
            />
          </div>
          <div>
            <TextField
              label="City"
              styles={halfWidthInput}
              value={objValues.ConsultantCity}
              onChange={(e) => {
                objValues.ConsultantCity = e.target["value"];
                setRenderObjValue(true);
              }}
            />
            <TextField
              label="Pin Code"
              styles={halfWidthInput}
              value={objValues.ConsultantPinCode}
              onChange={(e) => {
                objValues.ConsultantPinCode = e.target["value"];
                setRenderObjValue(true);
              }}
            />
          </div>
          <div>
            <TextField
              styles={halfWidthInput}
              label="Address"
              multiline
              resizable={false}
              value={objValues.ConsultantAddress}
              onChange={(e) => {
                objValues.ConsultantAddress = e.target["value"];
                setRenderObjValue(true);
              }}
            />
          </div>
        </div>
        {/* Client Section */}
        <h3 style={{ color: myTheme.palette.themePrimary }}>To (Client)</h3>
        <div className={styles.consultantClient}>
          <div>
            <TextField
              label="Name"
              styles={halfWidthInput}
              value={objValues.ClientName}
              onChange={(e) => {
                objValues.ClientName = e.target["value"];
                setRenderObjValue(true);
              }}
            />
            <TextField
              label="Contact No"
              styles={halfWidthInput}
              value={objValues.ClientContactNo}
              onChange={(e) => {
                objValues.ClientContactNo = e.target["value"];
                setRenderObjValue(true);
              }}
            />
          </div>
          <div>
            <TextField
              label="City"
              styles={halfWidthInput}
              value={objValues.ClientCity}
              onChange={(e) => {
                objValues.ClientCity = e.target["value"];
                setRenderObjValue(true);
              }}
            />
            <TextField
              label="Pin Code"
              styles={halfWidthInput}
              value={objValues.ClientPinCode}
              onChange={(e) => {
                objValues.ClientPinCode = e.target["value"];
                setRenderObjValue(true);
              }}
            />
          </div>
          <div>
            <TextField
              styles={halfWidthInput}
              label="Address"
              multiline
              resizable={false}
              value={objValues.ClientAddress}
              onChange={(e) => {
                objValues.ClientAddress = e.target["value"];
                setRenderObjValue(true);
              }}
            />
          </div>
        </div>
      </div>
      {/* Section */}
      {/* Section */}
      <div className={styles.section}>
        <div>
          <ChoiceGroup
            options={sentViaOptions}
            styles={choiceGroupStyles}
            label="Sent Via:"
            onChange={(e, selected) => (objValues.SentVai = selected.key)}
          />
        </div>
        <div style={{ display: "flex", justifyContent: "space-between" }}>
          <TextField
            styles={halfWidthInput}
            label="Services"
            multiline
            resizable={false}
            value={objValues.Services}
            onChange={(e) => {
              objValues.Services = e.target["value"];
              setRenderObjValue(true);
            }}
          />
          <TextField
            styles={halfWidthInput}
            label="Project Description"
            multiline
            resizable={false}
            value={objValues.ProjectDescription}
            onChange={(e) => {
              objValues.ProjectDescription = e.target["value"];
              setRenderObjValue(true);
            }}
          />
        </div>
        <div>
          <ChoiceGroup
            options={typesOfProposalOptions}
            label="Types of proposal"
            styles={choiceGroupStyles}
            onChange={(e, selected) => {
              console.log(selected.key);

              objValues.TypesOfProposal = selected.key;
            }}
          />
        </div>
      </div>
      {/* Section */}
      <div className={`${styles.section} ${styles.sectionPovit}`}>
        {/*  Pivot Section Start */}
        <div
          style={{
            display: "flex",
            justifyContent: "space-between",
            alignItems: "center",
          }}
        >
          <h3 style={{ color: myTheme.palette.themePrimary }}>
            Proposed Services and Fee
          </h3>
          <TextField
            className="yourMultiplier"
            type="number"
            value={objValues.Multiplier}
            onChange={(e) => {
              multiplierHander(e.target["value"]);
              objValues.Multiplier = e.target["value"];
              setRenderObjValue(true);
              setFetchPartsTable(true);
            }}
            styles={{
              root: {
                width: 200,
              },
            }}
            label="Your Multiplier"
          />
        </div>
        {/* Pivot */}
        <ThemeProvider dir="ltr">
          <Pivot
            aria-label="Pivot Overflow Menu Example"
            linkFormat="links"
            overflowBehavior="menu"
            overflowAriaLabel="more items"
            selectedKey={String(selectedKey)}
            onLinkClick={(e) => {
              setSelectedKey(+e.props.itemKey);
            }}
          >
            <PivotItem
              headerText="JENEsysEDGE™"
              itemCount={partsCount.JENEsysEDGE}
              itemKey="1"
            >
              <div>
                {selectedServices.JENEsysEDGE.map((service) => {
                  return <span className={styles.selectedItem}>{service}</span>;
                })}
              </div>
              <table>
                <thead>
                  <tr>
                    <th></th>
                    <th>Part Name</th>
                    <th>
                      <div
                        style={{
                          display: "flex",
                          justifyContent: "space-between",
                          alignItems: "center",
                        }}
                      >
                        <div>Part Description</div>
                        <Icon
                          iconName={
                            isFullDescription
                              ? "ChevronUpSmall"
                              : "ChevronDownSmall"
                          }
                          styles={{
                            root: {
                              fontSize: 20,
                              fontWeight: 400,
                              color: "#004fa2",
                              cursor: "pointer",
                            },
                          }}
                          onClick={() => {
                            isFullDescription
                              ? setIsFullDescription(false)
                              : setIsFullDescription(true);
                          }}
                        />
                      </div>
                    </th>
                    <th>List Price</th>
                    <th>Quantitiy</th>
                    <th>Net Price</th>
                    <th>Note</th>
                  </tr>
                </thead>
                <tbody>
                  {partsDetails.map((part) => {
                    if (part.itemFor == "JENEsysEDGE") {
                      return (
                        <tr
                          style={{
                            backgroundColor: part.isSelected
                              ? "#eef4fa"
                              : "#ffffff",
                          }}
                        >
                          <td>
                            {" "}
                            <Checkbox
                              checked={part.isSelected ? true : false}
                              onChange={() => {
                                tableCheckboxHandler(part.id, part);
                              }}
                            />
                          </td>
                          <td>
                            <div style={{ fontWeight: "bold" }}>
                              {part.PartNo}
                            </div>
                            <div>{part.PartName}</div>
                          </td>
                          <td>
                            <label title={part.PartDescription}>
                              {isFullDescription
                                ? part.PartDescription
                                : part.PartDescriptionSort}
                            </label>
                          </td>
                          <td style={{ textAlign: "center" }}>
                            <TextField
                              resizable={false}
                              value={`${part.ListPrice}`}
                              onChange={(e) => {
                                let updateID = part.id;
                                let updateValue = e.target["value"];
                                listPriceChangeHandler(updateID, +updateValue);
                              }}
                            />
                          </td>
                          <td style={{ textAlign: "center" }}>
                            <TextField
                              resizable={false}
                              value={`${part.quantity}`}
                              onChange={(e) => {
                                let updateID = part.id;
                                let updateValue = e.target["value"];
                                quantityChangeHandler(updateID, +updateValue);
                              }}
                            />
                          </td>
                          <td style={{ textAlign: "center" }}>
                            {part.NetPrice}
                          </td>
                       
                          <td>{part.Note}</td>
                        </tr>
                      );
                    }
                  })}
                </tbody>
              </table>
            </PivotItem>
            <PivotItem
              headerText="ONYXX"
              itemCount={partsCount.ONYXX}
              itemKey="2"
            >
              <div>
                {selectedServices.ONYXX.map((service) => {
                  return <span className={styles.selectedItem}>{service}</span>;
                })}
              </div>
              <table>
                <thead>
                  <tr>
                    <th></th>
                    <th>Part Name</th>
                    <th>
                      <div
                        style={{
                          display: "flex",
                          justifyContent: "space-between",
                          alignItems: "center",
                        }}
                      >
                        <div>Part Description</div>
                        <Icon
                          iconName={
                            isFullDescription
                              ? "ChevronUpSmall"
                              : "ChevronDownSmall"
                          }
                          styles={{
                            root: {
                              fontSize: 20,
                              fontWeight: 400,
                              color: "#004fa2",
                              cursor: "pointer",
                            },
                          }}
                          onClick={() => {
                            isFullDescription
                              ? setIsFullDescription(false)
                              : setIsFullDescription(true);
                          }}
                        />
                      </div>
                    </th>
                    <th>List Price</th>
                    <th>Quantity</th>
                    <th>Net Price</th>
                    <th>Note</th>
                  </tr>
                </thead>
                <tbody>
                  {partsDetails.map((part) => {
                    if (part.itemFor == "ONYXX") {
                      return (
                        <tr
                          style={{
                            backgroundColor: part.isSelected
                              ? "#eef4fa"
                              : "#ffffff",
                          }}
                        >
                          <td>
                            {" "}
                            <Checkbox
                              checked={part.isSelected ? true : false}
                              onChange={() => {
                                tableCheckboxHandler(part.id, part);
                              }}
                            />
                          </td>
                          <td>
                            <div style={{ fontWeight: "bold" }}>
                              {part.PartNo}
                            </div>
                            <div>{part.PartName}</div>
                          </td>
                          <td>
                            <label title={part.PartDescription}>
                              {isFullDescription
                                ? part.PartDescription
                                : part.PartDescriptionSort}
                            </label>
                          </td>
                          <td style={{ textAlign: "center" }}>
                            <TextField
                              resizable={false}
                              value={`${part.ListPrice}`}
                              onChange={(e) => {
                                let updateID = part.id;
                                let updateValue = e.target["value"];
                                listPriceChangeHandler(updateID, +updateValue);
                              }}
                            />
                          </td>
                          
                         
                          <td style={{ textAlign: "center" }}>
                            <TextField
                              resizable={false}
                              value={`${part.quantity}`}
                              onChange={(e) => {
                                let updateID = part.id;
                                let updateValue = e.target["value"];
                                quantityChangeHandler(updateID, +updateValue);
                              }}
                            />
                          </td>
                          <td style={{ textAlign: "center" }}>
                            {part.NetPrice}
                          </td>
                          <td>{part.Note}</td>
                        </tr>
                      );
                    }
                  })}
                </tbody>
              </table>
            </PivotItem>
            <PivotItem
              headerText="ONYXX LX"
              itemCount={partsCount.ONYXXLX}
              itemKey="3"
            >
              <div>
                {selectedServices.ONYXXLX.map((service) => {
                  return <span className={styles.selectedItem}>{service}</span>;
                })}
              </div>
              <table>
                <thead>
                  <tr>
                    <th></th>
                    <th>Part Name</th>
                    <th>
                      <div
                        style={{
                          display: "flex",
                          justifyContent: "space-between",
                          alignItems: "center",
                        }}
                      >
                        <div>Part Description</div>
                        <Icon
                          iconName={
                            isFullDescription
                              ? "ChevronUpSmall"
                              : "ChevronDownSmall"
                          }
                          styles={{
                            root: {
                              fontSize: 20,
                              fontWeight: 400,
                              color: "#004fa2",
                              cursor: "pointer",
                            },
                          }}
                          onClick={() => {
                            isFullDescription
                              ? setIsFullDescription(false)
                              : setIsFullDescription(true);
                          }}
                        />
                      </div>
                    </th>
                    <th>List Price</th>
                    <th>Quantity</th>
                    <th>Net Price</th>
                    <th>Note</th>
                  </tr>
                </thead>
                <tbody>
                  {partsDetails.map((part) => {
                    if (part.itemFor == "ONYXXLX") {
                      return (
                        <tr
                          style={{
                            backgroundColor: part.isSelected
                              ? "#eef4fa"
                              : "#ffffff",
                          }}
                        >
                          <td>
                            {" "}
                            <Checkbox
                              checked={part.isSelected ? true : false}
                              onChange={() => {
                                tableCheckboxHandler(part.id, part);
                              }}
                            />
                          </td>
                          <td>
                            <div style={{ fontWeight: "bold" }}>
                              {part.PartNo}
                            </div>
                            <div>{part.PartName}</div>
                          </td>
                          <td>
                            <label title={part.PartDescription}>
                              {isFullDescription
                                ? part.PartDescription
                                : part.PartDescriptionSort}
                            </label>
                          </td>
                          <td style={{ textAlign: "center" }}>
                            <TextField
                              resizable={false}
                              value={`${part.ListPrice}`}
                              onChange={(e) => {
                                let updateID = part.id;
                                let updateValue = e.target["value"];
                                listPriceChangeHandler(updateID, +updateValue);
                              }}
                            />{" "}
                          </td>
                       
                          <td style={{ textAlign: "center" }}>
                            <TextField
                              resizable={false}
                              value={`${part.quantity}`}
                              onChange={(e) => {
                                let updateID = part.id;
                                let updateValue = e.target["value"];
                                quantityChangeHandler(updateID, +updateValue);
                              }}
                            />
                          </td>
                          <td style={{ textAlign: "center" }}>
                            {part.NetPrice}
                          </td>
                          <td>{part.Note}</td>
                        </tr>
                      );
                    }
                  })}
                </tbody>
              </table>
            </PivotItem>
            <PivotItem
              headerText="Niagara 4"
              itemKey="4"
              itemCount={partsCount.Niagara4}
              // onRenderItemLink={_customRenderer}
            >
              <div>
                {selectedServices.Niagara4.map((service) => {
                  return <span className={styles.selectedItem}>{service}</span>;
                })}
              </div>
              <table>
                <thead>
                  <tr>
                    <th></th>
                    <th>Part Name</th>
                    <th>
                      <div
                        style={{
                          display: "flex",
                          justifyContent: "space-between",
                          alignItems: "center",
                        }}
                      >
                        <div>Part Description</div>
                        <Icon
                          iconName={
                            isFullDescription
                              ? "ChevronUpSmall"
                              : "ChevronDownSmall"
                          }
                          styles={{
                            root: {
                              fontSize: 20,
                              fontWeight: 400,
                              color: "#004fa2",
                              cursor: "pointer",
                            },
                          }}
                          onClick={() => {
                            isFullDescription
                              ? setIsFullDescription(false)
                              : setIsFullDescription(true);
                          }}
                        />
                      </div>
                    </th>
                    <th>List Price</th>
                    <th>Quantity</th>
                    <th>Net Price</th>
                    <th>Note</th>
                  </tr>
                </thead>
                <tbody>
                  {partsDetails.map((part) => {
                    if (part.itemFor == "Niagara4") {
                      return (
                        <tr
                          style={{
                            backgroundColor: part.isSelected
                              ? "#eef4fa"
                              : "#ffffff",
                          }}
                        >
                          <td>
                            {" "}
                            <Checkbox
                              checked={part.isSelected ? true : false}
                              onChange={() => {
                                tableCheckboxHandler(part.id, part);
                              }}
                            />
                          </td>
                          <td>
                            <div style={{ fontWeight: "bold" }}>
                              {part.PartNo}
                            </div>
                            <div>{part.PartName}</div>
                          </td>
                          <td>
                            <label title={part.PartDescription}>
                              {isFullDescription
                                ? part.PartDescription
                                : part.PartDescriptionSort}
                            </label>
                          </td>
                          <td style={{ textAlign: "center" }}>
                            <TextField
                              resizable={false}
                              value={`${part.ListPrice}`}
                              onChange={(e) => {
                                let updateID = part.id;
                                let updateValue = e.target["value"];
                                listPriceChangeHandler(updateID, +updateValue);
                              }}
                            />{" "}
                          </td>
                         
                          <td style={{ textAlign: "center" }}>
                            <TextField
                              resizable={false}
                              value={`${part.quantity}`}
                              onChange={(e) => {
                                let updateID = part.id;
                                let updateValue = e.target["value"];
                                quantityChangeHandler(updateID, +updateValue);
                              }}
                            />
                          </td>
                          <td style={{ textAlign: "center" }}>
                            {part.NetPrice}
                          </td>
                          <td>{part.Note}</td>
                        </tr>
                      );
                    }
                  })}
                </tbody>
              </table>
            </PivotItem>
            <PivotItem
              headerText="Hardware Accessories"
              itemKey="5"
              itemCount={partsCount.HardwareAccessories}
            >
              <div>
                {selectedServices.HardwareAccessories.map((service) => {
                  return <span className={styles.selectedItem}>{service}</span>;
                })}
              </div>
              <table>
                <thead>
                  <tr>
                    <th></th>
                    <th>Part Name</th>
                    <th>
                      <div
                        style={{
                          display: "flex",
                          justifyContent: "space-between",
                          alignItems: "center",
                        }}
                      >
                        <div>Part Description</div>
                        <Icon
                          iconName={
                            isFullDescription
                              ? "ChevronUpSmall"
                              : "ChevronDownSmall"
                          }
                          styles={{
                            root: {
                              fontSize: 20,
                              fontWeight: 400,
                              color: "#004fa2",
                              cursor: "pointer",
                            },
                          }}
                          onClick={() => {
                            isFullDescription
                              ? setIsFullDescription(false)
                              : setIsFullDescription(true);
                          }}
                        />
                      </div>
                    </th>
                    <th>List Price</th>
                    <th>Quantity</th>
                    <th>Net Price</th>
                    <th>Note</th>
                  </tr>
                </thead>
                <tbody>
                  {partsDetails.map((part) => {
                    if (part.itemFor == "HardwareAccessories") {
                      return (
                        <tr
                          style={{
                            backgroundColor: part.isSelected
                              ? "#eef4fa"
                              : "#ffffff",
                          }}
                        >
                          <td>
                            {" "}
                            <Checkbox
                              checked={part.isSelected ? true : false}
                              onChange={() => {
                                tableCheckboxHandler(part.id, part);
                              }}
                            />
                          </td>
                          <td>
                            <div style={{ fontWeight: "bold" }}>
                              {part.PartNo}
                            </div>
                            <div>{part.PartName}</div>
                          </td>
                          <td>
                            <label title={part.PartDescription}>
                              {isFullDescription
                                ? part.PartDescription
                                : part.PartDescriptionSort}
                            </label>
                          </td>
                          <td style={{ textAlign: "center" }}>
                            <TextField
                              resizable={false}
                              value={`${part.ListPrice}`}
                              onChange={(e) => {
                                let updateID = part.id;
                                let updateValue = e.target["value"];
                                listPriceChangeHandler(updateID, +updateValue);
                              }}
                            />{" "}
                          </td>
                        
                          <td style={{ textAlign: "center" }}>
                            <TextField
                              resizable={false}
                              value={`${part.quantity}`}
                              onChange={(e) => {
                                let updateID = part.id;
                                let updateValue = e.target["value"];
                                quantityChangeHandler(updateID, +updateValue);
                              }}
                            />
                          </td>
                          <td style={{ textAlign: "center" }}>
                            {part.NetPrice}
                          </td>
                          <td>{part.Note}</td>
                        </tr>
                      );
                    }
                  })}
                </tbody>
              </table>
            </PivotItem>
            <PivotItem
              headerText="JENEsys Engineering Tools"
              itemKey="6"
              itemCount={partsCount.JENEsysEngineeringTools}
            >
              <div>
                {selectedServices.JENEsysEngineeringTools.map((service) => {
                  return <span className={styles.selectedItem}>{service}</span>;
                })}
              </div>
              <table>
                <thead>
                  <tr>
                    <th></th>
                    <th>Part Name</th>
                    <th>
                      <div
                        style={{
                          display: "flex",
                          justifyContent: "space-between",
                          alignItems: "center",
                        }}
                      >
                        <div>Part Description</div>
                        <Icon
                          iconName={
                            isFullDescription
                              ? "ChevronUpSmall"
                              : "ChevronDownSmall"
                          }
                          styles={{
                            root: {
                              fontSize: 20,
                              fontWeight: 400,
                              color: "#004fa2",
                              cursor: "pointer",
                            },
                          }}
                          onClick={() => {
                            isFullDescription
                              ? setIsFullDescription(false)
                              : setIsFullDescription(true);
                          }}
                        />
                      </div>
                    </th>
                    <th>List Price</th>
                    <th>Quantity</th>
                    <th>Net Price</th>
                    <th>Note</th>
                  </tr>
                </thead>
                <tbody>
                  {partsDetails.map((part) => {
                    if (part.itemFor == "JENEsysEngineeringTools") {
                      return (
                        <tr
                          style={{
                            backgroundColor: part.isSelected
                              ? "#eef4fa"
                              : "#ffffff",
                          }}
                        >
                          <td>
                            {" "}
                            <Checkbox
                              checked={part.isSelected ? true : false}
                              onChange={() => {
                                tableCheckboxHandler(part.id, part);
                              }}
                            />
                          </td>
                          <td>
                            <div style={{ fontWeight: "bold" }}>
                              {part.PartNo}
                            </div>
                            <div>{part.PartName}</div>
                          </td>
                          <td>
                            <label title={part.PartDescription}>
                              {isFullDescription
                                ? part.PartDescription
                                : part.PartDescriptionSort}
                            </label>
                          </td>
                          <td style={{ textAlign: "center" }}>
                            <TextField
                              resizable={false}
                              value={`${part.ListPrice}`}
                              onChange={(e) => {
                                let updateID = part.id;
                                let updateValue = e.target["value"];
                                listPriceChangeHandler(updateID, +updateValue);
                              }}
                            />{" "}
                          </td>
                        
                          <td style={{ textAlign: "center" }}>
                            <TextField
                              resizable={false}
                              value={`${part.quantity}`}
                              onChange={(e) => {
                                let updateID = part.id;
                                let updateValue = e.target["value"];
                                quantityChangeHandler(updateID, +updateValue);
                              }}
                            />
                          </td>
                          <td style={{ textAlign: "center" }}>
                            {part.NetPrice}
                          </td>
                          <td>{part.Note}</td>
                        </tr>
                      );
                    }
                  })}
                </tbody>
              </table>
            </PivotItem>
            <PivotItem
              headerText="JENEsys Enclosures"
              itemKey="7"
              itemCount={partsCount.JENEsysEnclosures}
            >
              <div>
                {selectedServices.JENEsysEngineeringTools.map((service) => {
                  return <span className={styles.selectedItem}>{service}</span>;
                })}
              </div>
              <table>
                <thead>
                  <tr>
                    <th></th>
                    <th>Part Name</th>
                    <th>
                      <div
                        style={{
                          display: "flex",
                          justifyContent: "space-between",
                          alignItems: "center",
                        }}
                      >
                        <div>Part Description</div>
                        <Icon
                          iconName={
                            isFullDescription
                              ? "ChevronUpSmall"
                              : "ChevronDownSmall"
                          }
                          styles={{
                            root: {
                              fontSize: 20,
                              fontWeight: 400,
                              color: "#004fa2",
                              cursor: "pointer",
                            },
                          }}
                          onClick={() => {
                            isFullDescription
                              ? setIsFullDescription(false)
                              : setIsFullDescription(true);
                          }}
                        />
                      </div>
                    </th>
                    <th>List Price</th>
                    <th>Quantity</th>
                    <th>Net Price</th>
                    <th>Note</th>
                  </tr>
                </thead>
                <tbody>
                  {partsDetails.map((part) => {
                    if (part.itemFor == "JENEsysEnclosures") {
                      return (
                        <tr
                          style={{
                            backgroundColor: part.isSelected
                              ? "#eef4fa"
                              : "#ffffff",
                          }}
                        >
                          <td>
                            {" "}
                            <Checkbox
                              checked={part.isSelected ? true : false}
                              onChange={() => {
                                tableCheckboxHandler(part.id, part);
                              }}
                            />
                          </td>
                          <td>
                            <div style={{ fontWeight: "bold" }}>
                              {part.PartNo}
                            </div>
                            <div>{part.PartName}</div>
                          </td>
                          <td>
                            <label title={part.PartDescription}>
                              {isFullDescription
                                ? part.PartDescription
                                : part.PartDescriptionSort}
                            </label>
                          </td>
                          <td style={{ textAlign: "center" }}>
                            <TextField
                              resizable={false}
                              value={`${part.ListPrice}`}
                              onChange={(e) => {
                                let updateID = part.id;
                                let updateValue = e.target["value"];
                                listPriceChangeHandler(updateID, +updateValue);
                              }}
                            />{" "}
                          </td>
                         
                          <td style={{ textAlign: "center" }}>
                            <TextField
                              resizable={false}
                              value={`${part.quantity}`}
                              onChange={(e) => {
                                let updateID = part.id;
                                let updateValue = e.target["value"];
                                quantityChangeHandler(updateID, +updateValue);
                              }}
                            />
                          </td>
                          <td style={{ textAlign: "center" }}>
                            {part.NetPrice}
                          </td>
                          <td>{part.Note}</td>
                        </tr>
                      );
                    }
                  })}
                </tbody>
              </table>
            </PivotItem>
            <PivotItem
              headerText="Renewals"
              itemKey="8"
              itemCount={partsCount.Renewals}
            >
              <div>
                {selectedServices.Renewals.map((service) => {
                  return <span className={styles.selectedItem}>{service}</span>;
                })}
              </div>
              <table>
                <thead>
                  <tr>
                    <th></th>
                    <th>Part Name</th>
                    <th>
                      <div
                        style={{
                          display: "flex",
                          justifyContent: "space-between",
                          alignItems: "center",
                        }}
                      >
                        <div>Part Description</div>
                        <Icon
                          iconName={
                            isFullDescription
                              ? "ChevronUpSmall"
                              : "ChevronDownSmall"
                          }
                          styles={{
                            root: {
                              fontSize: 20,
                              fontWeight: 400,
                              color: "#004fa2",
                              cursor: "pointer",
                            },
                          }}
                          onClick={() => {
                            isFullDescription
                              ? setIsFullDescription(false)
                              : setIsFullDescription(true);
                          }}
                        />
                      </div>
                    </th>
                    <th>List Price</th>
                    <th>Quantity</th>
                    <th>Net Price</th>
                    <th>Note</th>
                  </tr>
                </thead>
                <tbody>
                  {partsDetails.map((part) => {
                    if (part.itemFor == "Renewals") {
                      return (
                        <tr
                          style={{
                            backgroundColor: part.isSelected
                              ? "#eef4fa"
                              : "#ffffff",
                          }}
                        >
                          <td>
                            {" "}
                            <Checkbox
                              checked={part.isSelected ? true : false}
                              onChange={() => {
                                tableCheckboxHandler(part.id, part);
                              }}
                            />
                          </td>
                          <td>
                            <div style={{ fontWeight: "bold" }}>
                              {part.PartNo}
                            </div>
                            <div>{part.PartName}</div>
                          </td>
                          <td>
                            <label title={part.PartDescription}>
                              {isFullDescription
                                ? part.PartDescription
                                : part.PartDescriptionSort}
                            </label>
                          </td>
                          <td style={{ textAlign: "center" }}>
                            <TextField
                              resizable={false}
                              value={`${part.ListPrice}`}
                              onChange={(e) => {
                                let updateID = part.id;
                                let updateValue = e.target["value"];
                                listPriceChangeHandler(updateID, +updateValue);
                              }}
                            />{" "}
                          </td>
                        
                          <td style={{ textAlign: "center" }}>
                            <TextField
                              resizable={false}
                              value={`${part.quantity}`}
                              onChange={(e) => {
                                let updateID = part.id;
                                let updateValue = e.target["value"];
                                quantityChangeHandler(updateID, +updateValue);
                              }}
                            />
                          </td>
                          <td style={{ textAlign: "center" }}>
                            {part.NetPrice}
                          </td>
                          <td>{part.Note}</td>
                        </tr>
                      );
                    }
                  })}
                </tbody>
              </table>
            </PivotItem>
            <PivotItem
              headerText="JENEsys Thermostats & Peripherals"
              itemKey="9"
              itemCount={partsCount.JENEsysThermostatsPeripherals}
            >
              <div>
                {selectedServices.JENEsysThermostatsPeripherals.map(
                  (service) => {
                    return (
                      <span className={styles.selectedItem}>{service}</span>
                    );
                  }
                )}
              </div>
              <table>
                <thead>
                  <tr>
                    <th></th>
                    <th>Part Name</th>
                    <th>
                      <div
                        style={{
                          display: "flex",
                          justifyContent: "space-between",
                          alignItems: "center",
                        }}
                      >
                        <div>Part Description</div>
                        <Icon
                          iconName={
                            isFullDescription
                              ? "ChevronUpSmall"
                              : "ChevronDownSmall"
                          }
                          styles={{
                            root: {
                              fontSize: 20,
                              fontWeight: 400,
                              color: "#004fa2",
                              cursor: "pointer",
                            },
                          }}
                          onClick={() => {
                            isFullDescription
                              ? setIsFullDescription(false)
                              : setIsFullDescription(true);
                          }}
                        />
                      </div>
                    </th>
                    <th>List Price</th>
                    <th>Quantity</th>
                    <th>Net Price</th>
                    <th>Note</th>
                  </tr>
                </thead>
                <tbody>
                  {partsDetails.map((part) => {
                    if (part.itemFor == "JENEsysThermostatsPeripherals") {
                      return (
                        <tr
                          style={{
                            backgroundColor: part.isSelected
                              ? "#eef4fa"
                              : "#ffffff",
                          }}
                        >
                          <td>
                            {" "}
                            <Checkbox
                              checked={part.isSelected ? true : false}
                              onChange={() => {
                                tableCheckboxHandler(part.id, part);
                              }}
                            />
                          </td>
                          <td>
                            <div style={{ fontWeight: "bold" }}>
                              {part.PartNo}
                            </div>
                            <div>{part.PartName}</div>
                          </td>
                          <td>
                            <label title={part.PartDescription}>
                              {isFullDescription
                                ? part.PartDescription
                                : part.PartDescriptionSort}
                            </label>
                          </td>
                          <td style={{ textAlign: "center" }}>
                            <TextField
                              resizable={false}
                              value={`${part.ListPrice}`}
                              onChange={(e) => {
                                let updateID = part.id;
                                let updateValue = e.target["value"];
                                listPriceChangeHandler(updateID, +updateValue);
                              }}
                            />{" "}
                          </td>
                          
                          <td style={{ textAlign: "center" }}>
                            <TextField
                              resizable={false}
                              value={`${part.quantity}`}
                              onChange={(e) => {
                                let updateID = part.id;
                                let updateValue = e.target["value"];
                                quantityChangeHandler(updateID, +updateValue);
                              }}
                            />
                          </td>
                          <td style={{ textAlign: "center" }}>
                            {part.NetPrice}
                          </td>
                          <td>{part.Note}</td>
                        </tr>
                      );
                    }
                  })}
                </tbody>
              </table>
            </PivotItem>
            <PivotItem
              headerText="BACnet Controllers"
              itemKey="10"
              itemCount={partsCount.BACnetControllers}
            >
              <div>
                {selectedServices.BACnetControllers.map((service) => {
                  return <span className={styles.selectedItem}>{service}</span>;
                })}
              </div>
              <table>
                <thead>
                  <tr>
                    <th></th>
                    <th>Part Name</th>
                    <th>
                      <div
                        style={{
                          display: "flex",
                          justifyContent: "space-between",
                          alignItems: "center",
                        }}
                      >
                        <div>Part Description</div>
                        <Icon
                          iconName={
                            isFullDescription
                              ? "ChevronUpSmall"
                              : "ChevronDownSmall"
                          }
                          styles={{
                            root: {
                              fontSize: 20,
                              fontWeight: 400,
                              color: "#004fa2",
                              cursor: "pointer",
                            },
                          }}
                          onClick={() => {
                            isFullDescription
                              ? setIsFullDescription(false)
                              : setIsFullDescription(true);
                          }}
                        />
                      </div>
                    </th>
                    <th>List Price</th>
                    <th>Quantity</th>
                    <th>Net Price</th>
                    <th>Note</th>
                  </tr>
                </thead>
                <tbody>
                  {partsDetails.map((part) => {
                    if (part.itemFor == "BACnetControllers") {
                      return (
                        <tr
                          style={{
                            backgroundColor: part.isSelected
                              ? "#eef4fa"
                              : "#ffffff",
                          }}
                        >
                          <td>
                            {" "}
                            <Checkbox
                              checked={part.isSelected ? true : false}
                              onChange={() => {
                                tableCheckboxHandler(part.id, part);
                              }}
                            />
                          </td>
                          <td>
                            <div style={{ fontWeight: "bold" }}>
                              {part.PartNo}
                            </div>
                            <div>{part.PartName}</div>
                          </td>
                          <td>
                            <label title={part.PartDescription}>
                              {isFullDescription
                                ? part.PartDescription
                                : part.PartDescriptionSort}
                            </label>
                          </td>
                          <td style={{ textAlign: "center" }}>
                            <TextField
                              resizable={false}
                              value={`${part.ListPrice}`}
                              onChange={(e) => {
                                let updateID = part.id;
                                let updateValue = e.target["value"];
                                listPriceChangeHandler(updateID, +updateValue);
                              }}
                            />{" "}
                          </td>
                         
                          <td style={{ textAlign: "center" }}>
                            <TextField
                              resizable={false}
                              value={`${part.quantity}`}
                              onChange={(e) => {
                                let updateID = part.id;
                                let updateValue = e.target["value"];
                                quantityChangeHandler(updateID, +updateValue);
                              }}
                            />
                          </td>
                          <td style={{ textAlign: "center" }}>
                            {part.NetPrice}
                          </td>
                          <td>{part.Note}</td>
                        </tr>
                      );
                    }
                  })}
                </tbody>
              </table>
            </PivotItem>
            <PivotItem
              headerText="Distech"
              itemKey="11"
              itemCount={partsCount.Distech}
            >
              <div>
                {selectedServices.Distech.map((service) => {
                  return <span className={styles.selectedItem}>{service}</span>;
                })}
              </div>
              <table>
                <thead>
                  <tr>
                    <th></th>
                    <th>Part Name</th>
                    <th>
                      <div
                        style={{
                          display: "flex",
                          justifyContent: "space-between",
                          alignItems: "center",
                        }}
                      >
                        <div>Part Description</div>
                        <Icon
                          iconName={
                            isFullDescription
                              ? "ChevronUpSmall"
                              : "ChevronDownSmall"
                          }
                          styles={{
                            root: {
                              fontSize: 20,
                              fontWeight: 400,
                              color: "#004fa2",
                              cursor: "pointer",
                            },
                          }}
                          onClick={() => {
                            isFullDescription
                              ? setIsFullDescription(false)
                              : setIsFullDescription(true);
                          }}
                        />
                      </div>
                    </th>
                    <th>List Price</th>
                    <th>Net Price</th>
                    <th>Quantity</th>
                    <th>Note</th>
                  </tr>
                </thead>
                <tbody>
                  {partsDetails.map((part) => {
                    if (part.itemFor == "Distech") {
                      return (
                        <tr
                          style={{
                            backgroundColor: part.isSelected
                              ? "#eef4fa"
                              : "#ffffff",
                          }}
                        >
                          <td>
                            {" "}
                            <Checkbox
                              checked={part.isSelected ? true : false}
                              onChange={() => {
                                tableCheckboxHandler(part.id, part);
                              }}
                            />
                          </td>
                          <td>
                            <div style={{ fontWeight: "bold" }}>
                              {part.PartNo}
                            </div>
                            <div>{part.PartName}</div>
                          </td>
                          <td>
                            <label title={part.PartDescription}>
                              {isFullDescription
                                ? part.PartDescription
                                : part.PartDescriptionSort}
                            </label>
                          </td>
                          <td style={{ textAlign: "center" }}>
                            <TextField
                              resizable={false}
                              value={`${part.ListPrice}`}
                              onChange={(e) => {
                                let updateID = part.id;
                                let updateValue = e.target["value"];
                                listPriceChangeHandler(updateID, +updateValue);
                              }}
                            />{" "}
                          </td>
                        
                          <td style={{ textAlign: "center" }}>
                            <TextField
                              resizable={false}
                              value={`${part.quantity}`}
                              onChange={(e) => {
                                let updateID = part.id;
                                let updateValue = e.target["value"];
                                quantityChangeHandler(updateID, +updateValue);
                              }}
                            />
                          </td>
                          <td style={{ textAlign: "center" }}>
                            {part.NetPrice}
                          </td>
                          <td>{part.Note}</td>
                        </tr>
                      );
                    }
                  })}
                </tbody>
              </table>
            </PivotItem>
            <PivotItem
              headerText="Veris"
              itemKey="12"
              itemCount={partsCount.Veris}
            >
              <div>
                {selectedServices.Veris.map((service) => {
                  return <span className={styles.selectedItem}>{service}</span>;
                })}
              </div>
              <table>
                <thead>
                  <tr>
                    <th></th>
                    <th>Part Name</th>
                    <th>
                      <div
                        style={{
                          display: "flex",
                          justifyContent: "space-between",
                          alignItems: "center",
                        }}
                      >
                        <div>Part Description</div>
                        <Icon
                          iconName={
                            isFullDescription
                              ? "ChevronUpSmall"
                              : "ChevronDownSmall"
                          }
                          styles={{
                            root: {
                              fontSize: 20,
                              fontWeight: 400,
                              color: "#004fa2",
                              cursor: "pointer",
                            },
                          }}
                          onClick={() => {
                            isFullDescription
                              ? setIsFullDescription(false)
                              : setIsFullDescription(true);
                          }}
                        />
                      </div>
                    </th>
                    <th>List Price</th>
                    <th>Quantity</th>
                                        <th>Net Price</th>
                    <th>Note</th>
                  </tr>
                </thead>
                <tbody>
                  {partsDetails.map((part) => {
                    if (part.itemFor == "Veris") {
                      return (
                        <tr
                          style={{
                            backgroundColor: part.isSelected
                              ? "#eef4fa"
                              : "#ffffff",
                          }}
                        >
                          <td>
                            {" "}
                            <Checkbox
                              checked={part.isSelected ? true : false}
                              onChange={() => {
                                tableCheckboxHandler(part.id, part);
                              }}
                            />
                          </td>
                          <td>
                            <div style={{ fontWeight: "bold" }}>
                              {part.PartNo}
                            </div>
                            <div>{part.PartName}</div>
                          </td>
                          <td>
                            <label title={part.PartDescription}>
                              {isFullDescription
                                ? part.PartDescription
                                : part.PartDescriptionSort}
                            </label>
                          </td>
                          <td style={{ textAlign: "center" }}>
                            <TextField
                              resizable={false}
                              value={`${part.ListPrice}`}
                              onChange={(e) => {
                                let updateID = part.id;
                                let updateValue = e.target["value"];
                                listPriceChangeHandler(updateID, +updateValue);
                              }}
                            />{" "}
                          </td>
                          
                          <td style={{ textAlign: "center" }}>
                            <TextField
                              resizable={false}
                              value={`${part.quantity}`}
                              onChange={(e) => {
                                let updateID = part.id;
                                let updateValue = e.target["value"];
                                quantityChangeHandler(updateID, +updateValue);
                              }}
                            />
                          </td>
                          <td style={{ textAlign: "center" }}>
                            {part.NetPrice}
                          </td>
                          <td>{part.Note}</td>
                        </tr>
                      );
                    }
                  })}
                </tbody>
              </table>
            </PivotItem>
            <PivotItem
              headerText="Belimo"
              itemKey="13"
              itemCount={partsCount.Belimo}
            >
              <div>
                {selectedServices.Belimo.map((service) => {
                  return <span className={styles.selectedItem}>{service}</span>;
                })}
              </div>
              <table>
                <thead>
                  <tr>
                    <th></th>
                    <th>Part Name</th>
                    <th>
                      <div
                        style={{
                          display: "flex",
                          justifyContent: "space-between",
                          alignItems: "center",
                        }}
                      >
                        <div>Part Description</div>
                        <Icon
                          iconName={
                            isFullDescription
                              ? "ChevronUpSmall"
                              : "ChevronDownSmall"
                          }
                          styles={{
                            root: {
                              fontSize: 20,
                              fontWeight: 400,
                              color: "#004fa2",
                              cursor: "pointer",
                            },
                          }}
                          onClick={() => {
                            isFullDescription
                              ? setIsFullDescription(false)
                              : setIsFullDescription(true);
                          }}
                        />
                      </div>
                    </th>
                    <th>List Price</th>
                    <th>Quantity</th>

                    <th>Net Price</th>
                    <th>Note</th>
                  </tr>
                </thead>
                <tbody>
                  {partsDetails.map((part) => {
                    if (part.itemFor == "Belimo") {
                      return (
                        <tr>
                          <td>
                            {" "}
                            <Checkbox />
                          </td>
                          <td>
                            <div style={{ fontWeight: "bold" }}>
                              {part.PartNo}
                            </div>
                            <div>{part.PartName}</div>
                          </td>
                          <td>
                            <label title={part.PartDescription}>
                              {isFullDescription
                                ? part.PartDescription
                                : part.PartDescriptionSort}
                            </label>
                          </td>
                          <td style={{ textAlign: "center" }}>
                            <TextField
                              resizable={false}
                              value={`${part.ListPrice}`}
                              onChange={(e) => {
                                let updateID = part.id;
                                let updateValue = e.target["value"];
                                listPriceChangeHandler(updateID, +updateValue);
                              }}
                            />{" "}
                          </td>
                         
                          <td style={{ textAlign: "center" }}>
                            <TextField
                              resizable={false}
                              value={`${part.quantity}`}
                              onChange={(e) => {
                                let updateID = part.id;
                                let updateValue = e.target["value"];
                                quantityChangeHandler(updateID, +updateValue);
                              }}
                            />
                          </td>
                          <td style={{ textAlign: "center" }}>
                            {part.NetPrice}
                          </td>
                          <td>{part.Note}</td>
                        </tr>
                      );
                    }
                  })}
                </tbody>
              </table>
            </PivotItem>
            <PivotItem
              headerText="Temperature / RH / CO2 Sensors"
              itemKey="14"
              itemCount={partsCount.TemperatureRHCO2Sensors}
            >
              <div>
                {selectedServices.TemperatureRHCO2Sensors.map((service) => {
                  return <span className={styles.selectedItem}>{service}</span>;
                })}
              </div>
              <table>
                <thead>
                  <tr>
                    <th></th>
                    <th>Part Name</th>
                    <th>
                      <div
                        style={{
                          display: "flex",
                          justifyContent: "space-between",
                          alignItems: "center",
                        }}
                      >
                        <div>Part Description</div>
                        <Icon
                          iconName={
                            isFullDescription
                              ? "ChevronUpSmall"
                              : "ChevronDownSmall"
                          }
                          styles={{
                            root: {
                              fontSize: 20,
                              fontWeight: 400,
                              color: "#004fa2",
                              cursor: "pointer",
                            },
                          }}
                          onClick={() => {
                            isFullDescription
                              ? setIsFullDescription(false)
                              : setIsFullDescription(true);
                          }}
                        />
                      </div>
                    </th>
                    <th>List Price</th>
                    <th>Quantity</th>

                    <th>Net Price</th>
                    <th>Note</th>
                  </tr>
                </thead>
                <tbody>
                  {partsDetails.map((part) => {
                    if (part.itemFor == "TemperatureRHCO2Sensors") {
                      return (
                        <tr>
                          <td>
                            {" "}
                            <Checkbox />
                          </td>
                          <td>
                            <div style={{ fontWeight: "bold" }}>
                              {part.PartNo}
                            </div>
                            <div>{part.PartName}</div>
                          </td>
                          <td>
                            {isFullDescription
                              ? part.PartDescription
                              : part.PartDescriptionSort}
                          </td>
                          <td style={{ textAlign: "center" }}>
                            <TextField
                              resizable={false}
                              value={`${part.ListPrice}`}
                              onChange={(e) => {
                                let updateID = part.id;
                                let updateValue = e.target["value"];
                                listPriceChangeHandler(updateID, +updateValue);
                              }}
                            />{" "}
                          </td>
                          
                          <td style={{ textAlign: "center" }}>
                            <TextField
                              resizable={false}
                              value={`${part.quantity}`}
                              onChange={(e) => {
                                let updateID = part.id;
                                let updateValue = e.target["value"];
                                quantityChangeHandler(updateID, +updateValue);
                              }}
                            />
                          </td>
                          <td style={{ textAlign: "center" }}>
                            {part.NetPrice}
                          </td>
                          <td>{part.Note}</td>
                        </tr>
                      );
                    }
                  })}
                </tbody>
              </table>
            </PivotItem>
            <PivotItem
              headerText="Power Meters"
              itemKey="15"
              itemCount={partsCount.PowerMeters}
            >
              <div>
                {selectedServices.PowerMeters.map((service) => {
                  return <span className={styles.selectedItem}>{service}</span>;
                })}
              </div>
              <table>
                <thead>
                  <tr>
                    <th></th>
                    <th>Part Name</th>
                    <th>
                      <div
                        style={{
                          display: "flex",
                          justifyContent: "space-between",
                          alignItems: "center",
                        }}
                      >
                        <div>Part Description</div>
                        <Icon
                          iconName={
                            isFullDescription
                              ? "ChevronUpSmall"
                              : "ChevronDownSmall"
                          }
                          styles={{
                            root: {
                              fontSize: 20,
                              fontWeight: 400,
                              color: "#004fa2",
                              cursor: "pointer",
                            },
                          }}
                          onClick={() => {
                            isFullDescription
                              ? setIsFullDescription(false)
                              : setIsFullDescription(true);
                          }}
                        />
                      </div>
                    </th>
                    <th>List Price</th>
                    <th>Quantity</th>

                    <th>Net Price</th>
                    <th>Note</th>
                  </tr>
                </thead>
                <tbody>
                  {partsDetails.map((part) => {
                    if (part.itemFor == "PowerMeters") {
                      return (
                        <tr
                          style={{
                            backgroundColor: part.isSelected
                              ? "#eef4fa"
                              : "#ffffff",
                          }}
                        >
                          <td>
                            {" "}
                            <Checkbox
                              checked={part.isSelected ? true : false}
                              onChange={() => {
                                tableCheckboxHandler(part.id, part);
                              }}
                            />
                          </td>
                          <td>
                            <div style={{ fontWeight: "bold" }}>
                              {part.PartNo}
                            </div>
                            <div>{part.PartName}</div>
                          </td>
                          <td>
                            {isFullDescription
                              ? part.PartDescription
                              : part.PartDescriptionSort}
                          </td>
                          <td style={{ textAlign: "center" }}>
                            <TextField
                              resizable={false}
                              value={`${part.ListPrice}`}
                              onChange={(e) => {
                                let updateID = part.id;
                                let updateValue = e.target["value"];
                                listPriceChangeHandler(updateID, +updateValue);
                              }}
                            />{" "}
                          </td>
                          
                          <td style={{ textAlign: "center" }}>
                            <TextField
                              resizable={false}
                              value={`${part.quantity}`}
                              onChange={(e) => {
                                let updateID = part.id;
                                let updateValue = e.target["value"];
                                quantityChangeHandler(updateID, +updateValue);
                              }}
                            />
                          </td>
                          <td style={{ textAlign: "center" }}>
                            {part.NetPrice}
                          </td>
                          <td>{part.Note}</td>
                        </tr>
                      );
                    }
                  })}
                </tbody>
              </table>
            </PivotItem>
            <PivotItem
              headerText="Differential Pressure Transmitters & Switches"
              itemKey="16"
              itemCount={partsCount.DifferentialPressureTransmittersSwitches}
            >
              <div>
                {selectedServices.DifferentialPressureTransmittersSwitches.map(
                  (service) => {
                    return (
                      <span className={styles.selectedItem}>{service}</span>
                    );
                  }
                )}
              </div>
              <table>
                <thead>
                  <tr>
                    <th></th>
                    <th>Part Name</th>
                    <th>
                      <div
                        style={{
                          display: "flex",
                          justifyContent: "space-between",
                          alignItems: "center",
                        }}
                      >
                        <div>Part Description</div>
                        <Icon
                          iconName={
                            isFullDescription
                              ? "ChevronUpSmall"
                              : "ChevronDownSmall"
                          }
                          styles={{
                            root: {
                              fontSize: 20,
                              fontWeight: 400,
                              color: "#004fa2",
                              cursor: "pointer",
                            },
                          }}
                          onClick={() => {
                            isFullDescription
                              ? setIsFullDescription(false)
                              : setIsFullDescription(true);
                          }}
                        />
                      </div>
                    </th>
                    <th>List Price</th>
                    <th>Quantity</th>

                    <th>Net Price</th>
                    <th>Note</th>
                  </tr>
                </thead>
                <tbody>
                  {partsDetails.map((part) => {
                    if (
                      part.itemFor == "DifferentialPressureTransmittersSwitches"
                    ) {
                      return (
                        <tr
                          style={{
                            backgroundColor: part.isSelected
                              ? "#eef4fa"
                              : "#ffffff",
                          }}
                        >
                          <td>
                            {" "}
                            <Checkbox
                              checked={part.isSelected ? true : false}
                              onChange={() => {
                                tableCheckboxHandler(part.id, part);
                              }}
                            />
                          </td>
                          <td>
                            <div style={{ fontWeight: "bold" }}>
                              {part.PartNo}
                            </div>
                            <div>{part.PartName}</div>
                          </td>
                          <td>
                            {isFullDescription
                              ? part.PartDescription
                              : part.PartDescriptionSort}
                          </td>
                          <td style={{ textAlign: "center" }}>
                            <TextField
                              resizable={false}
                              value={`${part.ListPrice}`}
                              onChange={(e) => {
                                let updateID = part.id;
                                let updateValue = e.target["value"];
                                listPriceChangeHandler(updateID, +updateValue);
                              }}
                            />{" "}
                          </td>
                        
                          <td style={{ textAlign: "center" }}>
                            <TextField
                              resizable={false}
                              value={`${part.quantity}`}
                              onChange={(e) => {
                                let updateID = part.id;
                                let updateValue = e.target["value"];
                                quantityChangeHandler(updateID, +updateValue);
                              }}
                            />
                          </td>
                          <td style={{ textAlign: "center" }}>
                            {part.NetPrice}
                          </td>
                          <td>{part.Note}</td>
                        </tr>
                      );
                    }
                  })}
                </tbody>
              </table>
            </PivotItem>
            <PivotItem
              headerText="Relays"
              itemKey="17"
              itemCount={partsCount.Relays}
            >
              <div>
                {selectedServices.Relays.map((service) => {
                  return <span className={styles.selectedItem}>{service}</span>;
                })}
              </div>
              <table>
                <thead>
                  <tr>
                    <th></th>
                    <th>Part Name</th>
                    <th>
                      <div
                        style={{
                          display: "flex",
                          justifyContent: "space-between",
                          alignItems: "center",
                        }}
                      >
                        <div>Part Description</div>
                        <Icon
                          iconName={
                            isFullDescription
                              ? "ChevronUpSmall"
                              : "ChevronDownSmall"
                          }
                          styles={{
                            root: {
                              fontSize: 20,
                              fontWeight: 400,
                              color: "#004fa2",
                              cursor: "pointer",
                            },
                          }}
                          onClick={() => {
                            isFullDescription
                              ? setIsFullDescription(false)
                              : setIsFullDescription(true);
                          }}
                        />
                      </div>
                    </th>
                    <th>List Price</th>
                    <th>Quantity</th>

                    <th>Net Price</th>
                    <th>Note</th>
                  </tr>
                </thead>
                <tbody>
                  {partsDetails.map((part) => {
                    if (part.itemFor == "Relays") {
                      return (
                        <tr
                          style={{
                            backgroundColor: part.isSelected
                              ? "#eef4fa"
                              : "#ffffff",
                          }}
                        >
                          <td>
                            {" "}
                            <Checkbox
                              checked={part.isSelected ? true : false}
                              onChange={() => {
                                tableCheckboxHandler(part.id, part);
                              }}
                            />
                          </td>
                          <td>
                            <div style={{ fontWeight: "bold" }}>
                              {part.PartNo}
                            </div>
                            <div>{part.PartName}</div>
                          </td>
                          <td>
                            {isFullDescription
                              ? part.PartDescription
                              : part.PartDescriptionSort}
                          </td>
                          <td style={{ textAlign: "center" }}>
                            <TextField
                              resizable={false}
                              value={`${part.ListPrice}`}
                              onChange={(e) => {
                                let updateID = part.id;
                                let updateValue = e.target["value"];
                                listPriceChangeHandler(updateID, +updateValue);
                              }}
                            />{" "}
                          </td>
                          
                          <td style={{ textAlign: "center" }}>
                            <TextField
                              resizable={false}
                              value={`${part.quantity}`}
                              onChange={(e) => {
                                let updateID = part.id;
                                let updateValue = e.target["value"];
                                quantityChangeHandler(updateID, +updateValue);
                              }}
                            />
                          </td>
                          <td style={{ textAlign: "center" }}>
                            {part.NetPrice}
                          </td>
                          <td>{part.Note}</td>
                        </tr>
                      );
                    }
                  })}
                </tbody>
              </table>
            </PivotItem>
            <PivotItem
              headerText="Current Sensors & Transmitters"
              itemKey="18"
              itemCount={partsCount.CurrentSensorsTransmitters}
            >
              <div>
                {selectedServices.CurrentSensorsTransmitters.map((service) => {
                  return <span className={styles.selectedItem}>{service}</span>;
                })}
              </div>
              <table>
                <thead>
                  <tr>
                    <th></th>
                    <th>Part Name</th>
                    <th>
                      <div
                        style={{
                          display: "flex",
                          justifyContent: "space-between",
                          alignItems: "center",
                        }}
                      >
                        <div>Part Description</div>
                        <Icon
                          iconName={
                            isFullDescription
                              ? "ChevronUpSmall"
                              : "ChevronDownSmall"
                          }
                          styles={{
                            root: {
                              fontSize: 20,
                              fontWeight: 400,
                              color: "#004fa2",
                              cursor: "pointer",
                            },
                          }}
                          onClick={() => {
                            isFullDescription
                              ? setIsFullDescription(false)
                              : setIsFullDescription(true);
                          }}
                        />
                      </div>
                    </th>
                    <th>List Price</th>
                    <th>Quantity</th>

                    <th>Net Price</th>
                    <th>Note</th>
                  </tr>
                </thead>
                <tbody>
                  {partsDetails.map((part) => {
                    if (part.itemFor == "CurrentSensorsTransmitters") {
                      return (
                        <tr
                          style={{
                            backgroundColor: part.isSelected
                              ? "#eef4fa"
                              : "#ffffff",
                          }}
                        >
                          <td>
                            {" "}
                            <Checkbox
                              checked={part.isSelected ? true : false}
                              onChange={() => {
                                tableCheckboxHandler(part.id, part);
                              }}
                            />
                          </td>
                          <td>
                            <div style={{ fontWeight: "bold" }}>
                              {part.PartNo}
                            </div>
                            <div>{part.PartName}</div>
                          </td>
                          <td>
                            {isFullDescription
                              ? part.PartDescription
                              : part.PartDescriptionSort}
                          </td>
                          <td style={{ textAlign: "center" }}>
                            <TextField
                              resizable={false}
                              value={`${part.ListPrice}`}
                              onChange={(e) => {
                                let updateID = part.id;
                                let updateValue = e.target["value"];
                                listPriceChangeHandler(updateID, +updateValue);
                              }}
                            />{" "}
                          </td>
                         
                          <td style={{ textAlign: "center" }}>
                            <TextField
                              resizable={false}
                              value={`${part.quantity}`}
                              onChange={(e) => {
                                let updateID = part.id;
                                let updateValue = e.target["value"];
                                quantityChangeHandler(updateID, +updateValue);
                              }}
                            />
                          </td>
                          <td style={{ textAlign: "center" }}>
                            {part.NetPrice}
                          </td>
                          <td>{part.Note}</td>
                        </tr>
                      );
                    }
                  })}
                </tbody>
              </table>
            </PivotItem>
            <PivotItem
              headerText="Power Supplies"
              itemKey="19"
              itemCount={partsCount.PowerSupplies}
            >
              <div>
                {selectedServices.PowerSupplies.map((service) => {
                  return <span className={styles.selectedItem}>{service}</span>;
                })}
              </div>
              <table>
                <thead>
                  <tr>
                    <th></th>
                    <th>Part Name</th>
                    <th>
                      <div
                        style={{
                          display: "flex",
                          justifyContent: "space-between",
                          alignItems: "center",
                        }}
                      >
                        <div>Part Description</div>
                        <Icon
                          iconName={
                            isFullDescription
                              ? "ChevronUpSmall"
                              : "ChevronDownSmall"
                          }
                          styles={{
                            root: {
                              fontSize: 20,
                              fontWeight: 400,
                              color: "#004fa2",
                              cursor: "pointer",
                            },
                          }}
                          onClick={() => {
                            isFullDescription
                              ? setIsFullDescription(false)
                              : setIsFullDescription(true);
                          }}
                        />
                      </div>
                    </th>
                    <th>List Price</th>
                    <th>Quantity</th>

                    <th>Net Price</th>
                    <th>Note</th>
                  </tr>
                </thead>
                <tbody>
                  {partsDetails.map((part) => {
                    if (part.itemFor == "PowerSupplies") {
                      return (
                        <tr
                          style={{
                            backgroundColor: part.isSelected
                              ? "#eef4fa"
                              : "#ffffff",
                          }}
                        >
                          <td>
                            {" "}
                            <Checkbox
                              checked={part.isSelected ? true : false}
                              onChange={() => {
                                tableCheckboxHandler(part.id, part);
                              }}
                            />
                          </td>
                          <td>
                            <div style={{ fontWeight: "bold" }}>
                              {part.PartNo}
                            </div>
                            <div>{part.PartName}</div>
                          </td>
                          <td>
                            {isFullDescription
                              ? part.PartDescription
                              : part.PartDescriptionSort}
                          </td>
                          <td style={{ textAlign: "center" }}>
                            <TextField
                              resizable={false}
                              value={`${part.ListPrice}`}
                              onChange={(e) => {
                                let updateID = part.id;
                                let updateValue = e.target["value"];
                                listPriceChangeHandler(updateID, +updateValue);
                              }}
                            />{" "}
                          </td>
                         
                          <td style={{ textAlign: "center" }}>
                            <TextField
                              resizable={false}
                              value={`${part.quantity}`}
                              onChange={(e) => {
                                let updateID = part.id;
                                let updateValue = e.target["value"];
                                quantityChangeHandler(updateID, +updateValue);
                              }}
                            />
                          </td>
                          <td style={{ textAlign: "center" }}>
                            {part.NetPrice}
                          </td>
                          <td>{part.Note}</td>
                        </tr>
                      );
                    }
                  })}
                </tbody>
              </table>
            </PivotItem>
            <PivotItem
              headerText="Transformers"
              itemKey="20"
              itemCount={partsCount.Transformers}
            >
              <div>
                {selectedServices.Transformers.map((service) => {
                  return <span className={styles.selectedItem}>{service}</span>;
                })}
              </div>
              <table>
                <thead>
                  <tr>
                    <th></th>
                    <th>Part Name</th>
                    <th>
                      <div
                        style={{
                          display: "flex",
                          justifyContent: "space-between",
                          alignItems: "center",
                        }}
                      >
                        <div>Part Description</div>
                        <Icon
                          iconName={
                            isFullDescription
                              ? "ChevronUpSmall"
                              : "ChevronDownSmall"
                          }
                          styles={{
                            root: {
                              fontSize: 20,
                              fontWeight: 400,
                              color: "#004fa2",
                              cursor: "pointer",
                            },
                          }}
                          onClick={() => {
                            isFullDescription
                              ? setIsFullDescription(false)
                              : setIsFullDescription(true);
                          }}
                        />
                      </div>
                    </th>
                    <th>List Price</th>
                    <th>Quantity</th>

                    <th>Net Price</th>
                    <th>Note</th>
                  </tr>
                </thead>
                <tbody>
                  {partsDetails.map((part) => {
                    if (part.itemFor == "Transformers") {
                      return (
                        <tr
                          style={{
                            backgroundColor: part.isSelected
                              ? "#eef4fa"
                              : "#ffffff",
                          }}
                        >
                          <td>
                            {" "}
                            <Checkbox
                              checked={part.isSelected ? true : false}
                              onChange={() => {
                                tableCheckboxHandler(part.id, part);
                              }}
                            />
                          </td>
                          <td>
                            <div style={{ fontWeight: "bold" }}>
                              {part.PartNo}
                            </div>
                            <div>{part.PartName}</div>
                          </td>
                          <td>
                            {isFullDescription
                              ? part.PartDescription
                              : part.PartDescriptionSort}
                          </td>
                          <td style={{ textAlign: "center" }}>
                            <TextField
                              resizable={false}
                              value={`${part.ListPrice}`}
                              onChange={(e) => {
                                let updateID = part.id;
                                let updateValue = e.target["value"];
                                listPriceChangeHandler(updateID, +updateValue);
                              }}
                            />{" "}
                          </td>
                          
                          <td style={{ textAlign: "center" }}>
                            <TextField
                              resizable={false}
                              value={`${part.quantity}`}
                              onChange={(e) => {
                                let updateID = part.id;
                                let updateValue = e.target["value"];
                                quantityChangeHandler(updateID, +updateValue);
                              }}
                            />
                          </td>
                          <td style={{ textAlign: "center" }}>
                            {part.NetPrice}
                          </td>
                          <td>{part.Note}</td>
                        </tr>
                      );
                    }
                  })}
                </tbody>
              </table>
            </PivotItem>
            <PivotItem
              headerText="Lynxspring University"
              itemKey="21"
              itemCount={partsCount.LynxspringUniversity}
            >
              <div>
                {selectedServices.LynxspringUniversity.map((service) => {
                  return <span className={styles.selectedItem}>{service}</span>;
                })}
              </div>
              <table>
                <thead>
                  <tr>
                    <th></th>
                    <th>Part Name</th>
                    <th>
                      <div
                        style={{
                          display: "flex",
                          justifyContent: "space-between",
                          alignItems: "center",
                        }}
                      >
                        <div>Part Description</div>
                        <Icon
                          iconName={
                            isFullDescription
                              ? "ChevronUpSmall"
                              : "ChevronDownSmall"
                          }
                          styles={{
                            root: {
                              fontSize: 20,
                              fontWeight: 400,
                              color: "#004fa2",
                              cursor: "pointer",
                            },
                          }}
                          onClick={() => {
                            isFullDescription
                              ? setIsFullDescription(false)
                              : setIsFullDescription(true);
                          }}
                        />
                      </div>
                    </th>
                    <th>List Price</th>
                    <th>Quantity</th>

                    <th>Net Price</th>
                    <th>Note</th>
                  </tr>
                </thead>
                <tbody>
                  {partsDetails.map((part) => {
                    if (part.itemFor == "LynxspringUniversity") {
                      return (
                        <tr
                          style={{
                            backgroundColor: part.isSelected
                              ? "#eef4fa"
                              : "#ffffff",
                          }}
                        >
                          <td>
                            {" "}
                            <Checkbox
                              checked={part.isSelected ? true : false}
                              onChange={() => {
                                tableCheckboxHandler(part.id, part);
                              }}
                            />
                          </td>
                          <td>
                            <div style={{ fontWeight: "bold" }}>
                              {part.PartNo}
                            </div>
                            <div>{part.PartName}</div>
                          </td>
                          <td>
                            {isFullDescription
                              ? part.PartDescription
                              : part.PartDescriptionSort}
                          </td>
                          <td style={{ textAlign: "center" }}>
                            <TextField
                              resizable={false}
                              value={`${part.ListPrice}`}
                              onChange={(e) => {
                                let updateID = part.id;
                                let updateValue = e.target["value"];
                                listPriceChangeHandler(updateID, +updateValue);
                              }}
                            />{" "}
                          </td>
                        
                          <td style={{ textAlign: "center" }}>
                            <TextField
                              resizable={false}
                              value={`${part.quantity}`}
                              onChange={(e) => {
                                let updateID = part.id;
                                let updateValue = e.target["value"];
                                quantityChangeHandler(updateID, +updateValue);
                              }}
                            />
                          </td>
                          <td style={{ textAlign: "center" }}>
                            {part.NetPrice}
                          </td>
                          <td>{part.Note}</td>
                        </tr>
                      );
                    }
                  })}
                </tbody>
              </table>
            </PivotItem>
            <PivotItem
              headerText="TAPA"
              itemKey="22"
              itemCount={partsCount.TAPA}
            >
              <div>
                {selectedServices.TAPA.map((service) => {
                  return <span className={styles.selectedItem}>{service}</span>;
                })}
              </div>
              <table>
                <thead>
                  <tr>
                    <th></th>
                    <th>Part Name</th>
                    <th>
                      <div
                        style={{
                          display: "flex",
                          justifyContent: "space-between",
                          alignItems: "center",
                        }}
                      >
                        <div>Part Description</div>
                        <Icon
                          iconName={
                            isFullDescription
                              ? "ChevronUpSmall"
                              : "ChevronDownSmall"
                          }
                          styles={{
                            root: {
                              fontSize: 20,
                              fontWeight: 400,
                              color: "#004fa2",
                              cursor: "pointer",
                            },
                          }}
                          onClick={() => {
                            isFullDescription
                              ? setIsFullDescription(false)
                              : setIsFullDescription(true);
                          }}
                        />
                      </div>
                    </th>
                    <th>List Price</th>
                    <th>Quantity</th>

                    <th>Net Price</th>
                    <th>Note</th>
                  </tr>
                </thead>
                <tbody>
                  {partsDetails.map((part) => {
                    if (part.itemFor == "TAPA") {
                      return (
                        <tr
                          style={{
                            backgroundColor: part.isSelected
                              ? "#eef4fa"
                              : "#ffffff",
                          }}
                        >
                          <td>
                            {" "}
                            <Checkbox
                              checked={part.isSelected ? true : false}
                              onChange={() => {
                                tableCheckboxHandler(part.id, part);
                              }}
                            />
                          </td>
                          <td>
                            <div style={{ fontWeight: "bold" }}>
                              {part.PartNo}
                            </div>
                            <div>{part.PartName}</div>
                          </td>
                          <td>
                            {isFullDescription
                              ? part.PartDescription
                              : part.PartDescriptionSort}
                          </td>
                          <td style={{ textAlign: "center" }}>
                            <TextField
                              resizable={false}
                              value={`${part.ListPrice}`}
                              onChange={(e) => {
                                let updateID = part.id;
                                let updateValue = e.target["value"];
                                listPriceChangeHandler(updateID, +updateValue);
                              }}
                            />{" "}
                          </td>
                      
                          <td style={{ textAlign: "center" }}>
                            <TextField
                              resizable={false}
                              value={`${part.quantity}`}
                              onChange={(e) => {
                                let updateID = part.id;
                                let updateValue = e.target["value"];
                                quantityChangeHandler(updateID, +updateValue);
                              }}
                            />
                          </td>
                          <td style={{ textAlign: "center" }}>
                            {part.NetPrice}
                          </td>
                          <td>{part.Note}</td>
                        </tr>
                      );
                    }
                  })}
                </tbody>
              </table>
            </PivotItem>
            <PivotItem
              headerText="DGLux"
              itemKey="23"
              itemCount={partsCount.DGLux}
            >
              <div>
                {selectedServices.DGLux.map((service) => {
                  return <span className={styles.selectedItem}>{service}</span>;
                })}
              </div>
              <table>
                <thead>
                  <tr>
                    <th></th>
                    <th>Part Name</th>
                    <th>
                      <div
                        style={{
                          display: "flex",
                          justifyContent: "space-between",
                          alignItems: "center",
                        }}
                      >
                        <div>Part Description</div>
                        <Icon
                          iconName={
                            isFullDescription
                              ? "ChevronUpSmall"
                              : "ChevronDownSmall"
                          }
                          styles={{
                            root: {
                              fontSize: 20,
                              fontWeight: 400,
                              color: "#004fa2",
                              cursor: "pointer",
                            },
                          }}
                          onClick={() => {
                            isFullDescription
                              ? setIsFullDescription(false)
                              : setIsFullDescription(true);
                          }}
                        />
                      </div>
                    </th>
                    <th>List Price</th>
                    <th>Quantity</th>

                    <th>Net Price</th>
                    <th>Note</th>
                  </tr>
                </thead>
                <tbody>
                  {partsDetails.map((part) => {
                    if (part.itemFor == "DGLux") {
                      return (
                        <tr
                          style={{
                            backgroundColor: part.isSelected
                              ? "#eef4fa"
                              : "#ffffff",
                          }}
                        >
                          <td>
                            {" "}
                            <Checkbox
                              checked={part.isSelected ? true : false}
                              onChange={() => {
                                tableCheckboxHandler(part.id, part);
                              }}
                            />
                          </td>
                          <td>
                            <div style={{ fontWeight: "bold" }}>
                              {part.PartNo}
                            </div>
                            <div>{part.PartName}</div>
                          </td>
                          <td>
                            {isFullDescription
                              ? part.PartDescription
                              : part.PartDescriptionSort}
                          </td>
                          <td style={{ textAlign: "center" }}>
                            <TextField
                              resizable={false}
                              value={`${part.ListPrice}`}
                              onChange={(e) => {
                                let updateID = part.id;
                                let updateValue = e.target["value"];
                                listPriceChangeHandler(updateID, +updateValue);
                              }}
                            />{" "}
                          </td>
                          
                          <td style={{ textAlign: "center" }}>
                            <TextField
                              resizable={false}
                              value={`${part.quantity}`}
                              onChange={(e) => {
                                let updateID = part.id;
                                let updateValue = e.target["value"];
                                quantityChangeHandler(updateID, +updateValue);
                              }}
                            />
                          </td>
                          <td style={{ textAlign: "center" }}>
                            {part.NetPrice}
                          </td>
                          <td>{part.Note}</td>
                        </tr>
                      );
                    }
                  })}
                </tbody>
              </table>
            </PivotItem>
            <PivotItem
              headerText="SkyFoundry"
              itemKey="24"
              itemCount={partsCount.SkyFoundry}
            >
              <div>
                {selectedServices.SkyFoundry.map((service) => {
                  return <span className={styles.selectedItem}>{service}</span>;
                })}
              </div>
              <table>
                <thead>
                  <tr>
                    <th></th>
                    <th>Part Name</th>
                    <th>
                      <div
                        style={{
                          display: "flex",
                          justifyContent: "space-between",
                          alignItems: "center",
                        }}
                      >
                        <div>Part Description</div>
                        <Icon
                          iconName={
                            isFullDescription
                              ? "ChevronUpSmall"
                              : "ChevronDownSmall"
                          }
                          styles={{
                            root: {
                              fontSize: 20,
                              fontWeight: 400,
                              color: "#004fa2",
                              cursor: "pointer",
                            },
                          }}
                          onClick={() => {
                            isFullDescription
                              ? setIsFullDescription(false)
                              : setIsFullDescription(true);
                          }}
                        />
                      </div>
                    </th>
                    <th>List Price</th>
                    <th>Quantity</th>

                    <th>Net Price</th>
                    <th>Note</th>
                  </tr>
                </thead>
                <tbody>
                  {partsDetails.map((part) => {
                    if (part.itemFor == "SkyFoundry") {
                      return (
                        <tr
                          style={{
                            backgroundColor: part.isSelected
                              ? "#eef4fa"
                              : "#ffffff",
                          }}
                        >
                          <td>
                            {" "}
                            <Checkbox
                              checked={part.isSelected ? true : false}
                              onChange={() => {
                                tableCheckboxHandler(part.id, part);
                              }}
                            />
                          </td>
                          <td>
                            <div style={{ fontWeight: "bold" }}>
                              {part.PartNo}
                            </div>
                            <div>{part.PartName}</div>
                          </td>
                          <td>
                            {isFullDescription
                              ? part.PartDescription
                              : part.PartDescriptionSort}
                          </td>
                          <td style={{ textAlign: "center" }}>
                            <TextField
                              resizable={false}
                              value={`${part.ListPrice}`}
                              onChange={(e) => {
                                let updateID = part.id;
                                let updateValue = e.target["value"];
                                listPriceChangeHandler(updateID, +updateValue);
                              }}
                            />{" "}
                          </td>
                         
                          <td style={{ textAlign: "center" }}>
                            <TextField
                              resizable={false}
                              value={`${part.quantity}`}
                              onChange={(e) => {
                                let updateID = part.id;
                                let updateValue = e.target["value"];
                                quantityChangeHandler(updateID, +updateValue);
                              }}
                            />
                          </td>
                          <td style={{ textAlign: "center" }}>
                            {part.NetPrice}
                          </td>
                          <td>{part.Note}</td>
                        </tr>
                      );
                    }
                  })}
                </tbody>
              </table>
            </PivotItem>
            <PivotItem
              headerText="Tridium Analytics"
              itemKey="25"
              itemCount={partsCount.TridiumAnalytics}
            >
              <div>
                {selectedServices.TridiumAnalytics.map((service) => {
                  return <span className={styles.selectedItem}>{service}</span>;
                })}
              </div>
              <table>
                <thead>
                  <tr>
                    <th></th>
                    <th>Part Name</th>
                    <th>
                      <div
                        style={{
                          display: "flex",
                          justifyContent: "space-between",
                          alignItems: "center",
                        }}
                      >
                        <div>Part Description</div>
                        <Icon
                          iconName={
                            isFullDescription
                              ? "ChevronUpSmall"
                              : "ChevronDownSmall"
                          }
                          styles={{
                            root: {
                              fontSize: 20,
                              fontWeight: 400,
                              color: "#004fa2",
                              cursor: "pointer",
                            },
                          }}
                          onClick={() => {
                            isFullDescription
                              ? setIsFullDescription(false)
                              : setIsFullDescription(true);
                          }}
                        />
                      </div>
                    </th>
                    <th>List Price</th>
                    <th>Quantity</th>

                    <th>Net Price</th>
                    <th>Note</th>
                  </tr>
                </thead>
                <tbody>
                  {partsDetails.map((part) => {
                    if (part.itemFor == "TridiumAnalytics") {
                      return (
                        <tr
                          style={{
                            backgroundColor: part.isSelected
                              ? "#eef4fa"
                              : "#ffffff",
                          }}
                        >
                          <td>
                            {" "}
                            <Checkbox
                              checked={part.isSelected ? true : false}
                              onChange={() => {
                                tableCheckboxHandler(part.id, part);
                              }}
                            />
                          </td>
                          <td>
                            <div style={{ fontWeight: "bold" }}>
                              {part.PartNo}
                            </div>
                            <div>{part.PartName}</div>
                          </td>
                          <td>
                            {isFullDescription
                              ? part.PartDescription
                              : part.PartDescriptionSort}
                          </td>
                          <td style={{ textAlign: "center" }}>
                            <TextField
                              resizable={false}
                              value={`${part.ListPrice}`}
                              onChange={(e) => {
                                let updateID = part.id;
                                let updateValue = e.target["value"];
                                listPriceChangeHandler(updateID, +updateValue);
                              }}
                            />{" "}
                          </td>
                         
                          <td style={{ textAlign: "center" }}>
                            <TextField
                              resizable={false}
                              value={`${part.quantity}`}
                              onChange={(e) => {
                                let updateID = part.id;
                                let updateValue = e.target["value"];
                                quantityChangeHandler(updateID, +updateValue);
                              }}
                            />
                          </td>
                          <td style={{ textAlign: "center" }}>
                            {part.NetPrice}
                          </td>
                          <td>{part.Note}</td>
                        </tr>
                      );
                    }
                  })}
                </tbody>
              </table>
            </PivotItem>
            <PivotItem headerText="Summary" itemKey="26">
              <table>
                <thead>
                  <tr>
                    <th></th>
                    <th>Part Name</th>
                    <th>
                      <div
                        style={{
                          display: "flex",
                          justifyContent: "space-between",
                          alignItems: "center",
                        }}
                      >
                        <div>Part Description</div>
                        <Icon
                          iconName={
                            isFullDescription
                              ? "ChevronUpSmall"
                              : "ChevronDownSmall"
                          }
                          styles={{
                            root: {
                              fontSize: 20,
                              fontWeight: 400,
                              color: "#004fa2",
                              cursor: "pointer",
                            },
                          }}
                          onClick={() => {
                            isFullDescription
                              ? setIsFullDescription(false)
                              : setIsFullDescription(true);
                          }}
                        />
                      </div>
                    </th>
                    <th>List Price</th>
                    <th>Quantity</th>

                    <th>Net Price</th>
                    <th>Note</th>
                  </tr>
                </thead>
                <tbody>
                  {partsDetails.map((part) => {
                    if (part.isSelected) {
                      return (
                        <tr
                          style={{
                            backgroundColor: part.isSelected
                              ? "#eef4fa"
                              : "#ffffff",
                          }}
                        >
                          <td>
                            {" "}
                            <Checkbox
                              checked={part.isSelected ? true : false}
                              onChange={() => {
                                tableCheckboxHandler(part.id, part);
                              }}
                            />
                          </td>
                          <td>
                            <div style={{ fontWeight: "bold" }}>
                              {part.PartNo}
                            </div>
                            <div>{part.PartName}</div>
                          </td>
                          <td>
                            {isFullDescription
                              ? part.PartDescription
                              : part.PartDescriptionSort}
                          </td>
                          <td style={{ textAlign: "center" }}>
                            <TextField
                              resizable={false}
                              value={`${part.ListPrice}`}
                              onChange={(e) => {
                                let updateID = part.id;
                                let updateValue = e.target["value"];
                                listPriceChangeHandler(updateID, +updateValue);
                              }}
                            />{" "}
                          </td>
                         
                          <td style={{ textAlign: "center" }}>
                            <TextField
                              resizable={false}
                              value={`${part.quantity}`}
                              onChange={(e) => {
                                let updateID = part.id;
                                let updateValue = e.target["value"];
                                quantityChangeHandler(updateID, +updateValue);
                              }}
                            />
                          </td>
                          <td style={{ textAlign: "center" }}>
                            {part.NetPrice}
                          </td>
                          <td>{part.Note}</td>
                        </tr>
                      );
                    }
                  })}
                </tbody>
              </table>
            </PivotItem>
          </Pivot>
        </ThemeProvider>
        <div style={{ display: "flex", justifyContent: "space-between" }}>
          <ActionButton
            iconProps={{ iconName: "DoubleChevronLeft8" }}
            allowDisabledFocus
            onClick={onPrevClick}
            disabled={selectedKey == 1 ? true : false}
          >
            Previous
          </ActionButton>
          <ActionButton
            styles={{ flexContainer: { flexDirection: "row-reverse" } }}
            iconProps={{ iconName: "DoubleChevronRight8" }}
            allowDisabledFocus
            onClick={onNextClick}
            disabled={selectedKey == 26 ? true : false}
          >
            Next
          </ActionButton>
        </div>
        {/* Pivot */}
        {/*  Pivot Section End */}
      </div>
      {/* Section */}
      <div className={styles.section}>
        <div style={{ display: "flex", justifyContent: "space-between" }}>
          <div>
            <h3
              style={{
                color: myTheme.palette.themePrimary,
              }}
            >
              Proposed for Consultant: Lynxspring, Inc
            </h3>
            <TextField
              styles={halfWidthInput}
              label="By"
              value={objValues.ProposedBy}
              onChange={(e) => {
                objValues.ProposedBy = e.target["value"];
                setRenderObjValue(true);
              }}
            />
            <TextField
              styles={halfWidthInput}
              label="Name (Printed)"
              value={objValues.ProposedName}
              onChange={(e) => {
                objValues.ProposedName = e.target["value"];
                setRenderObjValue(true);
              }}
            />
            <TextField
              styles={halfWidthInput}
              label="Title"
              value={objValues.ProposedTitle}
              onChange={(e) => {
                objValues.ProposedTitle = e.target["value"];
                setRenderObjValue(true);
              }}
            />
            <DatePicker
              styles={halfWidthInput}
              label="Date"
              formatDate={(date: Date): string => {
                return (
                  date.getMonth() +
                  1 +
                  "/" +
                  date.getDate() +
                  "/" +
                  date.getFullYear()
                );
              }}
              onSelectDate={(date) => {
                objValues.ProposedDate = date;
                console.log(objValues);
                setRenderObjValue(true);
              }}
            />
          </div>
          <div>
            <div>
              <div style={{ display: "flex", alignItems: "center" }}>
                <h3
                  style={{
                    color: myTheme.palette.themePrimary,
                    marginRight: "0.5rem",
                  }}
                >
                  Accepted for Client:
                </h3>
                <TextField
                  styles={{ root: { width: 200 } }}
                  placeholder="Client Name"
                  value={objValues.AcceptedForClient}
                  onChange={(e) => {
                    objValues.AcceptedForClient = e.target["value"];
                    setRenderObjValue(true);
                  }}
                />
              </div>
              <TextField
                styles={halfWidthInput}
                label="By"
                value={objValues.AcceptedBy}
                onChange={(e) => {
                  objValues.AcceptedBy = e.target["value"];
                  setRenderObjValue(true);
                }}
              />
              <TextField
                styles={halfWidthInput}
                label="Name (Printed)"
                value={objValues.AcceptedByName}
                onChange={(e) => {
                  objValues.AcceptedByName = e.target["value"];
                  setRenderObjValue(true);
                }}
              />
              <TextField
                styles={halfWidthInput}
                label="Title"
                value={objValues.AcceptedByTitle}
                onChange={(e) => {
                  objValues.AcceptedByTitle = e.target["value"];
                  setRenderObjValue(true);
                }}
              />
              <DatePicker
                formatDate={(date: Date): string => {
                  return (
                    date.getMonth() +
                    1 +
                    "/" +
                    date.getDate() +
                    "/" +
                    date.getFullYear()
                  );
                }}
                styles={halfWidthInput}
                label="Date"
                onSelectDate={(date) => {
                  objValues.AcceptedByDate = date;
                  setRenderObjValue(true);
                }}
              />
            </div>
          </div>
        </div>
      </div>
      {/* Section */}
      <div className={styles.section}>
        <div
          style={{
            display: "flex",
            flexDirection: "column",
            alignItems: "center",
            width: "100%",
          }}
        >
          <h3
            style={{ textAlign: "center", color: myTheme.palette.themePrimary }}
          >
            LYNXSPRING Schedules of Invoice and Statement of work
          </h3>
          <TextField
            styles={halfWidthInput}
            label="Statement of work:"
            value={objValues.StatementOfWork}
            onChange={(e) => {
              objValues.StatementOfWork = e.target["value"];
              setRenderObjValue(true);
              console.log(objValues);
            }}
          />
        </div>
        {/* Milestone Section */}
        <table className={styles.mileStone}>
          <thead>
            <tr>
              <th></th>
              <th>Description of Deliverables</th>
              <th>Estimated Start and End Date</th>
              <th>Amount</th>
              <th></th>
            </tr>
          </thead>
          <tbody>
            {milestones.map((milestone) => {
              return (
                <tr key={milestone.id}>
                  <td>{milestone.title}</td>
                  <td>
                    <TextField
                      // key={milestone.id}
                      id={`${milestone.id}`}
                      // styles={halfWidthInput}
                      value={milestone.description}
                      multiline
                      autoAdjustHeight
                      onChange={(e) => {
                        arrMilestones.filter(
                          (item) => item.id == +e.target["id"]
                        )[0].description = e.target["value"];
                        setFetchTable(true);
                      }}
                    />
                  </td>
                  <td>
                    <div
                      style={{
                        display: "flex",
                        alignItems: "center",
                        justifyContent: "center",
                      }}
                    >
                      <DatePicker
                        formatDate={(date: Date): string => {
                          return (
                            date.getMonth() +
                            1 +
                            "/" +
                            date.getDate() +
                            "/" +
                            date.getFullYear()
                          );
                        }}
                        // key={milestone.id}
                        id={`${milestone.id}`}
                        styles={{ root: { width: 150 } }}
                        firstDayOfWeek={firstDayOfWeek}
                        placeholder="Select a date..."
                        ariaLabel="Select a date"
                        value={new Date(milestone.startDate)}
                        onSelectDate={(date) => {
                          arrMilestones.filter(
                            (item) => item.id == milestone.id
                          )[0].startDate = date.toLocaleString();
                          setFetchTable(true);
                        }}
                      />
                      <label
                        style={{
                          display: "block",
                          width: "10px",
                          margin: "10px",
                        }}
                      >
                        to
                      </label>
                      <DatePicker
                        formatDate={(date: Date): string => {
                          return (
                            date.getMonth() +
                            1 +
                            "/" +
                            date.getDate() +
                            "/" +
                            date.getFullYear()
                          );
                        }}
                        styles={{ root: { width: 150 } }}
                        firstDayOfWeek={firstDayOfWeek}
                        placeholder="Select a date..."
                        ariaLabel="Select a date"
                        // key={milestone.id}
                        id={`${milestone.id}`}
                        value={new Date(milestone.endDate)}
                        onSelectDate={(date) => {
                          arrMilestones.filter(
                            (item) => item.id == milestone.id
                          )[0].endDate = date.toLocaleString();
                          setFetchTable(true);
                        }}
                      />
                    </div>
                  </td>
                  <td>
                    <TextField
                      value={milestone.amount}
                      // key={milestone.id}
                      id={`${milestone.id}`}
                      onChange={(e) => {
                        arrMilestones.filter(
                          (item) => item.id == e.target["id"]
                        )[0].amount = e.target["value"];
                        setFetchTable(true);
                      }}
                    />
                  </td>
                  <td>
                    <Icon
                      iconName="delete"
                      id={`${milestone.id}`}
                      styles={{
                        root: {
                          fontSize: 20,
                          fontWeight: 400,
                          color: "#D71E2B",
                          cursor: "pointer",
                        },
                      }}
                      onClick={(e) => {
                        arrMilestones = arrMilestones.filter(
                          (arritem) => arritem.id !== milestone.id
                        );
                        setFetchTable(true);
                      }}
                    />
                  </td>
                </tr>
              );
            })}
          </tbody>
        </table>
        <div
          style={{
            margin: "1rem auto 2rem auto",
            width: "90%",
            display: "flex",
            justifyContent: "flex-end",
          }}
        >
          <DefaultButton
            text="Add"
            iconProps={addIcon}
            styles={{ root: { marginLeft: "auto", textAlign: "right" } }}
            onClick={() => {
              addMilestoneHandler();
              setMilestones(arrMilestones);
            }}
          />
        </div>
        {/* Milestone Section */}
        <div className={styles.submitSection}>
          <PrimaryButton
            text="Submit"
            style={{ marginRight: "0.5rem" }}
            onClick={submitBtnHandler}
          />
          <DefaultButton text="Cancel" onClick={() => history.back()} />
        </div>
      </div>

      {/* Section */}
    </div>
  );
};

export default NWFQuoteForm;
