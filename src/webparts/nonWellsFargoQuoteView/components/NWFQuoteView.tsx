import * as React from 'react'
import styles from './NonWellsFargoQuoteView.module.scss'
import {
  loadTheme,
  createTheme,
  Theme,
  DayOfWeek,
  IChoiceGroupOption,
  PivotItem,
  Checkbox,
  Icon,
  ChoiceGroup,
  DatePicker,
  DefaultButton,
  Pivot,
  PrimaryButton,
  TextField,
  ThemeProvider,
} from '@fluentui/react'
import { Spinner, SpinnerSize } from '@fluentui/react/lib/Spinner'
import { useEffect, useState, useRef } from 'react'
import * as $ from 'jquery'
import * as moment from 'moment'

let formID = 0
const paramsString = window.location.href.split('?')[1].toLowerCase()
const searchParams = new URLSearchParams(paramsString)
searchParams.has('formid') ? (formID = Number(searchParams.get('formid'))) : ''
const choiceGroupStyles = {
  flexContainer: {
    display: 'flex',
    label: {
      marginRight: '1rem',
    },
  },
}
let arrnwParts = [
  {
    isSelected: false,
    PartNo: '',
    PartName: '',
    PartDescription: '',
    ListPrice: 0,
    itemFor: '',
    NetPrice: 0,
    Note: '',
    PartDescriptionSort: '',
    id: 0,
    quantity: '',
  },
]
let arrMilestones = []
let objValues = {
  ProjectNo: '',
  Date: new Date(),
  ConsultantName: '',
  ConsultantCity: '',
  ConsultantContactNo: '',
  ConsultantPinCode: '',
  ConsultantAddress: '',
  ClientName: '',
  ClientCity: '',
  ClientContactNo: '',
  ClientPinCode: '',
  ClientAddress: '',
  SentVai: '',
  ProjectDescription: '',
  TypesOfProposal: '',
  Multiplier: '',
  ProposedBy: '',
  ProposedName: '',
  ProposedTitle: '',
  ProposedDate: new Date(),
  AcceptedBy: '',
  AcceptedByName: '',
  AcceptedByDate: new Date(),
  AcceptedByTitle: '',
  StatementOfWork: '',
  Services: '',
  AcceptedForClient: '',
  CompanyName: '',
  Title: '',
  Requestor: '',
}
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
}
const myTheme = createTheme({
  palette: {
    themePrimary: '#004fa2',
    themeLighterAlt: '#f1f6fb',
    themeLighter: '#cadcf0',
    themeLight: '#9fc0e3',
    themeTertiary: '#508ac8',
    themeSecondary: '#155fae',
    themeDarkAlt: '#004793',
    themeDark: '#003c7c',
    themeDarker: '#002c5b',
    neutralLighterAlt: '#faf9f8',
    neutralLighter: '#f3f2f1',
    neutralLight: '#edebe9',
    neutralQuaternaryAlt: '#e1dfdd',
    neutralQuaternary: '#d0d0d0',
    neutralTertiaryAlt: '#c8c6c4',
    neutralTertiary: '#a19f9d',
    neutralSecondary: '#605e5c',
    neutralPrimaryAlt: '#3b3a39',
    neutralPrimary: '#323130',
    neutralDark: '#201f1e',
    black: '#000000',
    white: '#ffffff',
  },
})

let arrSentViaOptions = []
let arrTypesOfProposal = []
const NWFQuoteView = (props) => {
  let siteURLForFile = props.context.pageContext.web.absoluteUrl
  loadTheme(myTheme)
  const [Loader, setLoader] = useState(false)
  const [firstDayOfWeek, setFirstDayOfWeek] = useState(DayOfWeek.Sunday)
  const [selectedKey, setSelectedKey] = useState(1)
  const [milestones, setMilestones] = useState(arrMilestones)
  const [fetchTable, setFetchTable] = useState(true)
  const [partsDetails, setPartsDetails] = useState(arrnwParts)
  const [fetchPartsTable, setFetchPartsTable] = useState(true)
  const [objToPost, setObjToPost] = useState(objValues)
  const [renderObjValue, setRenderObjValue] = useState(true)
  const [selectedServices, setSelectedServices] = useState(objSelectedServices)
  const [fetchSelectedServices, setFetchSelectedServices] = useState(true)
  const [sentViaOptions, setSentViaOptions] = useState(arrSentViaOptions)
  const [typesOfProposalOptions, setTypesOfProposalOptions] = useState(
    arrTypesOfProposal,
  )

  const halfWidthInput = {
    root: { width: 300, margin: '0 1rem 0.5rem 0' },
  }
  useEffect(() => {
    props.spcontext.web.lists
      .getByTitle('GeneralQuoteRequestList')
      .fields.filter("EntityPropertyName eq 'SentVia'")
      .get()
      .then((SentVia) => {
        SentVia[0].Choices.forEach((option) => {
          arrSentViaOptions.push({
            key: option,
            text: option,
          })
        })
      })
    setSentViaOptions(arrSentViaOptions)
    props.spcontext.web.lists
      .getByTitle('GeneralQuoteRequestList')
      .fields.filter("EntityPropertyName eq 'TypesOfProposal'")
      .get()
      .then((types) => {
        types[0].Choices.forEach((option) => {
          arrTypesOfProposal.push({
            key: option,
            text: option,
          })
        })
        setTypesOfProposalOptions(arrTypesOfProposal)
      })
    props.spcontext.web.lists
      .getByTitle('GeneralQuoteRequestList')
      .items.getById(formID)
      .select('*', 'Author/Title')
      .expand('Author')
      .get()
      .then((li: any) => {
        objValues = {
          ProjectNo: li.ProjectNo,
          Date: new Date(li.Date),
          ConsultantName: li.ConsultantName,
          ConsultantCity: li.ConsultantCity,
          ConsultantContactNo: li.ConsultantContactNo,
          ConsultantPinCode: li.ConsultantPinCode,
          ConsultantAddress: li.ConsultantAddress,
          ClientName: li.ClientName,
          ClientCity: li.ClientCity,
          ClientContactNo: li.ClientContactNo,
          ClientPinCode: li.ClientPinCode,
          ClientAddress: li.ClientAddress,
          SentVai: li.SentVia,
          ProjectDescription: li.ProjectDescription,
          TypesOfProposal: li.TypesOfProposal,
          Multiplier: li.Multiplier ? li.Multiplier : 1,
          ProposedBy: li.ProposedBy,
          ProposedName: li.ProposedName,
          ProposedTitle: li.ProposedTitle,
          ProposedDate: new Date(li.ProposedDate),
          AcceptedBy: li.AcceptedBy,
          AcceptedByName: li.AcceptedByName,
          AcceptedByDate: new Date(li.AcceptedByDate),
          AcceptedByTitle: li.AcceptedByTitle,
          StatementOfWork: li.StatementOfWork,
          Services: li.Services,
          AcceptedForClient: li.AcceptedForClient,
          CompanyName: li.CompanyName,
          Title: li.Title,
          Requestor: li.Author.Title,
        }
        console.log(li)
        arrnwParts =
          li.ProposedServicesFees == ''
            ? []
            : JSON.parse(li.ProposedServicesFees)
        console.log(JSON.parse(li.ProposedServicesFees))

        arrMilestones =
          li.Milestones && li.Milestones != '' ? JSON.parse(li.Milestones) : []
        console.log(arrMilestones)
        setRenderObjValue(true)
        setFetchPartsTable(true)
        setFetchTable(true)
      })
      .catch((error) => console.log(error))
  }, [])
  useEffect(() => {
    if (fetchTable) {
      setMilestones([...arrMilestones])
      setFetchTable(false)
    }
  }, [fetchTable])
  useEffect(() => {
    if (fetchPartsTable) {
      arrnwParts && arrnwParts.length > 0
        ? setPartsDetails([...arrnwParts])
        : ''
      setFetchPartsTable(false)
    }
  }, [fetchPartsTable])
  useEffect(() => {
    if (fetchSelectedServices) {
      setSelectedServices(objSelectedServices)
      setFetchSelectedServices(false)
    }
  }, [fetchSelectedServices])
  useEffect(() => {
    if (renderObjValue) {
      setObjToPost(objValues)
      setRenderObjValue(false)
    }
  }, [renderObjValue])
  return (
    <div style={{ backgroundColor: '#F2F2F2', padding: '1rem 2rem' }}>
      {Loader && (
        <Spinner
          label="Loading items..."
          size={SpinnerSize.large}
          style={{
            width: '100vw',
            height: '100vh',
            position: 'fixed',
            top: 0,
            left: 0,
            backgroundColor: '#fff',
            zIndex: 10000,
          }}
        />
      )}
      <div className={styles.formHeader}>
        <Icon
          iconName="NavigateBack"
          styles={{
            root: {
              fontSize: 30,
              fontWeight: 600,
              color: myTheme.palette.themePrimary,
              marginRight: '1rem',
              cursor: 'Pointer',
            },
          }}
          onClick={() => {
            history.back()
          }}
        />
        <h2
          style={{
            textAlign: 'center',
            color: myTheme.palette.themePrimary,
            width: '100%',
          }}
        >
          Proposal of Services
        </h2>
      </div>
      <div
        style={{
          display: 'flex',
          justifyContent: 'flex-end',
          marginBottom: '1rem',
        }}
        className={styles.section}
      >
        <PrimaryButton
          text="Export docx"
          style={{ marginRight: '1rem' }}
          onClick={() =>
            downloadFile(
              'https://nonwellsfargo.azurewebsites.net/api/docx',
              'demo.docx',
              'docx',
            )
          }
        />
        <PrimaryButton
          text="Export Excel"
          onClick={() =>
            downloadFile(
              'https://nonwellsfargo.azurewebsites.net/api/excel',
              'demo.xlsx',
              'excel',
            )
          }
        />
      </div>
      <div className={`${styles.projectDetails} ${styles.section}`}>
        <TextField
          label="Project No"
          styles={halfWidthInput}
          value={objToPost.ProjectNo}
          disabled={true}
        />
        <DatePicker
          formatDate={(date: Date): string => {
            return (
              date.getMonth() +
              1 +
              '/' +
              date.getDate() +
              '/' +
              date.getFullYear()
            )
          }}
          styles={halfWidthInput}
          label="Date"
          firstDayOfWeek={firstDayOfWeek}
          placeholder="Select a date..."
          ariaLabel="Select a date"
          disabled={true}
          value={objToPost.Date}
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
              value={objToPost.ConsultantName}
              disabled={true}
            />
            <TextField
              label="Contact No"
              styles={halfWidthInput}
              value={objToPost.ConsultantContactNo}
              disabled={true}
            />
          </div>
          <div>
            <TextField
              label="City"
              styles={halfWidthInput}
              value={objToPost.ConsultantCity}
              disabled={true}
            />
            <TextField
              label="Pin Code"
              styles={halfWidthInput}
              value={objToPost.ClientPinCode}
              disabled={true}
            />
          </div>
          <div>
            <TextField
              styles={halfWidthInput}
              label="Address"
              multiline
              resizable={false}
              value={objToPost.ConsultantAddress}
              disabled={true}
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
              value={objToPost.ClientName}
              disabled={true}
            />
            <TextField
              label="Contact No"
              styles={halfWidthInput}
              value={objToPost.ClientContactNo}
              disabled={true}
            />
          </div>
          <div>
            <TextField
              label="City"
              styles={halfWidthInput}
              value={objToPost.ClientCity}
              disabled={true}
            />
            <TextField
              label="Pin Code"
              styles={halfWidthInput}
              value={objToPost.ClientPinCode}
              disabled={true}
            />
          </div>
          <div>
            <TextField
              styles={halfWidthInput}
              label="Address"
              multiline
              resizable={false}
              value={objToPost.ClientAddress}
              disabled={true}
            />
          </div>
        </div>
      </div>
      {/* Section */}
      {/* Section */}
      <div className={styles.section}>
        <div>
          <ChoiceGroup
            styles={choiceGroupStyles}
            options={sentViaOptions}
            label="Sent Via:"
            disabled={true}
            selectedKey={objToPost.SentVai}
          />
        </div>
        <div style={{ display: 'flex', justifyContent: 'space-between' }}>
          <TextField
            styles={halfWidthInput}
            label="Services"
            multiline
            resizable={false}
            value={objValues.Services}
            disabled={true}
          />
          <TextField
            styles={halfWidthInput}
            label="Project Description"
            multiline
            resizable={false}
            value={objValues.ProjectDescription}
            disabled={true}
          />
        </div>
        <div>
          <ChoiceGroup
            styles={choiceGroupStyles}
            options={typesOfProposalOptions}
            label="Types of proposal"
            disabled={true}
            selectedKey={objToPost.TypesOfProposal}
          />
        </div>
      </div>
      {/* Section */}
      <div className={`${styles.section} ${styles.sectionPovit}`}>
        {/*  Pivot Section Start */}
        <div
          style={{
            display: 'flex',
            justifyContent: 'space-between',
            alignItems: 'center',
          }}
        >
          <h3 style={{ color: myTheme.palette.themePrimary }}>
            Proposed Services and Fee
          </h3>
          <TextField
            type="number"
            value={objValues.Multiplier}
            styles={{
              root: {
                width: 100,
              },
            }}
            label="Your Multiplier"
            disabled={true}
          />
        </div>
        {/* Pivot */}
        <ThemeProvider dir="ltr">
          {partsDetails ? (
            <table>
              <thead>
                <tr>
                  <th></th>
                  <th>Part Name</th>
                  <th>Part Description</th>
                  <th>List Price</th>
                  <th>Quantity</th>
                  <th>Net Price</th>
                  <th>Note</th>
                </tr>
              </thead>
              <tbody>
                {partsDetails.filter((part) => part.isSelected).length > 0 ? (
                  partsDetails.map((part) => {
                    if (part.isSelected) {
                      return (
                        <tr
                          style={{
                            backgroundColor: part.isSelected
                              ? '#eef4fa'
                              : '#ffffff',
                          }}
                        >
                          <td>
                            {' '}
                            <Checkbox
                              checked={part.isSelected ? true : false}
                            />
                          </td>
                          <td>
                            <div>{part.PartNo}</div>
                            <div>{part.PartName}</div>
                          </td>
                          <td>
                            <label title={part.PartDescription}>
                              {part.PartDescriptionSort}
                            </label>
                          </td>
                          <td style={{ textAlign: 'center' }}>
                            {part.ListPrice}
                          </td>
                          <td style={{ textAlign: 'center' }}>
                            {part.quantity}
                          </td>
                          <td style={{ textAlign: 'center' }}>
                            {part.NetPrice}
                          </td>

                          <td>{part.Note}</td>
                        </tr>
                      )
                    }
                  })
                ) : (
                  <tr>
                    <td
                      colSpan={6}
                      style={{
                        textAlign: 'center',
                        fontWeight: 'bold',
                        padding: '1rem',
                      }}
                    >
                      No Data Found
                    </td>
                  </tr>
                )}
              </tbody>
            </table>
          ) : (
            ''
          )}
        </ThemeProvider>
        {/* Pivot */}
        {/*  Pivot Section End */}
      </div>
      {/* Section */}
      <div className={styles.section}>
        <div style={{ display: 'flex', justifyContent: 'space-between' }}>
          <div>
            <h3 style={{ color: myTheme.palette.themePrimary }}>
              Proposed for Consultant: Lynxspring, Inc
            </h3>
            <TextField
              styles={halfWidthInput}
              label="By"
              value={objValues.ProposedBy}
              disabled={true}
            />
            <TextField
              styles={halfWidthInput}
              label="Name (Printed)"
              value={objValues.ProposedName}
              disabled={true}
            />
            <TextField
              styles={halfWidthInput}
              label="Title"
              value={objValues.ProposedTitle}
              disabled={true}
            />
            <DatePicker
              formatDate={(date: Date): string => {
                return (
                  date.getMonth() +
                  1 +
                  '/' +
                  date.getDate() +
                  '/' +
                  date.getFullYear()
                )
              }}
              value={objValues.ProposedDate}
              styles={halfWidthInput}
              label="Date"
              disabled={true}
            />
          </div>
          <div>
            <div>
              <div style={{ display: 'flex', alignItems: 'center' }}>
                <h3 style={{ color: myTheme.palette.themePrimary }}>
                  Accepted for Client:
                </h3>
                <TextField
                  styles={halfWidthInput}
                  placeholder="Client Name"
                  disabled={true}
                  value={objValues.AcceptedForClient}
                />
              </div>
              <TextField
                styles={halfWidthInput}
                label="By"
                value={objValues.AcceptedBy}
                disabled={true}
              />
              <TextField
                styles={halfWidthInput}
                label="Name (Printed)"
                value={objValues.AcceptedByName}
                disabled={true}
              />
              <TextField
                styles={halfWidthInput}
                label="Title"
                value={objValues.AcceptedByTitle}
                disabled={true}
              />
              <DatePicker
                formatDate={(date: Date): string => {
                  return (
                    date.getMonth() +
                    1 +
                    '/' +
                    date.getDate() +
                    '/' +
                    date.getFullYear()
                  )
                }}
                value={objValues.AcceptedByDate}
                styles={halfWidthInput}
                label="Date"
                disabled={true}
              />
            </div>
          </div>
        </div>
      </div>
      {/* Section */}
      <div className={styles.section}>
        <div
          style={{
            display: 'flex',
            flexDirection: 'column',
            alignItems: 'center',
            width: '100%',
          }}
        >
          <h3
            style={{ textAlign: 'center', color: myTheme.palette.themePrimary }}
          >
            LYNXSPRING Schedules of Invoice and Statement of work
          </h3>
          <TextField
            styles={halfWidthInput}
            label="Statement of work:"
            value={objValues.StatementOfWork}
            disabled={true}
          />
        </div>
        {/* Milestone Section */}
        {milestones.length > 0 && (
          <table className={styles.mileStoneTable}>
            <thead>
              <tr>
                <th></th>
                <th>Description of Deliverables</th>
                <th>Estimated Start and End Date</th>
                <th>Amount</th>
              </tr>
            </thead>
            <tbody>
              {milestones.length > 0 &&
                milestones.map((milestone) => {
                  return (
                    <tr>
                      <td>{milestone.title}</td>
                      <td>
                        <TextField
                          key={milestone.id}
                          id={`${milestone.id}`}
                          multiline
                          value={milestone.description}
                          disabled={true}
                        />
                      </td>
                      <td>
                        <div
                          style={{
                            display: 'flex',
                            alignItems: 'center',
                            justifyContent: 'center',
                          }}
                        >
                          <DatePicker
                            formatDate={(date: Date): string => {
                              return (
                                date.getMonth() +
                                1 +
                                '/' +
                                date.getDate() +
                                '/' +
                                date.getFullYear()
                              )
                            }}
                            key={milestone.id}
                            id={`${milestone.id}`}
                            styles={{ root: { width: 150 } }}
                            firstDayOfWeek={firstDayOfWeek}
                            placeholder="Select a date..."
                            ariaLabel="Select a date"
                            disabled={true}
                            value={new Date(milestone.startDate)}
                          />
                          to
                          <DatePicker
                            formatDate={(date: Date): string => {
                              return (
                                date.getMonth() +
                                1 +
                                '/' +
                                date.getDate() +
                                '/' +
                                date.getFullYear()
                              )
                            }}
                            styles={{ root: { width: 150 } }}
                            firstDayOfWeek={firstDayOfWeek}
                            placeholder="Select a date..."
                            ariaLabel="Select a date"
                            key={milestone.id}
                            id={`${milestone.id}`}
                            disabled={true}
                            value={new Date(milestone.endDate)}
                          />
                        </div>
                      </td>
                      <td>
                        <TextField
                          value={milestone.amount}
                          key={milestone.id}
                          id={`${milestone.id}`}
                          disabled={true}
                        />
                      </td>
                    </tr>
                  )
                })}
            </tbody>
          </table>
        )}
        {/* Milestone Section */}
        <div className={styles.submitSection}>
          <DefaultButton text="Back" onClick={() => history.back()} />
        </div>
      </div>

      {/* Section */}
    </div>
  )

  function getExcelJsondata() {
    var excelJsonData = {
      quoteTo: objValues.CompanyName,
      name: objValues.Requestor,
      project: objValues.Title,
      date: moment(objValues.Date).format('DD/MM/YYYY'),
      expires: '30 days from the date generated',
      paymentTerms: '',
      preparedBy: '',
      multiplier: objValues.Multiplier,
      total: '$ ',

      installColumns: [
        {
          coloumn1: '1',
          coloumn2: '1',
          coloumn3: 'Lynxspring',
          coloumn4: 'JENE-EG414-VAV',
          coloumn5: 'JENEsys EDGE 414 Programmable VAV',
          coloumn6: '$ 328.75',
          coloumn7: '$ 328.75',
          coloumn8: '$ 328.75',
        },
        {
          coloumn1: '2',
          coloumn2: '1',
          coloumn3: 'Lynxspring',
          coloumn4: 'xxxxx',
          coloumn5: 'xxxxx',
          coloumn6: '$  0.00',
          coloumn7: '$  0.00',
          coloumn8: '$  0.00',
        },
        {
          coloumn1: '3',
          coloumn2: '1',
          coloumn3: 'Lynxspring',
          coloumn4: 'xxxxx',
          coloumn5: 'xxxxx',
          coloumn6: '$  0.00',
          coloumn7: '$  0.00',
          coloumn8: '$  0.00',
        },
        {
          coloumn1: '4',
          coloumn2: '1',
          coloumn3: 'Lynxspring',
          coloumn4: 'xxxxx',
          coloumn5: 'xxxxx',
          coloumn6: '$  0.00',
          coloumn7: '$  0.00',
          coloumn8: '$  0.00',
        },
      ],

      projectName: objValues.Title,
      clientName: objValues.ClientName,
      address: objValues.ClientAddress,
      cityState: objValues.ClientPinCode,
      address2: objValues.ClientCity,
      email: objValues.SentVai,
      services: objValues.Services,
      hourlySum: objValues.TypesOfProposal,
      projectDescription: objValues.ProjectDescription,
      proposedBy: objValues.ProposedBy,
      proposedName: objValues.ProposedName,
      proposedTitle: objValues.ProposedTitle,
      proposedDate: moment(objValues.ProposedDate).format('DD/MM/YYYY'),
      acceptedBy: objValues.AcceptedBy,
      acceptedName: objValues.AcceptedByName,
      acceptedTitle: objValues.AcceptedByTitle,
      acceptedDateSigned: moment(objValues.AcceptedByDate).format('DD/MM/YYYY'),
      amount1: '$XXXXXX',
      amount2: '$XXXXXX',
      amount3: '$XXXXXX',
      amount4: '$XXXXXX',
      totallast: '$XXXXXX',
    }

    var installationTable = []
    $.each(partsDetails, function (key, val) {
      installationTable.push({
        coloumn1: key,
        coloumn2: '1',
        coloumn3: 'Lynxspring',
        coloumn4: val.PartDescriptionSort,
        coloumn5: val.PartDescription,
        coloumn6: val.ListPrice,
        coloumn7: val.NetPrice,
        coloumn8: val.NetPrice,
      })
    })

    var sumofnetprice = 0
    $.each(installationTable, function (key, val) {
      sumofnetprice = sumofnetprice + val.coloumn8
    })

    excelJsonData.installColumns = installationTable
    excelJsonData.total = '$ ' + sumofnetprice.toString()
    var milestoneamounts = []

    var totalmilestoneamount = '$ XXXXXX'
    $.each(milestones, function (key, val) {
      milestoneamounts.push(val.amount)
    })

    if (milestoneamounts.length == 4) {
      excelJsonData.amount1 = '$' + milestoneamounts[0]
      excelJsonData.amount2 = '$' + milestoneamounts[1]
      excelJsonData.amount3 = '$' + milestoneamounts[2]
      excelJsonData.amount4 = '$' + milestoneamounts[3]
    }

    if (milestoneamounts.length == 3) {
      excelJsonData.amount1 = '$' + milestoneamounts[0]
      excelJsonData.amount2 = '$' + milestoneamounts[1]
      excelJsonData.amount3 = '$' + milestoneamounts[2]
    }

    if (milestoneamounts.length == 2) {
      excelJsonData.amount1 = '$' + milestoneamounts[0]
      excelJsonData.amount2 = '$' + milestoneamounts[1]
    }

    if (milestoneamounts.length == 1) {
      excelJsonData.amount1 = '$' + milestoneamounts[0]
    }

    var sumofmilestone = 0

    $.each(milestoneamounts, function (key, val) {
      sumofmilestone = Number(sumofmilestone) + Number(val)
    })

    excelJsonData.totallast = '$ ' + sumofmilestone.toString()

    return excelJsonData
  }

  async function downloadFile(URL, fileName, filetype) {
    setLoader(true)
    var jsonData = getExcelJsondata()
    console.log(jsonData)
    $.ajax({
      type: 'POST',
      cache: false,
      url: URL,
      data: jsonData,
      xhrFields: {
        responseType: 'arraybuffer',
      },
    })
      .done(async function (data, status, xmlHeaderRequest) {
        var blob = new Blob([data], {
          type: xmlHeaderRequest.getResponseHeader('Content-Type'),
        })

        let file = blob
        await props.spcontext.web
          .getFolderByServerRelativePath('Shared Documents')
          .files.addUsingPath(fileName, file, { Overwrite: true })
          .then(function (data) {
            //alert("success");
            console.log(data)
            setLoader(false)
            var link = document.createElement('a')

            if (filetype == 'excel')
              link.setAttribute(
                'href',
                data.data.ServerRelativeUrl + '?download=1',
              )
            else
              link.setAttribute(
                'href',
                siteURLForFile +
                  '/_layouts/download.aspx?SourceUrl=' +
                  siteURLForFile +
                  '/Shared%20Documents/' +
                  data.data.Name,
              )

            link.setAttribute('target', '_blank')
            link.setAttribute('download', fileName)
            //link.style = "visibility:hidden";
            document.body.appendChild(link)
            link.click()
            setTimeout(function () {
              document.body.removeChild(link)
            }, 500)
          })
          .catch(function () {
            setLoader(false)
            alert('Error while downloading File.Please contact admin')
          })
      })
      .catch(function (jqXHR, textStatus, errorThrown) {
        setLoader(false)
        alert('Error while downloading File.Please contact admin')
        console.log('Response from File API Failed')
        console.log(JSON.stringify(jqXHR))
        console.log(JSON.stringify(textStatus))
        console.log(JSON.stringify(errorThrown))
      })
  }
}
export default NWFQuoteView
