import * as React from 'react'
import { Fragment } from 'react'
import { useState, useEffect, useRef } from 'react'
import { Icon } from '@fluentui/react/lib/Icon'
import { TextField, MaskedTextField } from '@fluentui/react/lib/TextField'
import styles from './WellsFargoQuoteForm.module.scss'
import { DisplayMode } from '@microsoft/sp-core-library'
import { DefaultButton, PrimaryButton } from '@fluentui/react/lib/Button'
import {
  DatePicker,
  DayOfWeek,
  Dropdown,
  IDropdownOption,
  mergeStyles,
  defaultDatePickerStrings,
  Checkbox,
  ThemeProvider,
  IPersonaProps,
} from '@fluentui/react'
import {
  ContextualMenu,
  IContextualMenuProps,
  IIconProps,
} from '@fluentui/react'
import { loadTheme, createTheme, Theme } from '@fluentui/react'
import {
  PeoplePicker,
  PrincipalType,
} from '@pnp/spfx-controls-react/lib/PeoplePicker'

let formID = 0
const paramsString = window.location.href.split('?')[1].toLowerCase()
const searchParams = new URLSearchParams(paramsString)
searchParams.has('formid') ? (formID = Number(searchParams.get('formid'))) : ''
let ProjectManagerUser = []
const fullWidthInput = {
  root: { width: '70%', marginBottom: '0.5rem' },
}
const halfWidthInput = {
  root: { width: '50%', margin: '0 1rem 0.5rem 0' },
}
let arrTemplateMaster = []
let arrTemplateSelected = []
let arrInstall = [
  {
    id: 100000,
    detailedDescription: '',
    manufacture: 'LynxSpring',
    modelNo: '',
    serialNo: '',
    quantity: '',
    eachPrice: '',
    eachMarkup: '',
    totalProduct: '',
    taxable: false,
    people: '',
    hoursPerPerson: '',
    hourlyBillingRate: '',
    unionRate: false,
    totalLabour: '',
    labourTaxable: false,
    grandTotalProduct: '',
    templateOf: '',
  },
]
let objProjInfo = {
  projNoInput: '',
  projManagerInput: '',
  BENoInput: '',
  ProjNameInput: '',
  DeliveryAddInput: '',
  projAreaInput: '',
  EstimateStartDateInput: new Date(),
  EstimateEndDateInput: new Date(),
  InstallTotalProduct: '',
  InstallTotalPeople: '',
  InstallTotalHoursPerPerson: '',
  InstallSecondTotalProduct: '',
}
let objVendorInfo = {
  companyNameInput: '',
  wfVendorNoInput: '',
  remitAddInput: '',
  proposalNoInput: '',
  cityStateZipInput: '',
  wfContractNoInput: '',
  contactNameInput: '',
  changeOrderInput: '',
  phoneNoInput: '',
  changeOrderPOInput: '',
  cellInput: '',
  emailIdInput: '',
  scopeOfWork: '',
  assumptionsAndClarifications: '',
}
let objTaxes = {
  Product: {
    PreTax: '',
    Tax: '',
    Total: '',
  },
  Labour: {
    PreTax: '',
    Tax: '',
    Total: '',
  },
  ProductSubTotal: {
    PreTax: '',
    Tax: '',
    Total: '',
  },
  DemoProduct: {
    PreTax: '',
    Tax: '',
    Total: '',
  },
  DemoLabour: {
    PreTax: '',
    Tax: '',
    Total: '',
  },
  DemoSubTotal: {
    PreTax: '',
    Tax: '',
    Total: '',
  },
  Freight: {
    PreTax: '',
    Tax: '',
    Total: '',
  },
  SpringHandling: {
    PreTax: '',
    Tax: '',
    Total: '',
  },
  ProfitOH: {
    PreTax: '',
    Tax: '',
    Total: '',
  },
  Insurance: {
    PreTax: '',
    Tax: '',
    Total: '',
  },
  Total: {
    PreTax: '',
    Tax: '',
    Total: '',
  },
}
let MasterInstallationOptions
const addIcon: IIconProps = { iconName: 'Add' }

const redTheme = createTheme({
  palette: {
    themePrimary: '#d71e2b',
    themeLighterAlt: '#fdf5f5',
    themeLighter: '#f8d6d9',
    themeLight: '#f3b4b8',
    themeTertiary: '#e77078',
    themeSecondary: '#db3540',
    themeDarkAlt: '#c11b26',
    themeDark: '#a31720',
    themeDarker: '#781118',
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

const WFQuoteForm = (props) => {
  const [projectInfo, setProjectInfo] = useState(objProjInfo)
  const [renderProjInfo, setRenderProjInfo] = useState(false)
  const [vendorInfo, setVendorInfo] = useState(objVendorInfo)
  const [renderVendorInfo, setRenderVendorInfo] = useState(false)
  const [orderNo, setOrderNo] = useState('')
  const [installationtable, setInstallationTable] = useState([])
  const [fetchingTable, setFetchingTable] = useState(true)
  const [firstDayOfWeek, setFirstDayOfWeek] = useState(DayOfWeek.Sunday)
  const [taxesInfo, setTaxesInfo] = useState(objTaxes)
  const [renderTaxesInfo, setRenderTaxesInfo] = useState(false)
  const [templateOptions, setTemplateOptions] = useState([])
  const [tempApplied, setTempApplied] = useState(false)
  const [isShowDescription, setIsShowDescription] = useState(false)
  // Call
  useEffect(() => {
    props.spcontext.web.lists
      .getByTitle('WFQuoteRequestList')
      .items.select('*')
      .getById(formID)
      .get()
      .then((listItem: any) => {
        setOrderNo(listItem.OrderNo)
        objProjInfo = {
          projNoInput: listItem.OrderNo,
          projManagerInput: listItem.ManagerName,
          BENoInput: listItem.BENumber,
          ProjNameInput: listItem.Title,
          DeliveryAddInput: listItem.ShippingAddress,
          projAreaInput: listItem.ProjectArea,
          EstimateStartDateInput: new Date(),
          EstimateEndDateInput: new Date(),
          InstallTotalProduct: listItem.InstallTotalProduct,
          InstallTotalPeople: listItem.InstallTotalPeople,
          InstallTotalHoursPerPerson: listItem.InstallTotalHoursPerPerson,
          InstallSecondTotalProduct: listItem.InstallSecondTotalProduct,
        }
        setProjectInfo(objProjInfo)
      })
      .then(() => {
        props.spcontext.web.lists
          .getByTitle('InstallationTemplates')
          .items.select('*')
          .get()
          .then((installationDetails) => {
            arrTemplateMaster = []
            arrTemplateMaster = installationDetails.map(
              (installationItem, i) => {
                return {
                  id: i,
                  detailedDescription: installationItem.DetailedDescription,
                  manufacture: installationItem.Title,
                  modelNo: installationItem.Model,
                  serialNo: installationItem.Serial,
                  quantity: installationItem.Quantity,
                  eachPrice: installationItem.EachPrice,
                  eachMarkup: installationItem.EachMarkup,
                  totalProduct: installationItem.TotalProduct,
                  taxable: installationItem.Taxable,
                  people: installationItem.People,
                  hoursPerPerson: installationItem.HoursPerPerson,
                  hourlyBillingRate: installationItem.HourlyBillRate,
                  unionRate: installationItem.UnionRate,
                  totalLabour: installationItem.TotalLabour,
                  labourTaxable: installationItem.LabourTaxable,
                  grandTotalProduct: installationItem.TotalProducts,
                  templateOf: installationItem.templateOf,
                }
              },
            )
            let masterOptions = installationDetails.map((installItem) => {
              return installItem.templateOf
            })
            MasterInstallationOptions = masterOptions
              .filter((c, index) => {
                return masterOptions.indexOf(c) === index
              })
              .map((option) => {
                return { key: option, text: option }
              })
            setTemplateOptions(MasterInstallationOptions)
          })
      })
      .catch((error) => console.log(error))
  }, [])

  useEffect(() => {
    if (renderProjInfo) {
      console.log(objProjInfo)
      setProjectInfo(objProjInfo)
      setRenderProjInfo(false)
    }
  }, [renderProjInfo])

  useEffect(() => {
    if (renderVendorInfo) {
      console.log(objVendorInfo)
      setVendorInfo(objVendorInfo)
      setRenderVendorInfo(false)
    }
  }, [renderVendorInfo])

  useEffect(() => {
    if (renderTaxesInfo) {
      objTaxes.Total.PreTax = (
        (isNaN(parseInt(objTaxes.Product.PreTax))
          ? 0
          : parseInt(objTaxes.Product.PreTax)) +
        (isNaN(parseInt(objTaxes.Labour.PreTax))
          ? 0
          : parseInt(objTaxes.Labour.PreTax)) +
        (isNaN(parseInt(objTaxes.ProductSubTotal.PreTax))
          ? 0
          : parseInt(objTaxes.ProductSubTotal.PreTax)) +
        (isNaN(parseInt(objTaxes.DemoProduct.PreTax))
          ? 0
          : parseInt(objTaxes.DemoProduct.PreTax)) +
        (isNaN(parseInt(objTaxes.DemoLabour.PreTax))
          ? 0
          : parseInt(objTaxes.DemoLabour.PreTax)) +
        (isNaN(parseInt(objTaxes.DemoSubTotal.PreTax))
          ? 0
          : parseInt(objTaxes.DemoSubTotal.PreTax)) +
        (isNaN(parseInt(objTaxes.Freight.PreTax))
          ? 0
          : parseInt(objTaxes.Freight.PreTax)) +
        (isNaN(parseInt(objTaxes.SpringHandling.PreTax))
          ? 0
          : parseInt(objTaxes.SpringHandling.PreTax)) +
        (isNaN(parseInt(objTaxes.ProfitOH.PreTax))
          ? 0
          : parseInt(objTaxes.ProfitOH.PreTax)) +
        (isNaN(parseInt(objTaxes.Insurance.PreTax))
          ? 0
          : parseInt(objTaxes.Insurance.PreTax))
      ).toString()

      objTaxes.Total.Tax = (
        (isNaN(parseInt(objTaxes.Product.Tax))
          ? 0
          : parseInt(objTaxes.Product.Tax)) +
        (isNaN(parseInt(objTaxes.Labour.Tax))
          ? 0
          : parseInt(objTaxes.Labour.Tax)) +
        (isNaN(parseInt(objTaxes.ProductSubTotal.Tax))
          ? 0
          : parseInt(objTaxes.ProductSubTotal.Tax)) +
        (isNaN(parseInt(objTaxes.DemoProduct.Tax))
          ? 0
          : parseInt(objTaxes.DemoProduct.Tax)) +
        (isNaN(parseInt(objTaxes.DemoLabour.Tax))
          ? 0
          : parseInt(objTaxes.DemoLabour.Tax)) +
        (isNaN(parseInt(objTaxes.DemoSubTotal.Tax))
          ? 0
          : parseInt(objTaxes.DemoSubTotal.Tax)) +
        (isNaN(parseInt(objTaxes.Freight.Tax))
          ? 0
          : parseInt(objTaxes.Freight.Tax)) +
        (isNaN(parseInt(objTaxes.SpringHandling.Tax))
          ? 0
          : parseInt(objTaxes.SpringHandling.Tax)) +
        (isNaN(parseInt(objTaxes.ProfitOH.Tax))
          ? 0
          : parseInt(objTaxes.ProfitOH.Tax)) +
        (isNaN(parseInt(objTaxes.Insurance.Tax))
          ? 0
          : parseInt(objTaxes.Insurance.Tax))
      ).toString()

      objTaxes.Total.Total = (
        (isNaN(parseInt(objTaxes.Product.Total))
          ? 0
          : parseInt(objTaxes.Product.Total)) +
        (isNaN(parseInt(objTaxes.Labour.Total))
          ? 0
          : parseInt(objTaxes.Labour.Total)) +
        (isNaN(parseInt(objTaxes.ProductSubTotal.Total))
          ? 0
          : parseInt(objTaxes.ProductSubTotal.Total)) +
        (isNaN(parseInt(objTaxes.DemoProduct.Total))
          ? 0
          : parseInt(objTaxes.DemoProduct.Total)) +
        (isNaN(parseInt(objTaxes.DemoLabour.Total))
          ? 0
          : parseInt(objTaxes.DemoLabour.Total)) +
        (isNaN(parseInt(objTaxes.DemoSubTotal.Total))
          ? 0
          : parseInt(objTaxes.DemoSubTotal.Total)) +
        (isNaN(parseInt(objTaxes.Freight.Total))
          ? 0
          : parseInt(objTaxes.Freight.Total)) +
        (isNaN(parseInt(objTaxes.SpringHandling.Total))
          ? 0
          : parseInt(objTaxes.SpringHandling.Total)) +
        (isNaN(parseInt(objTaxes.ProfitOH.Total))
          ? 0
          : parseInt(objTaxes.ProfitOH.Total)) +
        (isNaN(parseInt(objTaxes.Insurance.Total))
          ? 0
          : parseInt(objTaxes.Insurance.Total))
      ).toString()

      setTaxesInfo(objTaxes)
      setRenderTaxesInfo(false)
    }
  }, [renderTaxesInfo])

  useEffect(() => {
    if (fetchingTable) {
      setInstallationTable([...arrInstall])
      setFetchingTable(false)
    }
  }, [fetchingTable])
  const templateChangeHandler = (key) => {
    arrInstall = arrInstall.filter((item) => item.templateOf == '')
    arrTemplateSelected = arrTemplateMaster.filter(
      (masItem) => masItem.templateOf == key,
    )
    arrInstall = [...arrTemplateSelected, ...arrInstall]
    setFetchingTable(true)
  }

  const addInstallationRowhandler = () => {
    let newId =
      arrInstall.length > 0 ? arrInstall[arrInstall.length - 1].id + 1 : 100000
    arrInstall.push({
      id: newId,
      detailedDescription: '',
      manufacture: 'LynxSpring',
      modelNo: '',
      serialNo: '',
      quantity: '',
      eachPrice: '',
      eachMarkup: '',
      totalProduct: '',
      taxable: false,
      people: '',
      hoursPerPerson: '',
      hourlyBillingRate: '',
      unionRate: false,
      totalLabour: '',
      labourTaxable: false,
      grandTotalProduct: '',
      templateOf: '',
    })

    setFetchingTable(true)
  }

  const SubmitHandler = () => {
    try {
      props.spcontext.web.lists
        .getByTitle('WFQuoteRequestList')
        .items.getById(formID)
        .update({
          // install
          installationDetails: JSON.stringify(arrInstall),
          // Proj
          OrderNo: objProjInfo.projNoInput,
          ManagerName: objProjInfo.projManagerInput,
          BENumber: objProjInfo.BENoInput,
          Title: objProjInfo.ProjNameInput,
          ShippingAddress: objProjInfo.DeliveryAddInput,
          ProjectArea: objProjInfo.projAreaInput,
          StartDate: objProjInfo.EstimateStartDateInput,
          EndDate: objProjInfo.EstimateEndDateInput,
          InstallTotalProduct: objProjInfo.InstallTotalProduct,
          InstallTotalPeople: objProjInfo.InstallTotalPeople,
          InstallTotalHoursPerPerson: objProjInfo.InstallTotalHoursPerPerson,
          InstallSecondTotalProduct: objProjInfo.InstallSecondTotalProduct,
          // vendor
          companyName: objVendorInfo.companyNameInput,
          wfVendorNo: objVendorInfo.wfVendorNoInput,
          proposalNo: objVendorInfo.proposalNoInput,
          cityStateZip: objVendorInfo.cityStateZipInput,
          wfContractNo: objVendorInfo.wfContractNoInput,
          contactName: objVendorInfo.contactNameInput,
          changeOrder: objVendorInfo.changeOrderInput,
          phoneNo: objVendorInfo.phoneNoInput,
          changeOrderPO: objVendorInfo.changeOrderPOInput,
          cell: objVendorInfo.cellInput,
          remitToAddress: objVendorInfo.remitAddInput,
          emailID: objVendorInfo.emailIdInput,
          ScopeOfWork: objVendorInfo.scopeOfWork,
          AssumptionsClarifications: objVendorInfo.assumptionsAndClarifications,
          // tax Info
          taxesInfo: JSON.stringify(objTaxes),
          Status: 'Order in production',
          WfProjectManagerId:
            ProjectManagerUser.length > 0
              ? { results: ProjectManagerUser }
              : { results: ProjectManagerUser },
        })
        .then((data) => history.back())
        //
        .catch((error) => console.log(error))
      // .then(() => history.back());
    } catch (error) {
      console.log(error)
    }
  }
  async function getProjManager(event): Promise<void> {
    console.log(event)
    ProjectManagerUser = []
    for (let i = 0; i < event.length; i++) {
      await props.spcontext.web.siteUsers
        .getByEmail(event[i].secondaryText.toLowerCase())
        .get()
        .then(async function (result): Promise<void> {
          if (result.Id) ProjectManagerUser.push(result.Id)
          console.log(ProjectManagerUser)
        })
        .catch((error): void => {
          alert(error)
        })
    }
  }
  return (
    <ThemeProvider
      theme={redTheme}
      style={{ backgroundColor: '#F2F2F2', padding: '1rem' }}
    >
      <div className={styles.formHeader}>
        <Icon
          iconName="NavigateBack"
          styles={{
            root: {
              fontSize: 30,
              fontWeight: 600,
              color: '#D71E2B',
              marginRight: '1rem',
            },
          }}
          onClick={() => {
            window.open(
              `https://chandrudemo.sharepoint.com/sites/LynxSpring/SitePages/AdminDashboard.aspx`,
            )
          }}
        />
        <div style={{ fontWeight: 'bold' }}>Order No: {orderNo}</div>
      </div>
      <h1 className={styles.heading}>Quote Form</h1>
      <div className={styles.quoteFormSectionOne}>
        <div className={styles.sectionOneSub} style={{ marginRight: '0.3rem' }}>
          <h3 className={styles.heading} style={{ margin: '0 0 0.5rem 0' }}>
            PROJECT / INFROMATION (Information provided by Wells Fargo)
          </h3>
          <div style={{ display: 'flex' }}>
            <TextField
              label="Project No or Wor Order No"
              styles={halfWidthInput}
              value={projectInfo.projNoInput}
              onChange={(e) => {
                objProjInfo.projNoInput = e.target['value']
                setRenderProjInfo(true)
              }}
            />
            <ThemeProvider
              theme={redTheme}
              className={styles.peoplePickerSection}
            >
              <PeoplePicker
                context={props.cont}
                titleText="WF Project/Property Manager"
                personSelectionLimit={1}
                styles={{ root: { width: '100%' } }}
                groupName={'WellsFargoGroup'} // Leave this blank in case you want to filter from all users
                showtooltip={true}
                required={false}
                disabled={false}
                onChange={(e) => getProjManager(e)}
                showHiddenInUI={false}
                principalTypes={[PrincipalType.User]}
                resolveDelay={1000}
              />
            </ThemeProvider>
          </div>
          <div style={{ display: 'flex' }}>
            <TextField
              label="BE Number"
              styles={halfWidthInput}
              value={projectInfo.BENoInput}
              onChange={(e) => {
                objProjInfo.BENoInput = e.target['value']
                setRenderProjInfo(true)
              }}
            />
            <TextField
              label="Building / Project Name"
              styles={halfWidthInput}
              value={projectInfo.ProjNameInput}
              onChange={(e) => {
                objProjInfo.ProjNameInput = e.target['value']
                setRenderProjInfo(true)
              }}
            />
          </div>
          <div style={{ display: 'flex' }}>
            <TextField
              label="BE Service or Delivery Address"
              styles={halfWidthInput}
              value={projectInfo.DeliveryAddInput}
              onChange={(e) => {
                objProjInfo.DeliveryAddInput = e.target['value']
                setRenderProjInfo(true)
              }}
            />
            <TextField
              label="Project Area (sq.ft)"
              styles={halfWidthInput}
              value={projectInfo.projAreaInput}
              onChange={(e) => {
                objProjInfo.projAreaInput = e.target['value']
                setRenderProjInfo(true)
              }}
            />
          </div>

          <div style={{ display: 'flex' }}>
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
              label="Estimate Start Date"
              firstDayOfWeek={firstDayOfWeek}
              placeholder="Select a date..."
              ariaLabel="Select a date"
              value={objProjInfo.EstimateStartDateInput}
              // DatePicker uses English strings by default. For localized apps, you must override this prop.
              strings={defaultDatePickerStrings}
              styles={halfWidthInput}
              onSelectDate={(date) => {
                objProjInfo.EstimateStartDateInput = date
                setRenderProjInfo(true)
              }}
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
              label="Estimate End Date"
              firstDayOfWeek={firstDayOfWeek}
              placeholder="Select a date..."
              ariaLabel="Select a date"
              // DatePicker uses English strings by default. For localized apps, you must override this prop.
              strings={defaultDatePickerStrings}
              styles={halfWidthInput}
              value={objProjInfo.EstimateEndDateInput}
              onSelectDate={(date) => {
                objProjInfo.EstimateEndDateInput = date
                setRenderProjInfo(true)
              }}
            />
          </div>
        </div>

        <div className={styles.sectionOneSub} style={{ marginLeft: '0.3rem' }}>
          <h3 className={styles.heading} style={{ margin: '0 0 0.5rem 0' }}>
            VENDOR'S AUTHORIZED REPRESENTATIVE
          </h3>
          <div style={{ display: 'flex' }}>
            <TextField
              label="Company Name"
              styles={halfWidthInput}
              value={vendorInfo.companyNameInput}
              onChange={(e) => {
                objVendorInfo.companyNameInput = e.target['value']
                setRenderVendorInfo(true)
              }}
            />
            <TextField
              label="WF Vendor No"
              styles={halfWidthInput}
              value={vendorInfo.wfVendorNoInput}
              onChange={(e) => {
                objVendorInfo.wfVendorNoInput = e.target['value']
                setRenderVendorInfo(true)
              }}
            />
          </div>
          <div style={{ display: 'flex' }}>
            <TextField
              label="Remit to Address"
              styles={halfWidthInput}
              value={vendorInfo.remitAddInput}
              onChange={(e) => {
                objVendorInfo.remitAddInput = e.target['value']
                setRenderVendorInfo(true)
              }}
            />
            <TextField
              label="Proposal No"
              styles={halfWidthInput}
              value={vendorInfo.proposalNoInput}
              onChange={(e) => {
                objVendorInfo.proposalNoInput = e.target['value']
                setRenderVendorInfo(true)
              }}
            />
          </div>
          <div style={{ display: 'flex' }}>
            <TextField
              label="City,State,Zip"
              styles={halfWidthInput}
              value={vendorInfo.cityStateZipInput}
              onChange={(e) => {
                objVendorInfo.cityStateZipInput = e.target['value']
                setRenderVendorInfo(true)
              }}
            />
            <TextField
              label="WF Contract Number"
              styles={halfWidthInput}
              value={vendorInfo.wfContractNoInput}
              onChange={(e) => {
                objVendorInfo.wfContractNoInput = e.target['value']
                setRenderVendorInfo(true)
              }}
            />
          </div>
          <div style={{ display: 'flex' }}>
            <TextField
              label="Contact Name"
              styles={halfWidthInput}
              value={vendorInfo.contactNameInput}
              onChange={(e) => {
                objVendorInfo.contactNameInput = e.target['value']
                setRenderVendorInfo(true)
              }}
            />
            <TextField
              label="Change Order"
              styles={halfWidthInput}
              value={vendorInfo.changeOrderInput}
              onChange={(e) => {
                objVendorInfo.changeOrderInput = e.target['value']
                setRenderVendorInfo(true)
              }}
            />
          </div>
          <div style={{ display: 'flex' }}>
            <TextField
              label="Phone Number"
              styles={halfWidthInput}
              value={vendorInfo.phoneNoInput}
              onChange={(e) => {
                objVendorInfo.phoneNoInput = e.target['value']
                setRenderVendorInfo(true)
              }}
            />
            <TextField
              label="Change Order Previous PO#"
              styles={halfWidthInput}
              value={vendorInfo.changeOrderPOInput}
              onChange={(e) => {
                objVendorInfo.changeOrderPOInput = e.target['value']
                setRenderVendorInfo(true)
              }}
            />
          </div>
          <div style={{ display: 'flex' }}>
            <TextField
              label="Cell"
              styles={halfWidthInput}
              value={vendorInfo.cellInput}
              onChange={(e) => {
                objVendorInfo.cellInput = e.target['value']
                setRenderVendorInfo(true)
              }}
            />
            <TextField
              label="Email ID"
              styles={halfWidthInput}
              value={vendorInfo.emailIdInput}
              onChange={(e) => {
                objVendorInfo.emailIdInput = e.target['value']
                setRenderVendorInfo(true)
              }}
            />
          </div>
        </div>
      </div>
      <div className={styles.quoteFormSection}>
        <h3 className={styles.heading} style={{ margin: '0 0 0.5rem 0' }}>
          Scope Of Work
        </h3>

        <TextField
          multiline
          rows={3}
          value={vendorInfo.scopeOfWork}
          onChange={(e) => {
            objVendorInfo.scopeOfWork = e.target['value']
            setRenderVendorInfo(true)
          }}
        />

        {/* <p>
          Detail Needed- Be as descriptive as possible and if there are
          multiples of items please specify how many. Details and numbers of
          units help fixed assets determine asset value and will help eliminate
          questions and the need to resubmit proposals or invoices.{" "}
        </p> */}
      </div>
      <div className={styles.quoteFormSection}>
        <h3 className={styles.heading} style={{ margin: '0 0 0.5rem 0' }}>
          Installation
        </h3>
        <div
          style={{
            display: 'flex',
            alignItems: 'center',
            marginBottom: '1rem',
          }}
        >
          <Dropdown
            options={templateOptions}
            disabled={tempApplied}
            placeholder="Select a template"
            styles={{ root: { width: 200, margin: '0 2rem 0 auto' } }}
            onChange={(e, selected) => {
              templateChangeHandler(selected.key)
              setTempApplied(true)
            }}
          />
          <Icon
            iconName="Refresh"
            styles={{
              root: {
                fontSize: 20,
                fontWeight: 400,
                color: '#D71E2B',
                cursor: 'pointer',
              },
            }}
            onClick={() => {
              setTempApplied(false)
            }}
          />
        </div>
        <table className={styles.installationTbl}>
          <thead>
            <tr>
              <th>
                <div
                  style={{
                    display: 'flex',
                    alignItems: 'center',
                    justifyContent: 'space-between',
                    padding: '0.5rem',
                  }}
                >
                  <Icon
                    iconName={
                      isShowDescription
                        ? 'ChevronDownSmall'
                        : 'ChevronRightSmall'
                    }
                    styles={{
                      root: {
                        fontSize: 20,
                        fontWeight: 400,
                        color: '#D71E2B',
                        cursor: 'pointer',
                        marginRight: '0.5rem',
                      },
                    }}
                    onClick={() => {
                      isShowDescription
                        ? setIsShowDescription(false)
                        : setIsShowDescription(true)
                    }}
                  />
                  {isShowDescription
                    ? 'WF needs to see the Labor Costs associated with the Product. If Quote is for labor only, leave Product Section blank. Include separate lines for HVAC components as well as separate lines for equipent and materials related to that component.'
                    : ''}
                </div>
              </th>
              <th>Manufacturer</th>
              <th>Model#</th>
              <th>Serial#</th>
              <th>Quantity</th>
              <th>Each Price</th>
              <th>Each Markup</th>
              <th>Total Product</th>
              <th>Taxable Y/N</th>
              <th>People#</th>
              <th>Hours per person</th>
              <th>Hourly Billing Rate</th>
              <th>Union rate Y/N</th>
              <th>Total Labor</th>
              <th>Labor taxable?(Y/N)</th>
              <th>Total Products</th>
              <th>Actions</th>
            </tr>
          </thead>
          <tbody>
            {installationtable.map((installItem, i) => {
              return (
                <tr>
                  <td style={{ width: isShowDescription ? '14rem' : '1rem' }}>
                    {isShowDescription ? (
                      <TextField
                        multiline
                        rows={3}
                        value={installItem.detailedDescription}
                        onChange={(e) => {
                          arrInstall.filter(
                            (arrItem) => arrItem.id === installItem.id,
                          )[0].detailedDescription = e.target['value']
                          // setInstallationTable(arrInstall);
                          setFetchingTable(true)
                        }}
                      />
                    ) : (
                      ''
                    )}
                  </td>
                  <td>
                    <TextField
                      key={i}
                      id={`${installItem.id}`}
                      value={installItem.manufacture}
                      onChange={(e) => {
                        arrInstall.filter(
                          (arrItem) => arrItem.id === +e.target['id'],
                        )[0].manufacture = e.target['value']
                        // setInstallationTable(arrInstall);
                        setFetchingTable(true)
                      }}
                    />
                  </td>
                  <td>
                    <TextField
                      key={i}
                      id={`${installItem.id}`}
                      value={installItem.modelNo}
                      onChange={(e) => {
                        arrInstall.filter(
                          (arrItem) => arrItem.id === +e.target['id'],
                        )[0].modelNo = e.target['value']
                        // setInstallationTable(arrInstall);
                        setFetchingTable(true)
                      }}
                    />
                  </td>
                  <td>
                    <TextField
                      key={i}
                      id={`${installItem.id}`}
                      value={installItem.serialNo}
                      onChange={(e) => {
                        arrInstall.filter(
                          (arrItem) => arrItem.id === +e.target['id'],
                        )[0].serialNo = e.target['value']
                        // setInstallationTable(arrInstall);
                        setFetchingTable(true)
                      }}
                    />
                  </td>
                  <td>
                    <TextField
                      key={i}
                      id={`${installItem.id}`}
                      value={installItem.quantity}
                      onChange={(e) => {
                        arrInstall.filter(
                          (arrItem) => arrItem.id === +e.target['id'],
                        )[0].quantity = e.target['value']
                        // setInstallationTable(arrInstall);
                        setFetchingTable(true)
                      }}
                    />
                  </td>
                  <td>
                    <p>$</p>
                    <TextField
                      key={i}
                      id={`${installItem.id}`}
                      type="number"
                      value={installItem.eachPrice}
                      onChange={(e) => {
                        arrInstall.filter(
                          (arrItem) => arrItem.id === +e.target['id'],
                        )[0].eachPrice = e.target['value']
                        // setInstallationTable(arrInstall);
                        setFetchingTable(true)
                      }}
                    />
                  </td>
                  <td>
                    <TextField
                      key={i}
                      id={`${installItem.id}`}
                      value={installItem.eachMarkup}
                      onChange={(e) => {
                        arrInstall.filter(
                          (arrItem) => arrItem.id === +e.target['id'],
                        )[0].eachMarkup = e.target['value']
                        // setInstallationTable(arrInstall);
                        setFetchingTable(true)
                      }}
                    />
                  </td>
                  <td>
                    <TextField
                      key={i}
                      value={installItem.totalProduct}
                      onChange={(e) => {
                        arrInstall.filter(
                          (arrItem) => arrItem.id === installItem.id,
                        )[0].totalProduct = e.target['value']
                        setFetchingTable(true)
                      }}
                    />
                  </td>
                  <td>
                    <Checkbox
                      styles={{ root: { justifyContent: 'center' } }}
                      checked={installItem.taxable ? true : false}
                      onChange={(e, checked) => {
                        arrInstall.filter(
                          (arrItem) => arrItem.id === installItem.id,
                        )[0].taxable = checked
                        setFetchingTable(true)
                      }}
                    />
                  </td>
                  <td>
                    <TextField
                      key={i}
                      id={`${installItem.id}`}
                      value={installItem.people}
                      onChange={(e) => {
                        arrInstall.filter(
                          (arrItem) => arrItem.id === +e.target['id'],
                        )[0].people = e.target['value']
                        // setInstallationTable(arrInstall);
                        setFetchingTable(true)
                      }}
                    />
                  </td>
                  <td>
                    <TextField
                      key={i}
                      id={`${installItem.id}`}
                      value={installItem.hoursPerPerson}
                      onChange={(e) => {
                        arrInstall.filter(
                          (arrItem) => arrItem.id === +e.target['id'],
                        )[0].hoursPerPerson = e.target['value']
                        // setInstallationTable(arrInstall);
                        setFetchingTable(true)
                      }}
                    />
                  </td>
                  <td>
                    <p>$</p>
                    <TextField
                      key={i}
                      type="number"
                      id={`${installItem.id}`}
                      value={installItem.hourlyBillingRate}
                      onChange={(e) => {
                        arrInstall.filter(
                          (arrItem) => arrItem.id === +e.target['id'],
                        )[0].hourlyBillingRate = e.target['value']
                        // setInstallationTable(arrInstall);
                        setFetchingTable(true)
                      }}
                    />
                  </td>
                  <td>
                    <Checkbox
                      styles={{ root: { justifyContent: 'center' } }}
                      checked={installItem.unionRate ? true : false}
                      onChange={(e, checked) => {
                        arrInstall.filter(
                          (arrItem) => arrItem.id === installItem.id,
                        )[0].unionRate = checked
                        setFetchingTable(true)
                      }}
                    />
                  </td>
                  <td>
                    <TextField
                      key={i}
                      id={`${installItem.id}`}
                      value={installItem.totalLabour}
                      onChange={(e) => {
                        arrInstall.filter(
                          (arrItem) => arrItem.id === +e.target['id'],
                        )[0].totalLabour = e.target['value']
                        // setInstallationTable(arrInstall);
                        setFetchingTable(true)
                      }}
                    />
                  </td>
                  <td>
                    <Checkbox
                      styles={{ root: { justifyContent: 'center' } }}
                      checked={installItem.labourTaxable ? true : false}
                      onChange={(e, checked) => {
                        arrInstall.filter(
                          (arrItem) => arrItem.id === installItem.id,
                        )[0].labourTaxable = checked
                        setFetchingTable(true)
                      }}
                    />
                  </td>
                  <td>
                    <TextField
                      key={i}
                      id={`${installItem.id}`}
                      value={installItem.grandTotalProduct}
                      onChange={(e) => {
                        arrInstall.filter(
                          (arrItem) => arrItem.id === +e.target['id'],
                        )[0].grandTotalProduct = e.target['value']
                        // setInstallationTable(arrInstall);
                        setFetchingTable(true)
                      }}
                    />
                  </td>
                  <td>
                    <Icon
                      iconName="delete"
                      styles={{
                        root: {
                          fontSize: 20,
                          fontWeight: 400,
                          color: '#D71E2B',
                          cursor: 'pointer',
                        },
                      }}
                      onClick={() => {
                        arrInstall = arrInstall.filter(
                          (arrItem) => arrItem.id !== installItem.id,
                        )
                        setFetchingTable(true)
                      }}
                    />
                  </td>
                </tr>
              )
            })}
          </tbody>
        </table>
        <div
          style={{
            display: 'flex',
            justifyContent: 'flex-end',
            margin: '1rem 0',
          }}
        >
          {' '}
          <DefaultButton
            text="Add"
            iconProps={addIcon}
            styles={{ root: { marginLeft: 'auto', textAlign: 'right' } }}
            onClick={() => {
              addInstallationRowhandler()
              console.log(arrInstall)
              setInstallationTable(arrInstall)
            }}
          />
        </div>
        <table className={styles.installationTbl}>
          <thead>
            <tr>
              <th></th>
              <th></th>
              <th></th>
              <th></th>
              <th></th>
              <th></th>
              <th></th>
              <th></th>
              <th></th>
              <th></th>
              <th></th>
              <th></th>
              <th></th>
              <th></th>
              <th></th>
              <th></th>
              <th></th>
            </tr>
          </thead>
          <tbody>
            <tr>
              <td></td>
              <td colSpan={7}>
                <div style={{ width: '3rem', marginLeft: 'auto' }}>
                 {/* <label htmlFor=""></label> */}
                  <TextField
                    value={projectInfo.InstallTotalProduct}
                    onChange={(e) => {
                      objProjInfo.InstallTotalProduct = e.target['value']
                      setRenderProjInfo(true)
                    }}
                  />
                </div>
              </td>

              <td colSpan={2}>
                <div style={{ width: '3rem', marginLeft: 'auto' }}>
                  {/* <label htmlFor=""></label> */}
                  <TextField
                    value={projectInfo.InstallTotalPeople}
                    onChange={(e) => {
                      objProjInfo.InstallTotalPeople = e.target['value']
                      setRenderProjInfo(true)
                    }}
                  />
                </div>
              </td>
              <td>
                <div style={{ width: '3rem', marginLeft: 'auto' }}>
                  {/* <label htmlFor=""></label> */}
                  <TextField
                    value={projectInfo.InstallTotalHoursPerPerson}
                    onChange={(e) => {
                      objProjInfo.InstallTotalHoursPerPerson = e.target['value']
                      setRenderProjInfo(true)
                    }}
                  />
                </div>
              </td>

              <td colSpan={5}>
                <div style={{ width: '3rem', marginLeft: 'auto' }}>
                  <TextField
                    value={projectInfo.InstallSecondTotalProduct}
                    onChange={(e) => {
                      objProjInfo.InstallSecondTotalProduct = e.target['value']
                      setRenderProjInfo(true)
                    }}
                  />
                </div>
              </td>
            </tr>
          </tbody>
        </table>
      </div>

      <div
        className={styles.quoteFormSection}
        style={{ opacity: 0.5, position: 'relative' }}
      >
        <div
          style={{
            position: 'absolute',
            height: '100%',
            width: '100%',
            zIndex: 100,
          }}
        ></div>
        <h3 className={styles.heading} style={{ margin: '0 0 0.5rem 0' }}>
          Demo / Removal / Patching / Repairs / Relo
        </h3>
        <table className={styles.installationTbl}>
          <thead>
            <tr>
              <th>Include costs of demolition and demolition labor.</th>
              <th>Quantity</th>
              <th>Price Each</th>
              <th>Total Demo</th>
              <th>Taxable? (Y/N)</th>
              <th>People #</th>
              <th>Hours Per Person #</th>
              <th>Hourly Bill rate #</th>
              <th>Union Rate Y/N</th>
              <th>Total Labor</th>
              <th>Labor taxable?(Y/N)</th>
              <th>Total Products</th>
            </tr>
          </thead>
          <tbody>
            <tr>
              <td>
                {' '}
                <TextField value="Lynx Spring" />
              </td>
              <td>
                <TextField
                  styles={{
                    root: {
                      width: 30,
                      margin: 'auto',
                    },
                  }}
                />
              </td>
              <td>
                <MaskedTextField mask="$\" />
              </td>
              <td>
                <MaskedTextField mask="$\" />
              </td>
              <td>
                <TextField />
              </td>
              <td>
                <TextField />
              </td>
              <td>
                <TextField />
              </td>
              <td>
                <MaskedTextField mask="$\" />
              </td>
              <td>
                <TextField />
              </td>
              <td>
                <MaskedTextField mask="$\" />
              </td>
              <td>
                <TextField />
              </td>
              <td>
                <MaskedTextField mask="$\" />
              </td>
            </tr>
          </tbody>
        </table>
        <DefaultButton
          text="Add"
          iconProps={addIcon}
          styles={{ root: { marginLeft: 'auto', textAlign: 'right' } }}
          onClick={() => {}}
        />
        <div className={styles.totalSection}>
          <MaskedTextField mask="$\" />
          <TextField />
          <TextField />
          <MaskedTextField mask="$\" />
        </div>
      </div>
      <div className={styles.quoteFormSection}>
        <h3 className={styles.heading} style={{ margin: '0 0 0.5rem 0' }}>
          Assumptions and Clarifications
        </h3>

        <TextField
          multiline
          rows={3}
          value={vendorInfo.assumptionsAndClarifications}
          onChange={(e) => {
            objVendorInfo.assumptionsAndClarifications = e.target['value']
            setRenderVendorInfo(true)
          }}
        />
        {/* <p>
          Clarify assumptions, clarifications, and exclusions in this space.
        </p> */}
      </div>
      <div className={styles.taxSection}>
        <table className={styles.taxTable}>
          <thead>
            <tr>
              <th></th>
              <th>Pre-Tax</th>
              <th>Tax</th>
              <th>Total</th>
            </tr>
          </thead>
          <tbody>
            <tr>
              <td>Product</td>
              <td>
                <TextField
                  value={taxesInfo.Product.PreTax}
                  onChange={(e) => {
                    taxesInfo.Product.PreTax = e.target['value']
                    setRenderTaxesInfo(true)
                  }}
                />
              </td>
              <td>
                <TextField
                  value={taxesInfo.Product.Tax}
                  onChange={(e) => {
                    taxesInfo.Product.Tax = e.target['value']
                    setRenderTaxesInfo(true)
                  }}
                />
              </td>
              <td>
                <TextField
                  value={taxesInfo.Product.Total}
                  onChange={(e) => {
                    taxesInfo.Product.Total = e.target['value']
                    setRenderTaxesInfo(true)
                  }}
                />
              </td>
            </tr>
            <tr>
              <td>Labor</td>
              <td>
                <TextField
                  value={taxesInfo.Labour.PreTax}
                  onChange={(e) => {
                    taxesInfo.Labour.PreTax = e.target['value']
                    setRenderTaxesInfo(true)
                  }}
                />
              </td>
              <td>
                <TextField
                  value={taxesInfo.Labour.Tax}
                  onChange={(e) => {
                    taxesInfo.Labour.Tax = e.target['value']
                    setRenderTaxesInfo(true)
                  }}
                />
              </td>
              <td>
                <TextField
                  value={taxesInfo.Labour.Total}
                  onChange={(e) => {
                    taxesInfo.Labour.Total = e.target['value']
                    setRenderTaxesInfo(true)
                  }}
                />
              </td>
            </tr>
            <tr style={{ backgroundColor: '#fdefeb' }}>
              <td>Sub Total</td>
              <td>
                <TextField
                  value={taxesInfo.ProductSubTotal.PreTax}
                  onChange={(e) => {
                    taxesInfo.ProductSubTotal.PreTax = e.target['value']
                    setRenderTaxesInfo(true)
                  }}
                />
              </td>
              <td>
                <TextField
                  value={taxesInfo.ProductSubTotal.Tax}
                  onChange={(e) => {
                    taxesInfo.ProductSubTotal.Tax = e.target['value']
                    setRenderTaxesInfo(true)
                  }}
                />
              </td>
              <td>
                <TextField
                  value={taxesInfo.ProductSubTotal.Total}
                  onChange={(e) => {
                    taxesInfo.ProductSubTotal.Total = e.target['value']
                    setRenderTaxesInfo(true)
                  }}
                />
              </td>
            </tr>
            <tr>
              <td>Demo Product</td>
              <td>
                <TextField
                  value={taxesInfo.DemoProduct.PreTax}
                  onChange={(e) => {
                    taxesInfo.DemoProduct.PreTax = e.target['value']
                    setRenderTaxesInfo(true)
                  }}
                />
              </td>
              <td>
                <TextField
                  value={taxesInfo.DemoProduct.Tax}
                  onChange={(e) => {
                    taxesInfo.DemoProduct.Tax = e.target['value']
                    setRenderTaxesInfo(true)
                  }}
                />
              </td>
              <td>
                <TextField
                  value={taxesInfo.DemoProduct.Total}
                  onChange={(e) => {
                    taxesInfo.DemoProduct.Total = e.target['value']
                    setRenderTaxesInfo(true)
                  }}
                />
              </td>
            </tr>
            <tr>
              <td>Demo Labor</td>
              <td>
                <TextField
                  value={taxesInfo.DemoLabour.PreTax}
                  onChange={(e) => {
                    taxesInfo.DemoLabour.PreTax = e.target['value']
                    setRenderTaxesInfo(true)
                  }}
                />
              </td>
              <td>
                <TextField
                  value={taxesInfo.DemoLabour.Tax}
                  onChange={(e) => {
                    taxesInfo.DemoLabour.Tax = e.target['value']
                    setRenderTaxesInfo(true)
                  }}
                />
              </td>
              <td>
                <TextField
                  value={taxesInfo.DemoLabour.Total}
                  onChange={(e) => {
                    taxesInfo.DemoLabour.Total = e.target['value']
                    setRenderTaxesInfo(true)
                  }}
                />
              </td>
            </tr>
            <tr style={{ backgroundColor: '#fdefeb' }}>
              <td>Sub Total</td>
              <td>
                <TextField
                  value={taxesInfo.DemoSubTotal.PreTax}
                  onChange={(e) => {
                    taxesInfo.DemoSubTotal.PreTax = e.target['value']
                    setRenderTaxesInfo(true)
                  }}
                />
              </td>
              <td>
                <TextField
                  value={taxesInfo.DemoSubTotal.Tax}
                  onChange={(e) => {
                    taxesInfo.DemoSubTotal.Tax = e.target['value']
                    setRenderTaxesInfo(true)
                  }}
                />
              </td>
              <td>
                <TextField
                  value={taxesInfo.DemoSubTotal.Total}
                  onChange={(e) => {
                    taxesInfo.DemoSubTotal.Total = e.target['value']
                    setRenderTaxesInfo(true)
                  }}
                />
              </td>
            </tr>
            <tr>
              <td>Freight</td>
              <td>
                <TextField
                  value={taxesInfo.Freight.PreTax}
                  onChange={(e) => {
                    taxesInfo.Freight.PreTax = e.target['value']
                    setRenderTaxesInfo(true)
                  }}
                />
              </td>
              <td>
                <TextField
                  value={taxesInfo.Freight.Tax}
                  onChange={(e) => {
                    taxesInfo.Freight.Tax = e.target['value']
                    setRenderTaxesInfo(true)
                  }}
                />
              </td>
              <td>
                <TextField
                  value={taxesInfo.Freight.Total}
                  onChange={(e) => {
                    taxesInfo.Freight.Total = e.target['value']
                    setRenderTaxesInfo(true)
                  }}
                />
              </td>
            </tr>
            <tr>
              <td>Spring Handling</td>
              <td>
                <TextField
                  value={taxesInfo.SpringHandling.PreTax}
                  onChange={(e) => {
                    taxesInfo.SpringHandling.PreTax = e.target['value']
                    setRenderTaxesInfo(true)
                  }}
                />
              </td>
              <td>
                <TextField
                  value={taxesInfo.SpringHandling.Tax}
                  onChange={(e) => {
                    taxesInfo.SpringHandling.Tax = e.target['value']
                    setRenderTaxesInfo(true)
                  }}
                />
              </td>
              <td>
                <TextField
                  value={taxesInfo.SpringHandling.Total}
                  onChange={(e) => {
                    taxesInfo.SpringHandling.Total = e.target['value']
                    setRenderTaxesInfo(true)
                  }}
                />
              </td>
            </tr>
            <tr>
              <td>Profit & OH</td>
              <td>
                <TextField
                  value={taxesInfo.ProfitOH.PreTax}
                  onChange={(e) => {
                    taxesInfo.ProfitOH.PreTax = e.target['value']
                    setRenderTaxesInfo(true)
                  }}
                />
              </td>
              <td>
                <TextField
                  value={taxesInfo.ProfitOH.Tax}
                  onChange={(e) => {
                    taxesInfo.ProfitOH.Tax = e.target['value']
                    setRenderTaxesInfo(true)
                  }}
                />
              </td>
              <td>
                <TextField
                  value={taxesInfo.ProfitOH.Total}
                  onChange={(e) => {
                    taxesInfo.ProfitOH.Total = e.target['value']
                    setRenderTaxesInfo(true)
                  }}
                />
              </td>
            </tr>
            <tr>
              <td>Insurance</td>
              <td>
                <TextField
                  value={taxesInfo.Insurance.PreTax}
                  onChange={(e) => {
                    taxesInfo.Insurance.PreTax = e.target['value']
                    setRenderTaxesInfo(true)
                  }}
                />
              </td>
              <td>
                <TextField
                  value={taxesInfo.Insurance.Tax}
                  onChange={(e) => {
                    taxesInfo.Insurance.Tax = e.target['value']
                    setRenderTaxesInfo(true)
                  }}
                />
              </td>
              <td>
                <TextField
                  value={taxesInfo.Insurance.Total}
                  onChange={(e) => {
                    taxesInfo.Insurance.Total = e.target['value']
                    setRenderTaxesInfo(true)
                  }}
                />
              </td>
            </tr>
            <tr style={{ backgroundColor: '#fdefeb' }}>
              <td>Total</td>
              <td>
                <TextField
                  value={taxesInfo.Total.PreTax}
                  onChange={(e) => {
                    taxesInfo.Total.PreTax = e.target['value']
                    setRenderTaxesInfo(true)
                  }}
                />
              </td>
              <td>
                <TextField
                  value={taxesInfo.Total.Tax}
                  onChange={(e) => {
                    taxesInfo.Total.Tax = e.target['value']
                    setRenderTaxesInfo(true)
                  }}
                />
              </td>
              <td>
                <TextField
                  value={taxesInfo.Total.Total}
                  onChange={(e) => {
                    taxesInfo.Total.Total = e.target['value']
                    setRenderTaxesInfo(true)
                  }}
                />
              </td>
            </tr>
          </tbody>
        </table>
      </div>
      <div className={styles.SubmitSection}>
        <PrimaryButton
          text="Submit"
          onClick={SubmitHandler}
          styles={{ root: { marginRight: '1rem' } }}
        />
        <DefaultButton
          text="Cancel"
          onClick={() => {
            history.back()
          }}
        />
      </div>
    </ThemeProvider>
  )
}

export default WFQuoteForm
