import * as React from 'react'
import { Fragment } from 'react'
import { useState, useEffect, useRef } from 'react'
import { Icon } from '@fluentui/react/lib/Icon'
import { TextField, MaskedTextField } from '@fluentui/react/lib/TextField'
import { Spinner, SpinnerSize } from '@fluentui/react/lib/Spinner'
import styles from './WellsFargoQuoteView.module.scss'
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
} from '@fluentui/react'
import {
  ContextualMenu,
  IContextualMenuProps,
  IIconProps,
} from '@fluentui/react'
import { loadTheme, createTheme, Theme } from '@fluentui/react'
import * as $ from 'jquery'
let formID = 0
let MasterInstallationOptions
const paramsString = window.location.href.split('?')[1].toLowerCase()
const searchParams = new URLSearchParams(paramsString)
searchParams.has('formid') ? (formID = Number(searchParams.get('formid'))) : ''

const fullWidthInput = {
  root: { width: '70%', marginBottom: '0.5rem' },
}
const halfWidthInput = {
  root: { width: '50%', margin: '0 1rem 0.5rem 0' },
}
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
declare global {
  interface Navigator {
    msSaveBlob?: (blob: any, defaultName?: string) => boolean
  }
}
let arrTemplateMaster = []
let arrTemplateSelected = []
let arrInstall = [
  {
    id: 100000,
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
const WFQuoteView = (props) => {
  let siteURLForFile = props.context.pageContext.web.absoluteUrl
  const [projectInfo, setProjectInfo] = useState(objProjInfo)
  const [vendorInfo, setVendorInfo] = useState(objVendorInfo)
  const [orderNo, setOrderNo] = useState('')
  const [installationtable, setInstallationTable] = useState([])
  const [firstDayOfWeek, setFirstDayOfWeek] = useState(DayOfWeek.Sunday)
  const [taxesInfo, setTaxesInfo] = useState(objTaxes)
  const [templateOptions, setTemplateOptions] = useState([])
  const [Loader, setLoader] = useState(false)
  useEffect(() => {
    props.spcontext.web.lists
      .getByTitle('WFQuoteRequestList')
      .items.select('*,WfProjectManager/Title,Author/Title,Author/EMail')
      .expand('WfProjectManager,Author')
      .getById(formID)
      .get()
      .then((li) => {
        console.log(li)
        objProjInfo = {
          projNoInput: li.OrderNo,
          projManagerInput: li.ManagerName,
          BENoInput: li.BENumber,
          ProjNameInput: li.Title,
          DeliveryAddInput: li.ShippingAddress,
          projAreaInput: li.ProjectArea,
          EstimateStartDateInput: new Date(li.StartDate),
          EstimateEndDateInput: new Date(li.EndDate),
          InstallTotalProduct: li.InstallTotalProduct,
          InstallTotalPeople: li.InstallTotalPeople,
          InstallTotalHoursPerPerson: li.InstallTotalHoursPerPerson,
          InstallSecondTotalProduct: li.InstallSecondTotalProduct,
        }
        objVendorInfo = {
          companyNameInput: li.companyName,
          wfVendorNoInput: li.wfVendorNo,
          remitAddInput: li.remitToAddress,
          proposalNoInput: li.proposalNo,
          cityStateZipInput: li.cityStateZip,
          wfContractNoInput: li.wfContractNo,
          contactNameInput: li.contactName,
          changeOrderInput: li.changeOrder,
          phoneNoInput: li.phoneNo,
          changeOrderPOInput: li.changeOrderPO,
          cellInput: li.cell,
          emailIdInput: li.emailID,
          scopeOfWork: li.ScopeOfWork,
          assumptionsAndClarifications: li.AssumptionsClarifications,
        }
        setProjectInfo(objProjInfo)
        setVendorInfo(objVendorInfo)
        arrInstall = JSON.parse(li.installationDetails)
        setInstallationTable(arrInstall)
        li.taxesInfo
          ? setTaxesInfo(JSON.parse(li.taxesInfo))
          : setTaxesInfo(objTaxes)
        setOrderNo(li.OrderNo)
      })
      .then(() => {
        props.spcontext.web.lists
          .getByTitle('InstallationTemplates')
          .items.select('*')
          .get()
          .then((installationDetails) => {
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
              cursor: 'pointer',
            },
          }}
          onClick={() => {
            history.back()
          }}
        />
        <div style={{ fontWeight: 'bold' }}>Order No: {orderNo}</div>
      </div>
      <h1 className={styles.heading}>Quote Form</h1>
      <div
        style={{
          display: 'flex',
          justifyContent: 'flex-end',
          marginBottom: '1rem',
        }}
      >
        <PrimaryButton
          text="Export Pdf"
          style={{ marginRight: '1rem' }}
          onClick={() =>
            downloadFile(
              'https://wfaapiservice.azurewebsites.net/api/pdf',
              'demo.pdf',
              'pdf',
            )
          }
        />
        <PrimaryButton
          text="Export Excel"
          onClick={() =>
            downloadFile(
              'https://wfaapiservice.azurewebsites.net/api/excel',
              'demo.xlsx',
              'excel',
            )
          }
        />
      </div>
      <div className={styles.quoteFormSectionOne}>
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
        <div className={styles.sectionOneSub} style={{ marginRight: '0.3rem' }}>
          <h3 className={styles.heading} style={{ margin: '0 0 0.5rem 0' }}>
            PROJECT / INFROMATION (Information provided by Wells Fargo)
          </h3>
          <div style={{ display: 'flex' }}>
            <TextField
              label="Project No or Wor Order No"
              styles={halfWidthInput}
              value={projectInfo.projNoInput}
              disabled={true}
            />
            <TextField
              label="WF Project/Property Manager"
              styles={halfWidthInput}
              value={projectInfo.projManagerInput}
              disabled={true}
            />
          </div>
          <div style={{ display: 'flex' }}>
            <TextField
              label="BE Number"
              styles={halfWidthInput}
              value={projectInfo.BENoInput}
              disabled={true}
            />
            <TextField
              label="Building / Project Name"
              styles={halfWidthInput}
              value={projectInfo.ProjNameInput}
              disabled={true}
            />
          </div>
          <div style={{ display: 'flex' }}>
            <TextField
              label="BE Service or Delivery Address"
              styles={halfWidthInput}
              value={projectInfo.DeliveryAddInput}
              disabled={true}
            />
            <TextField
              label="Project Area (sq.ft)"
              styles={halfWidthInput}
              value={projectInfo.projAreaInput}
              disabled={true}
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
              value={objProjInfo.EstimateEndDateInput}
              // DatePicker uses English strings by default. For localized apps, you must override this prop.
              strings={defaultDatePickerStrings}
              styles={halfWidthInput}
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
              label="Estimate End Date"
              firstDayOfWeek={firstDayOfWeek}
              placeholder="Select a date..."
              ariaLabel="Select a date"
              // DatePicker uses English strings by default. For localized apps, you must override this prop.
              strings={defaultDatePickerStrings}
              styles={halfWidthInput}
              disabled={true}
              value={objProjInfo.EstimateEndDateInput}
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
              disabled={true}
            />
            <TextField
              label="WF Vendor No"
              styles={halfWidthInput}
              value={vendorInfo.wfVendorNoInput}
              disabled={true}
            />
          </div>
          <div style={{ display: 'flex' }}>
            <TextField
              label="Remit to Address"
              styles={halfWidthInput}
              value={vendorInfo.remitAddInput}
              disabled={true}
            />
            <TextField
              label="Proposal No"
              styles={halfWidthInput}
              value={vendorInfo.proposalNoInput}
              disabled={true}
            />
          </div>
          <div style={{ display: 'flex' }}>
            <TextField
              label="City,State,Zip"
              styles={halfWidthInput}
              value={vendorInfo.cityStateZipInput}
              disabled={true}
            />
            <TextField
              label="WF Contract Number"
              styles={halfWidthInput}
              value={vendorInfo.wfContractNoInput}
              disabled={true}
            />
          </div>
          <div style={{ display: 'flex' }}>
            <TextField
              label="Contact Name"
              styles={halfWidthInput}
              value={vendorInfo.contactNameInput}
              disabled={true}
            />
            <TextField
              label="Change Order"
              styles={halfWidthInput}
              value={vendorInfo.changeOrderInput}
              disabled={true}
            />
          </div>
          <div style={{ display: 'flex' }}>
            <TextField
              label="Phone Number"
              styles={halfWidthInput}
              value={vendorInfo.phoneNoInput}
              disabled={true}
            />
            <TextField
              label="Change Order Previous PO#"
              styles={halfWidthInput}
              value={vendorInfo.changeOrderPOInput}
              disabled={true}
            />
          </div>
          <div style={{ display: 'flex' }}>
            <TextField
              label="Cell"
              styles={halfWidthInput}
              value={vendorInfo.cellInput}
              disabled={true}
            />
            <TextField
              label="Email ID"
              styles={halfWidthInput}
              value={vendorInfo.emailIdInput}
              disabled={true}
            />
          </div>
        </div>
      </div>
      <div className={styles.quoteFormSection}>
        <h3 className={styles.heading} style={{ margin: '0 0 0.5rem 0' }}>
          Scope Of Work
        </h3>
        <p>{vendorInfo.scopeOfWork}</p>
      </div>
      <div className={styles.quoteFormSection}>
        <h3 className={styles.heading} style={{ margin: '0 0 0.5rem 0' }}>
          Installation
        </h3>
        {/* <div
          style={{
            display: "flex",
            alignItems: "center",
            marginBottom: "1rem",
          }}
        >
          <Dropdown
            options={templateOptions}
            disabled={true}
            placeholder="Select a template"
            styles={{ root: { width: 200, margin: "0 2rem 0 auto" } }}
          />
        </div> */}
        <table className={styles.installationTbl}>
          <thead>
            <tr>
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
            </tr>
          </thead>
          <tbody>
            {installationtable && installationtable.length > 0
              ? installationtable.map((installItem, i) => {
                  return (
                    <tr>
                      <td>
                        <TextField
                          key={i}
                          value={installItem.manufacture}
                          disabled={true}
                        />
                      </td>
                      <td>
                        <TextField
                          key={i}
                          value={installItem.modelNo}
                          disabled={true}
                        />
                      </td>
                      <td>
                        <TextField
                          key={i}
                          value={installItem.serialNo}
                          disabled={true}
                        />
                      </td>
                      <td>
                        <TextField
                          key={i}
                          value={installItem.quantity}
                          disabled={true}
                        />
                      </td>
                      <td>
                        <TextField
                          key={i}
                          value={installItem.eachPrice}
                          disabled={true}
                        />
                      </td>
                      <td>
                        <TextField
                          key={i}
                          value={installItem.eachMarkup}
                          disabled={true}
                        />
                      </td>
                      <td>
                        <TextField
                          key={i}
                          value={installItem.totalProduct}
                          disabled={true}
                        />
                      </td>
                      <td>
                        <Checkbox
                          checked={installItem.taxable ? true : false}
                          disabled={true}
                          styles={{ root: { justifyContent: 'center' } }}
                        />
                      </td>
                      <td>
                        <TextField
                          key={i}
                          value={installItem.people}
                          disabled={true}
                        />
                      </td>
                      <td>
                        <TextField
                          key={i}
                          value={installItem.hoursPerPerson}
                          disabled={true}
                        />
                      </td>
                      <td>
                        <TextField
                          key={i}
                          value={installItem.hourlyBillingRate}
                          disabled={true}
                        />
                      </td>
                      <td>
                        <Checkbox
                          checked={installItem.unionRate ? true : false}
                          disabled={true}
                          styles={{ root: { justifyContent: 'center' } }}
                        />
                      </td>
                      <td>
                        <TextField
                          key={i}
                          value={installItem.totalLabour}
                          disabled={true}
                        />
                      </td>
                      <td>
                        <Checkbox
                          checked={installItem.labourTaxable ? true : false}
                          disabled={true}
                          styles={{ root: { justifyContent: 'center' } }}
                        />
                      </td>
                      <td>
                        <TextField
                          key={i}
                          value={installItem.grandTotalProduct}
                          disabled={true}
                        />
                      </td>
                    </tr>
                  )
                })
              : ''}
            <tr>
              <td colSpan={7}>
                <div style={{ width: '3rem', marginLeft: 'auto' }}>
                  <TextField
                    disabled={true}
                    value={objProjInfo.InstallTotalProduct}
                  />
                </div>
              </td>

              <td colSpan={2}>
                <div style={{ width: '3rem', marginLeft: 'auto' }}>
                  <TextField
                    disabled={true}
                    value={objProjInfo.InstallTotalPeople}
                  />
                </div>
              </td>
              <td>
                <div style={{ width: '3rem', marginLeft: 'auto' }}>
                  <TextField
                    disabled={true}
                    value={objProjInfo.InstallTotalHoursPerPerson}
                  />
                </div>
              </td>

              <td colSpan={5}>
                <div style={{ width: '3rem', marginLeft: 'auto' }}>
                  <TextField
                    disabled={true}
                    value={objProjInfo.InstallSecondTotalProduct}
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
                <TextField />
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
                <TextField />
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
                <TextField />
              </td>
              <td>
                <TextField />
              </td>
            </tr>
          </tbody>
        </table>
        <div className={styles.totalSection}>
          <TextField />
          <TextField />
          <TextField />
          <TextField />
        </div>
      </div>
      <div className={styles.quoteFormSection}>
        <h3 className={styles.heading} style={{ margin: '0 0 0.5rem 0' }}>
          Assumptions and Clarifications
        </h3>
        <p>{vendorInfo.assumptionsAndClarifications}</p>
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
                <TextField value={taxesInfo.Product.PreTax} disabled={true} />
              </td>
              <td>
                <TextField value={taxesInfo.Product.Tax} disabled={true} />
              </td>
              <td>
                <TextField value={taxesInfo.Product.Total} disabled={true} />
              </td>
            </tr>
            <tr>
              <td>Labor</td>
              <td>
                <TextField value={taxesInfo.Labour.PreTax} disabled={true} />
              </td>
              <td>
                <TextField value={taxesInfo.Labour.Tax} disabled={true} />
              </td>
              <td>
                <TextField value={taxesInfo.Labour.Total} disabled={true} />
              </td>
            </tr>
            <tr style={{ backgroundColor: '#fdefeb' }}>
              <td>Sub Total</td>
              <td>
                <TextField
                  value={taxesInfo.ProductSubTotal.PreTax}
                  disabled={true}
                />
              </td>
              <td>
                <TextField
                  value={taxesInfo.ProductSubTotal.Tax}
                  disabled={true}
                />
              </td>
              <td>
                <TextField
                  value={taxesInfo.ProductSubTotal.Total}
                  disabled={true}
                />
              </td>
            </tr>
            <tr>
              <td>Demo Product</td>
              <td>
                <TextField
                  value={taxesInfo.DemoProduct.PreTax}
                  disabled={true}
                />
              </td>
              <td>
                <TextField value={taxesInfo.DemoProduct.Tax} disabled={true} />
              </td>
              <td>
                <TextField
                  value={taxesInfo.DemoProduct.Total}
                  disabled={true}
                />
              </td>
            </tr>
            <tr>
              <td>Demo Labor</td>
              <td>
                <TextField
                  value={taxesInfo.DemoLabour.PreTax}
                  disabled={true}
                />
              </td>
              <td>
                <TextField value={taxesInfo.DemoLabour.Tax} disabled={true} />
              </td>
              <td>
                <TextField value={taxesInfo.DemoLabour.Total} disabled={true} />
              </td>
            </tr>
            <tr style={{ backgroundColor: '#fdefeb' }}>
              <td>Sub Total</td>
              <td>
                <TextField
                  value={taxesInfo.DemoSubTotal.PreTax}
                  disabled={true}
                />
              </td>
              <td>
                <TextField value={taxesInfo.DemoSubTotal.Tax} disabled={true} />
              </td>
              <td>
                <TextField
                  value={taxesInfo.DemoSubTotal.Total}
                  disabled={true}
                />
              </td>
            </tr>
            <tr>
              <td>Freight</td>
              <td>
                <TextField value={taxesInfo.Freight.PreTax} disabled={true} />
              </td>
              <td>
                <TextField value={taxesInfo.Freight.Tax} disabled={true} />
              </td>
              <td>
                <TextField value={taxesInfo.Freight.Total} disabled={true} />
              </td>
            </tr>
            <tr>
              <td>Spring Handling</td>
              <td>
                <TextField
                  value={taxesInfo.SpringHandling.PreTax}
                  disabled={true}
                />
              </td>
              <td>
                <TextField
                  value={taxesInfo.SpringHandling.Tax}
                  disabled={true}
                />
              </td>
              <td>
                <TextField
                  value={taxesInfo.SpringHandling.Total}
                  disabled={true}
                />
              </td>
            </tr>
            <tr>
              <td>Profit & OH</td>
              <td>
                <TextField value={taxesInfo.ProfitOH.PreTax} disabled={true} />
              </td>
              <td>
                <TextField value={taxesInfo.ProfitOH.Tax} disabled={true} />
              </td>
              <td>
                <TextField value={taxesInfo.ProfitOH.Total} disabled={true} />
              </td>
            </tr>
            <tr>
              <td>Insurance</td>
              <td>
                <TextField value={taxesInfo.Insurance.PreTax} disabled={true} />
              </td>
              <td>
                <TextField value={taxesInfo.Insurance.Tax} disabled={true} />
              </td>
              <td>
                <TextField value={taxesInfo.Insurance.Total} disabled={true} />
              </td>
            </tr>
            <tr style={{ backgroundColor: '#fdefeb' }}>
              <td>Total</td>
              <td>
                <TextField value={taxesInfo.Total.PreTax} disabled={true} />
              </td>
              <td>
                <TextField value={taxesInfo.Total.Tax} disabled={true} />
              </td>
              <td>
                <TextField value={taxesInfo.Total.Total} disabled={true} />
              </td>
            </tr>
          </tbody>
        </table>
      </div>
      <div className={styles.SubmitSection}>
        <DefaultButton
          text="Back"
          onClick={() => {
            history.back()
          }}
        />
      </div>
    </ThemeProvider>
  )

  function getJsonData() {
    var jsonObj = {
      mainTitle: 'Wells Fargo Quote Form for HVAC Projects',
      tableOneTitle:
        'PROJECT/ WO INFORMATION (Information provided by Wells Fargo)',
      tableTwoTitle: "VENDOR'S AUTHORIZED REPRESENTATIVE",
      dateSubmitted: 'XXXXX',
      projectOrWorkOrder: projectInfo.projNoInput,
      wfProjectOrPropertyManager: projectInfo.projManagerInput,
      beNumber: projectInfo.BENoInput,
      buildingOrProjectName: projectInfo.ProjNameInput,
      beServiceOrDeliveryAddress: projectInfo.DeliveryAddInput,
      projectArea: projectInfo.projAreaInput,
      estimatedStartDate: projectInfo.EstimateStartDateInput.toLocaleDateString(),
      estimatedCompleteDate: projectInfo.EstimateEndDateInput.toLocaleDateString(),
      companyName: 'Lynxspring Inc',
      remitToAddress: '2900 NE Independence Ave',
      cityStateZip: "Lee's Summit, MO 64086",
      contactName: vendorInfo.contactNameInput,
      phone: '816 347 3500',
      cell: '913 219 5513',
      email: vendorInfo.emailIdInput,
      wfVendOrNumber: vendorInfo.wfVendorNoInput,
      proposalNumber: vendorInfo.proposalNoInput,
      wfContractNumber: vendorInfo.wfContractNoInput,
      changeOrder: vendorInfo.changeOrderInput,
      changeOrderPreviousPO: vendorInfo.changeOrderPOInput,
      scopeTitle: 'SCOPE OF WORK',
      scopeSubHeadOne:
        'Detail Needed- Be as descriptive as possible and if there are multiples of items please specify how many.  Details and numbers of units help fixed assets determine asset value and will help eliminate questions and the need to resubmit proposals or invoices. ',
      scopeSubDescription:
        'Provide JENEsys Hardware as submitted by Wells Fargo on JENE order forms.  Pre-Mount equipment in v 2.54 Admin Panels and ship to site for Site Electrical/BMS Contractor to Install. Site Contractor responsible to provide on site integration, graphics, schedules, alarms and configuration.  Station to be returned to Lynxspring for review.  BExxx Site to have xxxxxxxxxxx.',
      scopeSubDescription_D1:
        'Provide JENEsys Hardware as submitted by Wells Fargo on JENE order forms.  Pre-Mount equipment in v 2.54 Admin Panels and ship to site for Site Electrical/BMS Contractor to Install. Site Contractor responsible to provide on site integration, graphics, schedules, alarms and configuration.  Station to be returned to Lynxspring for review.',
      scopeSubDescription_D2: 'BExxx Site to have xxxxxxxxxxx.',
      installHeading: 'INSTALLATION',
      detailSubHeadOne: 'DETAILED DESCRIPTION',
      productSubHeadTwo: 'PRODUCT / MATERIAL / EQUIPMENT',
      laborSubHeadThree: 'LABOR / SERVICE',
      installHeadingColumn1:
        'WF needs to see the Labor Costs associated with the Product.  If Quote is for labor only, leave Product Section blank.  Include separate lines for HVAC components as well  as separate lines for equipent and materials related to that component.Include separate lines for controls components (i.e. controller devices such as JENEsys, thermostats, miscellaneous materials). Include labor associated with specific equipment on the same line.',
      installHeadingColumn1_D1:
        'WF needs to see the Labor Costs associated with the Product.  If Quote is for labor only, leave Product Section blank.',
      installHeadingColumn1_D2:
        'Include separate lines for HVAC components as well  as separate lines for equipent and materials related to that component.\n\n',
      installHeadingColumn1_D3:
        'Include separate lines for controls components (i.e. controller devices such as JENEsys, thermostats, miscellaneous materials). Include labor associated with specific equipment on the same line.',
      installHeadingColumn2: 'Manufacturer',
      installHeadingColumn3: 'Model #',
      installHeadingColumn4: 'Serial # (if available)',
      installHeadingColumn5: 'Qty.',
      installHeadingColumn6: 'Each Price ',
      installHeadingColumn7: 'Each Mark-Up',
      installHeadingColumn8: 'TOTAL PRODUCT',
      installHeadingColumn9: 'Taxable? (Y/N)',
      installHeadingColumn10: '# People',
      installHeadingColumn11: '# Hours per Person',
      installHeadingColumn12: 'Hourly Bill Rate',
      installHeadingColumn13: 'Union Rate(Y/N)',
      installHeadingColumn14: 'TOTAL LABOR',
      installHeadingColumn15: 'Labor Taxable? (Y/N)',
      installHeadingColumn16: 'TOTAL PRODUCT & LABOR',
      installColumns: [],

      totalProduct: '$ - ' + projectInfo.InstallTotalProduct,
      totalPeople: projectInfo.InstallTotalPeople,
      totalHoursPerPerson: projectInfo.InstallTotalHoursPerPerson,
      totalHourlyBillRate: 'xxxxx',
      totalUnionRate: 'xxxxx',
      totalLabor: '$ -',
      totalLaborTaxable: 'xxxxx',
      totalProductAndLabor: '$ - ' + projectInfo.InstallSecondTotalProduct,

      demoHeading: 'DEMO / REMOVAL / PATCHING / REPAIRS / RELO ',
      demoSubHeadOne: 'DEMO - PRODUCT/ MATERIALS',
      demoSubHeadTwo: 'DEMO - LABOR',
      demoHeadingColumn1: 'Include costs of demolition and demolition labor.',
      demoHeadingColumn2: '',
      demoHeadingColumn3: '',
      demoHeadingColumn4: 'Qty.',
      demoHeadingColumn5: 'Price Each',
      demoHeadingColumn6: 'TOTAL DEMO',
      demoHeadingColumn7: 'Taxable? (Y/N)',
      demoHeadingColumn8: '# People',
      demoHeadingColumn9: '# Hours per Person',
      demoHeadingColumn10: 'Hourly Bill Rate',
      demoHeadingColumn11: 'Union Rate(Y/N)',
      demoHeadingColumn12: 'TOTAL LABOR',
      demoHeadingColumn13: 'Labor Taxable? (Y/N)',
      demoHeadingColumn14: 'TOTAL DEMO PRODUCT & LABOR',
      demoColumns: [],
      totalDemo: '$ -',
      totalDemoPeople: '0',
      totalDemoHoursPerPerson: '0',
      totalDemoHourlyBillRate: 'xxxxx',
      totalDemoUnionRate: 'xxxxx',
      totalDemoLabor: '$ -',
      totalDemoLaborTaxable: 'xxxxx',
      totalDemoProductAndLabor: '$ -',
      clarificationsHeading: 'Assumptions and Clarifications ',
      clarificationDescription:
        ' Assumes local contractor will integrate all existing devices and complete all necessary graphics.',
      taxHeading: 'Manually Insert Tax $ as required',
      taxHeadingColoumn1: 'ENTER TAX RATE ',
      taxHeadingColoumn2: 'PRE-TAX',
      taxHeadingColoumn3: 'TAX',
      totalTaxHeading: 'TOTAL',
      taxRate: '0.00% ',
      product: 'PRODUCT  ',
      productPreTax: '$   - ' + taxesInfo.Product.PreTax,
      productTax: '$   - ' + taxesInfo.Product.Tax,
      productTotal: '$   - ' + taxesInfo.Product.Total,
      labor: 'LABOR  ',
      laborPreTax: '$   - ' + taxesInfo.Labour.PreTax,
      laborTax: '$   - ' + taxesInfo.Labour.Tax,
      laborTotal: '$   - ' + taxesInfo.Labour.Total,
      subTotal1: 'SUBTOTAL  ',
      subTotal1PreTax: '$   - ' + taxesInfo.ProductSubTotal.PreTax,
      subTotal1Tax: '$   - ' + taxesInfo.ProductSubTotal.Tax,
      subTotal1Total: '$   -' + taxesInfo.ProductSubTotal.Total,
      demoProduct: 'DEMO PRODUCT  ',
      demoProductPreTax: '$   - ' + taxesInfo.DemoProduct.PreTax,
      demoProductTax: '$   - ' + taxesInfo.DemoProduct.Tax,
      demoProductTotal: '$   - ' + taxesInfo.DemoProduct.Total,
      demoLabor: 'DEMO LABOR  ',
      demoLaborPreTax: '$   - ' + taxesInfo.DemoLabour.PreTax,
      demoLaborTax: '$   - ' + taxesInfo.DemoLabour.Tax,
      demoLaborTotal: '$   - ' + taxesInfo.DemoLabour.Total,
      subTotal2: 'SUBTOTAL  ',
      subTotal2PreTax: '$   - ' + taxesInfo.DemoSubTotal.PreTax,
      subTotal2Tax: '$   - ' + taxesInfo.DemoSubTotal.Tax,
      subTotal2Total: '$   - ' + taxesInfo.DemoSubTotal.Total,
      freight: 'FREIGHT  ',
      freightPreTax: '$   - ' + taxesInfo.Freight.PreTax,
      freightTax: '$   - ' + taxesInfo.Freight.Tax,
      freightTotal: '$   - ' + taxesInfo.Freight.Total,
      shipping: 'SHIPPING & HANDLING  ',
      shippingPreTax: '$   - ' + taxesInfo.SpringHandling.PreTax,
      shippingTax: '$   - ' + taxesInfo.SpringHandling.Tax,
      shippingTotal: '$   - ' + taxesInfo.SpringHandling.Total,
      profit: 'PROFIT & OH  ',
      profitPreTax: '$   - ' + taxesInfo.ProfitOH.PreTax,
      profitTax: '$   - ' + taxesInfo.ProfitOH.Tax,
      profitTotal: '$   -' + taxesInfo.ProfitOH.Total,
      insurance: 'INSURANCE  ',
      insurancePreTax: '$   - ' + taxesInfo.Insurance.PreTax,
      insuranceTax: '$   - ' + taxesInfo.Insurance.Tax,
      insuranceTotal: '$   - ' + taxesInfo.Insurance.Total,
      allTotal: 'TOTAL  ',
      allTotalTaxPreTax: '$   - ' + taxesInfo.Total.PreTax,
      allTotalTax: '$   - ' + taxesInfo.Total.Tax,
      allMaterialTotal: '$   - ' + taxesInfo.Total.Total,
      rotationText: 'if \n applicable',
    }

    var installationTable = []
    $.each(installationtable, function (key, val) {
      installationTable.push({
        coloumn1: 'Detailed Description',
        coloumn2: val.manufacture,
        coloumn3: val.modelNo,
        coloumn4: val.serialNo,
        coloumn5: val.quantity,
        coloumn6: val.eachPrice,
        coloumn7: '$  - ' + val.eachMarkup,
        coloumn8: '$  - ' + val.totalProduct,
        coloumn9: val.taxable ? 'Y' : 'N',
        coloumn10: val.people,
        coloumn11: val.hoursPerPerson,
        coloumn12: val.hourlyBillingRate,
        coloumn13: val.unionRate ? 'Y' : 'N',
        coloumn14: '$  - ' + val.totalLabour,
        coloumn15: val.labourTaxable ? 'Y' : 'N',
        coloumn16: '$  - ' + val.grandTotalProduct,
      })
    })

    var demoTable = []

    demoTable.push({
      coloumn1: '',
      coloumn2: '',
      coloumn3: '',
      coloumn4: '',
      coloumn5: '$    -',
      coloumn6: '$    -',
      coloumn7: '',
      coloumn8: '',
      coloumn9: '',
      coloumn10: '$    -',
      coloumn11: '',
      coloumn12: '$    -',
      coloumn13: '',
      coloumn14: '$    -',
    })

    jsonObj.demoColumns = demoTable
    jsonObj.installColumns = installationTable

    return jsonObj
  }

  async function downloadFile(URL, fileName, filetype) {
    setLoader(true)
    var jsonData = getJsonData()
    console.log(jsonData)
    //var fileName = 'demo.pdf';
    $.ajax({
      type: 'POST',
      cache: false,
      url: URL,
      data: jsonData,
      /*headers:
    {
      "accept": "application/json;odata=verbose",
      "content-Type": "application/json;odata=verbose"
    },*/
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
            setLoader(false)
            console.log(data)
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

        /*var downloadLink = document.createElement('a');
    var blob = new Blob([data],
    {
    type: xmlHeaderRequest.getResponseHeader('Content-Type')
    });
    var url = window.URL || window.webkitURL;
    var downloadUrl = url.createObjectURL(blob);

    if (typeof window.navigator.msSaveBlob !== 'undefined') {
    window.navigator.msSaveBlob(blob, fileName);
    } else {
    if (fileName) {
    if (typeof downloadLink.download === 'undefined')
    {
      window.location.href = downloadUrl;

    } else {
    downloadLink.href = downloadUrl;
    downloadLink.download = fileName;
    document.body.appendChild(downloadLink);
    downloadLink.click();
    }
    } else
    {
      window.location.href = downloadUrl;
    }



    setTimeout(function () {
    url.revokeObjectURL(downloadUrl);
    },
    100);
    }*/

        /*var url = (window.URL || window.webkitURL).createObjectURL(blob);
    var link = document.createElement("a");
    link.setAttribute("href", url+"&download=1");
    link.setAttribute("target", "_blank");
    link.setAttribute("download", fileName);
    //link.style = "visibility:hidden";
    document.body.appendChild(link);
    link.click();
    setTimeout(function(){ document.body.removeChild(link); }, 500);*/
      })
      .catch(function (jqXHR, textStatus, errorThrown) {
        alert('Error while downloading File.Please contact admin')
        console.log('Response from File API Failed')
        console.log(JSON.stringify(jqXHR))
        console.log(JSON.stringify(textStatus))
        console.log(JSON.stringify(errorThrown))
      })
  }
}
export default WFQuoteView
