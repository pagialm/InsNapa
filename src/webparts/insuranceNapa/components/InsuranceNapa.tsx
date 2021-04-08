import * as React from "react";
import styles from "./InsuranceNapa.module.scss";
import { IInsuranceNapaProps } from "./IInsuranceNapaProps";
import { escape } from "@microsoft/sp-lodash-subset";
import MenuIcon from "./MenuIcon";
import HeaderInfo from "./HeaderInfo";
import {
  TextField,
  Checkbox,
  ICheckboxProps,
  Dropdown,
  DropdownMenuItemType,
  IDropdownStyles,
  IDropdownOption,
  DatePicker,
  IStackStyles,
  IStackProps,
  Stack,
  autobind,
  Toggle,
  Separator,
  DefaultButton,
  PrimaryButton,
  ITextFieldProps,
  getTheme,
  FontWeights,
  ITheme,
  Label,
  Button,
  BaseButton,
  MessageBar,
  MessageBarType,
  MessageBarButton,
} from "office-ui-fabric-react/lib";
import {
  PeoplePicker,
  PrincipalType,
} from "@pnp/spfx-controls-react/lib/PeoplePicker";
import { InsuranceNapaState } from "./InsuranceNapaState";
import {
  SPHttpClient,
  SPHttpClientResponse,
  ISPHttpClientOptions,
} from "@microsoft/sp-http";
import { IProposal } from "./IProposal";
import FilteredDropdown from "./FilteredDropdown";
import { HttpRequestError } from "@pnp/odata";
import * as HeadersDecor from "./Headers";
const stackTokens = { childrenGap: 50 };
const stackStyles: Partial<IStackStyles> = { root: { width: 784 } };
const columnProps: Partial<IStackProps> = {
  tokens: { childrenGap: 10 },
  styles: { root: { width: 450 } },
};
const dropdownStyles: Partial<IDropdownStyles> = { dropdown: { width: 300 } };
const proposalObj = {
  countriesListname: "Countries",
  productsListname: "Entities",
  jurisdictionListname: "Jurisdiction",
  targetClientSectorListname: "Target Client Sector",
  currenciesListname: "Currencies",
  natureOfTradeListname: "Nature Of Trade",
  bookingLegalEntitiesListname: "Booking Legal Entities",
  distributionChannelListname: "Distribution Channels",
  napaProposalsListname: "NAPA Proposals",
  userObjects: [],
  userObjectsCount: 0,
  //proposalObj: {},
};
const proposalRegions: IDropdownOption[] = [
  { key: "ARO", text: "ARO" },
  { key: "SA", text: "SA" },
  // { key: "UK", text: "UK" },
  // { key: "USA", text: "USA" },
];
const productFamRiskClass: IDropdownOption[] = [
  { key: "High", text: "High" },
  { key: "Medium", text: "Medium" },
  { key: "Low", text: "Low" },
];
const SuccessExample = () => (
  <MessageBar messageBarType={MessageBarType.success} isMultiline={false}>
    You NPA has been successfully submitted!
  </MessageBar>
);
const ErrorExample = (message: string) => (
  <MessageBar
    messageBarType={MessageBarType.error}
    isMultiline={false}
    dismissButtonAriaLabel="Close"
  >
    {message}
  </MessageBar>
);

export default class InsuranceNapa extends React.Component<
  IInsuranceNapaProps,
  InsuranceNapaState
> {
  constructor(props: IInsuranceNapaProps, state: InsuranceNapaState) {
    super(props);
    this.state = {
      allCountries: [],
      shortCountries: [],
      bookingCurrencies: [],
      tradeActivities: [],
      legalEntities: [],
      users: [],
      clientSectors: [],
      companies: [],
      businessAreas: [],
      productAreas: [],
      subProducts: [],
      proposalObject: {},
      applicationCompletedBy: "",
      sponser: "",
      tradingBookOwner: "",
      workstreamCoordinator: "",
      targetCompletionDate: null,
      proposalObj: {},
      ID: 0,
      Title: "",
      TargetCompletionDate: null,
      AppCreatedById: 0,
      SponsorId: 0,
      TradingBookOwnerId: 0,
      WorkStreamCoordinatorId: 0,
      Region: [],
      Country0: "",
      Company: "",
      BusinessArea: "",
      ExecutiveSummary: "",
      ProductArea0: "",
      SubProduct: "",
      NewForProposal: "",
      TransactionInPipeline: "",
      LinkToExistingProposal: "",
      TaxTreatment: "",
      LineOfCredit: "",
      ConductRiskIssuesComments: "",
      PrincipalRisks: "",
      IFCountry: [],
      SalesTeamLocation: [],
      ClientLocation: [],
      ClientSector: [],
      ProductOfferingCountry: [],
      BookingCurrencies: [],
      BookingLocation: [],
      NatureOfTrade: "",
      TraderLocation: [],
      BookingEntity: [],
      JointVenture: false,
      Status: "",
      distributionChannels: [],
      submitionStatus: "",
      errorMessage: [],
    };
  }
  public async componentDidMount(): Promise<void> {
    const handler = this;
    const allCountriesArr: IDropdownOption[] = [];
    const someCountries: IDropdownOption[] = [];
    const clientSectorsArr: IDropdownOption[] = [];
    const bookingCurrenciesArr: IDropdownOption[] = [];
    const tradeActivityArr: IDropdownOption[] = [];
    const legalEntityArr: IDropdownOption[] = [];
    const distributionChannelsArr: IDropdownOption[] = [];

    //load countries
    this._getListitems(proposalObj.countriesListname).then((allCountries) => {
      allCountries.forEach((country) => {
        if (country.IsProposalCountry)
          someCountries.push({ key: country.Title, text: country.Title });
        allCountriesArr.push({ key: country.Title, text: country.Title });
      });
      this.setState({ allCountries: allCountriesArr });
      this.setState({ shortCountries: someCountries });
    });
    //load Target Client Sector
    this._loadDropdownFromSP(
      proposalObj.targetClientSectorListname,
      clientSectorsArr,
      { clientSectors: clientSectorsArr }
    );

    //load Booking/Applicable Currencies
    this._loadDropdownFromSP(
      proposalObj.currenciesListname,
      bookingCurrenciesArr,
      { bookingCurrencies: bookingCurrenciesArr }
    );

    //load Nature of Trade Activity
    this._loadDropdownFromSP(
      proposalObj.natureOfTradeListname,
      tradeActivityArr,
      { tradeActivities: tradeActivityArr }
    );

    //load Booking Legal Entity
    this._loadDropdownFromSP(
      proposalObj.bookingLegalEntitiesListname,
      legalEntityArr,
      { legalEntities: legalEntityArr }
    );
    //this.setState({ legalEntities: legalEntityArr });
    // load distribution channels
    this._loadDropdownFromSP(
      proposalObj.distributionChannelListname,
      distributionChannelsArr,
      { distributionChannels: distributionChannelsArr }
    );

    //load Company --Entity
    this._loadCompany();

    // get Proposal
    if (this.props.itemId)
      this._getListitem(
        proposalObj.napaProposalsListname,
        this.props.itemId
      ).then((item) => {
        const _item: IProposal = item as IProposal;
        for (const key in _item) {
          if (Object.prototype.hasOwnProperty.call(_item, key)) {
            const element = _item[key];
            var newEl = {};
            newEl[key] = element;
            try {
              this.setState(newEl);
            } catch (err) {}
          }
        }
        this.setState({ proposalObject: _item });
        this.setState({ proposalObj: _item });
        const userObjects = [
          { stateName: "applicationCompletedBy", itemName: "AppCreatedById" },
          { stateName: "sponser", itemName: "SponsorId" },
          { stateName: "tradingBookOwner", itemName: "TradingBookOwnerId" },
          {
            stateName: "workstreamCoordinator",
            itemName: "WorkStreamCoordinatorId",
          },
        ];
        const dateObjects = [
          {
            stateName: "targetCompletionDate",
            itemName: "TargetCompletionDate",
          },
        ];
        const newstate = {};
        userObjects.forEach((userObject) => {
          this._getUserById(_item[userObject.itemName]).then((user) => {
            newstate[userObject.stateName] = user[0].Title;
            this.setState(newstate);
          });
        });
        dateObjects.forEach((dateObject) => {
          if (_item[dateObject.itemName]) {
            newstate[dateObject.stateName] = new Date(
              _item[dateObject.itemName]
            );
            this.setState(newstate);
          }
        });
        console.log(this.state);
        console.log("item", item);
      });
  }
  private _getUserById(userId: number | string): Promise<any> {
    const url: string =
      this.props.context.pageContext.site.absoluteUrl +
      "/_api/web/siteusers?$filter=ID eq " +
      userId;
    return this.props.context.spHttpClient
      .get(url, SPHttpClient.configurations.v1)
      .then((response) => {
        return response.json();
      })
      .then((jsonResponse) => {
        return jsonResponse.value;
      }) as Promise<any[]>;
  }
  private _loadDropdownFromSP(
    listName: string,
    dropdownArray: IDropdownOption[],
    stateObject: {}
  ) {
    this._getListitems(listName).then((responseItems) => {
      responseItems.forEach((responseItem) => {
        if (responseItem.Alphabetic_x0020_Code) {
          dropdownArray.push({
            key: responseItem.Alphabetic_x0020_Code,
            text: responseItem.Alphabetic_x0020_Code,
          });
        } else
          dropdownArray.push({
            key: responseItem.Title,
            text: responseItem.Title,
          });
        this.setState(stateObject);
      });
    });
  }
  private _getListitems(listName: string): Promise<any[]> {
    const url: string =
      this.props.context.pageContext.site.absoluteUrl +
      "/_api/web/lists/getbytitle('" +
      listName +
      "')/items";
    return this.props.context.spHttpClient
      .get(url, SPHttpClient.configurations.v1)
      .then((response) => {
        return response.json();
      })
      .then((jsonResponse) => {
        return jsonResponse.value;
      }) as Promise<any[]>;
  }
  private _getListitem(listName: string, itemId: number): Promise<any> {
    const url: string =
      this.props.context.pageContext.site.absoluteUrl +
      "/_api/web/lists/getbytitle('" +
      listName +
      "')/items(" +
      itemId +
      ")";
    return this.props.context.spHttpClient
      .get(url, SPHttpClient.configurations.v1)
      .then((response) => {
        return response.json();
      })
      .then((jsonResponse) => {
        return jsonResponse;
      }) as Promise<any>;
  }
  private _getListitemsFilter(
    listName: string,
    filter: string
  ): Promise<any[]> {
    const url: string =
      this.props.context.pageContext.site.absoluteUrl +
      "/_api/web/lists/getbytitle('" +
      listName +
      "')/items?$filter=" +
      filter;
    return this.props.context.spHttpClient
      .get(url, SPHttpClient.configurations.v1)
      .then((response) => {
        return response.json();
      })
      .then((jsonResponse) => {
        return jsonResponse.value;
      }) as Promise<any[]>;
  }
  private _loadCompany() {
    this._getListitems(proposalObj.productsListname).then((products) => {
      const productsArray: IDropdownOption[] = [];
      products.forEach((product) => {
        if (
          !productsArray.some((item) => {
            return item.text === product.Title;
          })
        )
          productsArray.push({ key: product.Title, text: product.Title });
      });
      this.setState({ companies: productsArray });
    });
  }
  @autobind
  private _laodDropdown(
    ev: React.FormEvent<HTMLDivElement>,
    option?: IDropdownOption,
    index?: number,
    extras?: any
  ): void {
    console.log(ev, option, index, extras);
  }
  @autobind
  private _onChangeToggle(
    ev: React.MouseEvent<HTMLElement, MouseEvent>,
    checked: boolean
  ): void {
    console.log(ev, checked);
    const el = ev.currentTarget;
    const stateEl = {};
    stateEl[el.id.split("_")[1]] = checked;
    console.log(stateEl);
    this.setState(stateEl);
  }
  @autobind
  private _loadFilteredDropdown(fieldName: string, columnName: string) {
    const companyName = document.getElementById("ddlCompany").firstChild
      .textContent!;
    this._getListitemsFilter(
      proposalObj.productsListname,
      "Title eq '" + companyName + "'"
    ).then((filteredItems) => {
      const arrayToLoad: IDropdownOption[] = [
        { key: "0", text: "", selected: true },
      ];
      filteredItems.forEach((filteredItem) => {
        //debugger;
        if (filteredItem[columnName]) {
          if (
            !arrayToLoad.some((item) => {
              return item.text === filteredItem[columnName];
            })
          )
            arrayToLoad.push({
              key: filteredItem[columnName],
              text: filteredItem[columnName],
            });
        }
      });

      if (fieldName === "ddlProductArea")
        this.setState({ productAreas: arrayToLoad });
      if (fieldName === "ddlBusinessArea")
        this.setState({ businessAreas: arrayToLoad });
      if (fieldName === "ddlSubProducts")
        this.setState({ subProducts: arrayToLoad });
    });
  }
  @autobind
  private _getPeoplePickerItems(items: any[]): any[] {
    debugger;
    let getSelectedUsers = [];
    for (let item in items) {
      getSelectedUsers.push(items[item].id);
    }
    console.log(getSelectedUsers);
    console.log(this);
    //this.setState({ users: getSelectedUsers });
    return getSelectedUsers;
  }
  @autobind
  private _onChange(
    ev: React.FormEvent<HTMLDivElement>,
    option?: IDropdownOption,
    index?: number
  ): void {
    const elementId = ev.currentTarget.id.split("_")[1].split("-")[0];
    const el = {};
    el[elementId] = option.key;
    this.setState(el);
  }
  @autobind
  private _onChangeText(
    ev: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>,
    newValue: string
  ): void {
    const element = ev.target as HTMLElement;

    const el = {};
    el[element.id.split("_")[1]] = newValue;
    this.setState(el);
  }
  private _validateForm(): boolean {
    let isValid = true;
    debugger;
    const elements = document.querySelectorAll("[required]");
    const dropElements = document.querySelectorAll("[aria-required]");
    let errMsg = "";
    const errMsgs = [];
    elements.forEach((el: HTMLInputElement | HTMLTextAreaElement) => {
      if (!el.value) {
        isValid = false;
        errMsgs.push(
          `${el.parentElement.previousSibling.textContent} is required.`
        );
      }
    });
    dropElements.forEach((el: HTMLDivElement) => {
      if (
        !el.firstChild.textContent ||
        el.firstChild.textContent === "Select options"
      ) {
        isValid = false;
        errMsgs.push(`${el.previousSibling.textContent} is required.`);
      }
    });

    if (isValid) {
      this.setState({ errorMessage: [] });
    } else {
      this.setState({
        errorMessage: errMsgs,
      });
    }

    return isValid;
  }
  @autobind
  private _saveApplicationProposal(e) {
    // debugger;
    console.log(this);
    const buttonClicked: string = e.target.innerText;
    let statusText = "NPS Determination";
    if (buttonClicked === "Save as Draft") statusText = "Enquiry";
    const _isFormValid =
      buttonClicked === "Save as Draft" ? true : this._validateForm();
    const proposal = {};

    proposal["Title"] = this.state.Title;
    proposal["TargetCompletionDate"] = this.state.targetCompletionDate;
    if (this.state.AppCreatedById)
      proposal["AppCreatedById"] = this.state.AppCreatedById;
    if (this.state.SponsorId) proposal["SponsorId"] = [this.state.SponsorId];
    if (this.state.TradingBookOwnerId)
      proposal["TradingBookOwnerId"] = this.state.TradingBookOwnerId;
    if (this.state.WorkStreamCoordinatorId)
      proposal["WorkStreamCoordinatorId"] = [
        this.state.WorkStreamCoordinatorId,
      ];
    if (this.state.Region.length > 0) proposal["Region"] = [this.state.Region];
    proposal["Country0"] = this.state.Country0;
    proposal["Company"] = this.state.Company;
    proposal["BusinessArea"] = this.state.BusinessArea;
    proposal["ExecutiveSummary"] = this.state.ExecutiveSummary;
    proposal["ProductArea0"] = this.state.ProductArea0;
    if (this.state.SubProduct) proposal["SubProduct"] = this.state.SubProduct;
    proposal["NewForProposal"] = this.state.NewForProposal;
    proposal["TransactionInPipeline"] = this.state.TransactionInPipeline;
    proposal["LinkToExistingProposal"] = this.state.LinkToExistingProposal;
    proposal["TaxTreatment"] = this.state.TaxTreatment;
    proposal["LineOfCredit"] = this.state.LineOfCredit;
    proposal[
      "ConductRiskIssuesComments"
    ] = this.state.ConductRiskIssuesComments;
    proposal["PrincipalRisks"] = this.state.PrincipalRisks;
    if (this.state.IFCountry.length > 0)
      proposal["IFCountry"] = [this.state.IFCountry];
    if (this.state.SalesTeamLocation.length > 0)
      proposal["SalesTeamLocation"] = [this.state.SalesTeamLocation];
    if (this.state.ClientLocation.length > 0)
      proposal["ClientLocation"] = [this.state.ClientLocation];
    if (this.state.ClientSector.length > 0)
      proposal["ClientSector"] = [this.state.ClientSector];
    if (this.state.ProductOfferingCountry.length > 0)
      proposal["ProductOfferingCountry"] = [this.state.ProductOfferingCountry];
    if (this.state.BookingCurrencies.length > 0)
      proposal["BookingCurrencies"] = [this.state.BookingCurrencies];
    if (this.state.BookingLocation.length > 0)
      proposal["BookingLocation"] = [this.state.BookingLocation];
    if (this.state.NatureOfTrade.length > 0)
      proposal["NatureOfTrade"] = this.state.NatureOfTrade;
    if (this.state.TraderLocation.length > 0)
      proposal["TraderLocation"] = [this.state.TraderLocation];
    if (this.state.BookingEntity.length > 0)
      proposal["BookingEntity"] = [this.state.BookingEntity];
    proposal["JointVenture"] = this.state.JointVenture;
    proposal["Status"] = statusText;

    const url: string =
      this.props.context.pageContext.web.absoluteUrl +
      "/_api/web/lists/getbytitle('" +
      proposalObj.napaProposalsListname +
      "')/items";
    const spHttpClientOptions: ISPHttpClientOptions = {
      body: JSON.stringify(proposal),
    };

    if (this.state.ID) {
      console.log("Item to be updated");
    } else {
      if (_isFormValid) {
        this.props.context.spHttpClient
          .post(url, SPHttpClient.configurations.v1, spHttpClientOptions)
          .then((response: SPHttpClientResponse) => {
            if (response.status === 201) {
              console.log("Ayeye!!!", response);
              this.setState({ submitionStatus: "Ok" });
              //location.href = this.props.context.pageContext.web.absoluteUrl;
            } else {
              this.setState({
                errorMessage: [
                  `Error: [HTTP]:${response.status} [CorrelationId]:${response.statusText}`,
                ],
              });
            }
          });
      }
    }
  }
  @autobind
  private _cancelProposal() {
    location.href = this.props.context.pageContext.web.absoluteUrl;
  }
  @autobind
  private _onSelectDate(date: Date | null | undefined): void {
    this.setState({ targetCompletionDate: date });
  }
  @autobind
  private _onFormatDate(date: Date): string {
    return (
      date.getDate() + "/" + (date.getMonth() + 1) + "/" + date.getFullYear()
    );
  }
  @autobind
  private _onFilteredDropdownChange(
    ev: React.FormEvent<HTMLDivElement>,
    option?: IDropdownOption,
    index?: number
  ): void {
    console.log("ev:", ev);
    console.log("option:", option);
    console.log("index:", index);
    debugger;
    let resultarr: IDropdownOption[] = [];
    const filteredDropdown = [
      {
        eleId: "ddl_Company",
        staticName: "Company",
        fieldName: "Title",
        fieldToFilter: "ProductFamily",
        fieldToFilterOptionsStateName: "businessAreas",
      },
    ];
    const selectedElement = ev.target as HTMLDivElement;
    const fieldObject = filteredDropdown.filter((dd) => {
      return dd.eleId === selectedElement.id;
    });
    const currElementState = {};
    const _key = fieldObject[0].staticName;
    currElementState[_key] = option.key;
    this.setState(currElementState);
    //this.setState({proposalObject.})
    this._getListitemsFilter(
      proposalObj.productsListname,
      `${fieldObject[0].fieldName} eq '${option.text}'`
    ).then((items) => {
      items.forEach((item) => {
        if (
          !resultarr.some((dd) => {
            return dd.key === item[fieldObject[0].fieldToFilter];
          })
        ) {
          resultarr.push({
            key: item[fieldObject[0].fieldToFilter],
            text: item[fieldObject[0].fieldToFilter],
          });
        }
      });
      const stateObj = {};
      stateObj[fieldObject[0].fieldToFilterOptionsStateName] = resultarr;
      this.setState(stateObj);
    });
  }
  public render(): React.ReactElement<IInsuranceNapaProps> {
    return (
      <div className={styles.insuranceNapa}>
        {/* <div className={styles.menuItem}>Enquiery</div> */}
        <MenuIcon iconName="Headset" stageName="Enquiry" activated={false} />

        <Stack>
          {this.state.ID > 0 && (
            <HeadersDecor.default
              proposalStatus={this.state.Status}
              proposalId={this.state.ID}
              selectedSection="Enquiry"
              title={this.state.Title}
            />
          )}

          <HeaderInfo
            title="Application Information"
            description="Provide the following administrative information"
          />
          <Stack horizontal tokens={stackTokens} styles={stackStyles}>
            <Stack {...columnProps}>
              <TextField
                label="Proposal Name:"
                required
                value={this.state.Title}
                id="txt_Title"
                onChange={this._onChangeText}
                description="The proposal name should contain the key distinguishing attributes associated with the proposal."
              />
              <PeoplePicker
                context={this.props.context}
                titleText="Application completed by"
                personSelectionLimit={3}
                showtooltip={true}
                disabled={false}
                defaultSelectedUsers={[this.state.applicationCompletedBy]}
                onChange={(items: any[]) => {
                  const _users = this._getPeoplePickerItems(items);
                  debugger;
                  if (_users.length > 0)
                    this.setState({
                      AppCreatedById: _users[0],
                    });
                }}
                // selectedItems={this._getPeoplePickerItems }
                showHiddenInUI={false}
                ensureUser={true}
                principalTypes={[PrincipalType.User]}
                resolveDelay={1000}
              />
              <Label className={styles.inputDesc}>
                The Product Originator is a business representative or business
                manager, or his/her delegate. The Originator will also be the
                person raising the Proposal Template.
              </Label>
              <PeoplePicker
                context={this.props.context}
                titleText="P&L Owner/ General Manager"
                personSelectionLimit={3}
                showtooltip={true}
                defaultSelectedUsers={[this.state.tradingBookOwner]}
                disabled={false}
                onChange={(items: any[]) => {
                  const _users = this._getPeoplePickerItems(items);
                  if (_users.length > 0)
                    this.setState({
                      TradingBookOwnerId: _users[0],
                    });
                }}
                showHiddenInUI={false}
                ensureUser={true}
                principalTypes={[PrincipalType.User]}
                resolveDelay={1000}
              />
            </Stack>
            <Stack {...columnProps}>
              <DatePicker
                label="Target Launch Date:"
                isRequired
                value={this.state.targetCompletionDate}
                onSelectDate={this._onSelectDate}
                formatDate={this._onFormatDate}
              />
              <PeoplePicker
                context={this.props.context}
                titleText="Sponsor"
                personSelectionLimit={3}
                showtooltip={true}
                defaultSelectedUsers={[this.state.sponser]}
                disabled={false}
                onChange={(items: any[]) => {
                  const _users = this._getPeoplePickerItems(items);
                  if (_users.length > 0)
                    this.setState({
                      SponsorId: _users[0],
                    });
                }}
                showHiddenInUI={false}
                ensureUser={true}
                principalTypes={[PrincipalType.User]}
                // resolveDelay={1000}
              />
              <Label className={styles.inputDesc}>
                The Sponsor generally is from the Business and must be Managing
                Director level (or a Desk Head in the case of Absa Capital).
              </Label>
              <PeoplePicker
                context={this.props.context}
                titleText="Product Owner"
                personSelectionLimit={3}
                showtooltip={true}
                defaultSelectedUsers={[this.state.workstreamCoordinator]}
                disabled={false}
                onChange={(items: any[]) => {
                  const _users = this._getPeoplePickerItems(items);
                  if (_users.length > 0)
                    this.setState({
                      WorkStreamCoordinatorId: _users[0],
                    });
                }}
                showHiddenInUI={false}
                ensureUser={true}
                principalTypes={[PrincipalType.User]}
                // resolveDelay={1000}
              />
            </Stack>
          </Stack>
          <HeaderInfo
            title="Product Description"
            description="This section captures the description of NAPA. It is to be concise in order for reviewers, regardless of their expertise, to understand the NAPA. More detail can be added when the full application is submitted. Please do not embed any documents in this section."
          />
          <Stack horizontal tokens={stackTokens} styles={stackStyles}>
            <Stack tokens={stackTokens} {...columnProps}>
              <Dropdown
                label="Region:"
                options={proposalRegions}
                selectedKey={this.state.Region}
                onChange={this._onChange}
                id="ddl_Region"
                required
              />
            </Stack>
            <Stack {...columnProps}>
              <Dropdown
                label="Country:"
                options={this.state.shortCountries}
                selectedKey={this.state.Country0}
                defaultSelectedKey={this.state.Country0}
                onChange={this._onChange}
                id="ddl_Country0"
                required
              />
            </Stack>
          </Stack>
          <Stack horizontal tokens={stackTokens} styles={stackStyles}>
            <Stack {...columnProps}>
              <Dropdown
                label="Entity:"
                // onChange={() => {
                //   this._loadFilteredDropdown(
                //     "ddlProductArea",
                //     "Product_x0020_Area"
                //   );
                // }}
                onChange={this._onFilteredDropdownChange}
                options={this.state.companies}
                id="ddl_Company"
                selectedKey={this.state.Company}
                required
                // selectedKey={this.state.proposalObject.Company}
              />
              <Dropdown
                label="Product Family:"
                options={this.state.businessAreas}
                // onChange={() => {
                //   this._loadFilteredDropdown("ddlSubProducts", "Product");
                // }}
                required
                onChange={this._onChange}
                id="ddl_BusinessArea"
                selectedKey={this.state.BusinessArea}
                // selectedKey={this.state.proposalObject.BusinessArea}
              />
              {/*  */}
            </Stack>
            <Stack {...columnProps}>
              <Dropdown
                label="Distribution Channel:"
                options={this.state.distributionChannels}
                // onChange={() => {
                //   this._loadFilteredDropdown("ddlBusinessArea", "Business");
                // }}
                title="Distribution Channels"
                onChange={this._onChange}
                id="ddl_ProductArea0"
                selectedKey={this.state.ProductArea0}
                required
                // selectedKey={this.state.proposalObject.ProductArea0}
              />
              {/* <FilteredDropdown
                label="Product Area:"
                context={this.props.context}
                listname="Products"
                field1={{
                  name: "Title",
                  value: this.state.proposalObject.ProductArea0,
                }}
              /> */}
              <Dropdown
                label="Product Family Risk Classification:"
                options={productFamRiskClass}
                id="ddl_SubProduct"
                // defaultSelectedKey="0"
                selectedKey={this.state.SubProduct}
                onChange={this._onChange}
              />
            </Stack>
          </Stack>
          <TextField
            label="Executive Summary:"
            multiline
            rows={5}
            value={this.state.ExecutiveSummary}
            id="txt_ExecutiveSummary"
            required
          />
          <Stack horizontal tokens={stackTokens} styles={stackStyles}>
            <Stack {...columnProps}>
              <TextField
                label="What is new for this Proposal?"
                multiline
                rows={3}
                value={this.state.NewForProposal}
                onChange={this._onChangeText}
                id="txt_NewForProposal"
              />
              <TextField
                label="Link to Existing Proposal:"
                multiline
                rows={3}
                value={this.state.LinkToExistingProposal}
                onChange={this._onChangeText}
                id="txt_LinkToExistingProposal"
              />
            </Stack>
            <Stack {...columnProps}>
              <TextField
                label="Is there a specific transaction in the pipeline?"
                multiline
                rows={3}
                value={this.state.TransactionInPipeline}
                onChange={this._onChangeText}
                id="txt_TransactionInPipeline"
                required
              />
            </Stack>
          </Stack>
          <Stack horizontal tokens={stackTokens} styles={stackStyles}>
            <Stack {...columnProps}>
              {/* <TextField
                label="Is the structure of the new product/transaction in any way influenced by
                          the anticipated tax treatment of any party to the transaction?"
                multiline
                rows={3}
                value={this.state.proposalObject.TaxTreatment}
              /> */}
              <TextField
                label="Are there any Reputational and/or Conduct Risk issues which arise from entering into
                          this new product or amended product/services? Please provide a rationale for your answer"
                multiline
                value={this.state.ConductRiskIssuesComments}
                rows={3}
                onChange={this._onChangeText}
                id="txt_ConductRiskIssuesComments"
              />
            </Stack>
            <Stack {...columnProps}>
              {/* <TextField
                label="Does this NAPA constitute issuing a line of credit/an extension of credit of any type to the client?"
                multiline
                rows={3}
                value={this.state.proposalObject.LineOfCredit}
              /> */}
              <TextField
                label="What do you consider to be the Principal Risks associated with this proposal?"
                multiline
                rows={3}
                value={this.state.proposalObject.PrincipalRisks}
                onChange={this._onChangeText}
                id="txt_PrincipalRisks"
              />
            </Stack>
          </Stack>
          <HeaderInfo
            title="Business Hierarchy"
            description="Provide the following country information"
          />
          <Stack horizontal styles={stackStyles} tokens={stackTokens}>
            <Stack {...columnProps}>
              <Dropdown
                placeholder="Select options"
                label="Infrastructure Support Country:"
                // selectedKeys={selectedKeys}
                // eslint-disable-next-line react/jsx-no-bind
                onChange={this._onChange}
                options={this.state.shortCountries}
                styles={dropdownStyles}
                selectedKey={this.state.IFCountry}
                id="ddl_IFCountry"
                required
              />
              <Dropdown
                placeholder="Select options"
                label="Target Client Location:"
                // selectedKeys={selectedKeys}
                // eslint-disable-next-line react/jsx-no-bind
                onChange={this._onChange}
                selectedKey={this.state.ClientLocation}
                options={this.state.shortCountries}
                styles={dropdownStyles}
                id="ddl_ClientLocation"
                required
              />
              <Dropdown
                placeholder="Select options"
                label="Country of Product Offering:"
                // selectedKeys={selectedKeys}
                // eslint-disable-next-line react/jsx-no-bind
                onChange={this._onChange}
                selectedKey={this.state.ProductOfferingCountry}
                options={this.state.shortCountries}
                styles={dropdownStyles}
                id="ddl_ProductOfferingCountry"
                required
              />
              <Dropdown
                placeholder="Select options"
                label="Booking Location:"
                // selectedKeys={selectedKeys}
                // eslint-disable-next-line react/jsx-no-bind
                onChange={this._onChange}
                selectedKey={this.state.BookingLocation}
                options={this.state.shortCountries}
                styles={dropdownStyles}
                id="ddl_BookingLocation"
                required
              />
              <Dropdown
                placeholder="Select options"
                label="Trader Location:"
                // selectedKeys={selectedKeys}
                // eslint-disable-next-line react/jsx-no-bind
                onChange={this._onChange}
                selectedKey={this.state.TraderLocation}
                options={this.state.shortCountries}
                styles={dropdownStyles}
                id="ddl_TraderLocation"
              />
            </Stack>
            <Stack {...columnProps}>
              <Dropdown
                placeholder="Select options"
                label="Sales/Coverage Team Location:"
                // selectedKeys={selectedKeys}
                // eslint-disable-next-line react/jsx-no-bind
                onChange={this._onChange}
                selectedKey={this.state.SalesTeamLocation}
                options={this.state.shortCountries}
                styles={dropdownStyles}
                id="ddl_SalesTeamLocation"
                required
              />
              <Dropdown
                placeholder="Select options"
                label="Target Client Sector:"
                // selectedKeys={selectedKeys}
                // eslint-disable-next-line react/jsx-no-bind
                onChange={this._onChange}
                selectedKey={this.state.ClientSector}
                options={this.state.clientSectors}
                styles={dropdownStyles}
                id="ddl_ClientSector"
                required
              />
              <Dropdown
                placeholder="Select options"
                label="Booking/Applicable Currencies:"
                // selectedKeys={selectedKeys}
                // eslint-disable-next-line react/jsx-no-bind
                onChange={this._onChange}
                selectedKeys={this.state.BookingCurrencies}
                id="ddl_BookingCurrencies"
                multiSelect
                options={this.state.bookingCurrencies}
                styles={dropdownStyles}
                required
              />
              <Dropdown
                placeholder="Select option"
                label="Nature of Trade Activity:"
                // selectedKeys={selectedKeys}
                selectedKey={this.state.NatureOfTrade}
                // eslint-disable-next-line react/jsx-no-bind
                id="ddl_NatureOfTrade"
                onChange={this._onChange}
                options={this.state.tradeActivities}
                styles={dropdownStyles}
              />
              <Dropdown
                placeholder="Select options"
                label="Booking Legal Entity:"
                selectedKey={this.state.BookingEntity}
                // eslint-disable-next-line react/jsx-no-bind
                onChange={this._onChange}
                id="ddl_BookingEntity"
                options={this.state.legalEntities}
                styles={dropdownStyles}
                required
              />
            </Stack>
          </Stack>
          <Separator />
          <Toggle
            label="Is this a joint venture divisions or business area?"
            // defaultChecked
            onText="Yes"
            offText="No"
            onChange={this._onChangeToggle}
            role="checkbox"
            checked={this.state.JointVenture}
            id="tgl_JointVenture"
          />
          <Separator />
          {this.state.submitionStatus === "Ok" && <SuccessExample />}
          {this.state.errorMessage.length > 0 &&
            this.state.errorMessage.map((msg) => (
              <MessageBar
                messageBarType={MessageBarType.error}
                isMultiline={false}
                dismissButtonAriaLabel="Close"
              >
                {msg}
              </MessageBar>
            ))}
          <Stack horizontal tokens={stackTokens}>
            <DefaultButton
              text="Cancel"
              onClick={this._cancelProposal}
              allowDisabledFocus
              className={styles.buttonsGroupInput}
            />
            <PrimaryButton
              text="Submit for NPS Determination"
              onClick={this._saveApplicationProposal}
              allowDisabledFocus
              className={styles.buttonsGroupInput}
            />
            <PrimaryButton
              text="Save as Draft"
              onClick={this._saveApplicationProposal}
              allowDisabledFocus
              className={styles.buttonsGroupInput}
            />
          </Stack>
        </Stack>
      </div>
    );
  }
}
