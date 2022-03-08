import * as React from "react";
import styles from "./InsuranceNapa.module.scss";
import { IInsuranceNapaProps } from "./IInsuranceNapaProps";
import MenuIcon from "./Common/MenuIcon";
import { IDropdownOption, Stack, autobind } from "office-ui-fabric-react/lib";
import { InsuranceNapaState } from "./InsuranceNapaState";
import {
  SPHttpClient,
  SPHttpClientResponse,
  ISPHttpClientOptions,
} from "@microsoft/sp-http";
import { IProposal } from "./IProposal";
import Proposal from "./Proposal/Proposal";
import Enquiry from "./Enquiry/Enquiry";
import SupportingDocuments from "./Common/SupportingDocuments";
import Pipeline from "./Pipeline/Pipeline";
import NPSPipelineReview from "./NPSPipelineReview/NPSPipelineReview";
import DatesAndUserUtil from "./Common/DatesAndUserUtil";
import Utility from "./Common/Utility";
import InfrastructureReview from "./InfrastructureReview/InfrastructureReview";
import ApprovalSummary from "./ApprovalSummary/ApprovalSummary";
import OtherStatus from "./OtherStatus/OtherStatus";
import FinalNPSReview from "./FinalNPSReview/FinalNPSReview";
import ApprovalToTrade from "./ApprovalToTrade/ApprovalToTrade";

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
  productFamilyListname: "Product Family",
  teamAssessmentReasonListname: "NAPA Team Assessment Reason",
  userObjects: [],
  userObjectsCount: 0,
  supportingDocsListname: "NAPA Supporting Documentation",
  napaApprovalsListname: "NAPA Infrastructure Approvals",
};
const productFamRiskClass: IDropdownOption[] = [
  { key: "High", text: "High" },
  { key: "Medium", text: "Medium" },
  { key: "Low", text: "Low" },
];

let menuObj = {};

const mainStatuses = [
  "Enquiry",
  "NPS Determination",
  "Pipeline",
  "NPS Pipeline Review",
  "Infrastructure Review",
  "Final NPS Review",
  "Chair Approval",
  "Approval to Trade",
  "Approved to Trade",
  "Approval Expired",
];

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
      sponser: [],
      tradingBookOwner: [],
      workstreamCoordinator: [],
      targetCompletionDate: null,
      proposalObj: {},
      ID: 0,
      Title: "",
      TargetCompletionDate: null,
      AppCreatedById: 0,
      SponsorId: [],
      TradingBookOwnerId: [],
      WorkStreamCoordinatorId: [],
      Region: "",
      Country0: "",
      Company: "",
      BusinessArea: "",
      ExecutiveSummary: "",
      ProductArea0: [],
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
      ClientSector: "",
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
      ProductFamilyOptions: [],
      TeamAssesmentReasonOptions: [],
      selectedSection: "",
      BUPRCDate: null,
      ExistingFamily: "",
      ActionsRasedByBUPRC: "",
      InfraAreaApprovedByBUPRCId: [],
      nAPATeamCoordinators: [],
      infraAreaApprovedByBUPRC: [],
      buttonClickedDisabled: false,
      SupportingDocs: [],
      attachmentAdded: "",
      ResetToNPSDComment: "",
      LegalReviewer: [],
      ITReviewer: [],
      FinancialCrimeReviewer: [],
      TaxReviewer: [],
      FraudRiskReviewer: [],
      ComplianceReviwer: [],
      OperationsReviewer: [],
      CRMReviewer: [],
      CreditRiskReviwer: [],
      MarketRiskReviewer: [],
      ProductControlReviewer: [],
      RegulatoryReportingReviewer: [],
      TreasuryReviewer: [],
      TreasuryRiskReviewer: [],
      IRMReviewer: [],
      GroupResilienceReviewer: [],
      FinancialReportingReviewer: [],
      ConductRiskReviewer: [],
      BusinessCaseApprovalFrom: "",
      ReinsuranceReviewer: [],
      CustomerExperienceReviewer: [],
      DistributionReviewer: [],
      isAttachmentAdded: false,
      Approval_x0020_withdrawn_x0020_d: null,
      ProposalDateWithdrawal: null,
      PIRComments: "",
      ResetFinalNPSComment: "",
    };
  }
  @autobind
  private menuClicked(e, menu) {
    console.log(menu);
    // debugger;
    if (e.target.innerText === "Infrastructure Review")
      this.setState({ selectedSection: "Enquiry" });
    else this.setState({ selectedSection: e.target.innerText });
    console.log(this.state);
    menuObj = menu;
  }
  public async componentDidMount(): Promise<void> {
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

    //NAPA Team Assessment Reason
    const arr = new Array<IDropdownOption>();
    this._loadDropdownFromSP(proposalObj.teamAssessmentReasonListname, arr, {
      TeamAssesmentReasonOptions: arr,
    });
    //Product Family
    const arr2 = new Array<IDropdownOption>();

    this._loadDropdownFromSP(proposalObj.productFamilyListname, arr2, {
      ProductFamilyOptions: arr2,
    });

    //Get User Groups
    this.GetCurrentUserGroups();

    //update the display mode    
    this.setState({EditMode : this.props.editMode});

    // get Proposal
    if (this.props.itemId && this.props.itemId > 0) {
      //Get Approved List items
      
      const approvedItemsFilterStr: string = `Proposal_ID eq '${this.props.itemId}'&$select=Proposal_ID,NAPA_Infra,Author/Title,Created&$expand=Author/Title`;
      if (this.state.Status !== mainStatuses[0] &&
          this.state.Status !== mainStatuses[1] &&
          this.state.Status !== mainStatuses[2] &&
          this.state.Status !== mainStatuses[3]){
        this._getListitemsFilter(
          proposalObj.napaApprovalsListname,
          approvedItemsFilterStr
        ).then((items) => {
          this.setState({ ApprovedItems: items });
        });        
      }
        
      //Get current proposal
      this._getListitem(proposalObj.napaProposalsListname, this.props.itemId)
        .then((item) => {
          const _item: IProposal = item as IProposal;
          return _item as Promise<IProposal>;
        })
        .then((itemAsProposal) => {
          this._getListitemsFilter(
            proposalObj.productsListname,
            `Title eq '${itemAsProposal.Company}'`
          ).then(async (filteredItems: any[]) => {
            const productAreasOptions = this.filterArrayOfOptions(
              filteredItems,
              "ProductFamily",
              "Title",
              itemAsProposal.Company
            );
            debugger;
            const _item = itemAsProposal;            
            this.setState({ businessAreas: productAreasOptions });
            if (_item["Status"] === "Infrastructure Review") {
              const totalApprovals = this.state.ApprovedItems ? this.state.ApprovedItems.length : 0;
              const totalReviewers = this._countReviews(_item);
              if (totalReviewers > 0 && totalApprovals >= totalReviewers)
                _item["Status"] = "Final NPS Review";
            }
            if (_item["Status"] === "Infrastructure Review")
              this.setState({ selectedSection: "Enquiry" });
            else {
              const isInMainItems = mainStatuses.filter(
                (s) => s === _item["Status"]
              );
              if (isInMainItems.length > 0)
                this.setState({ selectedSection: _item["Status"] });
              else this.setState({ selectedSection: "Other Status" });
            }

            const dateObjects = DatesAndUserUtil.GetDates();
            const userObjects = DatesAndUserUtil.GetDisplayNames(
              _item,
              this._getUserById
            );
            (await userObjects).forEach((element) => {
              this.setState(element);
            });

            dateObjects.forEach((dateObject) => {
              const newstate = {};
              if (_item[dateObject.itemName]) {
                newstate[dateObject.stateName] = new Date(
                  _item[dateObject.itemName]
                );
                this.setState(newstate);
              }
            });            
            this.setState({ 
              ..._item,
              proposalObject: _item,
              proposalObj: _item,
            });            
            console.log(_item);
            debugger;
            const _stageIndex = mainStatuses.indexOf(_item.Status);
            if(_stageIndex > 3){
              const filteredMenu:any[] = this.GetFilteredMenu(_item);
              this.setState({ExcludeMenuItems: filteredMenu});
            }
          });
        });
    }
  }
  @autobind
  private CheckApprovals():void{    
    const noOfApprovals = this.state.ApprovedItems.length + 1;
    const infrastructureCount = this.state.InfrastructureCount;
    if(noOfApprovals === infrastructureCount){
      const proposal = {};
      proposal["Status"] = mainStatuses[5];
      this.submitToSP(proposal);
    }
    else{
      location.href = this.props.context.pageContext.site.absoluteUrl;
    }
  }
  @autobind
  private GetCurrentUserGroups(){
    let userRole = "";
    let userInfraArea = [];

    const apiUri =this.props.context.pageContext.web.absoluteUrl +
    "/_api/web/GetUserById('" +
    this.props.context.pageContext.legacyPageContext["userId"] +
    "')/Groups";
    this._getLitsItems(apiUri)
    .then((groups: any[]) => {      
      const CurrentUserGroups = groups.map((a) =>
        a.Title.replace("NPS Country ", "").trim()
      );
      if(CurrentUserGroups.some(
        (r) => r == "NPS Reviewers"
      ))
        userRole = "Reviewer";
        if(CurrentUserGroups.some(
          (r) => r == "NPS Approvers"
        ))
          userRole = "Approver";
          if(CurrentUserGroups.some(
            (r) => r == "NPS Admins"
          ))
            userRole = "Admin";
            // if(CurrentUserGroups.some(
            //   (r) => r == "NPS Chairs"
            // ))
            //   userRole = "Chair";
      
     CurrentUserGroups.forEach((r) => {
        if(r.indexOf("NPS Infrastructure ") !== -1)
          userInfraArea.push(r.replace("NPS Infrastructure ","").trim());        
      });

      //Update state
      this.setState({
        CurrentUserRole:userRole,
        CurrentUserInfrastructureAreas: userInfraArea
      });
    });
    
  }
  private _getLitsItems(url: string): Promise<any[]> {
    return this.props.context.spHttpClient
      .get(url, SPHttpClient.configurations.v1)
      .then((response) => {
        return response.json();
      })
      .then((responsJson) => {
        return responsJson.value;
      }) as Promise<any[]>;
  }
  @autobind
  private GetFilteredMenu(item:IProposal):any[]{
    let filteredStages:any[] = [];
    if (!this.state.LegalReviewerId) filteredStages.push("Legal Risk");
    if (!this.state.ComplianceReviwerId) filteredStages.push("Compliance");
    if (!this.state.ITReviewerId) filteredStages.push("CRO");
    if (!this.state.OperationsReviewerId) filteredStages.push("Operations");
    if (!this.state.CreditRiskReviwerId) filteredStages.push("Data Privacy");
    if (!this.state.MarketRiskReviewerId) filteredStages.push("Information Security Risk and Cyber Risk");
    if (!this.state.TaxReviewerId) filteredStages.push("Tax Risk");
    if (!this.state.ProductControlReviewerId) filteredStages.push("Head of Actuarial and Statutory Actuary");
    if (!this.state.RegulatoryReportingReviewerId) filteredStages.push("Marketing and Communications");
    if (!this.state.CRMReviewerId) filteredStages.push("RBB CVM");
    if (!this.state.TreasuryReviewerId) filteredStages.push("Financial & Insurance Risk");
    if (!this.state.TreasuryRiskReviewerId) filteredStages.push("Supplier Risk");
    if (!this.state.IRMReviewerId) filteredStages.push("Technology Risk");
    if (!this.state.GroupResilienceReviewerId) filteredStages.push("Business Continuity Risk");
    if (!this.state.FraudRiskReviewerId) filteredStages.push("Fraud Risk");
    if (!this.state.FinancialCrimeReviewerId) filteredStages.push("Financial Crime");
    if (!this.state.FinancialReportingReviewerId) filteredStages.push("Financial Reporting/ Control Risk");
    if (!this.state.ConductRiskReviewerId) filteredStages.push("Valuations");
    if (!this.state.FinanceReviewerId) filteredStages.push("Finance"); 
    if (!this.state.ReinsuranceReviewerId) filteredStages.push("Reinsurance");
    if (!this.state.CustomerExperienceReviewerId) filteredStages.push("Customer Experience");
    if (!this.state.DistributionReviewerId) filteredStages.push("Distribution");
    return filteredStages;
  }
  @autobind
  private updateState(stateObject) {
    this.setState(stateObject);
  }
  private filterArrayOfOptions(
    filtereItems,
    fieldName,
    parentField,
    parentFieldValue
  ): IDropdownOption[] {
    const arrayOptions = filtereItems
      .map((fi) => {
        if (fi[parentField] === parentFieldValue) return fi[fieldName];
      })
      .filter((value, index, self) => self.indexOf(value) === index); //{

    const arrayOptionsObjs = arrayOptions.map((str) => ({
      key: str,
      text: str,
    }));
    return arrayOptionsObjs;
  }
  private _getUserIdsFilter(userIds: Array<number>): string {
    let userIdsfilter = "";
    userIds.forEach((userId) => {
      userIdsfilter += "ID eq " + userId + ",";
    });
    const retString = userIdsfilter
      .substr(0, userIdsfilter.length - 1)
      .split(",")
      .join(" or ");
    return retString;
  }
  @autobind
  private _getUserById(userId: number | string | Array<number>): Promise<any> {
    const url: string = !Array.isArray(userId)
      ? this.props.context.pageContext.site.absoluteUrl +
        "/_api/web/siteusers?$filter=ID eq " +
        userId
      : this.props.context.pageContext.site.absoluteUrl +
        "/_api/web/siteusers?$filter=" +
        this._getUserIdsFilter(userId);
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
    const ctx = this.props ? this.props.context : this.context;
    const url: string =
      ctx.pageContext.site.absoluteUrl +
      "/_api/web/lists/getbytitle('" +
      listName +
      "')/items?$filter=" +
      filter;
    return ctx.spHttpClient
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
    const el = ev.currentTarget;
    const stateEl = {};
    stateEl[el.id.split("_")[1]] = checked;
    this.setState(stateEl);
  }
  @autobind
  private _loadFilteredDropdown(fieldName: string, columnName: string) {
    const companyName =
      document.getElementById("ddlCompany").firstChild.textContent!;
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
    // console.log(getSelectedUsers);
    // console.log(this);
    //this.setState({ users: getSelectedUsers });
    return getSelectedUsers;
  }
  @autobind
  private _onChange(
    ev: React.FormEvent<HTMLDivElement>,
    option?: IDropdownOption,
    index?: number
  ): void {
    const selectedElement = ev.target as HTMLDivElement;
    const elementId = selectedElement.id.split("_")[1].split("-")[0];
    const el = {};

    debugger;
    //Todo: Find old element, check its type... array or primitive and act accodingly
    const oldState = this.state[elementId]; //el[elementId];
    if (
      Array.isArray(oldState) ||
      (selectedElement["type"] && selectedElement["type"] === "checkbox")
    ) {
      debugger;
      const newArray = oldState ? [...oldState] : [];
      if (
        oldState &&
        oldState.some((opt) => {
          return opt === option.key;
        })
      ) {
        const finalArray = newArray.filter((itemToFilter) => {
          return itemToFilter !== option.key;
        });
        el[elementId] = finalArray;
      } else {
        newArray.push(option.key);
        el[elementId] = newArray;
      }
    } else {
      el[elementId] = option.key;
    }
    this.setState(el);
  }
  /**
   * this._onChnge Alternative. Use this method if using state object that can be null from SharePoint.
   * @param stateName static name
   * @param eventObj event onject
   * @param option option
   * @param index index
   */
  @autobind
  private tansformNullArray(stateName, eventObj, option, index){
    
    if(this.state[stateName] === null){
      const resetState = {};
      resetState[stateName] = [];
      // this.updateState(resetState);
      this.setState(resetState, ()=>{this._onChange(eventObj,option,index)});
    }
    else{
       this._onChange(eventObj,option,index);
    }    
   
  }
  @autobind
  private _onChangeText(
    ev: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>,
    newValue: string
  ): void {
    const element = ev.target as HTMLElement;
    debugger;
    const el = {};
    el[element.id.split("_")[1]] = newValue;
    this.setState(el);
  }
  @autobind
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
  private _saveApplicationProposal(e): void {
    debugger;
    this.setState({ buttonClickedDisabled: true });
    const buttonClicked: string = e.target.innerText;
    let statusText = "Pipeline";
    if (buttonClicked === "Save") statusText = "Proposal";
    const _isFormValid = buttonClicked === "Save" ? true : this._validateForm();
    const proposal = {};
    proposal["Status"] = statusText; // Reset to Enquiry
    if (buttonClicked === "Reset to Enquiry") {
      proposal["ResetToEnqComment"] = this.state.ResetToEnqComment; // Reset to Enquiry
      proposal["Status"] = "Enquiry"; // Reset to Enquiry
    } else {
      proposal["NapaTeamAssessment"] = this.state.NapaTeamAssessment; // Insurance BU PRC Classification
      proposal["NapaTeamAssReason"] = this.state.NapaTeamAssReason; // Insurance BU PRU Outcome
      proposal["BUPRCDate"] = this.state.BUPRCDate; // BU PRC Date
      proposal["NAPATeamCoordinatorsId"] = this.state.NAPATeamCoordinatorsId
        ? this.state.NAPATeamCoordinatorsId
        : []; // Product Governance Team Coordinator
      proposal["ProductFamily"] = this.state.ProductFamily; // Product Family
      proposal["ProductFamilyRiskClassification"] =
        this.state.ProductFamilyRiskClassification; // Product Family Risk Classification
      proposal["ExistingFamily"] = this.state.ExistingFamily; // Existing family or new family
      proposal["ApprovalCapacity"] = this.state.ApprovalCapacity; // Approval Capacity
      proposal["ActionsRasedByBUPRC"] = this.state.ActionsRasedByBUPRC; // Actions/ conditions/ commets raised by BU PRC
      proposal["InfraAreaApprovedByBUPRCId"] = this.state
        .InfraAreaApprovedByBUPRCId
        ? this.state.InfraAreaApprovedByBUPRCId
        : []; // Infrustructures area approved by BU
    }
    if (_isFormValid) this.submitToSP(proposal);
    else this.setState({ buttonClickedDisabled: false });
  }
  /**
   * Save Enquiry stage to Proposal if Submit proposal was clicked.
   * @param e Event trigger
   */
  @autobind
  private _saveApplicationEnquiry(e) {
    debugger;
    this.setState({ buttonClickedDisabled: true });
    const buttonClicked: string = e.target.innerText;
    let statusText = "Proposal";
    if (buttonClicked === "Save as Draft") statusText = "Enquiry";
    const _isFormValid =
      buttonClicked === "Save as Draft" ? true : this._validateForm();
    const proposal = {};

    proposal["Title"] = this.state.Title;
    proposal["TargetCompletionDate"] = this.state.targetCompletionDate;
    if(this.state.AppCreatedById > 0)
      proposal["AppCreatedById"] = this.state.AppCreatedById;
    proposal["SponsorId"] = this.state.SponsorId ? this.state.SponsorId : [];
    proposal["TradingBookOwnerId"] = this.state.TradingBookOwnerId
      ? this.state.TradingBookOwnerId
      : [];
    proposal["WorkStreamCoordinatorId"] = this.state.WorkStreamCoordinatorId
      ? this.state.WorkStreamCoordinatorId
      : [];
    proposal["Region"] = this.state.Region ? this.state.Region : "";
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
    proposal["ConductRiskIssuesComments"] =
      this.state.ConductRiskIssuesComments;
    proposal["PrincipalRisks"] = this.state.PrincipalRisks;
    proposal["IFCountry"] = this.state.IFCountry ? this.state.IFCountry : [];
    proposal["SalesTeamLocation"] = this.state.SalesTeamLocation
      ? this.state.SalesTeamLocation
      : [];
    proposal["ClientLocation"] = this.state.ClientLocation
      ? this.state.ClientLocation
      : [];
    proposal["ClientSector"] = this.state.ClientSector
      ? this.state.ClientSector
      : "";
    proposal["ProductOfferingCountry"] = this.state.ProductOfferingCountry
      ? this.state.ProductOfferingCountry
      : [];
    proposal["BookingCurrencies"] = this.state.BookingCurrencies
      ? this.state.BookingCurrencies
      : [];
    proposal["BookingLocation"] = this.state.BookingLocation
      ? this.state.BookingLocation
      : [];
    proposal["NatureOfTrade"] = this.state.NatureOfTrade;
    proposal["TraderLocation"] = this.state.TraderLocation
      ? this.state.TraderLocation
      : [];
    proposal["BookingEntity"] = this.state.BookingEntity
      ? this.state.BookingEntity
      : [];
    proposal["JointVenture"] = this.state.JointVenture;
    proposal["Status"] = statusText;
    if (_isFormValid) this.submitToSP(proposal);
    else this.setState({ buttonClickedDisabled: false });
  }
  /**
   * Submit a Proposal to SharePoint NAPA Proposal list.
   * @param proposal Proposal JSON Object
   */
  private submitToSP(proposal: any): void {
    const url: string =
      this.state.ID == 0
        ? this.props.context.pageContext.web.absoluteUrl +
          "/_api/web/lists/getbytitle('" +
          proposalObj.napaProposalsListname +
          "')/items"
        : this.props.context.pageContext.web.absoluteUrl +
          "/_api/web/lists/getbytitle('" +
          proposalObj.napaProposalsListname +
          "')/items(" +
          this.state.ID +
          ")";
    const spHttpClientOptions: ISPHttpClientOptions = {
      body: JSON.stringify(proposal),
    };

    if (this.state.ID > 0) {
      const headers: any = {
        "X-HTTP-Method": "MERGE",
        "IF-MATCH": "*",
      };
      spHttpClientOptions["headers"] = headers;
    }
    // {
    // if (_isFormValid) {
    this.props.context.spHttpClient
      .post(url, SPHttpClient.configurations.v1, spHttpClientOptions)
      .then((response: SPHttpClientResponse) => {
        if (response.status >= 201 && response.status < 300) {
          this.setState({ submitionStatus: "Ok" });
          setTimeout(() => {
            location.href = this.props.context.pageContext.web.absoluteUrl;
          }, 500);
        } else {
          this.setState({
            errorMessage: [
              `Error: [HTTP]:${response.status} [CorrelationId]:${response.statusText}`,
            ],
            buttonClickedDisabled: false,
          });
        }
      });
    //}
  }
  /**
   * Submit to a different SharePoint list other than NAPA Proposals.
   * @param listname SharePoint list name to Submit to.
   * @param isNewItem Boolean to check if it's a new item or existing item.
   * @param listItem SharePoint item to be submitted to a list.
   * @param responseFunc Function that will be executed when submit has completed.
   */
  @autobind
  private submitToOtherSPList(
    listname: string,
    isNewItem: boolean,
    listItem: any,
    responseFunc: any
  ): void {
    const url: string = isNewItem
      ? this.props.context.pageContext.web.absoluteUrl +
        "/_api/web/lists/getbytitle('" +
        listname +
        "')/items"
      : this.props.context.pageContext.web.absoluteUrl +
        "/_api/web/lists/getbytitle('" +
        listname +
        "')/items(" +
        listItem.ID +
        ")";
    const spHttpClientOptions: ISPHttpClientOptions = {
      body: JSON.stringify(listItem),
    };

    if (!isNewItem) {
      const headers: any = {
        "X-HTTP-Method": "MERGE",
        "IF-MATCH": "*",
      };
      spHttpClientOptions["headers"] = headers;
    }

    this.props.context.spHttpClient
      .post(url, SPHttpClient.configurations.v1, spHttpClientOptions)
      .then((response: SPHttpClientResponse) => {
        responseFunc(response);
      });
  }
  /**
   * Delete a SharePoint list item.
   * @param listname SharePoint list name to delete from.
   * @param itemToDelete SharePoint item to be deleted.
   * @param returnFunc Function that will be executed when deletion has completed.
   */
  @autobind
  private _DeleteFromSP(
    listname: string,
    itemToDelete: any,
    returnFunc: any
  ): void {
    const url: string =
      this.props.context.pageContext.web.absoluteUrl +
      "/_api/web/lists/getbytitle('" +
      listname +
      "')/items(" +
      itemToDelete.ID +
      ")";
    const spHttpClientOptions: ISPHttpClientOptions = {
      body: JSON.stringify(itemToDelete),
    };
    const headers: any = {
      "X-HTTP-Method": "DELETE",
      "IF-MATCH": "*",
    };
    spHttpClientOptions["headers"] = headers;

    this.props.context.spHttpClient
      .post(url, SPHttpClient.configurations.v1, spHttpClientOptions)
      .then((response: SPHttpClientResponse) => {
        returnFunc(response);
      });
  }
  /**
   * Go back to the site landing page.
   */
  @autobind
  private _cancelProposal() {
    location.href = this.props.context.pageContext.web.absoluteUrl;
  }
  @autobind
  private _onSelectDate(stateObjectString: string, date: Date): void {
    const objectState = {};
    objectState[stateObjectString] = date;

    const firstLetter = stateObjectString[0].toLocaleLowerCase();
    const secondaryStateObjStr = firstLetter + stateObjectString.substr(1);
    objectState[secondaryStateObjStr] = date;

    this.setState(objectState);
  }

  @autobind
  private _onFormatDate(date: Date): string {
    const _date: Date = typeof date === "string" ? new Date(date) : date;
    return (
      _date.getDate() + "/" + (_date.getMonth() + 1) + "/" + _date.getFullYear()
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
  @autobind
  private _addAttachments(): void {
    debugger;
    const fileToUpload = (
      document.querySelector(
        "#btnAddAttachments_NBS_System"
      ) as HTMLInputElement
    ).files[0];
    this.GetFolderServerRelativeURL(
      fileToUpload,
      `${this.props.itemId}_${fileToUpload.name}`,
      "Shared Documents"
      // proposalObj.supportingDocsListname
    );
  }
  public getFileBuffer(uploadedFiles): Promise<any> {
    let promised = new Promise((resolve, reject) => {
      // debugger;
      let reader = new FileReader();
      reader.onload = (e) => {
        resolve(e.target["result"]);
      };
      reader.onerror = (e) => {
        reject(e.target["error"]);
      };
      reader.readAsArrayBuffer(uploadedFiles);
    });

    return promised;
  }
  public GetFolderServerRelativeURL(file, FinalName, DocumentLibPath): void {
    try {
      const spOpts: ISPHttpClientOptions = {
        body: file,
      };
      if (this.state.isAttachmentAdded)
        this.setState({ isAttachmentAdded: false });
      var redirectionURL =
        this.props.context.pageContext.site.absoluteUrl +
        "/_api/Web/GetFolderByServerRelativeUrl('" +
        DocumentLibPath +
        "')/Files/Add(url='" +
        FinalName +
        "',overwrite=true)?$select=*&$expand=ListItemAllFields";
      this.getFileBuffer(file).then(() => {
        this.props.context.spHttpClient
          .post(redirectionURL, SPHttpClient.configurations.v1, spOpts)
          .then((response: SPHttpClientResponse) => {
            response.json().then((docItem: any) => {
              // debugger;
              const docItemProps = {
                ProposalId: this.props.itemId.toString(),
                DocumentLink: `${this.props.context.pageContext.site.absoluteUrl}/${DocumentLibPath}/${docItem.Name}`,
                Document_x0020_Type: this.state.selectedSection,
                NewDocName: docItem.Name,
                DocumentName: docItem.Name,
              };
              this.props.context.spHttpClient
                .post(
                  `${this.props.context.pageContext.site.absoluteUrl}/_api/lists/getbytitle('${proposalObj.supportingDocsListname}')/Items(${docItem.ListItemAllFields.Id})`,
                  SPHttpClient.configurations.v1,
                  {
                    headers: {
                      Accept: "application/json;odata=nometadata",
                      "Content-type": "application/json;odata=nometadata",
                      "odata-version": "",
                      "IF-MATCH": "*",
                      "X-HTTP-Method": "MERGE",
                    },
                    body: JSON.stringify(docItemProps),
                  }
                )
                .then(
                  (attachResult: SPHttpClientResponse): void => {
                    if (attachResult.ok) {
                      this.setState({
                        attachmentAdded:
                          "Attachment added successfully!,Success",
                        isAttachmentAdded: true,
                      });
                      // location.href = location.href;
                    } else {
                      this.setState({
                        attachmentAdded:
                          "Attachment added, however problem updating the attachment properties. Contact you administrator.,PartialSuccess",
                      });
                      // location.href = location.href;
                    }
                  },
                  (error: any) => {
                    console.log("Error!", error);
                    this.setState({
                      attachmentAdded:
                        "Problem occured while attempting to load attachment. Contact you administrator.,Fail",
                    });
                    // location.href = location.href;
                  }
                );
            });
          });
      });
    } catch (error) {
      console.log("Error in GetFolderServerRelativeURL " + error);
    }
  }
  @autobind
  private _savePipeline(e): void {
    debugger;
    this.setState({ buttonClickedDisabled: true });
    const buttonClicked: string = e.target.innerText;
    let statusText = "NPS Pipeline Review";
    if (buttonClicked === "Save") statusText = "Pipeline";
    const _isFormValid = buttonClicked === "Save" ? true : this._validateForm();
    const proposal = {};
    if (statusText === "NPS Pipeline Review")
      this.setState({ ResetToNPSDComment: "N/A" });
    proposal["ResetToNPSDComment"] = this.state.ResetToNPSDComment;
    proposal["Status"] = statusText;

    if (_isFormValid) this.submitToSP(proposal);
    else this.setState({ buttonClickedDisabled: false });
  }
  /**
   * Checks if Proposal contains at least one Reviewer from Pipeline review.
   * @returns True / False
   */
  @autobind
  private checkInfraReviwers():boolean{
    let isValid = false;
    let reviewersCount = 0;
    const errorMsg = "You have to capture at least one Reviewer to release to Infrastructure approval.";

    if(this.state.TaxReviewerId)
      reviewersCount++;
    if(this.state.LegalReviewerId)
    reviewersCount++;
    if(this.state.FinanceReviewerId)
    reviewersCount++;
    if(this.state.ReinsuranceReviewerId)
    reviewersCount++;
    if(this.state.ITReviewerId)
    reviewersCount++;
    if(this.state.CRMReviewerId)
    reviewersCount++;
    if(this.state.FraudRiskReviewerId)
    reviewersCount++;
    if(this.state.IRMReviewerId)
    reviewersCount++;
    if(this.state.FinanceReviewerId)
    reviewersCount++;
    if(this.state.TreasuryReviewerId)
    reviewersCount++;
    if(this.state.ConductRiskReviewerId)
    reviewersCount++;
    if(this.state.FraudRiskReviewerId)
    reviewersCount++;
    if(this.state.MarketRiskReviewerId)
    reviewersCount++;
    if(this.state.OperationsReviewerId)
    reviewersCount++;
    if(this.state.TreasuryRiskReviewerId)
      reviewersCount++;
    if(this.state.DistributionReviewerId)
      reviewersCount++;
    if(this.state.FinancialCrimeReviewerId)
      reviewersCount++;
    if(this.state.ProductControlReviewerId)
      reviewersCount++;
    if(this.state.GroupResilienceReviewerId)
      reviewersCount++;
    if(this.state.RegulatoryReportingReviewerId)
      reviewersCount++;
    if(this.state.FinancialReportingReviewerId)
      reviewersCount++;
    if(this.state.CustomerExperienceReviewerId)
      reviewersCount++;

    if(reviewersCount > 0)
      isValid = true;
    else{
      this.setState({errorMessage:[errorMsg]});
    }
    return isValid;
  }
  /**
   * Submit Proposal from Pipeline Review to Infrastructure Review.
   * @param e Button that was clicked.
   */
  @autobind
  private _savePipelineReview(e): void {
    debugger;
    this.setState({ buttonClickedDisabled: true });
    const buttonClicked: string = e.target.innerText;
    let statusText = buttonClicked === "Reset to Pipeline" ? "Pipeline" : "Infrastructure Review";
    if (buttonClicked === "Save") statusText = "NPS Pipeline Review";
    let _isFormValid = (buttonClicked === "Save" || buttonClicked === "Reset to Pipeline") ? true : this._validateForm();
    _isFormValid = this.checkInfraReviwers();
    const proposal = {};
    if (statusText === "Infrastructure Review")
      this.setState({ ResetPipelineComment: "N/A" });
    if(statusText !== "Pipeline"){
      proposal["ResetPipelineComment"] = this.state.ResetPipelineComment;
      proposal["Status"] = statusText;
      proposal["LegalReviewerId"] = this.state.LegalReviewerId || [];
      proposal["ComplianceReviwerId"] = this.state.ComplianceReviwerId || [];
      proposal["ITReviewerId"] = this.state.ITReviewerId || [];
      proposal["OperationsReviewerId"] = this.state.OperationsReviewerId || [];
      proposal["CreditRiskReviwerId"] = this.state.CreditRiskReviwerId || [];
      proposal["MarketRiskReviewerId"] = this.state.MarketRiskReviewerId || [];
      proposal["TaxReviewerId"] = this.state.TaxReviewerId || [];
      proposal["ProductControlReviewerId"] =
        this.state.ProductControlReviewerId || [];
      proposal["RegulatoryReportingReviewerId"] =
        this.state.RegulatoryReportingReviewerId || [];
      proposal["CRMReviewerId"] = this.state.CRMReviewerId || [];
      proposal["TreasuryReviewerId"] = this.state.TreasuryReviewerId || [];
      proposal["TreasuryRiskReviewerId"] =
        this.state.TreasuryRiskReviewerId || [];
      proposal["IRMReviewerId"] = this.state.IRMReviewerId || [];
      proposal["GroupResilienceReviewerId"] =
        this.state.GroupResilienceReviewerId || [];
      proposal["FraudRiskReviewerId"] = this.state.FraudRiskReviewerId || [];
      proposal["FinancialCrimeReviewerId"] =
        this.state.FinancialCrimeReviewerId || [];
      proposal["FinancialReportingReviewerId"] =
        this.state.FinancialReportingReviewerId || [];
      proposal["ConductRiskReviewerId"] = this.state.ConductRiskReviewerId || [];
      proposal["FinanceReviewerId"] = this.state.FinanceReviewerId || [];
      proposal["ReinsuranceReviewerId"] = this.state.ReinsuranceReviewerId || [];
      proposal["CustomerExperienceReviewerId"] =
        this.state.CustomerExperienceReviewerId || [];
      proposal["DistributionReviewerId"] =
        this.state.DistributionReviewerId || [];
      proposal["RiskRanking"] = this.state.RiskRanking;
      if(this.state.BusinessCaseApprovalFromId)
        proposal["BusinessCaseApprovalFromId"] =
          this.state.BusinessCaseApprovalFromId[0];
      proposal["BusinessCaseApprovalDate"] = this.state.BusinessCaseApprovalDate;
      proposal["BusinessCaseApprovalComment"] =
        this.state.BusinessCaseApprovalComment;
      proposal["ResetPipelineComment"] = this.state.ResetPipelineComment;
      proposal["TargetSubmissionByBusiness"] =
        this.state.TargetSubmissionByBusiness; //TargetBusinessGoLive
      proposal["TargetBusinessGoLive"] = this.state.TargetBusinessGoLive; //TargetBusinessGoLive
      proposal["NAPABriefingDate"] = this.state.NAPABriefingDate; //TargetBusinessGoLive
      // proposal["NAPABriefingComments"] = this.state.NAPABriefingComments;//TargetBusinessGoLive
      debugger;
      const numberOfReviewers = this._countReviews(proposal);
      proposal["InfrastructureCount"] = numberOfReviewers;
    }
    else{
      proposal["ResetPipelineComment"] = this.state.ResetPipelineComment;
      proposal["Status"] = statusText;
    }

    if (_isFormValid) this.submitToSP(proposal);
    else this.setState({ buttonClickedDisabled: false });
  }
  /**
   * Counts the number of reviewers assinged to the current Proposal from Pipeline Review stage.
   * @param proposal Proposal Object in JSON format from SharePoint Napa Proposal list.
   * @returns A number of Reviewers assinged to a proposal
   */
  @autobind
  private _countReviews(proposal): number {
    let _count = 0;
    if (proposal.LegalReviewerId && proposal.LegalReviewerId.length > 0) _count++;
    if (proposal.ComplianceReviwerId && proposal.ComplianceReviwerId.length > 0) _count++;
    if (proposal.ITReviewerId && proposal.ITReviewerId.length > 0) _count++;
    if (proposal.OperationsReviewerId && proposal.OperationsReviewerId.length > 0) _count++;
    if (proposal.CreditRiskReviwerId && proposal.CreditRiskReviwerId.length > 0) _count++;
    if (proposal.CreditRiskReviwerId && proposal.CreditRiskReviwerId.length > 0) _count++;
    if (proposal.TaxReviewerId && proposal.TaxReviewerId.length > 0) _count++;
    if (proposal.ProductControlReviewerId && proposal.ProductControlReviewerId.length > 0) _count++;
    if (proposal.RegulatoryReportingReviewerId && proposal.RegulatoryReportingReviewerId.length > 0) _count++;
    if (proposal.CRMReviewerId && proposal.CRMReviewerId.length > 0) _count++;
    if (proposal.TreasuryReviewerId && proposal.TreasuryReviewerId.length > 0) _count++;
    if (proposal.TreasuryRiskReviewerId && proposal.TreasuryRiskReviewerId.length > 0) _count++;
    if (proposal.IRMReviewerId && proposal.IRMReviewerId.length > 0) _count++;
    if (proposal.GroupResilienceReviewerId && proposal.GroupResilienceReviewerId.length > 0) _count++;
    if (proposal.FraudRiskReviewerId && proposal.FraudRiskReviewerId.length > 0) _count++;
    if (proposal.FinancialCrimeReviewerId && proposal.FinancialCrimeReviewerId.length > 0) _count++;
    if (proposal.FinancialReportingReviewerId && proposal.FinancialReportingReviewerId.length > 0) _count++;
    if (proposal.ConductRiskReviewerId && proposal.ConductRiskReviewerId.length > 0) _count++;
    if (proposal.FinancekReviewerId && proposal.FinanceReviewerId.length > 0) _count++;
    if (proposal.ReinsuranceReviewerId && proposal.ReinsuranceReviewerId.length > 0) _count++;
    if (proposal.CustomerExperienceReviewerId && proposal.CustomerExperienceReviewerId.length > 0) _count++;
    if (proposal.DistributionReviewerId && proposal.DistributionReviewerId.length > 0) _count++;

    return _count;
  }
  @autobind
  private ClearErrors():void{
    this.setState({errorMessage:[]});
  }
  @autobind
  private _saveFinalNPSReview(e): void {
    debugger;
    this.setState({ buttonClickedDisabled: true });
    const buttonClicked: string = e.target.innerText;
    let statusText = "Approval to Trade";
    if (buttonClicked === "Save") statusText = "Final NPS Review";
    const _isFormValid = buttonClicked === "Save" ? true : this._validateForm();
    const proposal = {};
    if (statusText === "Approval to Trade")
      this.setState({ ResetFinalNPSComment: "N/A" });
    proposal["ActionsRaisedByExco"] = this.state.ActionsRaisedByExco;
    proposal["BIRORegionalHeadId"] = this.state.BIRORegionalHeadId;
    proposal["BIRORegionalHeadReviewDate"] =
      this.state.BIRORegionalHeadReviewDate;
    proposal["CROComment"] = this.state.CROComment;
    proposal["FinalRiskClassification"] = this.state.FinalRiskClassification;
    proposal["CROStatusDate"] = this.state.CROStatusDate;
    proposal["CROStatus"] = this.state.CROStatus;
    proposal["IsPostImplementationRequired"] =
      this.state.IsPostImplementationRequired;
    proposal["OperationalChecklistRequirement"] =
      this.state.OperationalChecklistRequirement;
    proposal["PIRComments"] = this.state.PIRComments;
    proposal["PIRDateCompleted"] = this.state.PIRDateCompleted;
    proposal["TargetDueDate"] = this.state.TargetDueDate;
    proposal["ResetFinalNPSComment"] = this.state.ResetFinalNPSComment;

    proposal["Status"] = statusText;

    if (_isFormValid) this.submitToSP(proposal);
    else this.setState({ buttonClickedDisabled: false });
  }
  @autobind
  private _ResetToInfrastructureReview(infraAreas: []) {
    console.log(infraAreas);
    const proposal = {};
    let newCounter = 0;
    infraAreas.forEach((infraArea) => {
      const infraObj = Utility.GetInfraObject(infraArea);
      proposal[infraObj.approvalDate] = null;
      proposal[infraObj.comment] = this.state.ResetFinalNPSComment;
      proposal[infraObj.approvedBy] = null;
      newCounter++;
    });
    proposal["Status"] = "Infrastructure Review";
    proposal["InfrastructureApprovalCount"] =
      this.state.InfrastructureApprovalCount - newCounter;
    console.log(proposal);
    this.submitToSP(proposal);
  }
  @autobind
  private _saveApprovalToTrade(e) {
    debugger;
    this.setState({ buttonClickedDisabled: true });
    const buttonClicked: string = e.target.innerText;
    let statusText = "Approved to Trade";
    if (buttonClicked === "Save") statusText = "Approval to Trade";
    if (buttonClicked === "Reset to Final NPS Review")
      statusText = "Final NPS Review";
    const _isFormValid = buttonClicked === "Save" ? true : this._validateForm();
    const proposal = {};
    if (statusText === "Approved to Trade")
      this.setState({ ResetFinalNPSComment: "N/A" });
    proposal["ATTChairId"] = this.state.ATTChairId;
    proposal["ChairComments"] = this.state.ChairComments;
    proposal["ResetFinalNPSComment"] = this.state.ResetFinalNPSComment;
    proposal["Status"] = statusText;
    console.log(proposal);
    if (_isFormValid) this.submitToSP(proposal);
    else {
      this.setState({ buttonClickedDisabled: false });
    }
  }
  public render(): React.ReactElement<IInsuranceNapaProps> {
    return (
      <div className={styles.insuranceNapa}>
        <Stack horizontal tokens={{ childrenGap: 5 }}>
          {/* <div className={styles.menuItem}>Enquiery</div> */}
          <MenuIcon
            iconName="Headset"
            stageName={this.state.selectedSection}
            activated={false}
            menuClickHandler={this.menuClicked}
            proposalStatus={this.state.Status}
            ApprovedItems={this.state.ApprovedItems}
            ExcludedMenuItems={this.state.ExcludeMenuItems}
          />
          {(this.state.selectedSection === mainStatuses[0] ||
            this.state.Status === "") && (
            <Enquiry
              allCountries={this.state.allCountries}
              applicationCompletedBy={this.state.applicationCompletedBy}
              bookingCurrencies={this.state.bookingCurrencies}
              businessAreas={this.state.businessAreas}
              buttonClickedDisabled={this.state.buttonClickedDisabled}
              BookingCurrencies={this.state.BookingCurrencies}
              // BookingEntity={this.state.BookingEntity}
              BookingLocation={this.state.BookingLocation}
              BusinessArea={this.state.BusinessArea}
              cancelProposal={this._cancelProposal}
              clientSectors={this.state.clientSectors}
              ClientSector={this.state.ClientSector}
              ClientLocation={this.state.ClientLocation}
              Country0={this.state.Country0}
              companies={this.state.companies}
              Company={this.state.Company}
              ConductRiskIssuesComments={this.state.ConductRiskIssuesComments}
              context={this.props.context}
              distributionChannels={this.state.distributionChannels}
              EditMode={this.state.EditMode}
              ExecutiveSummary={this.state.ExecutiveSummary}
              errorMessage={this.state.errorMessage}
              getPeoplePickerItems={this._getPeoplePickerItems}
              IFCountry={this.state.IFCountry}
              ID={this.state.ID}
              JointVenture={this.state.JointVenture}
              legalEntities={this.state.legalEntities}
              LinkToExistingProposal={this.state.LinkToExistingProposal}
              LineOfCredit={this.state.LineOfCredit}
              NewForProposal={this.state.NewForProposal}
              onChangeText={this._onChangeText}
              onFormatDate={this._onFormatDate}
              onChange={this._onChange}
              onChangeToggle={this._onChangeToggle}
              onFilteredDropdownChange={this._onFilteredDropdownChange}
              onSelectDate={this._onSelectDate}
              PrincipalRisks={this.state.PrincipalRisks}
              productAreas={this.state.productAreas}
              ProductArea0={this.state.ProductArea0}
              ProductOfferingCountry={this.state.ProductOfferingCountry}
              ProductFamilyOptions={this.state.ProductFamilyOptions}
              productFamRiskClass={productFamRiskClass}
              Region={this.state.Region}
              SalesTeamLocation={this.state.SalesTeamLocation}
              saveApplicationEnquiry={this._saveApplicationEnquiry}
              selectedSection={this.state.selectedSection}
              setParentState={this.updateState}
              shortCountries={this.state.shortCountries}
              sponser={this.state.sponser}
              subProducts={this.state.subProducts}
              submitionStatus={this.state.submitionStatus}
              Status={this.state.Status}
              SubProduct={this.state.SubProduct}
              targetCompletionDate={this.state.targetCompletionDate}
              TeamAssesmentReasonOptions={this.state.TeamAssesmentReasonOptions}
              Title={this.state.Title}
              tansformNullArray={this.tansformNullArray}
              tradeActivities={this.state.tradeActivities}
              tradingBookOwner={this.state.tradingBookOwner}
              users={this.state.users}
              workstreamCoordinator={this.state.workstreamCoordinator}
            />
          )}
          {this.state.selectedSection === mainStatuses[1] && (
            <Proposal
              EditMode={this.state.EditMode}
              errorMessage={this.state.errorMessage}
              teamAssessment={this.state.NapaTeamAssessment} // Insurance BU PRC Classification (NAPA Team Assessment)
              productFamily={this.state.BusinessArea} // Product Family
              teamAssesmentReason={this.state.NapaTeamAssReason} // Insurance BU PRU Outcome  (NAPA Team Assessment Reason)
              teamCoordinators={this.state.NAPATeamCoordinatorsId} // Product Governance Team Coordinator
              nAPATeamCoordinators={this.state.nAPATeamCoordinators}
              existingFamilyOrNewFamily={this.state.ExistingFamily} // Existing family or new family
              buPrcDate={this.state.bUPRCDate} // BU PRC Date
              approvalCapacity={this.state.ApprovalCapacity} // Approval Capacity
              actionsRaisedByBUPRC={this.state.ActionsRasedByBUPRC} // Actions/ conditions/ commets raised by BU PRC
              infraApprovedByBuPrc={this.state.InfraAreaApprovedByBUPRCId} // Infrustructures area approved by BU PRC
              infraAreaApprovedByBUPRC={this.state.infraAreaApprovedByBUPRC}
              productFamilyRiskClassification={this.state.SubProduct}
              resetEnquiryComment={this.state.ResetToEnqComment}
              title={this.state.Title}
              proposalStatus={this.state.Status}
              proposalId={this.state.ID}
              onDdlChange={this._onChange}
              context={this.props.context}
              getPeoplePickerItems={this._getPeoplePickerItems}
              // nAPATeamCoordinators={this.state.NAPATeamCoordinatorsId}
              onChangeText={this._onChangeText}
              teamAssesmentReasonOptions={this.state.TeamAssesmentReasonOptions}
              productFamilyOptions={this.state.businessAreas}
              productRiskFamilyOptions={productFamRiskClass}
              // productFamilyOptions={this.state.ProductFamilyOptions} productFamRiskClass
              validateForm={this._validateForm}
              napaProposalsListname={proposalObj.napaProposalsListname}
              getItemsFilter={this._getListitemsFilter}
              saveProposal={this._saveApplicationProposal}
              cancelProposal={this._cancelProposal}
              onFormatDate={this._onFormatDate}
              onSelectDate={this._onSelectDate}
              setParentState={this.updateState}
              buttonDisabled={this.state.buttonClickedDisabled}
              Status={this.state.Status}
            />
          )}
          {this.state.selectedSection === mainStatuses[2] && (
            <Pipeline
              EditMode={this.state.EditMode}
              title={this.state.Title}
              buttonDisabled={this.state.buttonClickedDisabled}
              cancelProposal={this._cancelProposal}
              proposalId={this.state.ID}
              proposalStatus={this.state.Status}
              resetToProposal={this.state.ResetToNPSDComment}
              savePipeline={this._savePipeline}
              siteUrl={this.props.context.pageContext.site.absoluteUrl}
              onChangeText={this._onChangeText}
              Status={this.state.Status}
            />
          )}
          {this.state.selectedSection === mainStatuses[3] && (
            <NPSPipelineReview
              EditMode={this.state.EditMode}
              errorMessage={this.state.errorMessage}
              title={this.state.Title}
              buttonDisabled={this.state.buttonClickedDisabled}
              cancelProposal={this._cancelProposal}
              proposalId={this.state.ID}
              proposalStatus={this.state.Status}
              // resetToProposal={this.state.ResetToNPSDComment}
              resetToPipeline={this.state.ResetPipelineComment}
              savePipelineReview={this._savePipelineReview}
              siteUrl={this.props.context.pageContext.site.absoluteUrl}
              onChangeText={this._onChangeText}
              RiskRanking={this.state.RiskRanking}
              briefingDate={null}
              BusinessCaseApprovalFrom={this.state.BusinessCaseApprovalFrom}
              context={this.props.context}
              getPeoplePickerItems={this._getPeoplePickerItems}
              onChangeDropdown={this._onChange}
              onFormatDate={this._onFormatDate}
              onSelectDate={this._onSelectDate}
              riskRankingOptions={productFamRiskClass}
              setParentState={this.updateState}
              // targetSubmissionByBusiness={null}
              BusinessCaseApprovalComment={
                this.state.BusinessCaseApprovalComment
              }
              addAttachments={this._addAttachments}
              LegalReviewer={this.state.LegalReviewer}
              ITReviewer={this.state.ITReviewer}
              FinancialCrimeReviewer={this.state.FinancialCrimeReviewer}
              FinanceReviewer={this.state.FinanceReviewer}
              TaxReviewer={this.state.TaxReviewer}
              FraudRiskReviewer={this.state.FraudRiskReviewer}
              ComplianceReviwer={this.state.ComplianceReviwer}
              OperationsReviewer={this.state.OperationsReviewer}
              CRMReviewer={this.state.CRMReviewer}
              CreditRiskReviwer={this.state.CreditRiskReviwer}
              MarketRiskReviewer={this.state.MarketRiskReviewer}
              ProductControlReviewer={this.state.ProductControlReviewer}
              RegulatoryReportingReviewer={
                this.state.RegulatoryReportingReviewer
              }
              TreasuryReviewer={this.state.TreasuryReviewer}
              TreasuryRiskReviewer={this.state.TreasuryRiskReviewer}
              IRMReviewer={this.state.IRMReviewer}
              GroupResilienceReviewer={this.state.GroupResilienceReviewer}
              FinancialReportingReviewer={this.state.FinancialReportingReviewer}
              ConductRiskReviewer={this.state.ConductRiskReviewer}
              ReinsuranceReviewer={this.state.ReinsuranceReviewer}
              CustomerExperienceReviewer={this.state.CustomerExperienceReviewer}
              DistributionReviewer={this.state.DistributionReviewer}
              businessCaseApprovalDate={this.state.businessCaseApprovalDate}
              targetBusinessGoLive={this.state.targetBusinessGoLive}
              nAPABriefingDate={this.state.nAPABriefingDate}
              targetSubmissionByBusiness={this.state.targetSubmissionByBusiness}
              Status={this.state.Status}
            />
          )}
          {(this.state.selectedSection === "CRO" ||
            this.state.selectedSection === "Legal Risk" ||
            this.state.selectedSection === "Financial Crime" ||
            this.state.selectedSection === "Data Privacy" ||
            this.state.selectedSection === "Fraud Risk" ||
            this.state.selectedSection === "Tax Risk" ||
            this.state.selectedSection ===
              "Information Security Risk and Cyber Risk" ||
            this.state.selectedSection === "Finance" ||
            this.state.selectedSection ===
              "Head of Actuarial and Statutory Actuary" ||
            this.state.selectedSection === "Marketing and Communications" ||
            this.state.selectedSection === "Financial & Insurance Risk" ||
            this.state.selectedSection === "Compliance" ||
            this.state.selectedSection === "Operations" ||
            this.state.selectedSection === "Supplier Risk" ||
            this.state.selectedSection ===
              "Financial Reporting/ Control Risk" ||
            this.state.selectedSection === "Technology Risk" ||
            this.state.selectedSection === "Business Continuity Risk" ||
            this.state.selectedSection === "RBB CVM" ||
            this.state.selectedSection === "Valuations" ||
            this.state.selectedSection === "Reinsurance" ||
            this.state.selectedSection === "Customer Experience" ||
            this.state.selectedSection === "Distribution" ||
            this.state.selectedSection === "CRO" ||
            this.state.selectedSection === "CRO" ||
            this.state.selectedSection === "CRO" ||
            this.state.selectedSection === "CRO" ||
            this.state.selectedSection === mainStatuses[4]) && (
            // console.log(this.state.selectedSection) &&
            <InfrastructureReview
              context={this.props.context}
              checkApprovals={this.CheckApprovals}
              DeleteFromSP={this._DeleteFromSP}
              EditMode={this.state.EditMode}
              title="title"
              Title={this.state.Title}
              ID={this.state.ID}
              SelectedSection={this.state.selectedSection}
              Status={this.state.Status}
              saveOnSharePoint={this.submitToOtherSPList}
              subtitle={this.state.selectedSection}
              mainItem={this.state.proposalObject}
              menuObject={menuObj}
              NoOfApprovalsRequired={this.state.InfrastructureCount}
              onChange={this._onChange}
              onChangeText={this._onChangeText}
              onSelectDate={this._onSelectDate}
              userRole={this.state.CurrentUserRole}
              userInfraAreas={this.state.CurrentUserInfrastructureAreas}
              internalMenuId={this.state.selectedSection}
              getPeoplePickerItems={this._getPeoplePickerItems}
              ValidateForm={this._validateForm}
              ErrorMessages={this.state.errorMessage}
              ClearErrors={this.ClearErrors}
            />
          )}
          {this.state.selectedSection === mainStatuses[5] && (
            <FinalNPSReview
              ActionsRaisedByExco={this.state.ActionsRaisedByExco}
              BusinessExecutiveId={this.state.BIRORegionalHeadId}
              BusinesExecutive={this.state.BIRORegionalHead}
              BusinesExecutiveApprovalDate={
                this.state.bIRORegionalHeadReviewDate
              }
              context={this.props.context}
              EditMode={this.state.EditMode}
              ExcoPrcCommitteeComment={this.state.CROComment}
              errorMessage={this.state.errorMessage}
              FinalRiskClassification={this.state.FinalRiskClassification}
              getPeoplePickerItems={this._getPeoplePickerItems}
              ID={this.state.ID}
              InduranceExcoPrcDate={this.state.cROStatusDate}
              InsuranceExcoPruOutcome={this.state.CROStatus}
              IsPostImplementationRequired={
                this.state.IsPostImplementationRequired
              }
              OperationalChecklistRequirement={
                this.state.OperationalChecklistRequirement
              }
              onChange={this._onChange}
              onChangeText={this._onChangeText}
              onSelectDate={this._onSelectDate}
              onFormatDate={this._onFormatDate}
              PirComments={this.state.PIRComments}
              PirDateCompleted={this.state.pIRDateCompleted}
              PirLaunchDate={this.state.targetDueDate}
              ResetFinalNPSComment={this.state.ResetFinalNPSComment}
              ResetToInfrastructureReview={this._ResetToInfrastructureReview}
              saveFinalNPSReview={this._saveFinalNPSReview}
              SelectedSection={this.state.selectedSection}
              setParentState={this.updateState}
              Status={this.state.Status}
              Title={this.state.Title}
            />
          )}
          {(this.state.selectedSection === "Approval to Trade" ||
            this.state.selectedSection === "Chair Approval") && (
            <ApprovalToTrade
              context={this.props.context}
              EditMode={this.state.EditMode}
              errorMessage={this.state.errorMessage}
              getPeoplePickerItems={this._getPeoplePickerItems}
              ID={this.state.ID}
              onChange={this._onChange}
              onChangeText={this._onChangeText}
              ProductGovernanceCustodians={
                this.state.ProductGovernanceCustodians
              }
              ResetFinalNPSComment={this.state.ResetFinalNPSComment}
              saveApprovalToTrade={this._saveApprovalToTrade}
              SelectedSection={this.state.selectedSection}
              setParentState={this.updateState}
              Status={this.state.Status}
              Title={this.state.Title}
            />
          )}
          {this.state.selectedSection === "Approval Summary" && (
            <ApprovalSummary
              context={this.props.context}
              ID={this.state.ID}
              Status={this.state.Status}
              Title={this.state.Title}
              ApprovedItems={this.state.ApprovedItems}
            />
          )}
          {this.state.selectedSection === "Other Status" && (
            <OtherStatus
              Approval_x0020_withdrawn_x0020_d={
                this.state.Approval_x0020_withdrawn_x0020_d
              }
              Status="Other Status"
              ID={this.state.ID}
              Title={this.state.Title}
              context={this.props.context}
              siteUrl={`${this.props.context.pageContext.site.absoluteUrl}`}
              OtherStatuses={this.state.OtherStatuses}
              OtherStatusComments={this.state.OtherStatusComments}
              OtherStatusDate={this.state.Approval_x0020_withdrawn_x0020_d}
              onFormatDate={this._onFormatDate}
              onChangeText={this._onChangeText}
              onChange={this._onChange}
              ProposalDateWithdrawal={this.state.ProposalDateWithdrawal}
            />
          )}
        </Stack>
        {this.state.ID > 0 && (
          <SupportingDocuments
            addAttachments={this._addAttachments}
            attachmentStatus={this.state.attachmentAdded}
            supportingDocs={this.state.SupportingDocs}
            siteUrl={`${this.props.context.pageContext.site.absoluteUrl}`}
            id={this.state.ID}
            isAttachmentAdded={this.state.isAttachmentAdded}
          />
        )}
      </div>
    );
  }
}
