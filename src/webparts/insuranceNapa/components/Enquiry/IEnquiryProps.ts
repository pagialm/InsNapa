import { WebPartContext } from "@microsoft/sp-webpart-base";
import { IDropdownOption } from "office-ui-fabric-react";

export interface IEnquiryProps {
  allCountries: IDropdownOption[];
  shortCountries: IDropdownOption[];
  clientSectors: IDropdownOption[];
  bookingCurrencies: IDropdownOption[];
  tradeActivities: IDropdownOption[];
  legalEntities: IDropdownOption[];
  users: [];
  companies: IDropdownOption[];
  businessAreas: IDropdownOption[];
  productAreas: IDropdownOption[];
  subProducts: IDropdownOption[];

  applicationCompletedBy: string;
  sponser: string[];
  tradingBookOwner: string[];
  workstreamCoordinator: string[];
  targetCompletionDate: Date;
  ID?: number;
  Title?: string; // Proposal Name
  TargetCompletionDate?: Date | string; // Target Launch Date
  AppCreatedById?: number; //Application completed by
  SponsorId?: number[]; // Sponsor
  TradingBookOwnerId?: number[]; // Trading Book/P&L Owner
  WorkStreamCoordinatorId?: number[]; // Workstream Coordinator
  Region?: string; // Region
  Country0?: string; // Country
  Company?: string; // Company
  BusinessArea?: string; // Business Area
  ExecutiveSummary?: string; // Executive Summary
  ProductArea0?: string[]; // Product Area
  SubProduct?: string; // Sub Product
  NewForProposal?: string; // What is new for this Proposal?
  TransactionInPipeline?: string; // Is there a specific transaction in the pipeline?
  LinkToExistingProposal?: string; // Link to Existing Proposal:
  TaxTreatment?: string; // Is the structure of the new product/transaction in any way influenced by the anticipated tax treatment of any party to the transaction?
  LineOfCredit?: string; // Does this NAPA constitute issuing a line of credit/an extension of credit of any type to the client?
  ConductRiskIssuesComments?: string; // Are there any Reputational and/or Conduct Risk issues which arise from entering into this new product or amended product/services? Please provide a rationale for your answer
  PrincipalRisks?: string; // What do you consider to be the Principal Risks associated with this proposal?
  IFCountry?: string[]; // Infrastructure Support Country
  SalesTeamLocation?: string[]; // Sales/Coverage Team Location
  ClientLocation?: string[]; // Target Client Location
  ClientSector?: string; // Target Client Sector
  ProductOfferingCountry?: string[]; // Country of Product Offering
  BookingCurrencies?: string[]; // Booking/Applicable Currencies
  BookingLocation?: string[]; // Booking Location
  NatureOfTrade?: string; // Nature of Trade Activity
  TraderLocation?: string[]; // Trader Location
  // BookingEntity?: string[]; // Booking Legal Entity
  JointVenture?: boolean; //Is this a joint venture divisions or business area?
  Status: string;
  distributionChannels: IDropdownOption[];
  submitionStatus: string;
  errorMessage: string[];
  NAPATeamCoordinatorsId?: any; // NAPA Team Coordinators
  NapaTeamAssessment?: string; // NAPA Team Assessment
  NapaTeamAssReason?: string; // NAPA Team Assessment Reason
  ProductFamily?: string; // Product Family
  ProductFamilyRiskClassification?: string; // Product Family Risk Classification

  ResetToEnqComment?: string; // Reset Enquiry Comment
  TeamAssesmentReasonOptions: IDropdownOption[];
  ProductFamilyOptions: IDropdownOption[];
  selectedSection: string;
  BUPRCDate?: Date; // BU PRC Date

  buttonClickedDisabled: boolean;
  onChangeText: any;
  context: WebPartContext;
  getPeoplePickerItems: any;
  onSelectDate: any;
  onFormatDate: any;
  onChange: any;
  onFilteredDropdownChange: any;
  cancelProposal: any;
  saveApplicationEnquiry: any;
  productFamRiskClass: IDropdownOption[];
  setParentState: any;
  onChangeToggle: any;
  EditMode?:boolean;
  tansformNullArray?:any;
}
