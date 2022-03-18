import { IDropdownOption } from "office-ui-fabric-react";
import { ISupportingDocItem } from "./Common/ISupportingDocItem";
import { IProposal } from "./IProposal";

export interface InsuranceNapaState {
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
  proposalObject: IProposal;
  proposalObj: IProposal;
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
  BookingEntity?: string[]; // Booking Legal Entity
  JointVenture?: boolean; //Is this a joint venture divisions or business area?
  Status?: string;
  distributionChannels: IDropdownOption[];
  submitionStatus: string;
  errorMessage: string[];
  NAPATeamCoordinatorsId?: any; // NAPA Team Coordinators
  NapaTeamAssessment?: string; // NAPA Team Assessment
  NapaTeamAssReason?: string; // NAPA Team Assessment Reason
  ProductFamily?: string; // Product Family
  ProductFamilyRiskClassification?: string; // Product Family Risk Classification
  ApprovalCapacity?: string; // Approval Capacity
  ResetToEnqComment?: string; // Reset Enquiry Comment
  TeamAssesmentReasonOptions: IDropdownOption[];
  ProductFamilyOptions: IDropdownOption[];
  selectedSection: string;
  BUPRCDate?: Date; // BU PRC Date
  bUPRCDate?: Date; // BU PRC Date
  ExistingFamily?: string; // Existing family or new family
  ActionsRasedByBUPRC?: string; // Actions/ conditions/ commets raised by BU PRC
  InfraAreaApprovedByBUPRCId?: number[]; // Infrustructures area approved by BU PRC
  nAPATeamCoordinators: string[];
  infraAreaApprovedByBUPRC: string[];
  buttonClickedDisabled: boolean;
  SupportingDocs: ISupportingDocItem[];
  attachmentAdded?: string;
  isAttachmentAdded?: boolean;
  ResetToNPSDComment?: string;
  RiskRanking?: string;
  ResetPipelineComment?: string;
  BusinessCaseApprovalComment?: string; //NPS Pipeline Review Comments
  LegalReviewer: string[];
  LegalReviewerId?: number[];
  ITReviewer: string[];
  ITReviewerId?: number[];
  FinancialCrimeReviewer: string[];
  FinancialCrimeReviewerId?: number[];
  TaxReviewer: string[];
  TaxReviewerId?: number[];
  FraudRiskReviewer: string[];
  FraudRiskReviewerId?: number[];
  ComplianceReviwer: string[];
  ComplianceReviwerId?: number[];
  OperationsReviewer: string[];
  OperationsReviewerId?: number[];
  CRMReviewer: string[];
  CRMReviewerId?: number[];
  CreditRiskReviwer: string[];
  CreditRiskReviwerId?: number[];
  MarketRiskReviewer: string[];
  MarketRiskReviewerId?: number[];
  ProductControlReviewer: string[];
  ProductControlReviewerId?: number[];
  RegulatoryReportingReviewer: string[];
  RegulatoryReportingReviewerId?: number[];
  TreasuryReviewer: string[];
  TreasuryReviewerId?: number[];
  TreasuryRiskReviewer: string[];
  TreasuryRiskReviewerId?: number[];
  IRMReviewer: string[];
  IRMReviewerId?: number[];
  GroupResilienceReviewer: string[];
  GroupResilienceReviewerId?: number[];
  FinancialReportingReviewer: string[];
  FinancialReportingReviewerId?: number[];
  FinanceReviewerId?:number[];
  FinanceReviewer?:string[];
  ConductRiskReviewer: string[];
  ConductRiskReviewerId?: number[];
  BusinessCaseApprovalFrom: string;
  BusinessCaseApprovalFromId?: number;
  ReinsuranceReviewer: string[];
  ReinsuranceReviewerId?: number[];
  CustomerExperienceReviewer: string[];
  CustomerExperienceReviewerId?: number[];
  DistributionReviewer: string[];
  DistributionReviewerId?: number[];
  businessCaseApprovalDate?: Date;
  BusinessCaseApprovalDate?: Date;
  targetBusinessGoLive?: Date;
  TargetBusinessGoLive?: Date;
  nAPABriefingDate?: Date;
  NAPABriefingDate?: Date;
  targetSubmissionByBusiness?: Date;
  TargetSubmissionByBusiness?: Date;
  Approval_x0020_withdrawn_x0020_d?: string | Date;
  OtherStatuses?: string;
  OtherStatusComments?: string;
  ProposalDateWithdrawal?: string | Date;
  ActionsRaisedByExco?: string;
  BIRORegionalHeadId?: number;
  BIRORegionalHead?: string;
  BIRORegionalHeadReviewDate?: string | Date;
  bIRORegionalHeadReviewDate?: string | Date;
  CROStatus?: string;
  CROStatusDate?: string | Date;
  cROStatusDate?: string | Date;
  CROComment?: string;
  FinalRiskClassification?: string;
  IsPostImplementationRequired?: string;
  TargetDueDate?: string | Date;
  targetDueDate?: string | Date;
  OperationalChecklistRequirement?: string;
  PIRDateCompleted?: string | Date;
  pIRDateCompleted?: string | Date;
  PIRComments: string;
  ResetFinalNPSComment?: string;
  InfrastructureApprovalCount?: number;
  ProductGovernanceCustodians?: string;
  ATTChairId?: number;
  ChairComments?: string;
  InfrastructureCount?: number;
  ApprovedItems?: any[];
  EditMode?:boolean;
  ExcludeMenuItems?:any[];
  CurrentUserRole?:string;
  CurrentUserInfrastructureAreas?:string[];
  ProposalScopeRestriction?:string;
  ProposalScopeClarification?:string;
  postApprovalDate?:string | Date;
  postApprovalExtensionDate?:string | Date;
  postApprovalFirstTradeDate?:string | Date;
  PostApprovalNPSComments?:string;
  Year1ActualGross?:string;
  Year1EstimatedGross?:string;
  Year2ActualGross?:string;
  Year2EstimatedGross?:string;
  // TargetSubmissionByBusiness?:string | Date;
  ApprovedToTradeDate?:Date;
  PreLaunchOpenConditions?:any[];
  isChairApprover?:boolean;
  ActionOwningAreas?:IDropdownOption[];
  InfrastructureAreasApprovedBPRC?:string[];
}
