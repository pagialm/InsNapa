import { WebPartContext } from "@microsoft/sp-webpart-base";
import { IDropdownOption } from "office-ui-fabric-react";

export interface IProposalProps {
  teamAssessment?: string; // Insurance BU PRC Classification
  teamAssesmentReason?: string; // Insurance BU PRU Outcome
  buPrcDate?: Date; //BU PRC Date
  teamCoordinators?: any[]; //Product Governance Team Coordinator
  productFamily?: string;
  productFamilyRiskClassification?: string;
  existingFamilyOrNewFamily?: string; //New
  approvalCapacity?: string; //New
  actionsRaisedByBUPRC: string; //New
  infraApprovedByBuPrc: any; //New
  resetEnquiryComment: string;
  title?: string;
  proposalStatus?: string;
  proposalId?: number;
  onDdlChange?: any;
  context?: WebPartContext;
  getPeoplePickerItems: any;
  nAPATeamCoordinators: string[];
  onChangeText: any;
  teamAssesmentReasonOptions: IDropdownOption[];
  productFamilyOptions: IDropdownOption[];
  productRiskFamilyOptions: IDropdownOption[];
  napaProposalsListname: string;
  validateForm: any;
  getItemsFilter: any;
  saveProposal: any;
  cancelProposal: any;
  onSelectDate: any;
  onFormatDate: any;
  setParentState: any;
  infraAreaApprovedByBUPRC: string[];
  buttonDisabled: boolean;
  Status: string;
  EditMode?:boolean;
  errorMessage?:string[];
}
