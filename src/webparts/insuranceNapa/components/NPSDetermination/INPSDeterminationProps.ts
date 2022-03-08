import { WebPartContext } from "@microsoft/sp-webpart-base";
import { IDropdownOption } from "office-ui-fabric-react";

export interface INPSDeterminationProps {
  teamAssessment?: string;
  teamAssesmentReason?: string;
  teamCoordinators?: number[];
  productFamily?: string;
  productFamilyRiskClassification?: string;
  resetEnquiryComment: string;
  title?: string;
  proposalStatus?: string;
  proposalId?: number;
  onDdlChange?: any;
  context?: WebPartContext;
  getPeoplePickerItems: any;
  nAPATeamCoordinators: any;
  onChangeText: any;
  teamAssesmentReasonOptions: IDropdownOption[];
  productFamilyOptions: IDropdownOption[];
  approvalCapacity: string;
  napaProposalsListname: string;
  validateForm: any;
}
