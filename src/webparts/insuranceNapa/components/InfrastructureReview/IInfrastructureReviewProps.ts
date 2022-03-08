import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface IInfrastructureReviewProps {
  ClearErrors?:any;
  DeleteFromSP?:any;
  NoOfApprovalsRequired?: number;
  title: string;
  subtitle: string;
  ID: number;
  IsStageApproved?: boolean;
  Status: string;
  Title: string;
  SelectedSection: string;
  RDARRImpact?: string;
  ReviewCompleted?:boolean;
  onChange?: any;
  mainItem?: any;
  context: WebPartContext;
  menuObject?: {};
  saveOnSharePoint?: any;
  onChangeText?: any;
  EditMode?:boolean;
  onSelectDate?:any;
  internalMenuId?:string;
  getPeoplePickerItems?:any;
  ValidateForm:any;
  ErrorMessages:string[];
  userRole?:string;
  userInfraAreas?:string[];
  checkApprovals?:any;
}
