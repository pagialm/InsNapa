import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface IApprovals {
  ApprovalComments?: string;
  context?: WebPartContext;
  ClearErrors?:any;
  DeleteFromSP?:any;
  Title?: string;
  Proposal_ID?: string;
  NAPA_Link?: string;
  NoOfApprovalsRequired?: number;
  Status?: string;
  siteUrl?: string;
  ApprovalInfrastructureColleaguesId?: number[];
  ApprovalInfrastructureColleagues?: any;
  ReviewCompleted?:boolean;
  SetIsStageApproved?: any;
  submenu?: {};
  SubmitToSP?: any;
  onChangeText?: any;
  userRole?:string;
  userInfraAreas?:string[];
  currentInfraArea?:string;
  canApprove?:boolean;
  CheckApprovals?:any;
  ValidateForm?:any;
}
