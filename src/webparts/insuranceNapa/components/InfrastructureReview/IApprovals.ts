import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface IApprovals {
  ApprovalComments?: string;
  context?: WebPartContext;
  DeleteFromSP?:any;
  Title?: string;
  Proposal_ID?: string;
  NAPA_Link?: string;
  NoOfApprovalsRequired?: number;
  Status?: string;
  siteUrl?: string;
  ApprovalInfrastructureColleaguesId?: number[];
  ApprovalInfrastructureColleagues?: any;
  SetIsStageApproved?: any;
  submenu?: {};
  SubmitToSP?: any;
  onChangeText?: any;
}
