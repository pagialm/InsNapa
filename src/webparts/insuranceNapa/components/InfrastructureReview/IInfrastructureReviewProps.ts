import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface IInfrastructureReviewProps {
  DeleteFromSP?:any
  NoOfApprovalsRequired?: number;
  title: string;
  subtitle: string;
  ID: number;
  IsStageApproved?: boolean;
  Status: string;
  Title: string;
  SelectedSection: string;
  RDARRImpact?: string;
  onChange?: any;
  mainItem?: any;
  context: WebPartContext;
  menuObject?: {};
  saveOnSharePoint?: any;
  onChangeText?: any;
}
