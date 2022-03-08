import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface IReviewQuestions {
  context?: WebPartContext;
  ClearErrors?:any;
  Title?: string;
  MemoireConsidered?: string;
  RiskAssessmentCompleted?: string;
  HeadcountConsidered?: string;
  WorkaroundsRequired?: string;
  ITDevRequired?: string;
  IsStageApproved?:boolean;
  OpRiskRequired?: string;
  ReviewComments?: string;
  Proposal_ID?: string;
  NAPA_Link?: string;
  RDARRImpact?: string; //Does this proposal have any impact on the existing RDARR , (Risk Data Aggregation and Risk Reporting) artifacts, processes or le
  RDARRRelevance?: string; //If yes, have the necessary changes / amendments been made to the relevant processes, documentation, reconciliations, controls or
  onChange: any;
  Status?: string;
  siteUrl?: string;
  ReviewInfrastructureColleaguesId?: number[];
  ReviewInfrastructureColleagues?: any;
  SetIsStageApproved?:any;
  submenu?: {};
  SubmitToSP?: any;
  onChangeText?: any;
  EditMode?:boolean;
  ValidateForm:any;
  ErrorMessages:string[];
  userRole?:string;
  userInfraAreas?:string[];
  currentInfraArea?:string;
  canReview?:boolean;
  SetReviewCompleted?:any;
  ReviewCompleted?:boolean;
}
