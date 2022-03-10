import {
  SPHttpClient,
  SPHttpClientResponse,
  ISPHttpClientOptions,
} from "@microsoft/sp-http";
import {  
  IStackStyles,
  Stack,  
} from "office-ui-fabric-react";
import * as React from "react";
import { render } from "react-dom";
import HeaderInfo from "../Common/HeaderInfo";
import HeadersDecor from "../Common/Headers";
import ShowConditions from "../Common/Conditions/ShowConditions";
import Approvals from "./Approvals";
import { IInfrastructureReviewProps } from "./IInfrastructureReviewProps";
import ReviewQuestions from "./ReviewQuestions";

const stackStyles: Partial<IStackStyles> = { root: { width: "100%" } };

const InfrastructureReview = (props: IInfrastructureReviewProps) => {
  const [canEditStage, setCanEditStage] = React.useState(false);
  const [reviewItems, setReviewItems] = React.useState([]);
  const [IsStageApproved, SetIsStageApproved] = React.useState(false);
  const [reviewCompleted, SetReviewCompleted] = React.useState(false);
  const getListIems = (listName: string, filter?: string) => {
    const url: string =
      props.context.pageContext.site.absoluteUrl +
      "/_api/web/lists/getbytitle('" +
      listName +
      "')/items" +
      filter;
    return props.context.spHttpClient
      .get(url, SPHttpClient.configurations.v1)
      .then((response) => {
        return response.json();
      })
      .then((jsonResponse) => {
        return jsonResponse.value;
      }) as Promise<any[]>;
  };
  React.useEffect(() => {
    getListIems(
      "NAPA Infrastructure Questions",
      "?$filter=Proposal_ID eq '" + props.ID + "'"
    ).then((items) => {
      // setReviewItems(items);      
      items.forEach((item) => {        
        setReviewItems([...reviewItems, item]);
      });      
    });
  }, []);
  React.useEffect(() => {
    const currentStage = props.SelectedSection === "Financial Reporting/ Control Risk" ? "Financial Reporting or Control Risk" : props.SelectedSection;
    const canEdit = props.userInfraAreas.some((iArea) => iArea == currentStage);
    
    setCanEditStage(props.userRole === "Admin" ? true : canEdit);    

  }, [props.SelectedSection]);

  const completeReview = (r:boolean) => {
    console.log("....ran...")
    SetReviewCompleted(r);
  }

  return (
    <Stack styles={stackStyles}>
      {props.ID > 0 && (
        <HeadersDecor
          proposalStatus={props.Status}
          proposalId={props.ID}
          selectedSection={props.SelectedSection}
          title={props.Title}
        />
      )}

      <HeaderInfo
        title={`${props.SelectedSection} Infrastructure Review`}
        description="All key stakeholders should do regular reviews on their signoff criteria for New & Amended
         Business Approvals to ensure benchmarking against local regulatory changes and material changes to processes,
          systems and people. It is required that during the approval of a proposal that you and your team do retain
           documentation and/or have clear rationale supportive of your assessment."
      />

      <ReviewQuestions
        ClearErrors={props.ClearErrors}
        context={props.context}
        currentInfraArea={props.SelectedSection}
        onChange={props.onChange}
        onChangeText={props.onChangeText}
        Proposal_ID={props.ID.toString()}
        ReviewCompleted={reviewCompleted}
        siteUrl={props.context.pageContext.site.absoluteUrl}
        submenu={props.menuObject}
        SetReviewCompleted={completeReview}
        Status={props.Status}
        SubmitToSP={props.saveOnSharePoint}
        IsStageApproved={IsStageApproved}
        EditMode={props.EditMode}
        ValidateForm={props.ValidateForm}
        ErrorMessages={props.ErrorMessages}
        canReview={canEditStage ? true : false}
      />

      <HeaderInfo
        title={`${props.SelectedSection} Conditions`}
        description=""
      />
      <ShowConditions        
        id={props.ID}
        siteUrl={props.context.pageContext.site.absoluteUrl}
        context={props.context}
        showAddPanel={props.EditMode ? true : false}
        SubmitToSP={props.saveOnSharePoint}
        submenu={props.menuObject}
        onSelectDate={props.onSelectDate}
        internalMenuId={props.internalMenuId}
        getPeoplePickerItems={props.getPeoplePickerItems}        
      />

      <HeaderInfo
        title={`${props.SelectedSection} Infrastructure Approval`}
        description="Provide the following Compliance Infrastructure Approval information"
      />
      <Approvals
        ClearErrors={props.ClearErrors}
        CheckApprovals={props.checkApprovals}
        context={props.context}
        currentInfraArea={props.SelectedSection}
        DeleteFromSP={props.DeleteFromSP}
        EditMode={props.EditMode}
        IsStageApproved={IsStageApproved}
        NoOfApprovalsRequired={props.NoOfApprovalsRequired}
        Proposal_ID={props.ID.toString()}
        ReviewCompleted={reviewCompleted}
        SetIsStageApproved={SetIsStageApproved}
        siteUrl={props.context.pageContext.site.absoluteUrl}
        submenu={props.menuObject}
        SubmitToSP={props.saveOnSharePoint}
        Status={props.Status}
        canApprove={(props.userRole === "Approver" && canEditStage || props.userRole === "Admin") ? true : false}
        ValidateForm={props.ValidateForm}
      />
    </Stack>
  );
};
export default InfrastructureReview;
