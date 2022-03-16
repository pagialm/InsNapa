import { IStackStyles, Stack } from "office-ui-fabric-react";
import * as React from "react";
import HeaderInfo from "../Common/HeaderInfo";
import Headers from "../Common/Headers";
import ShowConditions from "../Common/Conditions/ShowConditions";
import ApprovalSummaryList from "./ApprovalSummaryList";
import PostApproval from "./PostApproval";

const stackStyles: Partial<IStackStyles> = { root: { width: "100%" } };

const ApprovalSummary = (props) => {
  return (
    <Stack styles={stackStyles}>
      <Headers
        ApprovalDueDate={props.ApprovalDueDate}
        proposalStatus={props.Status}
        proposalId={props.ID}
        selectedSection="Approval Summary"
        title={props.Title}
      />
      <HeaderInfo
        title="Conditions Raised"
        description="Please see conditions raised per Infrastructure"
      />
      <ShowConditions
        id={props.ID}
        siteUrl={props.context.pageContext.site.absoluteUrl}
        SubmitToSP={props.SubmitToSP}
        context={props.context}
        showAddPanel={false}
      />
      <ApprovalSummaryList 
        Items={props.ApprovedItems} 
        ApprovedToTradeDate={props.ApprovedToTradeDate}
      />
      {(props.Status === "Approved to Trade" || props.Status === "Approval Expired" || props.Status === "Approved and Traded" || props.Status === "Amendment Approved") && (
        <PostApproval 
          buttonClickedDisabled={props.buttonClickedDisabled}
          cancelProposal={props.cancelProposal}
          onChangeText={props.onChangeText}
          onFormatDate={props.onFormatDate}
          onSelectDate={props.onSelectDate}
          postApprovalDate={props.postApprovalDate}
          postApprovalExtensionDate={props.postApprovalExtensionDate}
          postApprovalFirstTradeDate={props.postApprovalFirstTradeDate}
          PostApprovalNPSComments={props.PostApprovalNPSComments}
          savePostApprovalDetails={props.savePostApprovalDetails}
          Status={props.Status}        
          Year1ActualGross={props.Year1ActualGross}
          Year1EstimatedGross={props.Year1EstimatedGross}
          Year2ActualGross={props.Year2ActualGross}
          Year2EstimatedGross={props.Year2EstimatedGross}
        />
      )}
    </Stack>
  );
};

export default ApprovalSummary;
