import { IStackStyles, Stack } from "office-ui-fabric-react";
import * as React from "react";
import HeaderInfo from "../Common/HeaderInfo";
import Headers from "../Common/Headers";
import ShowConditions from "../Common/Conditions/ShowConditions";
import ApprovalSummaryList from "./ApprovalSummaryList";

const stackStyles: Partial<IStackStyles> = { root: { width: "100%" } };

const ApprovalSummary = (props) => {
  return (
    <Stack styles={stackStyles}>
      <Headers
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
        context={props.context}
        showAddPanel={false}
      />
      <ApprovalSummaryList Items={props.ApprovedItems} />
    </Stack>
  );
};

export default ApprovalSummary;
