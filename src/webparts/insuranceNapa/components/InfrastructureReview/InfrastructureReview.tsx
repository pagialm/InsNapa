import {
  SPHttpClient,
  SPHttpClientResponse,
  ISPHttpClientOptions,
} from "@microsoft/sp-http";
import {
  DefaultButton,
  DetailsList,
  DetailsListLayoutMode,
  Dropdown,
  IColumn,
  IDropdownOption,
  IStackProps,
  IStackStyles,
  PrimaryButton,
  SelectionMode,
  Separator,
  Stack,
  TextField,
  Text,
  Link,
} from "office-ui-fabric-react";
import * as React from "react";
import { render } from "react-dom";
import HeaderInfo from "../Common/HeaderInfo";
import HeadersDecor from "../Common/Headers";
import ShowConditions from "../Common/Conditions/ShowConditions";
import Approvals from "./Approvals";
import { IInfrastructureReviewProps } from "./IInfrastructureReviewProps";
import ReviewQuestions from "./ReviewQuestions";

const InfrastructureReview = (props: IInfrastructureReviewProps) => {
  const [conditionsItems, setConditionsItems] = React.useState([]);
  const [reviewItems, setReviewItems] = React.useState([]);
  const [IsStageApproved, SetIsStageApproved] = React.useState(false);
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
        // console.log("items:", reviewItems);
        setReviewItems([...reviewItems, item]);
      });
      console.log("items:", reviewItems);
    });
  }, []);

  return (
    <Stack>
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
        context={props.context}
        onChange={props.onChange}
        onChangeText={props.onChangeText}
        Proposal_ID={props.ID.toString()}
        siteUrl={props.context.pageContext.site.absoluteUrl}
        submenu={props.menuObject}
        Status={props.Status}
        SubmitToSP={props.saveOnSharePoint}
        IsStageApproved={IsStageApproved}
      />

      <HeaderInfo
        title={`${props.SelectedSection} Conditions`}
        description=""
      />
      <ShowConditions
        id={props.ID}
        siteUrl={props.context.pageContext.site.absoluteUrl}
        context={props.context}
        showAddPanel={true}
        SubmitToSP={props.saveOnSharePoint}
        submenu={props.menuObject}
      />

      <HeaderInfo
        title={`${props.SelectedSection} Infrastructure Approval`}
        description="Provide the following Compliance Infrastructure Approval information"
      />
      <Approvals
        context={props.context}
        DeleteFromSP={props.DeleteFromSP}
        IsStageApproved={IsStageApproved}
        NoOfApprovalsRequired={props.NoOfApprovalsRequired}
        Proposal_ID={props.ID.toString()}
        SetIsStageApproved={SetIsStageApproved}
        siteUrl={props.context.pageContext.site.absoluteUrl}
        submenu={props.menuObject}
        SubmitToSP={props.saveOnSharePoint}
        Status={props.Status}
      />
    </Stack>
  );
};
export default InfrastructureReview;
