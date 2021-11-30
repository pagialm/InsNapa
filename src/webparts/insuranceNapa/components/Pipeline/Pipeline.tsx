import {
  DefaultButton,
  IStackStyles,
  Link,
  PrimaryButton,
  Separator,
  Stack,
  TextField,
} from "office-ui-fabric-react";
import HeaderInfo from "../Common/HeaderInfo";
import Headers from "../Common/Headers";
import * as React from "react";

const stackTokens = { childrenGap: 50 };
const stackStyles: Partial<IStackStyles> = { root: { width: "100%" } };

const Pipeline = (props: IPipelineProps) => {
  return (
    <Stack>
      <Headers
        proposalId={props.proposalId}
        selectedSection="Pipeline"
        title={props.title}
        proposalStatus={props.proposalStatus}
      />
      <HeaderInfo
        title="RBB Insurance Product Approval Template"
        description="Download the below RBB Insurance Product Approval template and upload it once completed"
      />
      <p>
        <Link
          href={`${props.siteUrl}/_layouts/download.aspx?SourceUrl=${props.siteUrl}/PipelineTemplates/1. RBB Insurance Product Approval Template V2.8.docx`}
          title="Click here to download template"
        >
          Click here to download template
        </Link>
      </p>
      {/* <HeaderInfo
        title="RBB Insurance Product Specification Template"
        description="Download the below RBB Insurance Product Specification template and upload it once completed"
      />
      <p>
        <Link
          href={`${props.siteUrl}/_layouts/download.aspx?SourceUrl=${props.siteUrl}/PipelineTemplates/2. RBB Insurance Product Specification_v1.2.doc`}
          title="Click here to download template"
        >
          Click here to download template
        </Link>
      </p> */}
      <HeaderInfo
        title="Conduct Risk Assessment Template"
        description="Download the below conduct risk assessment template and upload it once completed"
      />
      <p>
        <Link
          href={`${props.siteUrl}/_layouts/download.aspx?SourceUrl=${props.siteUrl}/PipelineTemplates/3. AGL Conduct Risk Assessment.xls`}
          title="Click here to download template"
        >
          Click here to download template
        </Link>
      </p>
      {/* <HeaderInfo
        title="RBB Insurance Operational Checklist"
        description="Download the below RBB Insurance Operational Checklist and upload it once completed"
      />
      <p>
        <Link
          href={`${props.siteUrl}/_layouts/download.aspx?SourceUrl=${props.siteUrl}/PipelineTemplates/4. RBB Insurance Operational Checklist Template v1.6.xlsx`}
          title="Click here to download template"
        >
          Click here to download template
        </Link>
      </p> */}
      <HeaderInfo
        title="RBB Insurance Risk Classification"
        description="Download the below RBB Insurance Risk Classification and upload it once completed"
      />
      <p>
        <Link
          href={`${props.siteUrl}/_layouts/download.aspx?SourceUrl=${props.siteUrl}/PipelineTemplates/5. RBB Insurance Product Risk Classification Template.xlsx`}
          title="Click here to download template"
        >
          Click here to download template
        </Link>
      </p>
      <HeaderInfo
        title="Reset to NPS Determination"
        description="(only applicable if resetting to previous phase)"
      />
      <TextField
        title="Reset to Proposal Comment:"
        multiline
        rows={5}
        onChange={props.onChangeText}
        defaultValue={props.resetToProposal}
        id="txt_ResetNPSDComment"
      />
      <Separator />
      <Stack horizontal tokens={stackTokens} styles={stackStyles}>
        <DefaultButton onClick={props.cancelProposal} text="Cancel" />
        {props.Status === "Pipeline" && (
          <DefaultButton
            onClick={props.savePipeline}
            text="Save"
            disabled={props.buttonDisabled}
          />
        )}
        {props.Status === "Pipeline" && (
          <DefaultButton
            onClick={props.savePipeline}
            text="Reset to Proposal"
            disabled={props.buttonDisabled}
          />
        )}
        {props.Status === "Pipeline" && (
          <PrimaryButton
            text="Submit for Pipeline Review"
            onClick={props.savePipeline}
            allowDisabledFocus
            // className={styles.buttonsGroupInput}
          />
        )}
      </Stack>
    </Stack>
  );
};
export default Pipeline;
