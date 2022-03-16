import {
  PeoplePicker,
  PrincipalType,
} from "@pnp/spfx-controls-react/lib/PeoplePicker";
import {
  Checkbox,
  DatePicker,
  DefaultButton,
  Dropdown,
  IDropdownOption,
  IDropdownStyles,
  IStackProps,
  IStackStyles,
  PrimaryButton,
  Separator,
  Stack,
  TextField,
} from "office-ui-fabric-react";
import * as React from "react";
import AddAttachmentsPanel from "../Common/AddAttachmentsPanel";
import HeaderInfo from "../Common/HeaderInfo";
import Headers from "../Common/Headers";
import Utility from "../Common/Utility";

const stackTokens = { childrenGap: 50 };
const stackStyles: Partial<IStackStyles> = { root: { width: "100%" } };
const columnProps: Partial<IStackProps> = {
  tokens: { childrenGap: 10 },
  styles: { root: { width: 450 } },
};
const dropdownStyles: Partial<IDropdownStyles> = { dropdown: { width: 300 } };
const stackTokens2 = { childrenGap: 10 };

const FinalNPSReview = (props) => {
  const yesNoOptions: IDropdownOption[] = [
    { text: "", key: "0" },
    { text: "No", key: "No" },
    { text: "Yes", key: "Yes" },
  ];
  const FinalRiskClassificationOptions: IDropdownOption[] = [
    { text: "", key: "0" },
    { text: "High", key: "High" },
    { text: "Medium", key: "Medium" },
    { text: "Low", key: "Low" },
  ];
  const InsuranceExcoPruOutcomeOptions: IDropdownOption[] = [
    { text: "", key: "0" },
    { text: "Approved", key: "Approved" },
    { text: "Deffered", key: "Deffered" },
    { text: "Declined", key: "Declined" },
  ];
  const [resetInfraArray, setResetInfraArray] = React.useState([]);
  const _onChange = (
    ev: React.FormEvent<HTMLInputElement>,
    isChecked: boolean
  ) => {
    const infraArea = ev.currentTarget["ariaLabel"];
    console.log(ev.currentTarget["ariaLabel"]);
    debugger;
    if (isChecked) setResetInfraArray([...resetInfraArray, infraArea]);
    else {
      const tempArray = [...resetInfraArray];
      const tempArray2 = tempArray.filter((item) => {
        return item !== infraArea;
      });
      setResetInfraArray([...tempArray2]);
    }
  };
  const _resetToInfraReview = (ev) => {
    props.ResetToInfrastructureReview(resetInfraArray);
  };
  console.log("props...", props);
  return (
    <Stack styles={stackStyles}>
      <Headers
        ApprovalDueDate={props.ApprovalDueDate}
        proposalId={props.ID}
        selectedSection={props.SelectedSection}
        title={props.Title}
        proposalStatus={props.Status}
      />
      <HeaderInfo
        title="New Product Services"
        description="Provide the following product information"
      />
      <Stack horizontal tokens={stackTokens} styles={stackStyles}>
        <Stack {...columnProps}>
          <PeoplePicker
            context={props.context}
            titleText="Business Executive/ Sponsor:"
            personSelectionLimit={1}
            showtooltip={true}
            disabled={false}
            defaultSelectedUsers={[props.BusinesExecutive]}
            onChange={(items: any[]) => {
              const _users = props.getPeoplePickerItems(items);
              debugger;
              if (_users.length > 0)
                props.setParentState({
                  BIRORegionalHeadId: _users[0],
                });
            }}
            // selectedItems={this._getPeoplePickerItems }
            showHiddenInUI={false}
            ensureUser={true}
            principalTypes={[PrincipalType.User]}
            resolveDelay={1000}
          />
        </Stack>
        <Stack {...columnProps}>
          <DatePicker
            label="Sponsor/ Business Executive Approval Date:"
            isRequired
            value={props.BusinesExecutiveApprovalDate}
            onSelectDate={(d: Date) => {
              props.onSelectDate("BIRORegionalHeadReviewDate", d);
            }}
            formatDate={props.onFormatDate}
          />
        </Stack>
      </Stack>
      <HeaderInfo
        title="ARO PRC (Product Risk Committee)"
        description="Only required for ARO NAPA’s"
      />
      <Stack horizontal tokens={stackTokens} styles={stackStyles}>
        <Stack {...columnProps}>
          <DatePicker
            label="Insurance Exco PRC Date:"
            isRequired
            value={props.InduranceExcoPrcDate}
            onSelectDate={(d: Date) => {
              props.onSelectDate("CROStatusDate", d);
            }}
            formatDate={props.onFormatDate}
          />
          <TextField
            label="Exco PRC Committee Comment:"
            multiline
            rows={5}
            value={props.ExcoPrcCommitteeComment}
            id="txt_CROComment"
            required
            onChange={props.onChangeText}
          />
          <TextField
            label="Actions/ Conditions raised by Insurance Exco PRC (Pre/Post):"
            multiline
            rows={5}
            value={props.ActionsRaisedByExco}
            id="txt_ActionsRaisedByExco"
            required
            onChange={props.onChangeText}
          />
          <AddAttachmentsPanel
            addAttachments={props.addAttachments}
            attachmentsTitle="Attach Exco PRC minutes"
          />
        </Stack>
        <Stack {...columnProps}>
          <Dropdown
            label="Insurance Exco PRU Outcome:"
            options={InsuranceExcoPruOutcomeOptions}
            id="ddl_CROStatus"
            // defaultSelectedKey="0"
            selectedKey={props.InsuranceExcoPruOutcome}
            onChange={props.onChange}
          />
          <Dropdown
            label="Final Risk classification:"
            options={FinalRiskClassificationOptions}
            id="ddl_FinalRiskClassification"
            // defaultSelectedKey="0"
            selectedKey={props.FinalRiskClassification}
            onChange={props.onChange}
          />
          <Dropdown
            label="Operational Checklist requirements completed:"
            options={yesNoOptions}
            id="ddl_OperationalChecklistRequirement"
            // defaultSelectedKey="0"
            selectedKey={props.OperationalChecklistRequirement}
            onChange={props.onChange}
          />
        </Stack>
      </Stack>
      <HeaderInfo
        title="Post Implementation Review"
        description="Please complete the below PIR Information"
      />
      <Stack horizontal tokens={stackTokens} styles={stackStyles}>
        <Stack {...columnProps}>
          <Dropdown
            label="Is Post Implementation Required:"
            options={yesNoOptions}
            id="ddl_IsPostImplementationRequired"
            // defaultSelectedKey="0"
            selectedKey={props.IsPostImplementationRequired}
            onChange={props.onChange}
            required
          />
          <DatePicker
            label="PIR Date Completed:"
            value={props.PirDateCompleted}
            onSelectDate={(d: Date) => {
              props.onSelectDate("PIRDateCompleted", d);
            }}
            formatDate={props.onFormatDate}
          />
        </Stack>
        <Stack {...columnProps}>
          <DatePicker
            label="PIR Launch Date:"
            value={props.PirLaunchDate}
            onSelectDate={(d: Date) => {
              props.onSelectDate("TargetDueDate", d);
            }}
            formatDate={props.onFormatDate}
          />
          <TextField
            label="PIR Comments:"
            multiline
            rows={5}
            value={props.PirComments}
            id="txt_PIRComments"
            onChange={props.onChangeText}
          />
        </Stack>
      </Stack>
      <HeaderInfo
        title="Reset to Infrastructure Approval"
        description="(only applicable if resetting to previous phase)"
      />
      <Stack horizontal tokens={stackTokens} styles={stackStyles}>
        <Stack {...columnProps} tokens={stackTokens2}>
          {
            props.ApprovedItems.map(approvedItem => {              
              return <Checkbox label={Utility.GetMenuItemTitle(approvedItem["NAPA_Infra"])} onChange={_onChange} />
            })
          }          
        </Stack>
        <Stack {...columnProps} tokens={stackTokens2}>
          
        </Stack>
      </Stack>
      <Separator />
      <TextField
        label="Reset Comments:"
        multiline
        rows={5}
        value={props.ResetFinalNPSComment}
        id="txt_ResetFinalNPSComment"
        required
        onChange={props.onChangeText}
      />
      <Separator />
      <Stack horizontal tokens={stackTokens}>
        <DefaultButton
          text="Cancel"
          onClick={props.cancelProposal}
          allowDisabledFocus
          disabled={props.buttonClickedDisabled}
        />
        {props.EditMode && (props.Status === "Final NPS Review" || props.Status === "") && (
          <PrimaryButton
            text="Save"
            onClick={props.saveFinalNPSReview}
            allowDisabledFocus
            disabled={props.buttonClickedDisabled}
          />
        )}
        {props.EditMode && (props.Status === "Final NPS Review" || props.Status === "") && (
          <PrimaryButton
            text="Reset to Infrastructure Approval"
            onClick={_resetToInfraReview}
            allowDisabledFocus
            disabled={props.buttonClickedDisabled}
          />
        )}
        {(props.EditMode && (props.Status === "Final NPS Review" || props.Status === "")) && (
          <PrimaryButton
            text="Submit for Approval to Trade"
            onClick={props.saveFinalNPSReview}
            allowDisabledFocus
            disabled={props.buttonClickedDisabled}
          />
        )}
      </Stack>
    </Stack>
  );
};

export default FinalNPSReview;
