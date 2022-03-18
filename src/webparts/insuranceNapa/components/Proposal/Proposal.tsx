import {
  PeoplePicker,
  PrincipalType,
} from "@pnp/spfx-controls-react/lib/PeoplePicker";
import {
  DatePicker,
  DefaultButton,
  Dropdown,
  IDropdownOption,
  IStackProps,
  IStackStyles,
  PrimaryButton,
  Separator,
  Stack,
  TextField,
} from "office-ui-fabric-react";
import * as React from "react";
import HeaderInfo from "../Common/HeaderInfo";
import Headers from "../Common/Headers";
import { IProposalProps } from "./IProposalProps";
import {
  DateTimePicker,
  DateConvention,
  TimeConvention,
} from "@pnp/spfx-controls-react/lib/DateTimePicker";

const stackTokens = { childrenGap: 50 };
const stackStyles: Partial<IStackStyles> = { root: { width: "100%" } };
const columnProps: Partial<IStackProps> = {
  tokens: { childrenGap: 10 },
  styles: { root: { width: 450 } },
};
const Proposal = (props: IProposalProps) => {
  let [insBuPrcClassOptions, setInsBuPrcClassOptions] = React.useState([]);
  let [insBuPruOutcomeOptions, setInsBuPruOutcomeOptions] = React.useState([]);
  let [approvalCapacityOptions, setApprovalCapacityOptions] = React.useState(
    []
  );
  const _stageName = "NPS Determination";
  const emptyOption: IDropdownOption = { key: "0", text: "" };
  console.log(props.buPrcDate);
  // insBuPruOutcomeOptions: IDropdownOption[] = [];
  // React.useState(insBuPrcClassOptionsState)
  React.useEffect(() => {
    props
      .getItemsFilter("Insurance Dropdowns", "Stage eq 'Proposal'")
      .then((options: []) => {
        let classOptions: IDropdownOption[] = [emptyOption],
          outcomeOptions: IDropdownOption[] = [emptyOption],
          approvalCapOptions: IDropdownOption[] = [emptyOption];
        options.forEach((option) => {
          if (option["Title"] === "Insurance BU PRC Classification")
            classOptions.push({
              key: option["DropdownValue"],
              text: option["DropdownValue"],
            });
          else if (option["Title"] === "Insurance BU PRU Outcome")
            outcomeOptions.push({
              key: option["DropdownValue"],
              text: option["DropdownValue"],
            });
          else if (option["Title"] === "Approval Capacity")
            approvalCapOptions.push({
              key: option["DropdownValue"],
              text: option["DropdownValue"],
            });
        });
        debugger;
        setInsBuPrcClassOptions([...classOptions]);
        setInsBuPruOutcomeOptions([...outcomeOptions]);
        setApprovalCapacityOptions([...approvalCapOptions]);
      });
  }, []);
  return (
    
      <Stack styles={stackStyles}>
        <Headers
          ApprovalDueDate={props.ApprovalDueDate}
          proposalId={props.proposalId}
          selectedSection={_stageName}
          title={props.title}
          proposalStatus={props.proposalStatus}
        />
        <HeaderInfo
          title="New Product Services"
          description="Provide the following product information"
        />
        <Stack horizontal tokens={stackTokens} styles={stackStyles}>
          <Stack {...columnProps}>
            <Dropdown
              label="PRC Classification:"
              options={insBuPrcClassOptions}
              selectedKey={props.teamAssessment}
              onChange={props.onDdlChange}
              id="ddl_NapaTeamAssessment"
              required
            />
            <DatePicker
              label="PRC Date (In Principle Approval):"
              isRequired
              value={props.buPrcDate}
              onSelectDate={(date: Date) => {
                props.onSelectDate("BUPRCDate", date);
              }}
              formatDate={props.onFormatDate}
            />
            <Dropdown
              label="Product Family:"
              options={props.productFamilyOptions}
              selectedKey={props.productFamily}
              onChange={props.onDdlChange}
              id="ddl_ProductFamily"
              aria-readonly
              // disabled
            />

            <Dropdown
              label="Existing family or new family:"
              options={[
                { key: "Existing Family", text: "Existing Family" },
                { key: "New Family", text: "New Family" },
              ]}
              selectedKey={props.existingFamilyOrNewFamily}
              onChange={props.onDdlChange}
              id="ddl_ExistingFamily"
            />
            <TextField
              label="Actions/ Conditions/ Comments raised by PRC"
              defaultValue={props.actionsRaisedByBUPRC}
              multiline
              rows={3}
              onChange={props.onChangeText}
              id="txt_ActionsRasedByBUPRC"
            />
          </Stack>
          <Stack {...columnProps}>
            <Dropdown
              label="PRC Outcome:"
              options={insBuPruOutcomeOptions}
              selectedKey={props.teamAssesmentReason}
              onChange={props.onDdlChange}
              id="ddl_NapaTeamAssReason"
              required
            />
            <PeoplePicker
              context={props.context}
              titleText="Product Governance Team Coordinator"
              personSelectionLimit={3}
              showtooltip={true}
              defaultSelectedUsers={props.nAPATeamCoordinators}
              disabled={false}
              onChange={(items: any[]) => {
                // debugger;
                const _users = props.getPeoplePickerItems(items);
                if (_users.length > 0) props.nAPATeamCoordinators = _users;
                props.setParentState({ NAPATeamCoordinatorsId: _users });
              }}
              showHiddenInUI={false}
              ensureUser={true}
              principalTypes={[PrincipalType.User]}
              // resolveDelay={1000}
            />
            <Dropdown
              label="Product Family Risk Classification:"
              options={props.productRiskFamilyOptions}
              selectedKey={props.productFamilyRiskClassification}
              onChange={props.onDdlChange}
              id="ddl_ProductFamilyRiskClassification"
              aria-readonly
              // disabled
            />
            <Dropdown
              label="Approval Capacity:"
              options={approvalCapacityOptions}
              selectedKey={props.approvalCapacity}
              onChange={props.onDdlChange}
              id="ddl_ApprovalCapacity"
            />
            <Dropdown
              label="Infrastructure areas approved by PRC:"
              options={props.ActionOwiningAreas}
              multiSelect
              selectedKeys={props.InfrastructureAreasApprovedBPRC}
              onChange={(e,o,i)=>{
                props.tansformNullArray("InfrastructureAreasApprovedBPRC", e, o, i);
              }}
              id="ddl_InfrastructureAreasApprovedBPRC"
            />
            {/* <PeoplePicker
              context={props.context}
              titleText="Infrastructure areas approved by PRC"
              personSelectionLimit={10}
              showtooltip={true}
              defaultSelectedUsers={props.infraAreaApprovedByBUPRC}
              disabled={false}
              onChange={(items: any[]) => {
                const _users = props.getPeoplePickerItems(items);
                if (_users.length > 0) {
                  props.setParentState({ InfraAreaApprovedByBUPRCId: _users });
                }
              }}
              showHiddenInUI={false}
              ensureUser={true}
              principalTypes={[PrincipalType.User]}
              // resolveDelay={1000}
            /> */}
          </Stack>
        </Stack>
        <HeaderInfo
          title="Reset to Enquiry"
          description="(only applicable if resetting to previous phase)"
        />
        <TextField
          label="Reset Enquiry Comment:"
          multiline
          rows={3}
          value={props.resetEnquiryComment}
          onChange={props.onChangeText}
          id="txt_ResetToEnqComment"
        />
        <Separator />
        <Stack horizontal tokens={stackTokens} styles={stackStyles}>
          <DefaultButton 
            onClick={props.cancelProposal} 
            text="Cancel"
            disabled={props.buttonDisabled}
          />
          {props.EditMode && props.Status === _stageName && (
            <PrimaryButton
              onClick={props.saveProposal}
              text="Save"
              disabled={props.buttonDisabled}
            />
          )}
          {props.EditMode && props.Status === _stageName && (
            <PrimaryButton
              onClick={props.saveProposal}
              text="Reset to Enquiry"
              disabled={props.buttonDisabled}
            />
          )}
          {props.EditMode && props.Status === _stageName && (
            <PrimaryButton
              text="Submit for Pipeline"
              onClick={props.saveProposal}
              allowDisabledFocus
              disabled={props.buttonDisabled}
            />
          )}
        </Stack>
      </Stack>
    
  );
};
export default Proposal;
