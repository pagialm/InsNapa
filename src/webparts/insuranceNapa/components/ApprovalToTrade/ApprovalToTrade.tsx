import {
  PeoplePicker,
  PrincipalType,
} from "@pnp/spfx-controls-react/lib/PeoplePicker";
import {
  DefaultButton,
  IStackProps,
  IStackStyles,
  mergeStyleSets,
  PrimaryButton,
  Separator,
  Stack,
  TextField,
} from "office-ui-fabric-react";
import * as React from "react";
import DisplayErrors from "../Common/DisplayErrors";
import HeaderInfo from "../Common/HeaderInfo";
import Headers from "../Common/Headers";

const stackTokens = { childrenGap: 50 };
const stackStyles: Partial<IStackStyles> = { root: { width: "100%" } };
const columnProps: Partial<IStackProps> = {
  tokens: { childrenGap: 10 },
  styles: { root: { width: 450 } },
};

const customStyles = mergeStyleSets({
  errorColor:{
    padding:"1rem",
    border: "1px solid rgb(220,0,50)",
    backgroundColor:"rgba(220,0,50,0.5)",
    color:"#000",
    marginTop:"1rem",
    marginBottom:"1rem",
  }
});
const ApprovalToTrade = (props) => {
  console.log(props);
  const strCondition = props.PreLaunchOpenConditions.length > 1 ? "conditions" : "condition";
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
        title="Chair Approval"
        description="Provide the following product information"
      />
      <Stack horizontal tokens={stackTokens} styles={stackStyles}>
        <Stack {...columnProps}>
          <PeoplePicker
            context={props.context}
            titleText="NPS Chair Approval:"
            personSelectionLimit={1}
            showtooltip={true}
            disabled={false}
            defaultSelectedUsers={[props.ProductGovernanceCustodians]}
            onChange={(items: any[]) => {
              const _users = props.getPeoplePickerItems(items);
              debugger;
              if (_users.length > 0)
                props.setParentState({
                  ATTChairId: _users[0],
                });
            }}
            // selectedItems={this._getPeoplePickerItems }
            showHiddenInUI={false}
            ensureUser={true}
            principalTypes={[PrincipalType.User]}
            resolveDelay={1000}
            required
          />
        </Stack>
        <Stack {...columnProps}>
          <TextField
            label="Approval Comments:"
            multiline
            rows={5}
            value={props.ChairComments}
            id="txt_ChairComments"
            required
            onChange={props.onChangeText}
          />
        </Stack>
      </Stack>

      <Separator />
      <TextField
        label="Reset Final NPS Review Comment:"
        multiline
        rows={5}
        value={props.ResetFinalNPSComment}
        id="txt_ResetFinalNPSComment"
        required
        onChange={props.onChangeText}
      />
      <Separator />
      {(props.PreLaunchOpenConditions.length > 0) && (
        <p className={customStyles.errorColor}>
          *** Approval to trade button has been removed due to <strong>{props.PreLaunchOpenConditions.length}</strong> open <strong>Pre Trade</strong> {strCondition}. Click on
          Approval Summary to view and Close the {strCondition} before you can approve to trade.
        </p>
      )}
      <Stack horizontal tokens={stackTokens}>
        <DefaultButton
          text="Cancel"
          onClick={props.cancelProposal}
          allowDisabledFocus
          disabled={props.buttonClickedDisabled}
        />
        {props.EditMode && props.isChairApprover && (props.Status === "Approval to Trade" || props.Status === "Chair Approval") && (
          <PrimaryButton
            text="Save"
            onClick={props.saveApprovalToTrade}
            allowDisabledFocus
            disabled={props.buttonClickedDisabled}
          />
        )}
        {props.EditMode && props.isChairApprover && (props.Status === "Approval to Trade" || props.Status === "Chair Approval") && (
          <PrimaryButton
            text="Reset to Final NPS Review"
            onClick={props.saveApprovalToTrade}
            allowDisabledFocus
            disabled={props.buttonClickedDisabled}
          />
        )}
        {props.EditMode && props.isChairApprover && ((props.Status === "Approval to Trade" || props.Status === "Chair Approval") && props.PreLaunchOpenConditions.length === 0) && (
          <PrimaryButton
            text="Approve to Trade"
            onClick={props.saveApprovalToTrade}
            allowDisabledFocus
            disabled={props.buttonClickedDisabled}
          />
        )}

        {props.errorMessage.length > 0 && (
          <Stack>
            <p id="ErrorsDisplay"></p>
            <DisplayErrors
              ErrorMessages={props.errorMessage}
              Target={"#ErrorsDisplay"}
            />
          </Stack>
        )}
      </Stack>      
    </Stack>
  );
};
export default ApprovalToTrade;
