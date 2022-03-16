import {
  PeoplePicker,
  PrincipalType,
} from "@pnp/spfx-controls-react/lib/PeoplePicker";
import {
  DatePicker,
  DefaultButton,
  Dropdown,
  IDropdownOption,
  IDropdownStyles,
  IStackProps,
  IStackStyles,
  Label,
  Link,
  MessageBar,
  MessageBarType,
  PrimaryButton,
  Separator,
  Stack,
  TextField,
  Toggle,
} from "office-ui-fabric-react";
import * as React from "react";
import HeaderInfo from "../Common/HeaderInfo";
import Headers from "../Common/Headers";
import { IEnquiryProps } from "./IEnquiryProps";
import styles from "../InsuranceNapa.module.scss";
import ScopeClarification from "../InfrastructureReview/ScopeClarification";

const stackTokens = { childrenGap: 50 };
const stackStyles: Partial<IStackStyles> = { root: { width: "100%" } };
const columnProps: Partial<IStackProps> = {
  tokens: { childrenGap: 10 },
  styles: { root: { width: 450 } },
};
const dropdownStyles: Partial<IDropdownStyles> = { dropdown: { width: 300 } };
const proposalRegions: IDropdownOption[] = [
  { key: "ARO", text: "ARO" },
  { key: "SA", text: "SA" },  
];

const SuccessExample = () => (
  <MessageBar messageBarType={MessageBarType.success} isMultiline={false}>
    You NPA has been successfully submitted!
  </MessageBar>
);
const ErrorExample = (message: string) => (
  <MessageBar
    messageBarType={MessageBarType.error}
    isMultiline={false}
    dismissButtonAriaLabel="Close"
  >
    {message}
  </MessageBar>
);
const Enquiry = (props: IEnquiryProps) => {
  return (
    <Stack styles={stackStyles}>
      {props.ID > 0 && (
        <Headers
          ApprovalDueDate={props.ApprovalDueDate}
          proposalStatus={props.Status}
          proposalId={props.ID}
          selectedSection="Enquiry"
          title={props.Title}
        />
      )}
      {(props.Status !== "Enquiry" && props.Status !== "") && (
        <ScopeClarification            
          onChangeText={props.onChangeText}
          ProposalScopeRestriction={props.ProposalScopeRestriction}
          ProposalScopeClarification={props.ProposalScopeClarification}
        />
      )}
      <HeaderInfo
        title="Application Information"
        description="Provide the following administrative information"
      />
      <Stack horizontal tokens={stackTokens} styles={stackStyles}>
        <Stack {...columnProps}>
          <TextField
            label="Proposal Name:"
            required
            value={props.Title}
            id="txt_Title"
            onChange={props.onChangeText}
            description="The proposal name should contain the key distinguishing attributes associated with the proposal."
          />
          <PeoplePicker
            context={props.context}
            titleText="Application completed by"
            personSelectionLimit={1}
            showtooltip={true}
            disabled={false}
            defaultSelectedUsers={[props.applicationCompletedBy]}
            onChange={(items: any[]) => {
              const _users = props.getPeoplePickerItems(items);
              debugger;
              if (_users.length > 0)
                props.setParentState({
                  AppCreatedById: _users[0],
                });
            }}            
            showHiddenInUI={false}
            ensureUser={true}
            principalTypes={[PrincipalType.User]}
            resolveDelay={1000}
          />
          <Label className={styles.inputDesc}>
            The Product Originator is a business representative or business
            manager, or his/her delegate. The Originator will also be the person
            raising the Proposal Template.
          </Label>
          <PeoplePicker
            context={props.context}
            titleText="P&L Owner/ General Manager"
            personSelectionLimit={3}
            showtooltip={true}
            defaultSelectedUsers={props.tradingBookOwner}
            disabled={false}
            onChange={(items: any[]) => {
              const _users = props.getPeoplePickerItems(items);
              if (_users.length > 0)
                props.setParentState({
                  TradingBookOwnerId: _users,
                });
            }}
            showHiddenInUI={false}
            ensureUser={true}
            principalTypes={[PrincipalType.User]}
            resolveDelay={1000}
          />
        </Stack>
        <Stack {...columnProps}>
          <DatePicker
            label="Target Launch Date:"
            isRequired
            value={props.targetCompletionDate}
            onSelectDate={(d: Date) => {
              props.onSelectDate("targetCompletionDate", d);
            }}
            formatDate={props.onFormatDate}
          />
          <PeoplePicker
            context={props.context}
            titleText="Sponsor"
            personSelectionLimit={3}
            showtooltip={true}
            defaultSelectedUsers={props.sponser}
            disabled={false}
            onChange={(items: any[]) => {
              const _users = props.getPeoplePickerItems(items);
              if (_users.length > 0)
                props.setParentState({
                  SponsorId: _users,
                });
            }}
            showHiddenInUI={false}
            ensureUser={true}
            principalTypes={[PrincipalType.User]}
            // resolveDelay={1000}
          />
          <Label className={styles.inputDesc}>
            The Sponsor generally is from the Business and must be Managing
            Director level.
          </Label>
          <PeoplePicker
            context={props.context}
            titleText="Product Owner"
            personSelectionLimit={3}
            showtooltip={true}
            defaultSelectedUsers={props.workstreamCoordinator}
            disabled={false}
            onChange={(items: any[]) => {
              const _users = props.getPeoplePickerItems(items);
              if (_users.length > 0)
                props.setParentState({
                  WorkStreamCoordinatorId: _users,
                });
            }}
            showHiddenInUI={false}
            ensureUser={true}
            principalTypes={[PrincipalType.User]}
            // resolveDelay={1000}
          />
        </Stack>
      </Stack>
      <HeaderInfo
        title="Product Description"
        description="This section captures the description of NAPA. It is to be concise in order for reviewers, regardless of their expertise, to understand the NAPA. More detail can be added when the full application is submitted. Please do not embed any documents in this section."
      />
      <Stack horizontal tokens={stackTokens} styles={stackStyles}>
        <Stack tokens={stackTokens} {...columnProps}>
          <Dropdown
            label="Region:"
            options={proposalRegions}
            selectedKey={props.Region}
            onChange={props.onChange}
            id="ddl_Region"
            required
          />
        </Stack>
        <Stack {...columnProps}>
          <Dropdown
            label="Country:"
            options={props.shortCountries}
            selectedKey={props.Country0}
            defaultSelectedKey={props.Country0}
            onChange={props.onChange}
            id="ddl_Country0"
            required
          />
        </Stack>
      </Stack>
      <Stack horizontal tokens={stackTokens} styles={stackStyles}>
        <Stack {...columnProps}>
          <Dropdown
            label="Entity:"            
            onChange={props.onFilteredDropdownChange}
            options={props.companies}
            id="ddl_Company"
            selectedKey={props.Company}
            required           
          />
          <Dropdown
            label="Product Family:"
            options={props.businessAreas}            
            required
            onChange={props.onChange}
            id="ddl_BusinessArea"
            selectedKey={props.BusinessArea}            
          />
          {/*  */}
        </Stack>
        <Stack {...columnProps}>
          <Dropdown
            label="Distribution Channel:"
            options={props.distributionChannels}            
            title="Distribution Channels"
            onChange={(e,o,i)=>{
              props.tansformNullArray("ProductArea0", e, o, i);
            }}
            id="ddl_ProductArea0"
            selectedKeys={props.ProductArea0}
            multiSelect
            required            
          />          
          <Dropdown
            label="Product Family Risk Classification:"
            options={props.productFamRiskClass}
            id="ddl_SubProduct"
            // defaultSelectedKey="0"
            selectedKey={props.SubProduct}
            onChange={props.onChange}
          />
        </Stack>
      </Stack>
      <TextField
        label="Executive Summary:"
        multiline
        rows={5}
        value={props.ExecutiveSummary}
        id="txt_ExecutiveSummary"
        required
        onChange={props.onChangeText}
      />
      <Stack horizontal tokens={stackTokens} styles={stackStyles}>
        <Stack {...columnProps}>
          <TextField
            label="What is new for this Proposal?"
            multiline
            rows={3}
            value={props.NewForProposal}
            onChange={props.onChangeText}
            id="txt_NewForProposal"
          />
          <TextField
            label="Link to Existing Proposal:"
            multiline
            rows={3}
            value={props.LinkToExistingProposal}
            onChange={props.onChangeText}
            id="txt_LinkToExistingProposal"
          />
        </Stack>
        <Stack {...columnProps}>
          <TextField
            label="Request to the Committee"
            multiline
            rows={3}
            value={props.LineOfCredit}
            onChange={props.onChangeText}
            id="txt_LineOfCredit"
          />
          <TextField
            label="What do you consider to be the Principal Risks associated with this proposal?"
            multiline
            rows={3}
            value={props.PrincipalRisks}
            onChange={props.onChangeText}
            id="txt_PrincipalRisks"
          />
          <Link
            title="Principle Risks.docx"    
            target="_blank" 
            data-interception="off"       
            href="https://absacorp.sharepoint.com/:w:/s/ts_RBBIPPortal/ERSoRK55CvVBs7LxPgUZ0CsBPsbK5Qd2s8YRyMnTMA6onQ?e=0pvjCR">
              Principle Risks.docx
            </Link>
        </Stack>
      </Stack>
      <Stack horizontal tokens={stackTokens} styles={stackStyles}>
        <Stack {...columnProps}>
          {/* <TextField
                label="Is the structure of the new product/transaction in any way influenced by
                          the anticipated tax treatment of any party to the transaction?"
                multiline
                rows={3}
                value={props.proposalObject.TaxTreatment}
              /> */}
          <TextField
            label="Are there any Reputational and/or Conduct Risk issues which arise from entering into
                          this new product or amended product/services? Please provide a rationale for your answer"
            multiline
            value={props.ConductRiskIssuesComments}
            rows={3}
            onChange={props.onChangeText}
            id="txt_ConductRiskIssuesComments"
          />
        </Stack>
      </Stack>
      <HeaderInfo
        title="Business Hierarchy"
        description="Provide the following country information"
      />
      <Stack horizontal styles={stackStyles} tokens={stackTokens}>
        <Stack {...columnProps}>
          <Dropdown
            placeholder="Select options"
            label="Infrastructure Support Country:"            
            onChange={(e, o, i) => {
              props.setParentState({ IFCountry: [o.text] });
            }}
            options={props.shortCountries}
            styles={dropdownStyles}
            selectedKey={props.IFCountry}
            id="ddl_IFCountry"
            required
          />
          <Dropdown
            placeholder="Select options"
            label="Target Client Location:"            
            onChange={(e, o, i) => {
              props.setParentState({ ClientLocation: [o.text] });
            }}
            selectedKey={props.ClientLocation}
            options={props.shortCountries}
            styles={dropdownStyles}
            id="ddl_ClientLocation"
            required
          />
          <Dropdown
            placeholder="Select options"
            label="Country of Product Offering:"            
            onChange={(e, o, i) => {
              props.setParentState({ ProductOfferingCountry: [o.text] });
            }}
            selectedKey={props.ProductOfferingCountry}
            options={props.shortCountries}
            styles={dropdownStyles}
            id="ddl_ProductOfferingCountry"
            required
          />
        </Stack>
        <Stack {...columnProps}>
          <Dropdown
            placeholder="Select options"
            label="Sales/Coverage Team Location:"            
            onChange={(e, o, i) => {
              props.setParentState({ SalesTeamLocation: [o.text] });
            }}
            selectedKey={props.SalesTeamLocation}
            options={props.shortCountries}
            styles={dropdownStyles}
            id="ddl_SalesTeamLocation"
            required
          />          
          <TextField
            label="Target Market"
            value={props.ClientSector}
            onChange={props.onChangeText}
            id="txt_ClientSector"
          />
          <Dropdown
            placeholder="Select options"
            label="Applicable Currencies:"            
            onChange={(e,o,i)=>{
              props.tansformNullArray("BookingCurrencies", e, o, i);
            }}
            selectedKeys={props.BookingCurrencies}
            multiSelect
            id="ddl_BookingCurrencies"
            options={props.bookingCurrencies}
            styles={dropdownStyles}
            required
          />
        </Stack>
      </Stack>
      <Separator />
      <Toggle
        label="Is this a joint venture between divisions or business area?"
        // defaultChecked
        onText="Yes"
        offText="No"
        onChange={props.onChangeToggle}
        role="checkbox"
        checked={props.JointVenture}
        id="tgl_JointVenture"
      />
      <Separator />
      {props.submitionStatus === "Ok" && <SuccessExample />}
      {props.errorMessage.length > 0 &&
        props.errorMessage.map((msg) => (
          <MessageBar
            messageBarType={MessageBarType.error}
            isMultiline={false}
            dismissButtonAriaLabel="Close"
          >
            {msg}
          </MessageBar>
        ))}
      <Stack horizontal tokens={stackTokens}>
        <DefaultButton
          text="Cancel"
          onClick={props.cancelProposal}
          allowDisabledFocus          
          disabled={props.buttonClickedDisabled}
        />
        {(props.EditMode && (props.Status === "Enquiry")) && (
          <PrimaryButton
            text="Submit for NPS Determination"
            onClick={props.saveApplicationEnquiry}
            allowDisabledFocus
            disabled={props.buttonClickedDisabled}            
          />
        )}
        {(props.EditMode && (props.Status === "Enquiry" || props.Status === "")) && (
          <PrimaryButton
            text="Save as Draft"
            onClick={props.saveApplicationEnquiry}
            allowDisabledFocus
            disabled={props.buttonClickedDisabled}            
          />
        )}
        {(props.Status !== "Enquiry" && props.Status !== "") && (
          <PrimaryButton
          text="Save Scope Clarification/Restriction"
          onClick={props.saveClarification}
          allowDisabledFocus
          disabled={props.buttonClickedDisabled}            
        />
        )}
      </Stack>
    </Stack>
  );
};
export default Enquiry;
