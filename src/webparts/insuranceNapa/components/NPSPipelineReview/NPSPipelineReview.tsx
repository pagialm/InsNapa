import {
  DatePicker,
  DefaultButton,
  Dropdown,
  IStackProps,
  IStackStyles,
  PrimaryButton,
  Separator,
  Stack,
  TextField,
} from "office-ui-fabric-react";
import Headers from "../Common/Headers";
import HeaderInfo from "../Common/HeaderInfo";
import * as React from "react";
import {
  PeoplePicker,
  PrincipalType,
} from "@pnp/spfx-controls-react/lib/PeoplePicker";
import { INPSPipelineReviewProps } from "./INPSPipelineReviewProps";
import AddAttachmentsPanel from "../Common/AddAttachmentsPanel";
import DisplayErrors from "../Common/DisplayErrors";

const stackTokens = { childrenGap: 50 };
const stackStyles: Partial<IStackStyles> = { root: { width: "100%" } };
const columnProps: Partial<IStackProps> = {
  tokens: { childrenGap: 10 },
  styles: { root: { width: 450 } },
};

const NPSPipelineReview = (props: INPSPipelineReviewProps) => {
  return (
    <Stack styles={stackStyles}>
      <Headers
        proposalId={props.proposalId}
        selectedSection="NPS Pipeline Review"
        title={props.title}
        proposalStatus={props.proposalStatus}
      />
      <HeaderInfo
        title="New Product Services"
        description="Please complete the below"
      />
      <Stack horizontal tokens={stackTokens} styles={stackStyles}>
        <Stack {...columnProps}>
          <Dropdown
            id="ddl_RiskRanking"
            options={props.riskRankingOptions}
            selectedKey={props.RiskRanking}
            onChange={props.onChangeDropdown}
            label="Initial Risk Ranking:"
            required
          />
          <PeoplePicker
            context={props.context}
            titleText="Sponsor Sign Off:"
            personSelectionLimit={1}
            showtooltip={true}
            defaultSelectedUsers={[props.BusinessCaseApprovalFrom]}
            disabled={false}
            onChange={(items: any[]) => {
              const _users = props.getPeoplePickerItems(items);
              if (_users.length > 0)
                props.setParentState({
                  BusinessCaseApprovalFromId: _users,
                });
            }}
            showHiddenInUI={false}
            ensureUser={true}
            principalTypes={[PrincipalType.User]}
            resolveDelay={1000}
          />
          <TextField
            label="NPS Pipeline Review Comments:"
            multiline
            rows={3}
            value={props.BusinessCaseApprovalComment}
            onChange={props.onChangeText}
            id="txt_BusinessCaseApprovalComment"
          />
          <DatePicker
            label="Proposed date for PRC final Sanction"
            isRequired
            value={props.nAPABriefingDate}
            onSelectDate={(date: Date) => {
              props.onSelectDate("NAPABriefingDate", date);
            }}
            formatDate={props.onFormatDate}
          />
          <DatePicker
            label="Target Risk and Functional Area Sign-off Date"
            isRequired
            value={props.targetSubmissionByBusiness}
            onSelectDate={(date: Date) => {
              props.onSelectDate("TargetSubmissionByBusiness", date);
            }}
            formatDate={props.onFormatDate}
          />
          <AddAttachmentsPanel
            addAttachments={props.addAttachments}
            attachmentsTitle="Attach BU PRC minutes"
            isAttachmentAdded={props.isAttachmentAdded}
          />
        </Stack>
        <Stack {...columnProps}>
          <DatePicker
            label="Sponsor Approval Date:"
            isRequired
            value={props.businessCaseApprovalDate}
            onSelectDate={(date: Date) => {
              props.onSelectDate("BusinessCaseApprovalDate", date);
            }}
            formatDate={props.onFormatDate}
          />
          <DatePicker
            label="Target Business Go Live:"
            isRequired
            value={props.targetBusinessGoLive}
            onSelectDate={(date: Date) => {
              props.onSelectDate("TargetBusinessGoLive", date);
            }}
            formatDate={props.onFormatDate}
          />
          <AddAttachmentsPanel
            addAttachments={props.addAttachments}
            attachmentsTitle="Attach Sponsor approval"
            isAttachmentAdded={props.isAttachmentAdded}
          />
        </Stack>
      </Stack>
      <HeaderInfo
        title="Infrastructure Services"
        description="Provide the following addition participants information"
      />
      <Stack horizontal tokens={stackTokens} styles={stackStyles}>
        <Stack {...columnProps}>
          <PeoplePicker
            context={props.context}
            titleText="CRO"
            personSelectionLimit={3}
            showtooltip={true}
            defaultSelectedUsers={props.ITReviewer}
            disabled={false}
            onChange={(items: any[]) => {
              const _users = props.getPeoplePickerItems(items);
              if (_users.length > 0)
                props.setParentState({
                  ITReviewerId: _users,
                });
            }}
            showHiddenInUI={false}
            ensureUser={true}
            principalTypes={[PrincipalType.User]}
            resolveDelay={1000}
          />
          <PeoplePicker
            context={props.context}
            titleText="Legal Risk:"
            personSelectionLimit={10}
            showtooltip={true}
            defaultSelectedUsers={props.LegalReviewer}
            disabled={false}
            onChange={(items: any[]) => {
              const _users = props.getPeoplePickerItems(items);
              if (_users.length > 0)
                props.setParentState({
                  LegalReviewerId: _users,
                });
            }}
            showHiddenInUI={false}
            ensureUser={true}
            principalTypes={[PrincipalType.User]}
            resolveDelay={1000}
          />
          <PeoplePicker
            context={props.context}
            titleText="Financial Crime"
            personSelectionLimit={3}
            showtooltip={true}
            defaultSelectedUsers={props.FinancialCrimeReviewer}
            disabled={false}
            onChange={(items: any[]) => {
              const _users = props.getPeoplePickerItems(items);
              if (_users.length > 0)
                props.setParentState({
                  FinancialCrimeReviewerId: _users,
                });
            }}
            showHiddenInUI={false}
            ensureUser={true}
            principalTypes={[PrincipalType.User]}
            resolveDelay={1000}
          />
          <PeoplePicker
            context={props.context}
            titleText="Data Privacy"
            personSelectionLimit={3}
            showtooltip={true}
            defaultSelectedUsers={props.CreditRiskReviwer}
            disabled={false}
            onChange={(items: any[]) => {
              const _users = props.getPeoplePickerItems(items);
              if (_users.length > 0)
                props.setParentState({
                  CreditRiskReviwerId: _users,
                });
            }}
            showHiddenInUI={false}
            ensureUser={true}
            principalTypes={[PrincipalType.User]}
            resolveDelay={1000}
          />
          <PeoplePicker
            context={props.context}
            titleText="Fraud Risk"
            personSelectionLimit={3}
            showtooltip={true}
            defaultSelectedUsers={props.FraudRiskReviewer}
            disabled={false}
            onChange={(items: any[]) => {
              const _users = props.getPeoplePickerItems(items);
              if (_users.length > 0)
                props.setParentState({
                  FraudRiskReviewerId: _users,
                });
            }}
            showHiddenInUI={false}
            ensureUser={true}
            principalTypes={[PrincipalType.User]}
            resolveDelay={1000}
          />
          <PeoplePicker
            context={props.context}
            titleText="Tax Risk"
            personSelectionLimit={3}
            showtooltip={true}
            defaultSelectedUsers={props.TaxReviewer}
            disabled={false}
            onChange={(items: any[]) => {
              const _users = props.getPeoplePickerItems(items);
              if (_users.length > 0)
                props.setParentState({
                  TaxReviewerId: _users,
                });
            }}
            showHiddenInUI={false}
            ensureUser={true}
            principalTypes={[PrincipalType.User]}
            resolveDelay={1000}
          />
          <PeoplePicker
            context={props.context}
            titleText="Informantion Security Risk and Cyber Risk"
            personSelectionLimit={3}
            showtooltip={true}
            defaultSelectedUsers={props.MarketRiskReviewer}
            disabled={false}
            onChange={(items: any[]) => {
              const _users = props.getPeoplePickerItems(items);
              if (_users.length > 0)
                props.setParentState({
                  MarketRiskReviewerId: _users,
                });
            }}
            showHiddenInUI={false}
            ensureUser={true}
            principalTypes={[PrincipalType.User]}
            resolveDelay={1000}
          />
          <PeoplePicker
            context={props.context}
            titleText="Finance"
            personSelectionLimit={3}
            showtooltip={true}
            defaultSelectedUsers={props.FinanceReviewer}
            disabled={false}
            onChange={(items: any[]) => {
              const _users = props.getPeoplePickerItems(items);
              if (_users.length > 0)
                props.setParentState({
                  FinanceReviewerId: _users,
                });
            }}
            showHiddenInUI={false}
            ensureUser={true}
            principalTypes={[PrincipalType.User]}
            resolveDelay={1000}
          />
          <PeoplePicker
            context={props.context}
            titleText="Head of Actuarial and Statutory Actuary"
            personSelectionLimit={3}
            showtooltip={true}
            defaultSelectedUsers={props.ProductControlReviewer}
            disabled={false}
            onChange={(items: any[]) => {
              const _users = props.getPeoplePickerItems(items);
              if (_users.length > 0)
                props.setParentState({
                  ProductControlReviewerId: _users,
                });
            }}
            showHiddenInUI={false}
            ensureUser={true}
            principalTypes={[PrincipalType.User]}
            resolveDelay={1000}
          />
          <PeoplePicker
            context={props.context}
            titleText="Marketing and Communications"
            personSelectionLimit={3}
            showtooltip={true}
            defaultSelectedUsers={props.RegulatoryReportingReviewer}
            disabled={false}
            onChange={(items: any[]) => {
              const _users = props.getPeoplePickerItems(items);
              if (_users.length > 0)
                props.setParentState({
                  RegulatoryReportingReviewerId: _users,
                });
            }}
            showHiddenInUI={false}
            ensureUser={true}
            principalTypes={[PrincipalType.User]}
            resolveDelay={1000}
          />
          <PeoplePicker
            context={props.context}
            titleText="RBB CVM"
            personSelectionLimit={3}
            showtooltip={true}
            defaultSelectedUsers={props.CRMReviewer}
            disabled={false}
            onChange={(items: any[]) => {
              const _users = props.getPeoplePickerItems(items);
              if (_users.length > 0)
                props.setParentState({
                  CRMReviewerId: _users,
                });
            }}
            showHiddenInUI={false}
            ensureUser={true}
            principalTypes={[PrincipalType.User]}
            resolveDelay={1000}
          />
        </Stack>
        <Stack {...columnProps}>
          <PeoplePicker
            context={props.context}
            titleText="Financial & Insurance Risk"
            personSelectionLimit={3}
            showtooltip={true}
            defaultSelectedUsers={props.TreasuryReviewer}
            disabled={false}
            onChange={(items: any[]) => {
              const _users = props.getPeoplePickerItems(items);
              if (_users.length > 0)
                props.setParentState({
                  TreasuryReviewerId: _users,
                });
            }}
            showHiddenInUI={false}
            ensureUser={true}
            principalTypes={[PrincipalType.User]}
            resolveDelay={1000}
          />
          <PeoplePicker
            context={props.context}
            titleText="Compliance"
            personSelectionLimit={3}
            showtooltip={true}
            defaultSelectedUsers={props.ComplianceReviwer}
            disabled={false}
            onChange={(items: any[]) => {
              const _users = props.getPeoplePickerItems(items);
              if (_users.length > 0)
                props.setParentState({
                  ComplianceReviwerId: _users,
                });
            }}
            showHiddenInUI={false}
            ensureUser={true}
            principalTypes={[PrincipalType.User]}
            resolveDelay={1000}
          />
          <PeoplePicker
            context={props.context}
            titleText="Operations"
            personSelectionLimit={3}
            showtooltip={true}
            defaultSelectedUsers={props.OperationsReviewer}
            disabled={false}
            onChange={(items: any[]) => {
              const _users = props.getPeoplePickerItems(items);
              if (_users.length > 0)
                props.setParentState({
                  OperationsReviewerId: _users,
                });
            }}
            showHiddenInUI={false}
            ensureUser={true}
            principalTypes={[PrincipalType.User]}
            resolveDelay={1000}
          />
          <PeoplePicker
            context={props.context}
            titleText="Supplier Risk"
            personSelectionLimit={3}
            showtooltip={true}
            defaultSelectedUsers={props.TreasuryRiskReviewer}
            disabled={false}
            onChange={(items: any[]) => {
              const _users = props.getPeoplePickerItems(items);
              if (_users.length > 0)
                props.setParentState({
                  TreasuryRiskReviewerId: _users,
                });
            }}
            showHiddenInUI={false}
            ensureUser={true}
            principalTypes={[PrincipalType.User]}
            resolveDelay={1000}
          />
          <PeoplePicker
            context={props.context}
            titleText="Financial Reporting/ Control Risk"
            personSelectionLimit={3}
            showtooltip={true}
            defaultSelectedUsers={props.FinancialReportingReviewer}
            disabled={false}
            onChange={(items: any[]) => {
              const _users = props.getPeoplePickerItems(items);
              if (_users.length > 0)
                props.setParentState({
                  FinancialReportingReviewerId: _users,
                });
            }}
            showHiddenInUI={false}
            ensureUser={true}
            principalTypes={[PrincipalType.User]}
            resolveDelay={1000}
          />
          <PeoplePicker
            context={props.context}
            titleText="Technology Risk"
            personSelectionLimit={3}
            showtooltip={true}
            defaultSelectedUsers={props.IRMReviewer}
            disabled={false}
            onChange={(items: any[]) => {
              const _users = props.getPeoplePickerItems(items);
              if (_users.length > 0)
                props.setParentState({
                  IRMReviewerId: _users,
                });
            }}
            showHiddenInUI={false}
            ensureUser={true}
            principalTypes={[PrincipalType.User]}
            resolveDelay={1000}
          />
          <PeoplePicker
            context={props.context}
            titleText="Business Continuity Risk"
            personSelectionLimit={3}
            showtooltip={true}
            defaultSelectedUsers={props.GroupResilienceReviewer}
            disabled={false}
            onChange={(items: any[]) => {
              const _users = props.getPeoplePickerItems(items);
              if (_users.length > 0)
                props.setParentState({
                  GroupResilienceReviewerId: _users,
                });
            }}
            showHiddenInUI={false}
            ensureUser={true}
            principalTypes={[PrincipalType.User]}
            resolveDelay={1000}
          />
          <PeoplePicker
            context={props.context}
            titleText="Valuations"
            personSelectionLimit={3}
            showtooltip={true}
            defaultSelectedUsers={props.ConductRiskReviewer}
            disabled={false}
            onChange={(items: any[]) => {
              const _users = props.getPeoplePickerItems(items);
              if (_users.length > 0)
                props.setParentState({
                  ConductRiskReviewerId: _users,
                });
            }}
            showHiddenInUI={false}
            ensureUser={true}
            principalTypes={[PrincipalType.User]}
            resolveDelay={1000}
          />
          <PeoplePicker
            context={props.context}
            titleText="Reinsurance"
            personSelectionLimit={3}
            showtooltip={true}
            defaultSelectedUsers={props.ReinsuranceReviewer}
            disabled={false}
            onChange={(items: any[]) => {
              const _users = props.getPeoplePickerItems(items);
              if (_users.length > 0)
                props.setParentState({
                  ReinsuranceReviewerId: _users,
                });
            }}
            showHiddenInUI={false}
            ensureUser={true}
            principalTypes={[PrincipalType.User]}
            resolveDelay={1000}
          />
          <PeoplePicker
            context={props.context}
            titleText="Customer Experience"
            personSelectionLimit={3}
            showtooltip={true}
            defaultSelectedUsers={props.CustomerExperienceReviewer}
            disabled={false}
            onChange={(items: any[]) => {
              const _users = props.getPeoplePickerItems(items);
              if (_users.length > 0)
                props.setParentState({
                  CustomerExperienceReviewerId: _users,
                });
            }}
            showHiddenInUI={false}
            ensureUser={true}
            principalTypes={[PrincipalType.User]}
            resolveDelay={1000}
          />
          <PeoplePicker
            context={props.context}
            titleText="Distribution"
            personSelectionLimit={3}
            showtooltip={true}
            defaultSelectedUsers={props.DistributionReviewer}
            disabled={false}
            onChange={(items: any[]) => {
              const _users = props.getPeoplePickerItems(items);
              if (_users.length > 0)
                props.setParentState({
                  DistributionReviewerId: _users,
                });
            }}
            showHiddenInUI={false}
            ensureUser={true}
            principalTypes={[PrincipalType.User]}
            resolveDelay={1000}
          />
        </Stack>
      </Stack>
      <HeaderInfo
        title="Reset to Pipeline"
        description="(only applicable if resetting to previous phase)"
      />
      <TextField
        title="Reset to Pipeline Comment:"
        multiline
        rows={5}
        onChange={props.onChangeText}
        defaultValue={props.resetToPipeline}
        id="txt_ResetNPSDComment"
      />
      <Separator />
      <Stack horizontal tokens={stackTokens} styles={stackStyles}>
        <DefaultButton onClick={props.cancelProposal} text="Cancel" />
        {props.EditMode && props.Status === "NPS Pipeline Review" && (
          <DefaultButton
            onClick={props.savePipelineReview}
            text="Save"
            disabled={props.buttonDisabled}
          />
        )}
        {props.EditMode && props.Status === "NPS Pipeline Review" && (
          <DefaultButton
            onClick={props.savePipelineReview}
            text="Reset to Pipeline"
            disabled={props.buttonDisabled}
          />
        )}
        {props.EditMode && props.Status === "NPS Pipeline Review" && (
          <PrimaryButton
            text="Release Infrastructure Approval"
            onClick={props.savePipelineReview}
            allowDisabledFocus
            // className={styles.buttonsGroupInput}
          />
        )}
      </Stack>
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
  );
};
export default NPSPipelineReview;
