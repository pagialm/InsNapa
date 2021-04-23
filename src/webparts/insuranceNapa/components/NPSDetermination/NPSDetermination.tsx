import * as React from "react";
import { INPSDeterminationProps } from "./INPSDeterminationProps";
import Headers from "../Headers";
import HeaderInfo from "../HeaderInfo";
import { IStackStyles, Stack } from "office-ui-fabric-react";
import {
  TextField,
  Checkbox,
  ICheckboxProps,
  Dropdown,
  DropdownMenuItemType,
  IDropdownStyles,
  IDropdownOption,
  DatePicker,
  IStackProps,
  autobind,
  Toggle,
  Separator,
  DefaultButton,
  PrimaryButton,
  ITextFieldProps,
  getTheme,
  FontWeights,
  ITheme,
  Label,
  Button,
  BaseButton,
  MessageBar,
  MessageBarType,
  MessageBarButton,
} from "office-ui-fabric-react/lib";
import {
  SPHttpClient,
  SPHttpClientResponse,
  ISPHttpClientOptions,
} from "@microsoft/sp-http";
import {
  PeoplePicker,
  PrincipalType,
} from "@pnp/spfx-controls-react/lib/PeoplePicker";
const stackTokens = { childrenGap: 50 };
const stackStyles: Partial<IStackStyles> = { root: { width: 784 } };
const columnProps: Partial<IStackProps> = {
  tokens: { childrenGap: 10 },
  styles: { root: { width: 450 } },
};
export default class NPSDetermination extends React.Component<INPSDeterminationProps> {
  constructor(props: INPSDeterminationProps) {
    super(props);
  }
  @autobind
  private _saveApplicationProposal(e) {
    debugger;
    // console.log(this);
    const buttonClicked: string = e.target.innerText;
    let statusText = "NPS Determination";
    if (buttonClicked === "Reset to Enquiry") statusText = "Enquiry";
    if (buttonClicked === "Submit for Pipeline") statusText = "Pipeline";
    const _isFormValid =
      buttonClicked === "Save" ? true : this.props.validateForm();
    const proposal = {};

    // proposal["Title"] = this.state.Title;
    // proposal["TargetCompletionDate"] = this.state.targetCompletionDate;
    // if (this.state.AppCreatedById)
    //   proposal["AppCreatedById"] = this.state.AppCreatedById;
    // if (this.state.SponsorId) proposal["SponsorId"] = [this.state.SponsorId];
    // if (this.state.TradingBookOwnerId)
    //   proposal["TradingBookOwnerId"] = this.state.TradingBookOwnerId;
    // if (this.state.WorkStreamCoordinatorId)
    //   proposal["WorkStreamCoordinatorId"] = [
    //     this.state.WorkStreamCoordinatorId,
    //   ];
    // if (this.state.Region && this.state.Region.length > 0)
    //   proposal["Region"] = [this.state.Region];
    // proposal["Country0"] = this.state.Country0;
    // proposal["Company"] = this.state.Company;
    // proposal["BusinessArea"] = this.state.BusinessArea;
    // proposal["ExecutiveSummary"] = this.state.ExecutiveSummary;
    // proposal["ProductArea0"] = this.state.ProductArea0;
    // if (this.state.SubProduct) proposal["SubProduct"] = this.state.SubProduct;
    // proposal["NewForProposal"] = this.state.NewForProposal;
    // proposal["TransactionInPipeline"] = this.state.TransactionInPipeline;
    // proposal["LinkToExistingProposal"] = this.state.LinkToExistingProposal;
    // proposal["TaxTreatment"] = this.state.TaxTreatment;
    // proposal["LineOfCredit"] = this.state.LineOfCredit;
    // proposal[
    //   "ConductRiskIssuesComments"
    // ] = this.state.ConductRiskIssuesComments;
    // proposal["PrincipalRisks"] = this.state.PrincipalRisks;
    // if (this.state.IFCountry && this.state.IFCountry.length > 0)
    //   proposal["IFCountry"] = this.state.IFCountry;
    // if (this.state.SalesTeamLocation && this.state.SalesTeamLocation.length > 0)
    //   proposal["SalesTeamLocation"] = this.state.SalesTeamLocation;
    // if (this.state.ClientLocation && this.state.ClientLocation.length > 0)
    //   proposal["ClientLocation"] = this.state.ClientLocation;
    // if (this.state.ClientSector)
    //   proposal["ClientSector"] = this.state.ClientSector;
    // if (this.state.ProductOfferingCountry)
    //   proposal["ProductOfferingCountry"] = this.state.ProductOfferingCountry;
    // if (this.state.BookingCurrencies)
    //   proposal["BookingCurrencies"] = this.state.BookingCurrencies;
    // if (this.state.BookingLocation)
    //   proposal["BookingLocation"] = this.state.BookingLocation;
    // if (this.state.NatureOfTrade)
    //   proposal["NatureOfTrade"] = this.state.NatureOfTrade;
    // if (this.state.TraderLocation)
    //   proposal["TraderLocation"] = this.state.TraderLocation;
    // if (this.state.BookingEntity)
    //   proposal["BookingEntity"] = this.state.BookingEntity;
    // proposal["JointVenture"] = this.state.JointVenture;
    // proposal["Status"] = statusText;

    const url: string =
      this.props.context.pageContext.web.absoluteUrl +
      "/_api/web/lists/getbytitle('" +
      this.props.napaProposalsListname +
      "')/items(" +
      this.props.proposalId +
      ")";
    const spHttpClientOptions: ISPHttpClientOptions = {
      body: JSON.stringify(proposal),
    };

    if (this.props.proposalId > 0) {
      const headers: any = {
        "X-HTTP-Method": "MERGE",
        "IF-MATCH": "*",
      };
      spHttpClientOptions["headers"] = headers;
    }
    // {
    if (_isFormValid) {
      this.props.context.spHttpClient
        .post(url, SPHttpClient.configurations.v1, spHttpClientOptions)
        .then((response: SPHttpClientResponse) => {
          if (response.status >= 201 && response.status < 300) {
            // this.setState({ submitionStatus: "Ok" });
            // setTimeout(() => {
            location.href = this.props.context.pageContext.web.absoluteUrl;
            // }, 500);
          } else {
            // this.setState({
            //   errorMessage: [
            //     `Error: [HTTP]:${response.status} [CorrelationId]:${response.statusText}`,
            //   ],
            // });
          }
        });
    }
    // }
  }
  @autobind
  private _cancelProposal() {
    location.href = this.props.context.pageContext.web.absoluteUrl;
  }
  public render(): React.ReactElement<INPSDeterminationProps> {
    return (
      <div>
        <Stack>
          <Headers
            proposalId={this.props.proposalId}
            selectedSection="NPS Determination"
            title={this.props.title}
            proposalStatus={this.props.proposalStatus}
          />
          <HeaderInfo
            title="New Product Services"
            description="Provide the following product information"
          />
          <Stack horizontal tokens={stackTokens} styles={stackStyles}>
            <Stack {...columnProps}>
              <Dropdown
                label="NAPA Team Assessment:"
                options={[
                  { key: "New", text: "New" },
                  { key: "Ammended", text: "Ammended" },
                  { key: "Withdrawal", text: "Withdrawal" },
                ]}
                selectedKey={this.props.teamAssessment}
                onChange={this.props.onDdlChange}
                id="ddl_NapaTeamAssessment"
              />
              <PeoplePicker
                context={this.props.context}
                titleText="NAPA Team Coordinators"
                personSelectionLimit={3}
                showtooltip={true}
                defaultSelectedUsers={[this.props.nAPATeamCoordinators]}
                disabled={false}
                onChange={(items: any[]) => {
                  const _users = this.props.getPeoplePickerItems(items);
                  if (_users.length > 0)
                    // this.setState({
                    //   WorkStreamCoordinatorId: _users[0],
                    // });
                    debugger;
                }}
                showHiddenInUI={false}
                ensureUser={true}
                principalTypes={[PrincipalType.User]}
                // resolveDelay={1000}
              />
              <Dropdown
                label="Product Family Risk Classification:"
                options={[
                  { key: "Low", text: "Low" },
                  { key: "Medium", text: "Medium" },
                  { key: "High", text: "High" },
                ]}
                selectedKey={this.props.productFamilyRiskClassification}
                onChange={this.props.onDdlChange}
                id="ddl_ProductFamilyRiskClassification"
              />
            </Stack>
            <Stack {...columnProps}>
              <Dropdown
                label="NAPA Team Assessment Reason:"
                options={this.props.teamAssesmentReasonOptions}
                selectedKey={this.props.teamAssesmentReason}
                onChange={this.props.onDdlChange}
                id="ddl_NapaTeamAssReason"
              />
              <Dropdown
                label="Product Family:"
                options={this.props.productFamilyOptions}
                selectedKey={this.props.productFamily}
                onChange={this.props.onDdlChange}
                id="ddl_ProductFamily"
              />
              <Dropdown
                label="Approval Capacity:"
                options={[
                  { key: "0", text: "" },
                  { key: "Manufacturer", text: "Manufacturer" },
                  { key: "Distributor", text: "Distributor" },
                  {
                    key: "ManufacturerandDistributor",
                    text: "Manufacturer and Distributor",
                  },
                ]}
                selectedKey={this.props.approvalCapacity}
                onChange={this.props.onDdlChange}
                id="ddl_ApprovalCapacity"
              />
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
            value={this.props.resetEnquiryComment}
            onChange={this.props.onChangeText}
            id="txt_ResetToEnqComment"
          />
          <Separator />
          <Stack horizontal tokens={stackTokens} styles={stackStyles}>
            <DefaultButton onClick={this._cancelProposal} text="Cancel" />
            <DefaultButton
              onClick={this._saveApplicationProposal}
              text="Save"
            />
            <DefaultButton
              onClick={this._saveApplicationProposal}
              text="Reset to Enquiry"
            />
            <PrimaryButton
              text="Submit for Pipeline"
              onClick={this._saveApplicationProposal}
              allowDisabledFocus
              // className={styles.buttonsGroupInput}
            />
          </Stack>
        </Stack>
      </div>
    );
  }
}
