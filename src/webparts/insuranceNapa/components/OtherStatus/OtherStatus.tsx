import {
  DatePicker,
  DefaultButton,
  Dropdown,
  IDropdownStyles,
  IStackProps,
  IStackStyles,
  PrimaryButton,
  Separator,
  Stack,
  TextField,
} from "office-ui-fabric-react";
import * as React from "react";
import HeaderInfo from "../Common/HeaderInfo";
import HeadersDecor from "../Common/Headers";
import { IOtherStatus } from "./IOtherStatus";

const stackTokens = { childrenGap: 50 };
const stackStyles: Partial<IStackStyles> = { root: { width: "100%" } };
const columnProps: Partial<IStackProps> = {
  tokens: { childrenGap: 10 },
  styles: { root: { width: 450 } },
};
const dropdownStyles: Partial<IDropdownStyles> = { dropdown: { width: 300 } };

const os_proposalObj = {
  otherStatusListname: "Other Statuses",
};

const OtherStatus = (props) => {
  const [otherStatusOptions, setOtherStatusOption] = React.useState([]);
  const [isButtonEnabled, setIsButtonEnabled] = React.useState(false);
  const [otherStatusDate, setOtherStatusDate] = React.useState(null);
  const [approvalStatusDate, setApprovalStatusDate] = React.useState(null);

  //   console.log("props...", props);
  if (props.OtherStatusDate)
    setOtherStatusDate(new Date(props.OtherStatusDate));
  if (props.ProposalDateWithdrawal)
    setApprovalStatusDate(new Date(props.ProposalDateWithdrawal));

  //   debugger;
  console.log("Date...", approvalStatusDate);

  React.useEffect(() => {
    const loadOtherStatus = (data: any[]) => {
      const otherStatuses = data.map((d) => {
        return { key: d["Title"], text: d["Title"] };
      });
      setOtherStatusOption([...otherStatuses]);
    };

    fetch(
      props.siteUrl +
        "/_api/lists/getbytitle('" +
        os_proposalObj.otherStatusListname +
        "')/items",
      {
        method: "GET",
        headers: {
          "Content-Type": "application/JSON; odata=verbose",
          accept: "application/JSON; odata=verbose",
        },
      }
    )
      .then((response) => response.json())
      .then((responseData) => loadOtherStatus(responseData.d.results));
  }, []);

  const submitOtherStatus = (e) => {
    debugger;
  };
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
        title="Other Status"
        description="Please complete the below"
      />
      <Stack horizontal tokens={stackTokens} styles={stackStyles}>
        <Stack {...columnProps}>
          <Dropdown
            label="Other Status:"
            options={otherStatusOptions}
            selectedKey={props.OtherStatuses}
            onChange={props.onChange}
            id="ddl_OtherStatuses"
            required
          />
          <DatePicker
            label="Approval Withdrawn Date:"
            isRequired
            value={approvalStatusDate}
            onSelectDate={(d: Date) => {
              setApprovalStatusDate(d);
            }}
            formatDate={props.onFormatDate}
          />
          <TextField
            label="Comments:"
            multiline
            rows={5}
            value={props.OtherStatusComments}
            id="txt_OtherStatusComments"
            required
            onChange={props.onChangeText}
          />
        </Stack>
        <Stack {...columnProps}>
          <DatePicker
            label="Other Status Date:"
            isRequired
            value={otherStatusDate}
            onSelectDate={(d: Date) => {
              setOtherStatusDate(d);
            }}
            formatDate={props.onFormatDate}
          />
        </Stack>
      </Stack>
      <Separator />
      <Stack horizontal tokens={stackTokens} styles={stackStyles}>
        <DefaultButton
          onClick={() => {
            location.href = props.siteUrl;
          }}
          text="Cancel"
        />

        {true && (
          <PrimaryButton
            onClick={submitOtherStatus}
            text="Submit to Other Status"
            disabled={isButtonEnabled}
          />
        )}
      </Stack>
    </Stack>
  );
};

export default OtherStatus;
