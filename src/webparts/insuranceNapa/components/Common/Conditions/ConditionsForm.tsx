import {
  PeoplePicker,
  PrincipalType,
} from "@pnp/spfx-controls-react/lib/PeoplePicker";
import {
  DatePicker,
  DefaultButton,
  Dropdown,
  Text,
  Panel,
  PrimaryButton,
  Separator,
  Stack,
  TextField,
  Toggle,
  IDropdownOption,
  PanelType,
} from "office-ui-fabric-react";
import * as React from "react";
import { useBoolean } from "@uifabric/react-hooks";

import { IConditionItem } from "./IConditionItem";
import {
  SPHttpClient,
  SPHttpClientResponse,
  ISPHttpClientOptions,
} from "@microsoft/sp-http";

const ConditionsForm = (props) => {
  let attachemntMessage,
    attachmentCode = null;
  const [riskCategories, setRiskCategories] = React.useState([]);
  const [actionCategories, setActionCategories] = React.useState([]);
  const [actionOwningArea, setActionOwningArea] = React.useState([]);
  const [controlCategory, setControlCategory] = React.useState([]);
  const [controlOwningArea, setControlOwningArea] = React.useState([]);
  const [isButtonEnabled, setIsButtonEnabled] = React.useState(false);

  let riskCategoriesOptions: IDropdownOption[] = [];
  let actionCategoriesOptions: IDropdownOption[] = [];
  let actionOwningAreaOptions: IDropdownOption[] = [];
  let controlCategoryOptions: IDropdownOption[] = [];
  let controlOwningAreaOptions: IDropdownOption[] = [];
  const dropDownsItems = [
    {
      listName: "Risk Categories",
      varName: riskCategoriesOptions,
      stateName: riskCategories,
      stateFn: setRiskCategories,
    },
    {
      listName: "Action Categories",
      varName: actionCategoriesOptions,
      stateName: actionCategories,
      stateFn: setActionCategories,
    },
    {
      listName: "Action Owning Area",
      varName: actionOwningAreaOptions,
      stateName: actionOwningArea,
      stateFn: setActionOwningArea,
    },
    {
      listName: "Control Category",
      varName: controlCategoryOptions,
      stateName: controlCategory,
      stateFn: setControlCategory,
    },
    {
      listName: "Contro Owning Area",
      varName: controlOwningAreaOptions,
      stateName: controlOwningArea,
      stateFn: setControlOwningArea,
    },
  ];
  const [conditionsItems, setConditionsItems] = React.useState([]);
  const [actionOwnerBy, setActionOwnerBy] = React.useState(null);
  const [conditionsItem, setConditionsItem] = React.useState({});
  const [actionDueDate, setActionDueDate] = React.useState(null);
  const [dateConditionedRaised, setDateConditionedRaised] =
    React.useState(null);
  const [dateOfClosure, setDateOfClosure] = React.useState(null);

  if (props.attachmentStatus) {
    attachemntMessage = props.attachmentStatus.split(",")[0];
    attachmentCode = props.attachmentStatus.split(",")[1];
  }
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

  const getUsers = (filter?: string) => {
    const url: string =
      props.context.pageContext.site.absoluteUrl +
      "/_api/web/siteusers" +
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

  const getListIem = (listName: string, _itemID?: number) => {
    const url: string =
      props.context.pageContext.site.absoluteUrl +
      "/_api/web/lists/getbytitle('" +
      listName +
      "')/items(" +
      _itemID +
      ")";
    return props.context.spHttpClient
      .get(url, SPHttpClient.configurations.v1)
      .then((response) => {
        return response.json();
      })
      .then((jsonResponse) => {
        return jsonResponse;
      }) as Promise<any>;
  };

  React.useEffect(() => {
    // load dropdowns
    dropDownsItems.forEach((dropD) => {
      getListIems(dropD.listName, "").then((items) => {
        // setReviewItems(items);
        items.forEach((item) => {
          dropD.varName = [
            ...dropD.varName,
            { key: item.Title, text: item.Title },
          ];
        });
        dropD.stateFn([...dropD.varName]);
      });
    });
    getListIems(
      "Infrastructure Conditions",
      "?$filter=NAPA_ID eq '" + props.ID + "'"
    ).then((items) => {
      // setReviewItems(items);
      items.forEach((item) => {
        // console.log("items:", reviewItems);
        setConditionsItems([...conditionsItems, item]);
      });
      // console.log("items:", conditionsItems);
    });
  }, []);
  // console.log("itemID...", props.itemID);
  React.useEffect(() => {
    getListIem("Infrastructure Conditions", props.itemID).then(
      (_item: IConditionItem) => {
        console.log(_item);
        setConditionsItem({ ..._item });
        if (_item["ActionDueDate"]) {
          setActionDueDate(new Date(_item["ActionDueDate"] as string));
        }
        if (_item["DateConditionedRaised"]) {
          setDateConditionedRaised(
            new Date(_item["DateConditionedRaised"] as string)
          );
        }
        if (_item["DateOfClosure"]) {
          setDateOfClosure(new Date(_item["DateOfClosure"] as string));
        }
        // console.log(conditionDates);
        let usersFilter = `?$filter=Id eq ${_item["ActionOwnerById"]} or Id eq ${_item["ConditionRaisedById"]} or Id eq ${_item["ControlOwnedById"]}`;
        if (_item["ClosedById"])
          usersFilter += ` or Id eq ${_item["ClosedById"]}`;
        getUsers(usersFilter).then((users) => {
          console.log("batho ke ba", users);
          if (conditionsItem["ActionOwnerById"]) {
            const _actionOwnerBy = users.filter(
              (user) => user.Id === conditionsItem["ActionOwnerById"]
            );
            if (_actionOwnerBy.length > 0)
              setActionOwnerBy(_actionOwnerBy[0].Title);
          }
        });
      }
    );
  }, [props.itemID]);

  const submitCondition = (e) => {
    // console.log(e);
    setIsButtonEnabled(!isButtonEnabled);
    const isNewItem: boolean = conditionsItem["ID"] ? false : true;
    const apiReport = (data: SPHttpClientResponse) => {
      // console.log(data);
      setIsButtonEnabled(!isButtonEnabled);
    };
    props.SubmitToSP(
      "NAPA Infrastructure Questions",
      isNewItem,
      conditionsItem,
      apiReport
    );
  };

  return (
    <Panel
      headerText="Add Condition"
      isOpen={props.isPanelOpen}
      onDismiss={props.closePanel}
      closeButtonAriaLabel="Close"
      type={PanelType.medium}
    >
      <Stack tokens={{ childrenGap: 5 }}>
        <h3>Contition of Approval</h3>
        <hr />
        <TextField
          label="Description of Condition of Approval:"
          required
          value={conditionsItem["DescriptionOfCOA"]}
          id="txt_DescriptionOfCOA"
          onChange={props.onChangeText}
        />
        <Toggle
          label="Does the condition need  to be satisfied prior to
        go live?"
          // defaultChecked
          onText="Yes"
          offText="No"
          onChange={props.onChangeToggle}
          role="checkbox"
          checked={conditionsItem["ApprovalBeforeGoLIve"]}
          id="tgl_ApprovalBeforeGoLIve"
        />
        <DatePicker
          label="Date Condition Raised:"
          isRequired
          value={dateConditionedRaised}
          onSelectDate={(d: Date) => {
            props.onSelectDate("targetCompletionDate", d);
          }}
          formatDate={props.onFormatDate}
        />
        <PeoplePicker
          context={props.context}
          titleText="Condition Raised By:"
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
        <Dropdown
          label="Condition Status:"
          options={[
            { key: "Open", text: "Open" },
            { key: "Closed", text: "Closed" },
          ]}
          selectedKey={conditionsItem["ConditionStatus"]}
          onChange={props.onChange}
          id="ddl_ConditionStatus"
          required
        />
        <Dropdown
          label="Type:"
          options={[
            { key: "CO", text: "CO" },
            { key: "PA", text: "PA" },
          ]}
          selectedKey={conditionsItem["Type"]}
          onChange={props.onChange}
          id="ddl_Type"
          required
        />

        <Dropdown
          label="Risk Category:"
          options={riskCategories}
          selectedKey={conditionsItem["RiskCategory"]}
          onChange={props.onChange}
          id="ddl_RiskCategory"
          required
        />
        <TextField
          label="Description of Risk:"
          multiline
          rows={5}
          value={conditionsItem["DescriptionOfRisk"]}
          id="txt_DescriptionOfRisk"
          required
          onChange={props.onChangeText}
        />
        <TextField
          label="Action To Remove Risk:"
          multiline
          rows={5}
          value={conditionsItem["ActionToRemoveRisk"]}
          id="txt_ActionToRemoveRisk"
          required
          onChange={props.onChangeText}
        />

        <Dropdown
          label="Action Category:"
          options={actionCategories}
          selectedKey={conditionsItem["ActionCategory"]}
          onChange={props.onChange}
          id="ddl_ActionCategory"
          required
        />
        <PeoplePicker
          context={props.context}
          titleText="Action Owned By:"
          personSelectionLimit={3}
          showtooltip={true}
          defaultSelectedUsers={actionOwnerBy}
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
        <DatePicker
          label="Action Due Date:"
          isRequired
          value={actionDueDate}
          onSelectDate={(d: Date) => {
            props.onSelectDate("targetCompletionDate", d);
          }}
          formatDate={props.onFormatDate}
        />

        <Dropdown
          label="Action Owning Area:"
          options={actionOwningArea}
          selectedKey={conditionsItem["ActionOwningArea"]}
          onChange={props.onChange}
          id="ddl_ActionOwningArea"
          required
        />
        <TextField
          label="What Control is in place until this action can be delivered?:"
          multiline
          rows={5}
          value={conditionsItem["ControlInPlace"]}
          id="txt_ControlInPlace"
          required
          onChange={props.onChangeText}
        />
        <TextField
          label="How is this control being monitored?:"
          multiline
          rows={5}
          value={conditionsItem["ControlMonitoring"]}
          id="txt_ControlMonitoring"
          required
          onChange={props.onChangeText}
        />
        <Dropdown
          label="Control Category:"
          options={controlCategory}
          selectedKey={conditionsItem["ControlCategory"]}
          onChange={props.onChange}
          id="ddl_ControlCategory"
          required
        />
        <Dropdown
          label="Control Owning Area:"
          options={controlOwningArea}
          selectedKey={conditionsItem["ControlOwningArea"]}
          onChange={props.onChange}
          id="ddl_Region"
          required
        />
        <PeoplePicker
          context={props.context}
          titleText="Control Owned By:"
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
        <PeoplePicker
          context={props.context}
          titleText="Closed By:"
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
        <DatePicker
          label="Date of Closure:"
          isRequired
          value={dateOfClosure}
          onSelectDate={(d: Date) => {
            props.onSelectDate("targetCompletionDate", d);
          }}
          formatDate={props.onFormatDate}
        />
        <Toggle
          label="Evidence of accepted responsibility attached?"
          // defaultChecked
          onText="Yes"
          offText="No"
          onChange={props.onChangeToggle}
          role="checkbox"
          checked={conditionsItem["Attachments"]}
          id="tgl_Attachments"
        />
        <h3>Attach a file</h3>
        <input type="file" id="btnAddAttachments_NBS_System" />
        {/* <label htmlFor="btnUploadFile"> */}

        <PrimaryButton onClick={submitCondition}>Save Condition</PrimaryButton>
        <DefaultButton onClick={props.closePanel}>Close</DefaultButton>
        {attachmentCode !== null && <Separator />}
        <Text>{attachemntMessage}</Text>
      </Stack>
    </Panel>
  );
};

export default ConditionsForm;
