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
  
  const [actionOwningArea, setActionOwningArea] = React.useState([]);  
  const [controlOwningArea, setControlOwningArea] = React.useState([]);
  const [isButtonEnabled, setIsButtonEnabled] = React.useState(false);

  
  let actionOwningAreaOptions: IDropdownOption[] = [];  
  let controlOwningAreaOptions: IDropdownOption[] = [];
  const dropDownsItems = [       
    {
      listName: "Action Owning Area",
      varName: actionOwningAreaOptions,
      stateName: actionOwningArea,
      stateFn: setActionOwningArea,
    },    
    {
      listName: "Contro Owning Area",
      varName: controlOwningAreaOptions,
      stateName: controlOwningArea,
      stateFn: setControlOwningArea,
    },
  ];
  const [conditionsItems, setConditionsItems] = React.useState([]);
  const [actionOwnerBy, setActionOwnerBy] = React.useState("");
  const [conditionRaisedBy, setConditionRaisedBy] = React.useState("");
  const [closedBy, setClosedBy] = React.useState("");
  const [conditionsItem, setConditionsItem] = React.useState({});
  const [actionDueDate, setActionDueDate] = React.useState(null);
  const [dateConditionedRaised, setDateConditionedRaised] =
    React.useState(null);
  const [dateOfClosure, setDateOfClosure] = React.useState(null);
  const [viewSubmit, setViewSubmit] = React.useState(true);

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
    // debugger
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
    // debugger;
    //Set default fields --Fields will aways share these values
    const fieldIdentifier = {
      NAPA_Link : `${props.itemId}_${props.internalMenuId}`,
      NAPA_ID : `${props.itemId}`,
      
    };
      setConditionsItem({...conditionsItem, ...fieldIdentifier});
    // load dropdowns
    dropDownsItems.forEach((dropD) => {
      getListIems(dropD.listName, "").then((items) => {        
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
        items.forEach((item) => {        
        setConditionsItems([...conditionsItems, item]);
      });
         
    });
  }, []);
  
  React.useEffect(() => {
    if(props.itemID)
      getListIem("Infrastructure Conditions", props.itemID).then(
        (_item: IConditionItem) => {
          console.log(_item);
          if(_item.ID)
            setViewSubmit(false);
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
          let usersFilter = `?$filter=Id eq ${_item["ActionOwnerById"]} or Id eq ${_item["ConditionRaisedById"]}`;
          if (_item["ClosedById"])
            usersFilter += ` or Id eq ${_item["ClosedById"]}`;
          getUsers(usersFilter).then((users) => {
            console.log("batho ke ba", users);
            console.log("item...", _item)
            if (_item["ActionOwnerById"]) {
              const _actionOwnerBy = users.filter(
                (user) => user.Id === _item["ActionOwnerById"]
              );
              if (_actionOwnerBy.length > 0)
                setActionOwnerBy(_actionOwnerBy[0].Title);
            }
            if (_item["ConditionRaisedById"]) {
              const _ConditionRaisedBy = users.filter(
                (user) => user.Id === _item["ConditionRaisedById"]
              );
              if (_ConditionRaisedBy.length > 0)
                setConditionRaisedBy(_ConditionRaisedBy[0].Title);
            }
            if (_item["ClosedById"]) {
              const _closedBy = users.filter(
                (user) => user.Id === _item["ClosedById"]
              );
              if (_closedBy.length > 0){
                setClosedBy(_closedBy[0].Title);
                setConditionsItem({...conditionsItem, ConditionStatus:"Closed"})
              }
            }
          });
        }
      );
  }, [props.itemID]);
  /**
   * Submit the Condition to SharePoint.
   * @param e Event 
   */
  const submitCondition = (e) => {
    // Set the SharePoint Contidion Object 
    const _spItem = {
      ActionDueDate: conditionsItem["ActionDueDate"],
      ActionOwnerById: conditionsItem["ActionOwnerById"],      
      ActionOwningArea: conditionsItem["ActionOwningArea"],
      ActionToRemoveRisk: conditionsItem["ActionToRemoveRisk"],
      ApprovalBeforeGoLIve: conditionsItem["ApprovalBeforeGoLIve"],           
      ClosedById: conditionsItem["ClosedById"],      
      ConditionRaisedById: conditionsItem["ConditionRaisedById"],      
      ConditionStatus: conditionsItem["ConditionStatus"],      
      ControlInPlace: conditionsItem["ControlInPlace"],      
      DateConditionedRaised: conditionsItem["DateConditionedRaised"],
      DateOfClosure: conditionsItem["DateOfClosure"],
      DescriptionOfCOA: conditionsItem["DescriptionOfCOA"],
      DescriptionOfRisk: conditionsItem["DescriptionOfRisk"],
      ItemID: conditionsItem["ItemID"],
      NAPA_ID: conditionsItem["NAPA_ID"],
      NAPA_Link: conditionsItem["NAPA_Link"],      
      RaisingArea: conditionsItem["RaisingArea"],     
      Type: conditionsItem["Type"],
    };
    //Disable the buttons to avoide double clicking
    setIsButtonEnabled(!isButtonEnabled);
    const isNewItem: boolean = conditionsItem["ID"] ? false : true;
    //Set the item Id for updates
    if(!isNewItem)
      _spItem["ID"] = conditionsItem["ID"];
    //Return function after the item was sent to SharePoint
    const apiReport = (data: SPHttpClientResponse) => {      
      setIsButtonEnabled(!isButtonEnabled);
      props.closePanel();
    };
    //Submit condition to SharePoint
    props.SubmitToSP(
      "Infrastructure Conditions",
      isNewItem,
      _spItem,
      apiReport
    );
  };

  const onChangeText = (
    ev: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>,
    newValue: string
  ): void => {
    const element = ev.target as HTMLElement;
    // debugger;
    const el = {};
    el[element.id.split("_")[1]] = newValue;
    setConditionsItem({ ...conditionsItem, ...el });
  };

  const onChange = (
    ev: React.FormEvent<HTMLDivElement>,
    option?: IDropdownOption,
    index?: number
  ): void => {
    const selectedElement = ev.target as HTMLDivElement;
    const elementId = selectedElement.id.split("_")[1].split("-")[0];
    const el = {};

    // debugger;
    //Todo: Find old element, check its type... array or primitive and act accodingly
    const oldState = conditionsItem[elementId]; //el[elementId];
    if (
      Array.isArray(oldState) ||
      (selectedElement["type"] && selectedElement["type"] === "checkbox")
    ) {
      // debugger;
      const newArray = oldState ? [...oldState] : [];
      if (
        oldState &&
        oldState.some((opt) => {
          return opt === option.key;
        })
      ) {
        const finalArray = newArray.filter((itemToFilter) => {
          return itemToFilter !== option.key;
        });
        el[elementId] = finalArray;
      } else {
        newArray.push(option.key);
        el[elementId] = newArray;
      }
    } else {
      el[elementId] = elementId === "Region" ? [option.key] : option.key;
    }
    setConditionsItem({ ...conditionsItem, ...el });
  };

  /**
   * Get users from People picker control
   * @param items People picker items
   * @returns People Picker objects
   */
  const getPeoplePickerItems = (items: any[]): any[] => {
    // debugger;
    let getSelectedUsers = [];
    for (let item in items) {
      getSelectedUsers.push(items[item].id);
    }    
    return getSelectedUsers;
  }
  /**
   * Reset the Condition item state.
   * @param e  Close button
   */
  const resetClosePanel = (e):void =>{
    const emptyItem = {
      ActionDueDate: null,
      ActionOwnerById: null,      
      ActionOwningArea: "",
      ActionToRemoveRisk: "",
      ApprovalBeforeGoLIve: false,           
      ClosedById: null,      
      ConditionRaisedById: null,      
      ConditionStatus: "",      
      ControlInPlace: "",      
      DateConditionedRaised: "",
      DateOfClosure: null,
      DescriptionOfCOA: "",
      DescriptionOfRisk: "",
      ItemID: "",
      NAPA_ID: "",
      NAPA_Link: "",      
      RaisingArea: "",     
      Type: "",
    };
    setConditionsItem(emptyItem);
    setActionOwnerBy("");
    setConditionRaisedBy("");
    setClosedBy("");
    props.closePanel();
  };

  const onFormatDate = (date: Date): string => {
    const _date: Date = typeof date === "string" ? new Date(date) : date;
    return (
      _date.getDate() + "/" + (_date.getMonth() + 1) + "/" + _date.getFullYear()
    );
  }
  
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
          onChange={onChangeText}
        />
        <Toggle
          label="Does the condition need  to be satisfied prior to
        go live?"
          // defaultChecked={conditionsItem["ApprovalBeforeGoLIve"]}
          onText="Yes"
          offText="No"
          onChange={(e,c) => {
            setConditionsItem({...conditionsItem, ApprovalBeforeGoLIve: c})
          }}
          role="checkbox"
          checked={conditionsItem["ApprovalBeforeGoLIve"]}
          id="tgl_ApprovalBeforeGoLIve"
        />
        <DatePicker
          label="Date Condition Raised:"
          isRequired
          value={dateConditionedRaised}
          onSelectDate={(d: Date) => {
            setDateConditionedRaised(d);
            setConditionsItem({...conditionsItem, DateConditionedRaised : d});
          }}
          formatDate={onFormatDate}
        />
        <PeoplePicker
          context={props.context}
          titleText="Condition Raised By:"
          personSelectionLimit={1}
          showtooltip={true}
          defaultSelectedUsers={[conditionRaisedBy]}
          disabled={false}
          onChange={(items: any[]) => {
            const _users = getPeoplePickerItems(items);
            if (_users.length > 0)
              setConditionsItem({...conditionsItem, ConditionRaisedById: _users[0]});              
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
          selectedKey={conditionsItem["ConditionStatus"]?conditionsItem["ConditionStatus"]:"Open"}
          disabled={true}
          onChange={onChange}
          id="ddl_ConditionStatus"
          required
        />
        <Dropdown
          label="Condition Type:"
          options={[
            { key: "PreLaunch", text: "Pre Launch" },
            { key: "PostLaunch", text: "Post Launch" },
          ]}
          selectedKey={conditionsItem["Type"]}
          onChange={onChange}
          id="ddl_Type"
          required
        />
        
        <TextField
          label="Description of Risk:"
          multiline
          rows={5}
          value={conditionsItem["DescriptionOfRisk"]}
          id="txt_DescriptionOfRisk"
          required
          onChange={onChangeText}
        />
        <TextField
          label="Action To Remove Risk:"
          multiline
          rows={5}
          value={conditionsItem["ActionToRemoveRisk"]}
          id="txt_ActionToRemoveRisk"
          required
          onChange={onChangeText}
        />
        
        <PeoplePicker
          context={props.context}
          titleText="Action Owned By:"
          personSelectionLimit={1}
          showtooltip={true}
          defaultSelectedUsers={[actionOwnerBy]}
          disabled={false}
          onChange={(items: any[]) => {
            const _users = getPeoplePickerItems(items);
            if (_users.length > 0)
              setConditionsItem({...conditionsItem, ActionOwnerById: _users[0]});
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
            setActionDueDate(d);
            setConditionsItem({...conditionsItem, ActionDueDate : d});                 
          }}
          formatDate={onFormatDate}
        />

        <Dropdown
          label="Action Owning Area:"
          options={actionOwningArea}
          selectedKey={conditionsItem["ActionOwningArea"]}
          onChange={onChange}
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
          onChange={onChangeText}
        />
       
        <PeoplePicker
          context={props.context}
          titleText="Closed By:"
          personSelectionLimit={1}
          showtooltip={true}
          defaultSelectedUsers={[closedBy]}
          disabled={false}
          onChange={(items: any[]) => {
            const _users = getPeoplePickerItems(items);
            if (_users.length > 0)
              setConditionsItem({...conditionsItem, ClosedById: _users[0], ConditionStatus: "Closed"});
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
            setDateOfClosure(d);
            setConditionsItem({...conditionsItem, DateOfClosure : d}); 
          }}
          formatDate={onFormatDate}
        />        
        {(props.isSubmitVisible || viewSubmit) && (
          <PrimaryButton onClick={submitCondition}>Save Condition</PrimaryButton>
        )}
        
        <DefaultButton onClick={resetClosePanel}>Close</DefaultButton>
        {attachmentCode !== null && <Separator />}
        <Text>{attachemntMessage}</Text>
      </Stack>
    </Panel>
  );
};

export default ConditionsForm;
