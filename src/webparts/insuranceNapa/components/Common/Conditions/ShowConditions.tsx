import {
  CommandButton,
  DetailsList,
  DetailsListLayoutMode,
  IColumn,
  IContextualMenuItem,
  IContextualMenuProps,
  IIconProps,
  SelectionMode,
  Text,
} from "office-ui-fabric-react";
import * as React from "react";
import AddConditionPanel from "./AddConditionPanel";
import ConditionsForm from "./ConditionsForm";
import { IConditionItem } from "./IConditionItem";
import { useBoolean } from "@uifabric/react-hooks";

const ShowConditions = (props) => {
  const [isOpen, { setTrue: openPanel, setFalse: dismissPanel }] =
    useBoolean(false);
  const [isSubmitVisible, setIsSubmitVisible] = React.useState(true);
  const [conditionsItems, setConditionsItems] = React.useState([]);
  const [stageConditionsItems, setStageConditionsItems] = React.useState([]);
  const [itemID, setItemID] = React.useState(0);
  const subMenu = props.submenu;

  const SetConditions = (proposalConditions?: any[]) => {
    debugger;
    const allConditions = proposalConditions
      ? proposalConditions
      : conditionsItems;
    let currStageConditions: any[] = [];

    currStageConditions = allConditions.filter(
      (condition) => condition.RaisingArea === subMenu["subtile"]
    );

    setStageConditionsItems([...currStageConditions]);
  };

  const CollectConditions = () => {
    fetch(
      props.siteUrl +
        "/_api/lists/getbytitle('Infrastructure Conditions')/items?$filter=NAPA_ID eq " +
        props.id +
        "&$select=ID,ConditionStatus,ActionDueDate,ActionOwningArea,RaisingArea,ActionOwnerBy/Title,ConditionRaisedBy/Title,Created,Type,NAPA_ID,DescriptionOfCOA&$expand=ConditionRaisedBy/Title,ActionOwnerBy/Title",
      {
        method: "GET",
        headers: {
          "Content-Type": "application/JSON; odata=verbose",
          accept: "application/JSON; odata=verbose",
        },
      }
    )
      .then((response) => response.json())
      .then((responseJson) => {
        const dataConditions: IConditionItem[] = responseJson.d.results;
        if (dataConditions.length > 0) {
          setConditionsItems([...dataConditions]);
          SetConditions(dataConditions);
          // console.log(conditionsItems);
        }
      });
  };

  React.useEffect(() => {
    CollectConditions();
  }, []);
  if (props.submenu)
    React.useEffect(() => {
      SetConditions();
    }, [subMenu["subtile"]]);

  const formatDate = (date: string) => {
    return new Intl.DateTimeFormat("en-ZA", {
      year: "numeric",
      month: "numeric",
      day: "numeric",
    }).format(new Date(date));
  };
  const actionCondition = (ev?:React.MouseEvent<HTMLElement,MouseEvent>, item?:IContextualMenuItem) => {
    if(item.key === "viewCondition")
      setIsSubmitVisible(false);
    else
      setIsSubmitVisible(true);
    debugger;
    setItemID(parseInt(item.itemProps.itemID));
    openPanel();
  };

  const renderColumn = (
    item: { ConditionRaisedBy: string; ActionOwnerBy: string },
    index: number,
    col: IColumn
  ) => {
    // debugger;
    const colValue =
      col.name === "Action Due Date"
        ? formatDate(item[col.key])
        : col.name === "Raised By"
        ? item.ConditionRaisedBy["Title"]
        : col.name === "Action Owner By"
        ? item.ActionOwnerBy["Title"]
        : item[col.key];
    // debugger;
    // console.log(colValue);
    const menuProps: IContextualMenuProps = {
      items: [
        {
          key: "viewCondition",
          text: "View Condition",
          iconProps: { iconName: "ViewIcon" },
          onClick: actionCondition,
          itemProps: { itemID: item["ID"] },
        },
        {
          key: "editCondition",
          text: "Edit Condition",
          iconProps: { iconName: "EditIcon" },
          onClick: actionCondition,
          itemProps: { itemID: item["ID"] },
        },
      ],
    };
    const addIcon: IIconProps = { iconName: "Add" };
    return col.key === "ConditionAction" ? (
      <CommandButton iconProps={addIcon} text="Action" menuProps={menuProps} />
    ) : (
      <span>{colValue}</span>
    );
  };
  const columns: IColumn[] = [
    {
      key: "ID",
      name: "ID",
      // className: classNames.fileIconCell,
      minWidth: 25,
      maxWidth: 32,
      // iconClassName: classNames.fileIconHeaderIcon,
      onRender: renderColumn,
      isResizable: true,
    },
    {
      key: "DescriptionOfCOA",
      name: "Description of COA",
      minWidth: 100,
      maxWidth: 200,
      onRender: renderColumn,
      isResizable: true,
    },
    {
      key: "Type",
      name: "Type",
      minWidth: 20,
      maxWidth: 40,
      data: "string",
      onRender: renderColumn,
      isResizable: true,
    },
    {
      key: "RaisingArea",
      name: "Raising Area",
      minWidth: 80,
      maxWidth: 90,
      data: "string",
      onRender: renderColumn,
      isResizable: true,
    },
    {
      key: "ConditionRaisedBy",
      name: "Raised By",
      minWidth: 80,
      maxWidth: 110,
      data: "string",
      onRender: renderColumn,
      isResizable: true,
    },
    {
      key: "ActionOwnerBy",
      name: "Action Owner By",
      minWidth: 80,
      maxWidth: 110,
      data: "string",
      onRender: renderColumn,
      isResizable: true,
    },
    {
      key: "ActionOwningArea",
      name: "Action Owning Area",
      minWidth: 80,
      maxWidth: 100,
      data: "string",
      onRender: renderColumn,
      isResizable: true,
    },
    {
      key: "ActionDueDate",
      name: "Action Due Date",
      minWidth: 80,
      maxWidth: 90,
      data: "Date",
      onRender: renderColumn,
      isResizable: true,
    },
    {
      key: "ConditionStatus",
      name: "Status",
      minWidth: 60,
      maxWidth: 70,
      data: "string",
      onRender: renderColumn,
      isResizable: true,
    },
    {
      key: "ConditionAction",
      name: "Action",
      minWidth: 80,
      onRender: renderColumn,
      isResizable: true,
    },
  ];

  return (
    <div>
      {props.showAddPanel && (
        <AddConditionPanel
          addAttachments={props.addAttachments}
          attachmentsTitle="Add Condition"
          context={props.context}
          siteUrl={props.siteUrl}
          SubmitToSP={props.SubmitToSP}
          onSelectDate={props.onSelectDate}
          itemID={props.id}
          internalMenuId={props.internalMenuId}
          getPeoplePickerItems={props.getPeoplePickerItems}
        />
      )}

      <DetailsList
        items={!props.showAddPanel ? conditionsItems : stageConditionsItems}
        compact={true}
        columns={columns}
        selectionMode={SelectionMode.none}
        setKey="set"
        layoutMode={DetailsListLayoutMode.justified}
        isHeaderVisible={true}
        enterModalSelectionOnTouch={true}
        ariaLabelForSelectionColumn="Toggle selection"
        ariaLabelForSelectAllCheckbox="Toggle selection for all items"
        checkButtonAriaLabel="Row checkbox"
        data-is-scrollable={true}
      />

      <ConditionsForm
        context={props.context}
        siteUrl={props.siteUrl}
        closePanel={dismissPanel}
        isPanelOpen={isOpen}
        itemID={itemID}
        SubmitToSP={props.SubmitToSP}
        isSubmitVisible={isSubmitVisible}
      />
    </div>
  );
};

export default ShowConditions;
