import { ActionButton, IIconProps, Stack } from "office-ui-fabric-react";
import * as React from "react";
import { useBoolean } from "@uifabric/react-hooks";
import ConditionsForm from "./ConditionsForm";

const AddConditionPanel = (props) => {
  const addIcon: IIconProps = { iconName: "Add" };
  const [isOpen, setIsOpen] = React.useState(false);
  // const [isOpen, { setTrue: openPanel, setFalse: dismissPanel }] =
    // useBoolean(false);
    const dismissPanel = (e) => {
      debugger;
      setIsOpen(false);
      props.RefreshConditions();
    }
  return (
    <Stack>
      <ActionButton iconProps={addIcon} allowDisabledFocus onClick={() => setIsOpen(true)}>
        {props.attachmentsTitle}
      </ActionButton>
      <ConditionsForm
        context={props.context}
        siteUrl={props.siteUrl}
        closePanel={dismissPanel}
        isPanelOpen={isOpen}
        SubmitToSP={props.SubmitToSP}
        onSelectDate={props.onSelectDate}
        internalMenuId={props.internalMenuId}
        itemId={props.itemID}
        getPeoplePickerItems={props.getPeoplePickerItems}
      />
    </Stack>
  );
};

export default AddConditionPanel;
