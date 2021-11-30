import { ActionButton, IIconProps, Stack } from "office-ui-fabric-react";
import * as React from "react";
import { useBoolean } from "@uifabric/react-hooks";
import ConditionsForm from "./ConditionsForm";

const AddConditionPanel = (props) => {
  const addIcon: IIconProps = { iconName: "Add" };
  const [isOpen, { setTrue: openPanel, setFalse: dismissPanel }] =
    useBoolean(false);
  return (
    <Stack>
      <ActionButton iconProps={addIcon} allowDisabledFocus onClick={openPanel}>
        {props.attachmentsTitle}
      </ActionButton>
      <ConditionsForm
        context={props.context}
        siteUrl={props.siteUrl}
        closePanel={dismissPanel}
        isPanelOpen={isOpen}
        SubmitToSP={props.SubmitToSP}
      />
    </Stack>
  );
};

export default AddConditionPanel;
