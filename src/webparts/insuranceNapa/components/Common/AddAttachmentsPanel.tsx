import {
  ActionButton,
  Text,
  IIconProps,
  Panel,
  PrimaryButton,
  Stack,
  Link,
  Separator,
  DefaultButton,
} from "office-ui-fabric-react";
import * as React from "react";
import { useBoolean } from "@uifabric/react-hooks";
import { set } from "@microsoft/sp-lodash-subset";

export interface IAddAttachmentsProps {
  addAttachments: any;
  attachmentStatus?: string;
  attachmentsTitle: string;
  isAttachmentAdded?: boolean;
}
const AddAttachmentsPanel = (props: IAddAttachmentsProps) => {
  const addIcon: IIconProps = { iconName: "Add" };
  const [isOpen, setIsOpen] = React.useState(false);
  // useBoolean(false);
  let attachemntMessage,
    attachmentCode = null;
  if (props.attachmentStatus) {
    attachemntMessage = props.attachmentStatus.split(",")[0];
    attachmentCode = props.attachmentStatus.split(",")[1];
  }
  const actionPanel = () => {
    setIsOpen(!isOpen);
  };
  React.useEffect(() => {
    setIsOpen(false);
  }, [props.isAttachmentAdded]);
  return (
    <Stack>
      <ActionButton
        iconProps={addIcon}
        allowDisabledFocus
        onClick={actionPanel}
      >
        {props.attachmentsTitle}
      </ActionButton>

      <Panel
        headerText="Add Supporting Documentation"
        isOpen={isOpen}
        onDismiss={actionPanel}
        closeButtonAriaLabel="Close"
      >
        <Stack tokens={{ childrenGap: 5 }}>
          <h3>Choose a file</h3>
          <input type="file" id="btnAddAttachments_NBS_System" />
          {/* <label htmlFor="btnUploadFile"> */}
          <PrimaryButton onClick={props.addAttachments}>Upload</PrimaryButton>
          {/* </label> */}
          {attachmentCode !== null && <Separator />}
          <Text>
            {attachemntMessage}
            <DefaultButton
              onClick={() => {
                location.href = location.href;
              }}
            >
              Close
            </DefaultButton>
          </Text>
        </Stack>
      </Panel>
    </Stack>
  );
};
export default AddAttachmentsPanel;
