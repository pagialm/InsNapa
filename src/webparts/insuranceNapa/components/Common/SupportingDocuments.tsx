import { ConsoleListener } from "@pnp/logging";
import {
  classNamesFunction,
  DetailsList,
  DetailsListLayoutMode,
  IColumn,
  Link,
  mergeStyleSets,
  SelectionMode,
  Stack,
  Text,
} from "office-ui-fabric-react";
import * as React from "react";
import AddAttachmentsPanel from "./AddAttachmentsPanel";
import HeaderInfo from "./HeaderInfo";
import { ISupportingDocItem } from "./ISupportingDocItem";

export interface ISupportingDocumentsProps {
  supportingDocs?: ISupportingDocItem[];
  addAttachments: any;
  attachmentStatus?: string;
  siteUrl: string;
  id: number;
  isAttachmentAdded?: boolean;
}
const classNames = mergeStyleSets({
  fileIconHeaderIcon: {
    padding: 0,
    fontSize: "16px",
  },
  fileIconCell: {
    textAlign: "center",
    selectors: {
      "&:before": {
        content: ".",
        display: "inline-block",
        verticalAlign: "middle",
        height: "100%",
        width: "0px",
        visibility: "hidden",
      },
    },
  },
  fileIconImg: {
    verticalAlign: "middle",
    maxHeight: "16px",
    maxWidth: "16px",
  },
  controlWrapper: {
    display: "flex",
    flexWrap: "wrap",
  },
  exampleToggle: {
    display: "inline-block",
    marginBottom: "10px",
    marginRight: "30px",
  },
  selectionDetails: {
    marginBottom: "20px",
  },
});
let siteUrl = "";
const formatDate = (date: string) => {
  return new Intl.DateTimeFormat("en-ZA", {
    year: "numeric",
    month: "numeric",
    day: "numeric",
  }).format(new Date(date));
};
const renderColumn = (
  item: ISupportingDocItem,
  index: number,
  col: IColumn
) => {
  const colValue =
    col.name === "Created"
      ? formatDate(item[col.name])
      : col.name === "Created By"
      ? item.Author["Title"]
      : item[col.key];
  // debugger;
  // console.log(colValue);
  return col.name === "Document" ? (
    <Text>
      <Link href={`${siteUrl}/Shared Documents/${item.DocumentName}`}>
        {colValue}
      </Link>
    </Text>
  ) : (
    <span>{colValue}</span>
  );
};
const renderIcon = (item: ISupportingDocItem, index: number, col: IColumn) => {
  const fileIcon =
    item.DocumentName.split(".")[1] === "msg"
      ? "presentation"
      : item.DocumentName.split(".")[1] === "xls"
      ? "xlsx"
      : item.DocumentName.split(".")[1] === "xlsm"
      ? "xlsx"
      : item.DocumentName.split(".")[1];
  const fileUrl = `https://static2.sharepointonline.com/files/fabric/assets/item-types/16/${fileIcon}.svg`;
  return (
    <img
      src={fileUrl}
      className={classNames.fileIconImg}
      alt={fileIcon + " file icon"}
    />
  );
};
const columns: IColumn[] = [
  {
    key: "document",
    name: "File Type",
    className: classNames.fileIconCell,
    isIconOnly: true,
    minWidth: 16,
    maxWidth: 16,
    iconClassName: classNames.fileIconHeaderIcon,
    onRender: renderIcon,
  },
  {
    key: "DocumentName",
    name: "Document",
    minWidth: 350,
    onRender: renderColumn,
  },
  {
    key: "Document_x0020_Type",
    name: "Document Type",
    minWidth: 150,
    data: "string",
    onRender: renderColumn,
  },
  {
    key: "Created",
    name: "Created",
    minWidth: 150,
    data: "Date",
    onRender: renderColumn,
  },
  {
    key: "Author",
    name: "Created By",
    minWidth: 170,
    data: "string",
    onRender: renderColumn,
  },
];
const SupportingDocuments = (props: ISupportingDocumentsProps) => {
  siteUrl = props.siteUrl;
  const [supportingDocs, setSupportingDocs] = React.useState([]);
  React.useEffect(() => {
    fetch(
      props.siteUrl +
        "/_api/lists/getbytitle('NAPA Supporting Documentation')/items?$filter=ProposalId eq " +
        props.id +
        "&$select=DocumentName,Author/Title,Created,Document_x0020_Type,ProposalId&$expand=Author/Title",
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
        const dataDocs: ISupportingDocItem[] = responseJson.d.results;
        if (dataDocs.length > 0) {
          setSupportingDocs([...dataDocs]);
          console.log(supportingDocs);
        }
      });
  }, [props.isAttachmentAdded]);
  return (
    <div style={{ marginLeft: "12rem", marginTop: "2rem" }}>
      <HeaderInfo title="Supporting Documentation" />
      <AddAttachmentsPanel
        addAttachments={props.addAttachments}
        attachmentsTitle="Add Attachments"
        isAttachmentAdded={props.isAttachmentAdded}
      />
      <DetailsList
        items={supportingDocs}
        compact={true}
        columns={columns}
        selectionMode={SelectionMode.none}
        layoutMode={DetailsListLayoutMode.justified}
        isHeaderVisible={true}
        enterModalSelectionOnTouch={true}
        ariaLabelForSelectionColumn="Toggle selection"
        ariaLabelForSelectAllCheckbox="Toggle selection for all items"
        checkButtonAriaLabel="Row checkbox"
      />
    </div>
  );
};

export default SupportingDocuments;
