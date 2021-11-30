import {
  DetailsList,
  DetailsListLayoutMode,
  IColumn,
  SelectionMode,
  Stack,
} from "office-ui-fabric-react";
import * as React from "react";
import HeaderInfo from "../Common/HeaderInfo";
import Utility from "../Common/Utility";

const ApprovalSummaryList = (props) => {
  const columns: IColumn[] = [
    {
      key: "InfrastructureArea",
      name: "Infrastructure Area",
      fieldName: "NAPA_Infra",
      minWidth: 100,
      maxWidth: 350,
      onRender: (item) => {
        return <span>{Utility.GetMenuItemTitle(item["NAPA_Infra"])}</span>;
      },
    },
    {
      key: "ApprovedBy",
      name: "Approved By",
      // fieldName: "Author.Title",
      minWidth: 100,
      maxWidth: 200,
      onRender: (item) => {
        return <span>{item.Author.Title}</span>;
      },
    },
    {
      key: "ApprovedDate",
      name: "Approved Date",
      // fieldName: "Created",
      minWidth: 100,
      maxWidth: 200,
      onRender: (item) => {
        return <span>{new Date(item["Created"]).toLocaleDateString()}</span>;
      },
    },
    {
      key: "ApprovedTime",
      name: "Approved Time",
      // fieldName: "Created",
      minWidth: 100,
      maxWidth: 200,
      onRender: (item) => {
        return <span>{new Date(item["Created"]).toLocaleTimeString()}</span>;
      },
    },
  ];
  if (props.Items) console.log("approvals...", props.Items);
  return (
    <Stack>
      <HeaderInfo
        title="Approval Summary"
        description="Please see conditions raised per Infrastructure"
      />
      <DetailsList
        items={props.Items}
        compact={true}
        columns={columns}
        selectionMode={SelectionMode.none}
        setKey="none"
        layoutMode={DetailsListLayoutMode.justified}
        isHeaderVisible={true}
      />
    </Stack>
  );
};

export default ApprovalSummaryList;
