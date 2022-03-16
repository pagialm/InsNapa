import { IStackStyles, Stack, mergeStyleSets } from "office-ui-fabric-react";
import * as React from "react";
import styles from "../InsuranceNapa.module.scss";

const stackTokens = { childrenGap: 50 };
const stackStyles: Partial<IStackStyles> = { root: { width: "100%" } };
const customStyles = mergeStyleSets({
  headerContainer:{
    border: "1px solid rgb(175,20,75)",
    color:"#fff",
    backgroundColor:"rgb(175,20,75)",
    padding: "0.2rem",
  },
  proposalTitle:{
    fontSize:"1rem",
    fontWeight:"bold",
    padding:"0.2rem",
    textAlign:"center",
    textTransform:"uppercase"
  },
  statusText:{
    color:"#fff",
    fontSize:"1rem",
  },
  dividerLine:{
    width:"50%",
    textAlign:"center",
    color:"#fff",
    borderTop:"1px solid #fff",
  },
  headerItems:{
    justifyContent:"center",
    alignItems:"center",
  },
  approvalDueDate:{
    paddingLeft:"2rem"
  }
});

export interface IHeadersState {
  proposalTitle: string;
}

export interface IHeaders {
  ApprovalDueDate?:string|Date;
  proposalId: number | string | undefined;
  selectedSection: string;
  proposalStatus: string;
  title: string;

}
class HeadersDecor extends React.Component<IHeaders, IHeadersState> {
  constructor(props: IHeaders, state: IHeadersState) {
    super(props);
    this.state = {
      proposalTitle: "",
    };
  }
  public render(): React.ReactElement<IHeaders> {
    const _dueDate:Date = typeof(this.props.ApprovalDueDate) === "string" ? new Date(this.props.ApprovalDueDate) : this.props.ApprovalDueDate;
    const approvalDueDate = this.props.ApprovalDueDate ? 
    <span className={customStyles.approvalDueDate}>Approval Due Date: {_dueDate.getDate()}/{_dueDate.getMonth() + 1}/{_dueDate.getFullYear()}</span> : <></>;
    
    return (
      <div className={customStyles.headerContainer}>
        <Stack horizontal tokens={stackTokens} styles={stackStyles} className={customStyles.headerItems}>
          <h4 className={customStyles.statusText}>
            Proposal ID: {this.props.proposalId}
          </h4>
          <h4 className={customStyles.statusText}>
            Selected Section: {this.props.selectedSection}
          </h4>
          <h4 className={customStyles.statusText}>
            Proposal Status: {this.props.proposalStatus}
          </h4>
        </Stack>
        <hr className={customStyles.dividerLine} />
        <h3 className={customStyles.proposalTitle}>{this.props.title} {approvalDueDate}</h3>
      </div>
    );
  }
}
export default HeadersDecor;
