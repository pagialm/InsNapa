import { IStackStyles, Stack } from "office-ui-fabric-react";
import * as React from "react";
import styles from "../InsuranceNapa.module.scss";

const stackTokens = { childrenGap: 50 };
const stackStyles: Partial<IStackStyles> = { root: { width: 784 } };

export interface IHeadersState {
  proposalTitle: string;
}

export interface IHeaders {
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
    return (
      <div>
        <Stack horizontal tokens={stackTokens} styles={stackStyles}>
          <h4 className={styles["statusText"]}>
            Proposal ID: {this.props.proposalId}
          </h4>
          <h4 className={styles["statusText"]}>
            Selected Section: {this.props.selectedSection}
          </h4>
          <h4 className={styles["statusText"]}>
            Proposal Status: {this.props.proposalStatus}
          </h4>
        </Stack>
        <h3 style={{ fontWeight: 700, fontSize: 13.5 }}>{this.props.title}</h3>
      </div>
    );
  }
}
export default HeadersDecor;
