import { IStackProps, IStackStyles, Stack, TextField } from "office-ui-fabric-react";
import * as React from "react";
import HeaderInfo from "../Common/HeaderInfo";

const stackTokens = { childrenGap: 50 };
const stackStyles: Partial<IStackStyles> = { root: { width: "100%" } };
const columnProps: Partial<IStackProps> = {
    tokens: { childrenGap: 10 },
    styles: { root: { width: 450 } },
  };

const ScopeClarification = (props) => {
    return (
        <Stack styles={stackStyles}>
            <HeaderInfo
                title="Proposal Scope Clarification/Restriction"
                description="Provide the following scope clarification information"
            />
            <Stack horizontal tokens={stackTokens} styles={stackStyles}>
                <Stack {...columnProps}>
                    <TextField
                        label="Proposal Scope Clarification:"
                        multiline
                        rows={6}
                        value={props.ProposalScopeClarification}
                        onChange={props.onChangeText}
                        id="txt_ProposalScopeClarification"
                    />
                </Stack>
                <Stack {...columnProps}>
                    <TextField
                        label="Proposal Scope Restriction:"
                        multiline
                        rows={6}
                        value={props.ProposalScopeRestriction}
                        onChange={props.onChangeText}
                        id="txt_ProposalScopeRestriction"
                    />
                </Stack>
            </Stack>
        </Stack>
    );
}

export default ScopeClarification;