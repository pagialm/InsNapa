import { DatePicker, DefaultButton, IStackProps, IStackStyles, PrimaryButton, Stack, TextField } from "office-ui-fabric-react";
import * as React from "react";
import HeaderInfo from "../Common/HeaderInfo";

const stackTokens = { childrenGap: 50 };
const stackStyles: Partial<IStackStyles> = { root: { width: "100%" } };
const columnProps: Partial<IStackProps> = {
    tokens: { childrenGap: 10 },
    styles: { root: { width: 450 } },
  };

const PostApproval = (props) => {
    return (
        <Stack styles={stackStyles}>
            <HeaderInfo
                title="Post Approval Details"
                description="Please see conditions raised per Infrastructure"
            />
            <Stack horizontal tokens={stackTokens} styles={stackStyles}>
                <Stack {...columnProps}>
                    <DatePicker
                        label="Post Approval Date:"                        
                        value={props.postApprovalDate}
                        onSelectDate={(d: Date) => {
                        props.onSelectDate("postApprovalDate", d);
                        }}
                        formatDate={props.onFormatDate}
                    />
                    <DatePicker
                        label="Extension Date:"                        
                        value={props.postApprovalExtensionDate}
                        onSelectDate={(d: Date) => {
                            props.onSelectDate("postApprovalExtensionDate", d);
                        }}
                        formatDate={props.onFormatDate}
                    />
                    <TextField
                        label="Estimated Year 2 Revenue(Gross):"                        
                        value={props.Year2EstimatedGross}
                        onChange={props.onChangeText}
                        id="txt_Year2EstimatedGross"
                    />
                    <TextField
                        label="Actual Year 2 Revenue(Gross):"                        
                        value={props.Year2ActualGross}
                        onChange={props.onChangeText}
                        id="txt_Year2ActualGross"
                    />
                </Stack>
                <Stack {...columnProps}>
                    <DatePicker
                        label="First Trade Date:"                        
                        value={props.postApprovalFirstTradeDate}
                        onSelectDate={(d: Date) => {
                        props.onSelectDate("postApprovalFirstTradeDate", d);
                        }}
                        formatDate={props.onFormatDate}
                    />
                    <TextField
                        label="Estimated Year 1 Revenue(Gross):"                        
                        value={props.Year1EstimatedGross}
                        onChange={props.onChangeText}
                        id="txt_Year1EstimatedGross"
                    />
                    <TextField
                        label="Actual Year 1 Revenue(Gross):"                       
                        value={props.Year1ActualGross}
                        onChange={props.onChangeText}
                        id="txt_Year1ActualGross"
                    />
                    <TextField
                        label="Post Approval NPS Comments:"
                        multiline
                        rows={4}
                        value={props.PostApprovalNPSComments}
                        onChange={props.onChangeText}
                        id="txt_PostApprovalNPSComments"
                    />
                </Stack>
            </Stack>
            <Stack horizontal tokens={stackTokens}>
                <DefaultButton
                    text="Cancel"
                    onClick={props.cancelProposal}
                    allowDisabledFocus          
                    disabled={props.buttonClickedDisabled}
                />                
                <PrimaryButton
                    text="Submit Post Approval Details"
                    onClick={props.savePostApprovalDetails}
                    allowDisabledFocus
                    disabled={props.buttonClickedDisabled}            
                />                
            </Stack>
        </Stack>
    );
};

export default PostApproval;