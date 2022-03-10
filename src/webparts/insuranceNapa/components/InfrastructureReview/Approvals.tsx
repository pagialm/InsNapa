import {
  PeoplePicker,
  PrincipalType,
} from "@pnp/spfx-controls-react/lib/PeoplePicker";
import {
  DefaultButton,
  IStackProps,
  IStackStyles,
  PrimaryButton,
  Separator,
  Stack,
  TextField,
} from "office-ui-fabric-react";
import {
  SPHttpClient,
  SPHttpClientResponse,
  ISPHttpClientOptions,
} from "@microsoft/sp-http";
import * as React from "react";
import { IApprovals } from "./IApprovals";

const stackTokens = { childrenGap: 50 };
const stackStyles: Partial<IStackStyles> = { root: { width: 784 } };
const columnProps: Partial<IStackProps> = {
  tokens: { childrenGap: 10 },
  styles: { root: { width: 450 } },
};

const Approvals = (props) => {
  const [approvalItem, setApprovalItem] = React.useState([]);
  const [currentApprovalItem, setCurrentApprovalItem] = React.useState({});
  const [isButtonEnabled, setIsButtonEnabled] = React.useState(false);
  const [isStageApproved, setIsStageApproved] = React.useState(false);
  const subMenu = props.submenu;
  let currApproval: IApprovals[] = [];

  const SetStageApproval = (proposalApprovals?: any[]) => {
    // debugger;
    console.log("can approve:",props.canApprove);
    console.log("review completed:",props.ReviewCompleted);
    setIsButtonEnabled(false);
    const allApprovals = proposalApprovals ? proposalApprovals : approvalItem;
    currApproval = allApprovals.filter((rq) => {
      return rq.Title === `${props.Proposal_ID}_${subMenu["internalName"]}`;
    });
    if (currApproval.length > 0) {
      let userarr = [];
      if (currApproval[0].ApprovalInfrastructureColleagues.results)
        currApproval[0].ApprovalInfrastructureColleagues.results.forEach(
          (user) => {
            userarr.push(user.Title);
          }
        );
      const infraColleagues = { ApprovalInfrastructureColleagues: userarr };
      setCurrentApprovalItem({
        ...currApproval[0],
        ...infraColleagues,
      });
      setIsStageApproved(true);
      props.SetIsStageApproved(true);
    } 
    else {
      setCurrentApprovalItem({
        Title: `${props.Proposal_ID}_${subMenu["internalName"]}`,
        Proposal_ID: props.Proposal_ID,
        ApprovalComments: "",
        ApprovalInfrastructureColleagues: [],
      });
      setIsStageApproved(false);
      props.SetIsStageApproved(false);
    }
  };
  const CollectApprovals = () => {
    fetch(
      props.siteUrl +
        "/_api/lists/getbytitle('NAPA Infrastructure Approvals')/items?$filter=Proposal_ID eq " +
        props.Proposal_ID +
        "&$select=Title,ApprovalInfrastructureColleagues/Title,ApprovalComments,NAPA_Infra,\
                Proposal_ID,ID&$expand=ApprovalInfrastructureColleagues/Title",
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
        const approveItem: IApprovals[] = responseJson.d.results;
        if (approveItem.length > 0) {
          setApprovalItem([...approveItem]);
        }
        SetStageApproval(approveItem);
      });
  };

  //   Fetch Review questions once off when the control loads for the first time.
  React.useEffect(() => {
    CollectApprovals();
  }, []);

  //   Filter the reuslts of the review questions per page based on which infrastructure area is selected
  React.useEffect(() => {
    SetStageApproval();
    
  }, [subMenu["internalName"]]); // The menu object has internalName which is used to filter review questions based on the infrastructure area selected

  const submitApproval = (e) => {
    console.log(e);
    setIsButtonEnabled(!isButtonEnabled);
    debugger;
    const isValid = props.ValidateForm(".nps_approvals");
    const isNewItem: boolean = currentApprovalItem["ID"] ? false : true;
    const apiReport = (data: SPHttpClientResponse) => {
      console.log(data);
      setIsButtonEnabled(!isButtonEnabled);

      //TODO: Update Approval Field from main item
      props.SetIsStageApproved(true);

      //TODO: Check if all approvals completed so that the proposal can be moved to Final NPS Review
      props.CheckApprovals();
    };
    if(isValid)
      props.SubmitToSP(
        "NAPA Infrastructure Approvals",
        isNewItem,
        currentApprovalItem,
        apiReport
      );
    else
      setIsButtonEnabled(false);
  };

  const removeApproval = (e) => {
    // debugger;
    //Variables
    let approvalsCount: number = props.ApprovalsCount;
    const proposalObj = {
      ID: props.Proposal_ID,
    };
    const approvalSPItem = {
      ID: currentApprovalItem["ID"],
    };
    const SusccessfulUpdate = () => {
      console.info(`InfrastructureApprovalCount for ${props.submenu.subtile}`);
    };
    const SusccessfulRemovalUpdate = () => {
      console.info(
        `Approval ID: ${currentApprovalItem["ID"]} has been deleted for ${props.submenu.subtile}`
      );
      location.href = props.context.pageContext.site.absoluteUrl;
    };

    //Update the Infrastructure Count
    approvalsCount--;
    proposalObj["InfrastructureApprovalCount"] = approvalsCount;
    props.SubmitToSP("NAPA Proposals", false, proposalObj, SusccessfulUpdate);

    //TODO: Delete Approval Item
    props.DeleteFromSP(
      "NAPA Infrastructure Approvals",
      approvalSPItem,
      SusccessfulRemovalUpdate
    );

    props.SetIsStageApproved(false);
  };

  const onChangeText = (
    ev: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>,
    newValue: string
  ): void => {
    const element = ev.target as HTMLElement;
    debugger;
    const el = {};
    el[element.id.split("_")[1]] = newValue;
    setCurrentApprovalItem({ ...currentApprovalItem, ...el });
  };

  const getPeoplePickerItems = (items: any[]) => {
    let userarr = [];
    items.forEach((user) => {
      userarr.push(user.id);
    });
    const infraColleagues = { ApprovalInfrastructureColleaguesId: userarr };
    setCurrentApprovalItem({
      ...currentApprovalItem,
      ...infraColleagues,
    });
  };  

  return (
    <Stack className="nps_approvals">
      <Stack horizontal tokens={stackTokens} styles={stackStyles}>
        <Stack {...columnProps}>
          <TextField
            label="Approval Comments?"
            multiline
            rows={3}
            value={currentApprovalItem["ApprovalComments"]}
            id="txt_ApprovalComments"
            required={props.canApprove && props.ReviewCompleted ? true : false}
            onChange={onChangeText}
          />
        </Stack>
        <Stack {...columnProps}>
          <PeoplePicker
            context={props.context}
            titleText="Infrastructure colleague(s) consulted prior to approval? :"
            personSelectionLimit={3}
            groupName={""} // Leave this blank in case you want to filter from all users
            showtooltip={true}
            disabled={false}
            ensureUser={true}
            defaultSelectedUsers={
              currentApprovalItem["ApprovalInfrastructureColleagues"]
            }
            onChange={getPeoplePickerItems}
            showHiddenInUI={false}
            principalTypes={[PrincipalType.User]}
            resolveDelay={1000}
          />
        </Stack>
      </Stack>
      <Stack>
        <Separator />
        <Stack horizontal tokens={stackTokens} styles={stackStyles}>
          <DefaultButton
            onClick={() => {
              location.href = props.siteUrl;
            }}
            text="Cancel"
          />
          {(props.Status === "Infrastructure Review" && !isStageApproved && props.canApprove && props.ReviewCompleted) && (
            <PrimaryButton
              onClick={submitApproval}
              text="Submit Approval"
              disabled={isButtonEnabled}
            />
          )}
          {(props.Status === "Infrastructure Review" && isStageApproved && props.canApprove) && (
            <PrimaryButton
              onClick={removeApproval}
              text="Remove Approval"
              disabled={isButtonEnabled}
            />
          )}
        </Stack>
      </Stack>
    </Stack>
  );
};

export default Approvals;
