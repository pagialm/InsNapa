import {
  DefaultButton,
  Dropdown,
  IDropdownOption,
  IStackProps,
  IStackStyles,
  PrimaryButton,
  Separator,
  Stack,
  TextField,
} from "office-ui-fabric-react";
import * as React from "react";
import { IReviewQuestions } from "./IReviewQuestions";
import {
  SPHttpClient,
  SPHttpClientResponse,
  ISPHttpClientOptions,
} from "@microsoft/sp-http";
import {
  PeoplePicker,
  PrincipalType,
} from "@pnp/spfx-controls-react/lib/PeoplePicker";
import DisplayErrors from "../Common/DisplayErrors";

const stackTokens = { childrenGap: 50 };
const stackStyles: Partial<IStackStyles> = { root: { width: 784 } };
const columnProps: Partial<IStackProps> = {
  tokens: { childrenGap: 10 },
  styles: { root: { width: 450 } },
};
const ddOptions: IDropdownOption[] = [
  { key: "0", text: "" },
  { key: "Yes", text: "Yes" },
  { key: "No", text: "No" },
];

const ReviewQuestions = (props: IReviewQuestions) => {
  let currReviewQuestion: IReviewQuestions[] = [];
  const [reviewQuestion, setReviewQuestion] = React.useState([]);
  const [currentReviewQuestions, setCurrentReviewQuestions] = React.useState(
    {}
  );
  const [isButtonEnabled, setIsButtonEnabled] = React.useState(false);  
  const subMenu = props.submenu;
  const onChange = (
    ev: React.FormEvent<HTMLDivElement>,
    option?: IDropdownOption,
    index?: number
  ): void => {
    const selectedElement = ev.target as HTMLDivElement;
    const elementId = selectedElement.id.split("_")[1].split("-")[0];
    const el = {};
    //Todo: Find old element, check its type... array or primitive and act accodingly
    const oldState = currentReviewQuestions[elementId]; //el[elementId];
    if (
      Array.isArray(oldState) ||
      (selectedElement["type"] && selectedElement["type"] === "checkbox")
    ) {      
      const newArray = oldState ? [...oldState] : [];
      if (
        oldState &&
        oldState.some((opt) => {
          return opt === option.key;
        })
      ) {
        const finalArray = newArray.filter((itemToFilter) => {
          return itemToFilter !== option.key;
        });
        el[elementId] = finalArray;
      } else {
        newArray.push(option.key);
        el[elementId] = newArray;
      }
    } else {
      el[elementId] = elementId === "Region" ? [option.key] : option.key;
    }
    setCurrentReviewQuestions({ ...currentReviewQuestions, ...el });
  };

  const onChangeText = (
    ev: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>,
    newValue: string
  ): void => {
    const element = ev.target as HTMLElement;    
    const el = {};
    el[element.id.split("_")[1]] = newValue;
    setCurrentReviewQuestions({ ...currentReviewQuestions, ...el });
  };

  const SetInfrastructureReviewQuestion = (ProposalQuestions?: any[]) => {
    console.log("submenu...", subMenu);
    debugger;
    const reviewQuestions = ProposalQuestions
      ? ProposalQuestions
      : reviewQuestion;
    setIsButtonEnabled(false);
    currReviewQuestion = reviewQuestions.filter((rq) => {
      return rq.NAPA_Link === `${props.Proposal_ID}_${subMenu["internalName"]}`;
    });
    if (currReviewQuestion.length > 0) {
      let userarr = [];
      if(currReviewQuestion[0].ReviewInfrastructureColleaguesId){
          currReviewQuestion[0].ReviewInfrastructureColleagues.results.forEach(
            (user) => {
              userarr.push(user.Title);
            }
          );        
      }    
      const infraColleagues = { ReviewInfrastructureColleagues: userarr };
      setCurrentReviewQuestions({
        ...currReviewQuestion[0],
        ...infraColleagues,
      });
      //Set review completed.
      props.SetReviewCompleted(true);
    } 
    else{
      setCurrentReviewQuestions({
        HeadcountConsidered: "",
        ITDevRequired: "",
        MemoireConsidered: "",
        NAPA_Link: `${props.Proposal_ID}_${subMenu["internalName"]}`,
        OpRiskRequired: "",
        Proposal_ID: props.Proposal_ID,
        RDARRRelevance: "",
        ReviewComments: "",
        ReviewInfrastructureColleagues: [],
        RiskAssessmentCompleted: "",
        WorkaroundsRequired: "",
        RDARRImpact: "",
      });
      props.SetReviewCompleted(false);
    }
  };

  const CollectInfrastructureReviewQuestions = () => {
    fetch(
      props.siteUrl +
        "/_api/lists/getbytitle('NAPA Infrastructure Questions')/items?$filter=Proposal_ID eq " +
        props.Proposal_ID +
        "&$select=MemoireConsidered,ReviewInfrastructureColleagues/Title,ReviewInfrastructureColleaguesId,RiskAssessmentCompleted,ReviewComments,\
              ITDevRequired,HeadcountConsidered,OpRiskRequired,WorkaroundsRequired,Proposal_ID,RDARRRelevance\
              ,NAPA_Link,RDARRImpact,ID&$expand=ReviewInfrastructureColleagues/Title",
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
        const reviewQuestions: IReviewQuestions[] = responseJson.d.results;
        if (reviewQuestions.length > 0) {
          setReviewQuestion([...reviewQuestions]);
        }
        SetInfrastructureReviewQuestion(reviewQuestions);
      });
  };
  //   Fetch Review questions once off when the control loads for the first time.
  React.useEffect(() => {
    CollectInfrastructureReviewQuestions();
  }, []);

  //   Filter the reuslts of the review questions per page based on which infrastructure area is selected
  React.useEffect(() => {
    SetInfrastructureReviewQuestion();
  }, [subMenu["internalName"]]); // The menu object has internalName which is used to filter review questions based on the infrastructure area selected

  const submitReviewQuestions = (e) => {
    setIsButtonEnabled(!isButtonEnabled);
    debugger;
    const isValid = props.ValidateForm();
    if(isValid){
      console.log("Form Valid");
    }
    else{
      setIsButtonEnabled(false);
      return false;
    }
      
    const isNewItem: boolean = currentReviewQuestions["ID"] ? false : true;
    const apiReport = (data: SPHttpClientResponse) => {
      console.log(data);
      setIsButtonEnabled(!isButtonEnabled);
      location.href = props.context.pageContext.site.absoluteUrl;
    };
    const _InfraObject = {
      HeadcountConsidered: currentReviewQuestions["HeadcountConsidered"],
      ITDevRequired: currentReviewQuestions["ITDevRequired"],
      MemoireConsidered: currentReviewQuestions["MemoireConsidered"],
      NAPA_Link: currentReviewQuestions["NAPA_Link"],
      OpRiskRequired: currentReviewQuestions["OpRiskRequired"],
      Proposal_ID: currentReviewQuestions["Proposal_ID"],
      RDARRImpact: currentReviewQuestions["RDARRImpact"],
      RDARRRelevance: currentReviewQuestions["RDARRRelevance"],
      ReviewComments: currentReviewQuestions["ReviewComments"],
      RiskAssessmentCompleted:
        currentReviewQuestions["RiskAssessmentCompleted"],
      WorkaroundsRequired: currentReviewQuestions["WorkaroundsRequired"],
    };
    if (
      !Array.isArray(currentReviewQuestions["ReviewInfrastructureColleaguesId"])
    ) {
      let userarr = [];
      if(currentReviewQuestions["ReviewInfrastructureColleaguesId"]){
        currentReviewQuestions["ReviewInfrastructureColleaguesId"][
          "results"
        ].forEach((user) => {
          userarr.push(user);
        });
      }
      _InfraObject["ReviewInfrastructureColleaguesId"] = userarr;
    } else {
      _InfraObject["ReviewInfrastructureColleaguesId"] =
        currentReviewQuestions["ReviewInfrastructureColleaguesId"];
    }
    if (!isNewItem) {
      _InfraObject["ID"] = currentReviewQuestions["ID"];
    }
    props.SubmitToSP(
      "NAPA Infrastructure Questions",
      isNewItem,
      _InfraObject,
      apiReport
    );
  };

  const getPeoplePickerItems = (items: any[]) => {
    let userarr = [];
    items.forEach((user) => {
      userarr.push(user.id);
    });
    const infraColleagues = { ReviewInfrastructureColleaguesId: userarr };
    setCurrentReviewQuestions({
      ...currentReviewQuestions,
      ...infraColleagues,
    });
  };

  return (
    <Stack>
      <Stack horizontal tokens={stackTokens} styles={stackStyles}>
        <Stack {...columnProps}>
          <Dropdown
            label="Does this proposal have any impact on the existing RDARR , (Risk Data Aggregation and Risk Reporting) artifacts, processes or levels of compliance?"
            options={ddOptions}
            onChange={onChange}
            id="ddl_RDARRImpact"
            required
            selectedKey={currentReviewQuestions["RDARRImpact"]}
          />

          <Dropdown
            label="Have you considered all questions in your assessment and retained appropriate evidence to demonstrate consideration against these?"
            options={ddOptions}
            onChange={onChange}
            id="ddl_MemoireConsidered"
            required
            selectedKey={currentReviewQuestions["MemoireConsidered"]}
          />

          <Dropdown
            label="Is incremental Headcount required?"
            options={ddOptions}
            onChange={onChange}
            id="ddl_HeadcountConsidered"
            required
            selectedKey={currentReviewQuestions["HeadcountConsidered"]}
          />

          <Dropdown
            label="Is internal IT development required?"
            options={ddOptions}
            onChange={onChange}
            id="ddl_ITDevRequired"
            required
            selectedKey={currentReviewQuestions["ITDevRequired"]}
          />

          <TextField
            label="Review Comments?"
            multiline
            rows={3}
            value={currentReviewQuestions["ReviewComments"]}
            id="txt_ReviewComments"
            required
            onChange={onChangeText}
          />
        </Stack>
        <Stack {...columnProps}>
          <Dropdown
            label="If yes, have the necessary changes / amendments been made to the relevant processes, documentation, reconciliations, controls or other RDARR relevant artifacts?"
            onChange={onChange}
            options={ddOptions}
            id="ddl_RDARRRelevance"
            required
            selectedKey={currentReviewQuestions["RDARRRelevance"]}
          />
          <Dropdown
            label="Have you attached the completed risk assessment template?"
            options={ddOptions}
            onChange={onChange}
            id="ddl_RiskAssessmentCompleted"
            required
            selectedKey={currentReviewQuestions["RiskAssessmentCompleted"]}
          />
          <Dropdown
            label="Are manual Workarounds required?"
            options={ddOptions}
            onChange={onChange}
            id="ddl_WorkaroundsRequired"
            required
            selectedKey={currentReviewQuestions["WorkaroundsRequired"]}
          />
          <Dropdown
            label="Are there any other Operational Risks relevant to your Infrastructure Area?"
            options={ddOptions}
            onChange={onChange}
            id="ddl_OpRiskRequired"
            required
            selectedKey={currentReviewQuestions["OpRiskRequired"]}
          />

          <PeoplePicker
            context={props.context}
            titleText="Have you obtained sufficient input from infrastructure area representative for all locations in scope?"
            personSelectionLimit={3}
            groupName={""} // Leave this blank in case you want to filter from all users
            showtooltip={true}
            disabled={false}
            ensureUser={true}
            defaultSelectedUsers={
              currentReviewQuestions["ReviewInfrastructureColleagues"]
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
          {props.Status === "Infrastructure Review" &&
            (!props.IsStageApproved && props.canReview) && (
              <PrimaryButton
                onClick={submitReviewQuestions}
                text={props.ReviewCompleted? "Update Review" : "Submit Review"}
                disabled={isButtonEnabled}
              />
            )}

            {props.ErrorMessages.length > 0 && (
              <Stack>
                <p id="ErrorsDisplay"></p>
                <DisplayErrors
                  clearErrors={props.ClearErrors}
                  ErrorMessages={props.ErrorMessages}
                  Target={"#ErrorsDisplay"}
                />
              </Stack>
            )}
        </Stack>
      </Stack>
    </Stack>
  );
};
export default ReviewQuestions;
