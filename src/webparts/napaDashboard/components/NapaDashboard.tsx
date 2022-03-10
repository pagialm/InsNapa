import * as React from "react";
import styles from "./NapaDashboard.module.scss";
import { INapaDashboardProps } from "./INapaDashboardProps";
import { escape } from "@microsoft/sp-lodash-subset";
// import { initializeIcons } from "@fluentui/react/lib/Icons";
import {
  CommandButton,
  DetailsHeader,
  DetailsList,
  IColumn,
  IContextualMenuProps,
  IDetailsHeaderProps,
  IGroup,
  Link,
  SelectionMode,
  // initializeIcons,
} from "office-ui-fabric-react";
import { SPHttpClient } from "@microsoft/sp-http";

const proposalObj = {
  napaListname: "NAPA Proposals",
};
const napaStages = [
  "Enquiry",
  "NPS Determination",
  "Pipeline",
  "NPS Pipeline Review",
  "Infrastructure Review",
  "Final NPS Review",
  "Approval to Trade",
  "Approved to Trade",
  "Approved and Traded",
  "Approved Expired",
  "Amendment Approved",
];
let siteUrl: string;
let UserHasAccess: boolean,
  ShowEditLink: boolean,
  userIsInCountrySpesificGroup: boolean = false;
// initializeIcons(undefined, { disableWarnings: true });
export default class NapaDashboard extends React.Component<
  INapaDashboardProps,
  IDashboardState
> {
  private _groups: IGroup[];
  private _columns: IColumn[];

  constructor(props: INapaDashboardProps, state: IDashboardState) {
    super(props);
    this.state = {
      items: [],
    };
    this._columns = [
      {
        key: "Id",
        name: "ID",
        fieldName: "ID",
        minWidth: 80,
        maxWidth: 120,
        isResizable: true,
      },
      {
        key: "Title",
        name: "Title",
        fieldName: "Title",
        minWidth: 300,
        maxWidth: 500,
        isResizable: true,
      },
      {
        key: "Country0",
        name: "Country",
        fieldName: "Country0",
        minWidth: 100,
        maxWidth: 200,
        isResizable: true,
      },
      {
        key: "BusinessArea",
        name: "Business Area",
        fieldName: "BusinessArea",
        minWidth: 100,
        maxWidth: 350,
        isResizable: true,
      },
      {
        key: "ProductArea0",
        name: "Product Area",
        fieldName: "ProductArea0",
        minWidth: 100,
        maxWidth: 250,
        isResizable: true,
      },
      {
        key: "Region",
        name: "Region",
        fieldName: "Region",
        minWidth: 50,
        maxWidth: 200,
        isResizable: true,
      },
    ];
    this._groups = [];
    siteUrl = this.props.context.pageContext.web.absoluteUrl;
  }

  private _getLitsItems(url: string): Promise<any[]> {
    return this.props.context.spHttpClient
      .get(url, SPHttpClient.configurations.v1)
      .then((response) => {
        return response.json();
      })
      .then((responsJson) => {
        return responsJson.value;
      }) as Promise<any[]>;
  }
  public componentDidMount(): void {
    const _NapaItemsUrl =
      this.props.context.pageContext.web.absoluteUrl +
      "/_api/web/lists/getbytitle('" +
      proposalObj.napaListname +
      "')/items?$select=ID,Title,Status,Country0,Region,BusinessArea,ProductArea0&$orderby=Status,ID&$top=1000";

    const _permissionsUrl =
      this.props.context.pageContext.web.absoluteUrl +
      "/_api/web/GetUserById('" +
      this.props.context.pageContext.legacyPageContext["userId"] +
      "')/Groups";
    this._getLitsItems(_permissionsUrl)
      .then((groups: any[]) => {
        const CurrentUserGroups = groups.map((a) =>
          a.Title.replace("NPS Country ", "").trim()
        );
        proposalObj["isAdminUser"] = CurrentUserGroups.some(
          (r) => r == "NPS Admins"
        );
        proposalObj["isChair"] = CurrentUserGroups.some((r) => {
          return r == "NPS Chair";
        });
        proposalObj["isOriginator"] = CurrentUserGroups.some((r) => {
          return r == "NPS Originators";
        });
        proposalObj["isReviewer"] = CurrentUserGroups.some((r) => {
          return r == "NPS Reviewers";
        });
        proposalObj["isApprover"] = CurrentUserGroups.some((r) => {
          return r == "NPS Approvers";
        });
        const checkInfraReviewers = (groupName) => {
          return groupName.indexOf("NPS Infrastructure ") != -1;
        };
        proposalObj["infraReviewers"] =
          CurrentUserGroups.filter(checkInfraReviewers);
        const ListofCountries = [
          "Botswana",
          "Ghana",
          "Kenya",
          "Mauritius",
          "Mozambique",
          "Seychelles",
          "South Africa",
          "Tanzania (ABT)",
          "Tanzania (NBC)",
          "Uganda",
          "Zambia",
        ];
        const UserCountrySpesificGroups = ListofCountries.filter((r) => {
          return CurrentUserGroups.indexOf(r) >= 0;
        });
        if (UserCountrySpesificGroups) {
          if (UserCountrySpesificGroups.length > 0) {
            console.log("isInCountrySpesificGroup True");
            console.log(UserCountrySpesificGroups);

            userIsInCountrySpesificGroup = true;
          } else {
            console.log("isInCountrySpesificGroup False");
            userIsInCountrySpesificGroup = false;
          }
        }
        proposalObj["myCountries"] = UserCountrySpesificGroups;

        console.log(proposalObj);
      })
      .then(() => {
        this._getLitsItems(_NapaItemsUrl).then((items: IListItem[]) => {
          // console.log(items);
          napaStages.forEach((napaStage, idx: number) => {
            const _stageItems: IListItem[] = items.filter((item) => {
              this.setAccess(item);
              if (UserHasAccess) return item.Status === napaStage;
            });
            const _startIndex =
              items.indexOf(_stageItems[0]) >= 0
                ? items.indexOf(_stageItems[0])
                : 0;
            const _count = _stageItems.length;
            const _isCollapsed = idx > 0 ? true : false;
            this._groups.push({
              key: napaStage,
              name: napaStage,
              startIndex: _startIndex,
              isCollapsed: _isCollapsed,
              count: _count,
              level: 0,
            });
          });
          console.log(this._groups);
          this.setState({ items: items });
        });
      });
  }
  private setAccess(item: IListItem): void {
    //First Check If the User is in any of the 5 main groups
    if (
      proposalObj["isAdminUser"] ||
      proposalObj["isChair"] ||
      proposalObj["isOriginator"] ||
      proposalObj["isReviewer"] ||
      proposalObj["isApprover"]
    ) {
      //User Can View All Items In the Dashboard
      UserHasAccess = true;
      ShowEditLink = true;

      //Filter out user access based on country group
      //Show The Chairs and all the Admins
      if (userIsInCountrySpesificGroup) {
        if (item.Country) {
          ShowEditLink =
            proposalObj["CurrentUserGroups"].indexOf(item.Country) !== -1;
          if (item.Country === "UK")
            ShowEditLink =
              proposalObj["CurrentUserGroups"].indexOf("South Africa") !== -1 &&
              item.Country === "UK";
          if (item.Country === "USA")
            ShowEditLink =
              proposalObj["CurrentUserGroups"].indexOf("South Africa") !== -1 &&
              item.Country === "USA";
          //debugger;
          //Filter Out the NBC / BBT not alowed to see each other
          const isBBT = proposalObj["CurrentUserGroups"].some((r) => {
            return r == "Tanzania (ABT)";
          });
          const isNBC = proposalObj["CurrentUserGroups"].some((r) => {
            return r == "Tanzania (NBC)";
          });
          const itemCountryIsBBT = item.Country == "Tanzania (ABT)";
          const itemCountryIsNBC = item.Country == "Tanzania (NBC)";
          if (itemCountryIsBBT && isNBC) {
            UserHasAccess = false;
          }
          if (itemCountryIsNBC && isBBT) {
            UserHasAccess = false;
          }
        } else {
          //Item has no countries show item to the user
          ShowEditLink = true;
        }
      }

      if (
        proposalObj["isAdminUser"] === true ||
        proposalObj["isChair"] === true
      ) {
        //debugger;
        if (
          proposalObj["isAdminUser"] === true &&
          proposalObj["isChair"] === false
        ) {
        }
      } else {
        if (
          proposalObj["isOriginator"] === true &&
          (proposalObj["isReviewer"] === true ||
            proposalObj["isApprover"] === true)
        ) {
          if (
            item.Status != "Enquiry" &&
            item.Status != "Pipeline" &&
            item.Status != "Infrastructure Review"
          ) {
            ShowEditLink = false;
          }
        } else {
          //Filter out by role
          if (userIsInCountrySpesificGroup == true) {
            if (proposalObj["isOriginator"] === true) {
              if (item.Status != "Enquiry" && item.Status != "Pipeline") {
                ShowEditLink = false;
              }
            }
            if (
              proposalObj["isReviewer"] === true ||
              proposalObj["isApprover"] === true
            ) {
              if (item.Status != "Infrastructure Review") {
                ShowEditLink = false;
              }
              if (item.Status == "Approved to Trade") {
                ShowEditLink = true;
              }
              if (item.Status == "Approved and Traded") {
                ShowEditLink = true;
              }
              if (item.Status == "Approval to Trade") {
                ShowEditLink = true;
              }
            }
          } else {
            ShowEditLink = false;
          }
        }
      }

      //Set item access
      item.canView = UserHasAccess;
      item.canEdit = ShowEditLink;
      if (item.Status === "Infrastructure Review") console.log(item);
    }
  }
  private _renderItemColumn(item: IListItem, index: number, column: IColumn) {
    const fieldContent = item[column.fieldName as keyof IListItem] as string;
    const napaLink =
      siteUrl + "/SitePages/NapaProposal.aspx?ProposalId=" + item.ID + "&Mode=";

    let menuProps: IContextualMenuProps = {
      items: [
        {
          key: "viewProposal",
          text: "View Proposal",
          iconProps: { iconName: "View" },
          href: napaLink + "View",
        },
      ],
      // By default, the menu will be focused when it opens. Uncomment the next line to prevent this.
      // shouldFocusOnMount: false
    };
    //this.setAccess(item);
    if (item.canEdit)
      menuProps.items.push({
        key: "editProposal",
        text: "Edit Proposal",
        iconProps: { iconName: "Edit" },
        href: napaLink + "Edit",
      });
    switch (column.key) {
      case "Id":
        return <CommandButton text={fieldContent} menuProps={menuProps} />;

      default:
        return <span style={{ verticalAlign: "center" }}>{fieldContent}</span>;
    }
  }
  public render(): React.ReactElement<INapaDashboardProps> {
    return (
      <div className={styles.napaDashboard}>
        <DetailsList
          items={this.state.items}
          columns={this._columns}
          groups={this._groups}
          selectionMode={SelectionMode.none}
          onRenderItemColumn={this._renderItemColumn}
        />
      </div>
    );
  }
}
