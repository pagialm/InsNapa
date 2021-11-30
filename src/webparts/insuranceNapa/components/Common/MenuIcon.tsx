import * as React from "react";
import { FontIcon } from "office-ui-fabric-react/lib/Icon";
import "office-ui-fabric-react/dist/css/fabric.css";
// import stylesMenu from "./MenuIcon.module.scss";
import { mergeStyles } from "office-ui-fabric-react/lib/Styling";
import { IStackStyles, Stack } from "office-ui-fabric-react/lib/Stack";
import * as styles from "./MenuIcon.module.scss";

interface IIconName {
  iconName: string;
  stageName: string;
  activated: boolean;
  approved?: boolean;
  menuClickHandler: any;
  proposalStatus: string;
  ApprovedItems?: any[];
}
interface MenuItem {
  id: number;
  title: string;
  subtile?: string;
  selected: boolean;
  type: string;
  approved?: boolean;
  enabled: boolean;
  internalName?: string;
  hidden?: boolean;
}
const _PERMENANT = "permanant";

const stages: MenuItem[] = [
  { id: 0, title: "Enquiry", selected: false, type: "menu", enabled: true },
  {
    id: 1,
    title: "Proposal",
    selected: false,
    type: "menu",
    enabled: false,
  },
  {
    id: 2,
    title: "Pipeline",
    selected: false,
    type: "menu",
    enabled: false,
  },
  {
    id: 3,
    title: "NPS Pipeline Review",
    selected: false,
    type: "menu",
    enabled: false,
  },
  {
    id: 4,
    title: "Infrastructure Review",
    selected: false,
    type: "stick",
    enabled: false,
  },
  {
    id: 4,
    title: "Infrastructure Review",
    subtile: "CRO",
    selected: false,
    type: "submenu",
    internalName: "IT",
    approved: false,
    enabled: false,
  },
  {
    id: 4,
    title: "Infrastructure Review",
    subtile: "Legal Risk",
    selected: false,
    type: "submenu",
    internalName: "Legal",
    approved: false,
    enabled: false,
  },
  {
    id: 4,
    title: "Infrastructure Review",
    subtile: "Financial Crime",
    selected: false,
    type: "submenu",
    internalName: "FinCrime",
    approved: false,
    enabled: false,
  },
  {
    id: 4,
    title: "Infrastructure Review",
    subtile: "Data Privacy",
    selected: false,
    type: "submenu",
    internalName: "CreditRisk",
    approved: false,
    enabled: false,
  },
  {
    id: 4,
    title: "Infrastructure Review",
    subtile: "Fraud Risk",
    selected: false,
    type: "submenu",
    internalName: "FraudRisk",
    approved: false,
    enabled: false,
  },
  {
    id: 4,
    title: "Infrastructure Review",
    subtile: "Tax Risk",
    selected: false,
    type: "submenu",
    internalName: "Tax",
    approved: false,
    enabled: false,
  },
  {
    id: 4,
    title: "Infrastructure Review",
    subtile: "Information Security Risk and Cyber Risk",
    selected: false,
    type: "submenu",
    internalName: "MarketRisk",
    approved: false,
    enabled: false,
  },
  {
    id: 4,
    title: "Infrastructure Review",
    subtile: "Finance",
    selected: false,
    type: "submenu",
    internalName: "Finance",
    approved: false,
    enabled: false,
  },
  {
    id: 4,
    title: "Infrastructure Review",
    subtile: "Head of Actuarial and Statutory Actuary",
    selected: false,
    type: "submenu",
    internalName: "ProductControl",
    approved: false,
    enabled: false,
  },
  {
    id: 4,
    title: "Infrastructure Review",
    subtile: "Marketing and Communications",
    selected: false,
    type: "submenu",
    internalName: "RegulatoryReporting",
    approved: false,
    enabled: false,
  },
  {
    id: 4,
    title: "Infrastructure Review",
    subtile: "Financial & Insurance Risk",
    selected: false,
    type: "submenu",
    internalName: "Treasury",
    approved: false,
    enabled: false,
  },
  {
    id: 4,
    title: "Infrastructure Review",
    subtile: "Compliance",
    selected: false,
    type: "submenu",
    internalName: "Compliance",
    approved: false,
    enabled: false,
  },
  {
    id: 4,
    title: "Infrastructure Review",
    subtile: "Operations",
    selected: false,
    type: "submenu",
    internalName: "Operations",
    approved: false,
    enabled: false,
  },
  {
    id: 4,
    title: "Infrastructure Review",
    subtile: "Supplier Risk",
    selected: false,
    type: "submenu",
    internalName: "TreasuryRisk",
    approved: false,
    enabled: false,
  },
  {
    id: 4,
    title: "Infrastructure Review",
    subtile: "Financial Reporting/ Control Risk",
    selected: false,
    type: "submenu",
    internalName: "FinTag",
    approved: false,
    enabled: false,
  },
  // {
  //   id: 4,
  //   title: "Infrastructure Review",
  //   subtile: "Risk",
  //   selected: false,
  //   type: "submenu",
  //   approved: false,
  //   enabled: false,
  // },
  {
    id: 4,
    title: "Infrastructure Review",
    subtile: "Technology Risk",
    selected: false,
    type: "submenu",
    internalName: "IRM",
    approved: false,
    enabled: false,
  },
  {
    id: 4,
    title: "Infrastructure Review",
    subtile: "Business Continuity Risk",
    selected: false,
    type: "submenu",
    internalName: "GroupReslience",
    approved: false,
    enabled: false,
  },
  {
    id: 4,
    title: "Infrastructure Review",
    subtile: "RBB CVM",
    selected: false,
    type: "submenu",
    internalName: "CRM",
    approved: false,
    enabled: false,
  },
  {
    id: 4,
    title: "Infrastructure Review",
    subtile: "Valuations",
    selected: false,
    type: "submenu",
    internalName: "ConductRisk",
    approved: false,
    enabled: false,
  },
  {
    id: 4,
    title: "Infrastructure Review",
    subtile: "Reinsurance",
    selected: false,
    type: "submenu",
    internalName: "Reinsurance",
    approved: false,
    enabled: false,
  },
  {
    id: 4,
    title: "Infrastructure Review",
    subtile: "Customer Experience",
    selected: false,
    type: "submenu",
    internalName: "CustomerExperience",
    approved: false,
    enabled: false,
  },
  {
    id: 4,
    title: "Infrastructure Review",
    subtile: "Distribution",
    selected: false,
    type: "submenu",
    internalName: "Distribution",
    approved: false,
    enabled: false,
  },
  {
    id: 5,
    title: "Final NPS Review",
    selected: false,
    type: "menu",
    enabled: false,
  },
  {
    id: 6,
    title: "Chair Review",
    selected: false,
    type: "menu",
    enabled: false,
    hidden: true,
  },
  {
    id: 61,
    title: "Approval to Trade",
    selected: false,
    type: "menu",
    enabled: false,
  },
  {
    id: 62,
    title: "Approved to Trade",
    selected: false,
    type: "menu",
    enabled: false,
    hidden: true,
  },
  {
    id: 7,
    title: "Approval Summary",
    selected: false,
    type: _PERMENANT,
    enabled: false,
  },
  {
    id: 8,
    title: "Other Status",
    selected: false,
    type: _PERMENANT,
    enabled: false,
  },
];

const menuIconCls = mergeStyles({
  // borderRadius: "25px 0 0 25px;",
  padding: "2px;",
  backgroundColor: "#dc0032",
  color: "#fff",
  display: "inline-block",
  width: "11.8rem",
  borderBottom: "1px solid #fff",
  cursor: "pointer",
  textAlign: "center",
});
const selectedStage = mergeStyles(menuIconCls, {
  backgroundColor: "#777777",
});
const approvedStage = mergeStyles(menuIconCls, {
  backgroundColor: "rgb(255,120,15)",
});
const subMenu = mergeStyles(menuIconCls, {
  backgroundColor: "rgb(245, 45, 40)",
});
const amber = mergeStyles(menuIconCls, {
  backgroundColor: "rgb(255,120,15)",
});
const stick = mergeStyles(menuIconCls, {
  backgroundColor: "rgb(150,5,40)",
});
const menuIconClsHover = mergeStyles(menuIconCls, {
  selectors: {
    ":hover": {
      backgroundColor: "rgba(220,0,50,0.5)",
    },
  },
});
const alswaysOn = mergeStyles(menuIconClsHover, {
  backgroundColor: "#b9255d",
});
const appMenu = mergeStyles({
  display: "inline-block",
  width: "12rem",
});
const nonShrinkingStackItemStyles: IStackStyles = {
  root: {
    alignItems: "center",
    display: "flex",
    justifyContent: "center",
  },
};

class MenuIcon extends React.Component<IIconName, {}> {
  public render(): React.ReactElement<{}> {
    const stageItem = stages.filter(
      (stage) => stage.title === this.props.proposalStatus
    );
    const stageId = stageItem.length > 0 ? stageItem[0].id : 0;
    const maxStage = this.props.proposalStatus ? stageId : 0;
    if (this.props.proposalStatus === "Infrastructure Review") {
      // stageItem
    }
    const _approvedItems: any[] = this.props.ApprovedItems
      ? this.props.ApprovedItems
      : [];

    if (
      this.props.ApprovedItems &&
      this.props.proposalStatus === "Infrastructure Review"
    )
      console.log("menu... approved items:", this.props.ApprovedItems);

    return (
      <div className={appMenu}>
        {stages.map((menu) => {
          if (menu.id <= maxStage || (menu.id > 0 && menu.type === _PERMENANT))
            if (!menu.hidden)
              return (
                <div
                  className={
                    this.props.stageName === menu.title ||
                    (menu.title === "Infrastructure Review" &&
                      menu.subtile === this.props.stageName)
                      ? selectedStage
                      : menu.type === "submenu" &&
                        _approvedItems.some(
                          (approvedItem) =>
                            approvedItem.NAPA_Infra === menu.internalName
                        )
                      ? approvedStage
                      : menu.type === _PERMENANT
                      ? alswaysOn
                      : menu.title === "Infrastructure Review" &&
                        menu.type !== "submenu"
                      ? stick
                      : menuIconClsHover
                  }
                  onClick={(e) => {
                    this.props.menuClickHandler(e, menu);
                  }}
                >
                  <Stack>
                    <Stack
                      horizontal
                      styles={nonShrinkingStackItemStyles}
                      className={styles["menuItem"]}
                    >
                      <Stack.Item>
                        {/* <FontIcon
                      iconName={this.props.iconName}
                      className={stylesMenu.iconClass}
                      aria-label={menu.title}
                    /> */}
                      </Stack.Item>
                      <Stack.Item>
                        {menu.type !== "submenu" ? menu.title : menu.subtile}
                      </Stack.Item>
                    </Stack>
                  </Stack>
                </div>
              );
        })}
      </div>
    );
  }
}
export default MenuIcon;
