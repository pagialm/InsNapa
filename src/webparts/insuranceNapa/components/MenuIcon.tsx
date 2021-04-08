import * as React from "react";
import { FontIcon } from "office-ui-fabric-react/lib/Icon";
import { mergeStyles } from "office-ui-fabric-react/lib/Styling";
// import * as styles from "./MenuIcon.module.scss";

const iconClass = mergeStyles({
  fontSize: 25,
  height: 25,
  width: 25,
  //   margin: "0 25px",
});
const menuIconClass = mergeStyles({
  //   border: "1px solid #dc0032",
  borderRadius: "25px 0 0 25px;",
  padding: "5px 0px 0 7px;",
  backgroundColor: "#dc0032",
  color: "#fff",
});
interface IIconName {
  iconName: string;
  stageName: string;
  activated: boolean;
}

class MenuIcon extends React.Component<IIconName, {}> {
  public render(): React.ReactElement<{}> {
    return (
      <div
        className={menuIconClass}
        style={{
          backgroundColor: this.props.activated ? "grey" : "#dc0032",
        }}
      >
        <div className="ms-Grid">
          <div className="ms-Grid-row">
            <div className="ms-Grid-col ms-sm6 ms-md4 ms-lg2">
              <FontIcon
                iconName={this.props.iconName}
                className={iconClass}
                aria-label="Enquiery"
              />
            </div>
            <div
              className="ms-Grid-col ms-sm6 ms-md8 ms-lg10"
              style={{ fontSize: 18 }}
            >
              {this.props.stageName}
            </div>
          </div>
        </div>
      </div>
    );
  }
}
export default MenuIcon;
