import { mergeStyles } from "office-ui-fabric-react/lib/Styling";
import * as React from "react";
interface IHeaderInfoProps {
  title: string;
  description?: string;
}
const headerCover = mergeStyles({
  backgroundColor: "#f2f2f2",
  marginTop: "1rem",
});
const h2Style = mergeStyles({
  color: "#aa052d",
  fontWeight: 700,
  fontSize: "13.5px",
  padding: "5px 0 0 5px",
});
const inputDesc = mergeStyles({
  fontSize: "0.70rem",
  padding: "0 5px",
});
const lineBreak = mergeStyles({
  border: "#aa052d 2px solid",
});
class HeaderInfo extends React.Component<IHeaderInfoProps, {}> {
  public render(): React.ReactElement<IHeaderInfoProps> {
    return (
      <div className={headerCover}>
        <h2 className={h2Style}>{this.props.title}</h2>
        <p className={inputDesc}>{this.props.description}</p>
        <hr className={lineBreak} />
      </div>
    );
  }
}
export default HeaderInfo;
