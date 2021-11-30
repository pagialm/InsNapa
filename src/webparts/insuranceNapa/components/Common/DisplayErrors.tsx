import {
  Callout,
  FontWeights,
  getTheme,
  mergeStyleSets,
  Text,
} from "office-ui-fabric-react";
import * as React from "react";

const theme = getTheme();
const styles = mergeStyleSets({
  buttonArea: {
    verticalAlign: "top",
    display: "inline-block",
    textAlign: "center",
    margin: "0 100px",
    minWidth: 130,
    height: 32,
  },
  callout: {
    maxWidth: 300,
  },
  header: {
    padding: "18px 24px 12px",
  },
  title: [
    theme.fonts.xLarge,
    {
      margin: 0,
      fontWeight: FontWeights.semilight,
    },
  ],
  inner: {
    height: "100%",
    padding: "0 24px 20px",
  },
  actions: {
    position: "relative",
    marginTop: 20,
    width: "100%",
    whiteSpace: "nowrap",
  },
  subtext: [
    theme.fonts.small,
    {
      margin: 0,
      fontWeight: FontWeights.semilight,
    },
  ],
  link: [
    theme.fonts.medium,
    {
      color: theme.palette.neutralPrimary,
    },
  ],
});

const DisplayErrors = (props) => {
  console.log("zzzz", props);
  return (
    <Callout
      className={styles.callout}
      role="alertdialog"
      gapSpace={0}
      target={props.Target}
      //   onDismiss={toggleIsCalloutVisible}
      setInitialFocus
    >
      <div className={styles.header}>
        <Text className={styles.title}>The following issues were detected</Text>
      </div>
      <div className={styles.inner}>
        <ul>
          {props.ErrorMessages.map((errorMsg) => {
            return <li>{errorMsg}</li>;
          })}
        </ul>
      </div>
    </Callout>
  );
};

export default DisplayErrors;
