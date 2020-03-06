import * as React from "react";
import { Button, ButtonType } from "office-ui-fabric-react";
import Header from "./Header";
import Progress from "./Progress";
import * as GeneralUtils from "./../../utilities/general";
import * as OfficeUtils from "../../utilities/office";

export interface AppProps {
  title: string;
  isOfficeInitialized: boolean;
}

export default class App extends React.Component<AppProps, null> {
  constructor(props, context) {
    super(props, context);
  }

  public componentDidMount() {
    if (Office.context) {
      // Office.context.document.addHandlerAsync(
      //   Office.EventType.BindingDataChanged,
      //   "bindingDataChanged",
      //   this.whenBindingDataChanged
      // );
      // Office.context.document.addHandlerAsync(
      //   Office.EventType.BindingSelectionChanged,
      //   "bindingSelectionChanged",
      //   this.whenBindingSelected
      // );
    }
  }

  // private whenBindingDataChanged = (value: any): void => {
  //   console.log("whenBindingDataChanged");
  //   console.log(value);
  // };

  // private whenBindingSelected = (value: any): void => {
  //   console.log("whenBindingSelected");
  //   console.log(value);
  // };

  private textBindingClick = async () => {
    OfficeUtils.addBindingFromSelection();
  };

  private getSelectedTextClick = async () => {
    OfficeUtils.getTextFromSelection();
  };

  private write = (message: string): void => {
    Office.context.document.setSelectedDataAsync(message, asyncResult => {
      if (asyncResult.status === Office.AsyncResultStatus.Failed) {
        console.log(asyncResult.error.message);
      }
    });
  };

  private doCombo = async () => {
    // test
    Office.context.document.bindings.addFromSelectionAsync(
      Office.BindingType.Text,
      { id: GeneralUtils.generateUuid() },
      asyncResult => {
        if (asyncResult.status === Office.AsyncResultStatus.Failed) {
          console.log("Action failed. Error: " + asyncResult.error.message);
        } else {
          console.log("Added new binding with type: " + asyncResult.value.type + " and id: " + asyncResult.value.id);
        }
      }
    );
  };

  public render() {
    const { title, isOfficeInitialized } = this.props;

    if (!isOfficeInitialized) {
      return (
        <Progress title={title} logo="assets/logo-filled.png" message="Please sideload your addin to see app body." />
      );
    }

    return (
      <div className="ms-welcome">
        <Header logo="assets/logo-filled.png" title={this.props.title} message="Welcome" />
        <main className="ms-welcome__main">
          <p className="ms-font-l">
            To test out the features, click on any <b>Button</b>.
          </p>
          <Button className="ms-welcome__action" buttonType={ButtonType.hero} onClick={this.textBindingClick}>
            Text-Binding
          </Button>
          <br />
          <Button className="ms-welcome__action" buttonType={ButtonType.hero} onClick={this.getSelectedTextClick}>
            Get Selected Data
          </Button>
          <br />
          <Button className="ms-welcome__action" buttonType={ButtonType.hero} onClick={this.doCombo}>
            Get Selected, Bind, Answer
          </Button>
        </main>
      </div>
    );
  }
}
