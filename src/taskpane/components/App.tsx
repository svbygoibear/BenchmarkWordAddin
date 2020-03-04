import * as React from "react";
import { Button, ButtonType } from "office-ui-fabric-react";
import Header from "./Header";
import Progress from "./Progress";
import * as GeneralUtils from "./../../utilities/general";

export interface AppProps {
  title: string;
  isOfficeInitialized: boolean;
}

// export interface AppState {}

export default class App extends React.Component<AppProps, null> {
  constructor(props, context) {
    super(props, context);
    // this.state = {};
  }

  public componentDidMount() {
    Office.context.document.addHandlerAsync(
      Office.EventType.BindingDataChanged,
      "bindingDataChanged",
      this.whenBindingDataChanged
    );
    Office.context.document.addHandlerAsync(
      Office.EventType.BindingSelectionChanged,
      "bindingSelectionChanged",
      this.whenBindingSelected
    );
  }

  private whenBindingDataChanged = (value: any): void => {
    console.log("whenBindingDataChanged");
    console.log(value);
  };

  private whenBindingSelected = (value: any): void => {
    console.log("whenBindingSelected");
    console.log(value);
  };

  private textBindingClick = async () => {
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

  private getSelectedTextClick = async () => {
    Office.context.document.getSelectedDataAsync(Office.CoercionType.Text, asyncResult => {
      if (asyncResult.status === Office.AsyncResultStatus.Failed) {
        console.log("Action failed. Error: " + asyncResult.error.message);
      } else {
        console.log("Selected data: " + asyncResult.value);
      }
    });
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
            Modify the source files, then click <b>Run</b>.
          </p>
          <Button
            className="ms-welcome__action"
            buttonType={ButtonType.hero}
            iconProps={{ iconName: "ChevronRight" }}
            onClick={this.textBindingClick}
          >
            Text-Binding
          </Button>
          <Button
            className="ms-welcome__action"
            buttonType={ButtonType.hero}
            iconProps={{ iconName: "ChevronRight" }}
            onClick={this.getSelectedTextClick}
          >
            Get Selected Data
          </Button>
        </main>
      </div>
    );
  }
}
