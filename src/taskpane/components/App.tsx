import * as React from "react";
import { Button, ButtonType } from "office-ui-fabric-react";
import Header from "./Header";
import Progress from "./Progress";

export interface AppProps {
  title: string;
  isOfficeInitialized: boolean;
}

export interface AppState {}

export default class App extends React.Component<AppProps, AppState> {
  constructor(props, context) {
    super(props, context);
    this.state = {};
  }

  componentDidMount() {}

  click = async () => {
    // return Word.run(async context => {
    //   /**
    //    * Insert your Word code here
    //    */
    //   // insert a paragraph at the end of the document.
    //   const paragraph = context.document.body.insertParagraph("Hello World", Word.InsertLocation.end);
    //   // change the paragraph color to blue.
    //   paragraph.font.color = "blue";
    //   await context.sync();
    // });
  };

  render() {
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
            onClick={this.click}
          >
            Text-Binding
          </Button>
        </main>
      </div>
    );
  }
}
