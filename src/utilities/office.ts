import * as GeneralUtils from "./general";

export async function addBindingFromSelection(successAction: () => Promise<void>): Promise<string> {
  let textValue = null;
  Office.context.document.bindings.addFromSelectionAsync(
    Office.BindingType.Text,
    { id: GeneralUtils.generateUuid() },
    asyncResult => {
      if (asyncResult.status === Office.AsyncResultStatus.Failed) {
        // console.log("Action failed. Error: " + asyncResult.error.message);
        throw new Error(asyncResult.error.message);
      } else {
        // console.log("Added new binding with type: " + asyncResult.value.type + " and id: " + asyncResult.value.id);
        textValue = successAction();
      }
    }
  );
  return textValue;
}

export async function getTextFromSelection(): Promise<string> {
  let selectedText = null;
  Office.context.document.getSelectedDataAsync(Office.CoercionType.Text, asyncResult => {
    if (asyncResult.status === Office.AsyncResultStatus.Failed) {
      //   console.log("Action failed. Error: " + asyncResult.error.message);
      throw new Error(asyncResult.error.message);
    } else {
      selectedText = asyncResult.value as string;
      // console.log("Selected data: " + selectedText);
      // this.write(`${selectedText.trimRight()}  I am an intelligent answer to a question`);
    }
  });
  return selectedText;
}
