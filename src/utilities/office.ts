import * as GeneralUtils from "./general";

export async function addBindingFromSelection(): Promise<string> {
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

  return null;
}
