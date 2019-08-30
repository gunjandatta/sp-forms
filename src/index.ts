import { Helper } from "gd-sprest-bs";
import { Configuration } from "./cfg";
import { DisplayForm, EditForm, NewForm } from "./forms";

/**
 * Global Variable
 */
window["FormsDemo"] = {
    Configuration,
    DisplayForm,
    EditForm,
    NewForm
}

// Notify SharePoint that this library is loaded
Helper.SP.SOD.notifyScriptLoadedAndExecuteWaitingJobs("forms-demo");