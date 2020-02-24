import { Helper } from "gd-sprest-bs";
import { Configuration } from "./cfg";
import * as Forms from "./forms";

/**
 * Global Variable
 */
window["FormsDemo"] = {
    Configuration,
    DisplayForm: Forms.DisplayForm,
    EditForm: Forms.EditForm,
    NewForm: Forms.EditForm
}

// Notify SharePoint that this library is loaded
Helper.SP.SOD.notifyScriptLoadedAndExecuteWaitingJobs("forms-demo");