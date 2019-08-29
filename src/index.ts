import { Configuration } from "./cfg";
import { Helper } from "gd-sprest-bs";

/**
 * Global Variable
 */
window["FormsDemo"] = {
    Configuration
}

// Notify SharePoint that this library is loaded
Helper.SP.SOD.notifyScriptLoadedAndExecuteWaitingJobs("forms-demo");