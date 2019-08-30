import { Components, Helper, List } from "gd-sprest-bs";
import { ISessionInfo } from "./index.d";
import * as Common from "./common";

/**
 * New Form
 */
export const NewForm = (el: HTMLElement) => {
    let _ddlSession: Components.IDropdown;
    let _btnCancel: HTMLElement;
    let _sessionInfo: { [key: string]: Array<ISessionInfo> } = {}

    // Define the click event for the form
    let onSave = () => {
        // Ensure the form is valid
        if (form.isValid()) {
            // Display the saving dialog
            Helper.SP.ModalDialog.showWaitScreenWithNoClose("Saving the Registration", "This will close automatically");

            // Get the values
            let values = form.getValues();

            // Add a new item
            List("Custom Form Demo").Items().add({
                Title: "View Registration",
                SessionsLUId: values["SelectedSession"].value
            }).execute(
                // Success
                item => {
                    // Close the dialog
                    Helper.SP.ModalDialog.commonModalDialogClose();

                    // Show the display form
                    document.location.href = "DispForm.aspx?ID=" + item.Id;
                },
                // Error
                () => {
                    // Close the dialog
                    Helper.SP.ModalDialog.commonModalDialogClose();

                    // Display the alert
                    errorMessage.el.classList.remove("d-none");
                }
            );
        }
    }

    // Initialize the form
    _btnCancel = Common.initForm(onSave);

    // Render a jumbotron
    Components.Jumbotron({
        el,
        title: "New Registration",
        content: "Please select a session and time slot."
    });

    // Error message for new form
    let errorMessage = Components.Alert({
        el,
        className: "d-none",
        type: Components.AlertTypes.Danger,
        header: "Error Registering Session",
        content: "There was an error saving your registration. Please try again."
    });

    // Generate the form
    let form = Components.Form({
        el,
        rows: [
            {
                control: {
                    name: "Session",
                    label: "Select a Session",
                    type: Components.FormControlTypes.Dropdown,
                    required: true,
                    loadingMessage: "Loading the Session Information",
                    onValidate: (control, item: Components.IDropdownItem) => {
                        return {
                            isValid: item && item.value.length > 0 ? true : false,
                            invalidMessage: "Please select a session."
                        };
                    },
                    onChange: (item: Components.IDropdownItem) => {
                        // Update the sessions
                        Common.updateSessions(_ddlSession, item ? item.data : null);
                    },
                    onControlRendering: (props: Components.IFormControlPropsDropdown) => {
                        // Return a promise, while we load the session information
                        return new Promise((resolve, reject) => {
                            // Query the session list
                            List("Sessions").Items().query({
                                OrderBy: ["Title", "SessionDay", "SessionTime"],
                                Select: ["Id", "Title", "SessionDay", "SessionTime"]
                            }).execute(items => {
                                // Set the default item
                                props.items = [{
                                    text: "Select a Session",
                                    value: ""
                                }];

                                // Parse the items
                                for (let i = 0; i < items.results.length; i++) {
                                    let item = items.results[i];

                                    // Get the session name
                                    let session = item["Title"];

                                    // Ensure the key exists
                                    if (_sessionInfo[session] == null) {
                                        // Create an array for this session
                                        _sessionInfo[session] = [];

                                        // Create a dropdown option for this session
                                        props.items.push({
                                            data: _sessionInfo[session],
                                            text: session,
                                            value: session
                                        });
                                    }

                                    // Append the day and time
                                    _sessionInfo[session].push({
                                        itemId: item.Id,
                                        day: item["SessionDay"],
                                        time: item["SessionTime"]
                                    });
                                }

                                // Resolve the promise
                                resolve(props);
                            });
                        });
                    }
                } as Components.IFormControlPropsDropdown
            },
            {
                control: {
                    name: "SelectedSession",
                    label: "Available Sessions",
                    type: Components.FormControlTypes.Dropdown,
                    onValidate: (control, item: Components.IDropdownItem) => {
                        // Ensure a value exists
                        return {
                            isValid: item != null,
                            invalidMessage: "Please select an available slot."
                        };
                    },
                    onControlRendered: control => {
                        // Save a reference to this control
                        _ddlSession = control.get() as any;
                    }
                } as Components.IFormControlPropsDropdown
            }
        ]
    });

    // Render the form buttons
    Components.ButtonGroup({
        el,
        buttons: [
            {
                className: "mr-2",
                text: "Cancel",
                type: Components.ButtonTypes.Danger,
                onClick: () => {
                    // Cancel the request
                    _btnCancel.click();
                }
            },
            {
                text: "Create Registration",
                type: Components.ButtonTypes.Primary,
                onClick: onSave
            }
        ]
    });
}