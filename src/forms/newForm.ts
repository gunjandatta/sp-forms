import { Components, Helper, List } from "gd-sprest-bs";

/**
 * New Form
 */
export const NewForm = (el: HTMLElement) => {
    let _ddlSession: Components.IDropdown;
    let _btnCancel: HTMLElement;
    let _sessionInfo: {
        [key: string]: Array<{
            itemId: number;
            day: string;
            time: string;
        }>
    } = {}

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

    // Get the list form webpart element
    let elWebpart = document.querySelector(".ms-webpart-zone > div[id^='MSOZoneCell']") as HTMLElement;
    if (elWebpart) {
        // Hide the default list form
        elWebpart.style.display = "none";

        /*********************************************************************************
         * The ribbon "Save" button click event is linked to the save buttons.
         * We will need to update the click events of the default save buttons to the
         * new custom one.
         ********************************************************************************/

        // Get the "Save" buttons
        let elButtons = elWebpart.querySelectorAll("input[value='Save']");
        for (let i = 0; i < elButtons.length; i++) {
            // Remove the default click event
            (elButtons[i] as HTMLInputElement).removeAttribute("onclick");

            // Set the click event to the custom save method
            (elButtons[i] as HTMLInputElement).addEventListener("click", onSave);
        }

        // Get the "Cancel" button
        _btnCancel = elWebpart.querySelector("input[value='Cancel']");
    }

    // Render a jumbotron
    Components.Jumbotron({
        el,
        title: "Session Registration Form",
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
                    onControlRendering: (props: Components.IFormControlPropsDropdown) => {
                        // Return a promise, while we load the session information
                        return new Promise((resolve, reject) => {
                            // Query the session list
                            List("Sessions").Items().query({
                                OrderBy: ["Title", "SessionDay", "SessionTime"],
                                Select: ["Title", "SessionDay", "SessionTime"]
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
                    },
                    onChange: (item: Components.IDropdownItem) => {
                        let items: Array<Components.IDropdownItem> = [];

                        // Ensure an item was selected
                        if (item != null && item.data != null) {
                            // Parse the data
                            for (let i = 0; i < item.data.length; i++) {
                                let sessionInfo = item.data[i];

                                // Add a dropdown item for this session
                                items.push({
                                    text: sessionInfo.day + " - " + sessionInfo.time,
                                    value: sessionInfo.itemId
                                });
                            }
                        }

                        // Update the dropdown
                        _ddlSession.setItems(items);
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
                text: "Register",
                type: Components.ButtonTypes.Primary,
                onClick: onSave
            }
        ]
    });
}