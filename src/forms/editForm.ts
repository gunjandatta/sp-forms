import { Components, Helper, Web } from "gd-sprest-bs";
import { ListNames } from "../cfg";
import * as DataSource from "./ds";

/**
 * Edit Form
 */
export class EditForm {
    // Constructor
    constructor(el: HTMLElement) {
        // Initialize the form
        this._btnCancel = DataSource.initForm(this.onSave);

        // Set the item id
        this._itemId = DataSource.getItemID();

        // See if this is a new form
        if (this._itemId > 0) {
            // Get the item information
            DataSource.getSession(this._itemId).then(session => {
                // Render the form
                this.render(el, session);
            });
        } else {
            // Render the form
            this.render(el);
        }
    }

    // Global Variables
    private _btnCancel: HTMLElement;
    private _ddlSession: Components.IDropdown;
    private _errorMessage: Components.IAlert;
    private _form: Components.IForm;
    private _itemId: number;
    private _sessionInfo: { [key: string]: Array<Components.IDropdownItem> } = {}

    // Define the click event for the form
    private onSave() {
        // Ensure the form is valid
        if (this._form.isValid()) {
            // Display the saving dialog
            Helper.SP.ModalDialog.showWaitScreenWithNoClose("Saving the Registration", "This will close automatically");

            // Get the values
            let values = this._form.getValues();

            // on Success event
            let onSuccess = () => {
                // Close the dialog
                Helper.SP.ModalDialog.commonModalDialogClose();

                // Show the display form
                document.location.href = "DispForm.aspx?ID=" + this._itemId;
            };

            // on Error event
            let onError = () => {
                // Close the dialog
                Helper.SP.ModalDialog.commonModalDialogClose();

                // Display the alert
                this._errorMessage.show();
            }

            // See if this is an existing item
            if (this._itemId > 0) {
                // Update the selected session
                Web().Lists(ListNames.Main).Items(this._itemId).update({
                    SessionsLUId: values["SelectedSession"].value
                }).execute(onSuccess, onError);
            } else {
                // Add the session
                Web().Lists(ListNames.Main).Items().add({
                    Title: "View Registration",
                    SessionsLUId: values["SelectedSession"].value
                }).execute(item => {
                    // Set the item id
                    this._itemId = item.Id;

                    // Call the success event
                    onSuccess();
                }, onError);
            }
        }
    }

    // Renders the form
    private render(el: HTMLElement, item?) {
        // Render a jumbotron
        Components.Jumbotron({
            el,
            title: (item ? "Update" : "Create") + " Registration",
            content: "Please select a session and time slot."
        });

        // Error message for new form
        this._errorMessage = Components.Alert({
            el,
            className: "d-none",
            type: Components.AlertTypes.Danger,
            header: "Error Registering Session",
            content: "There was an error updating your registration. Please try again."
        });

        // Generate the form
        this._form = Components.Form({
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
                            this._ddlSession.setItems(item.data);
                        },
                        onControlRendering: (props: Components.IFormControlPropsDropdown) => {
                            // Return a promise, while we load the session information
                            return new Promise((resolve, reject) => {
                                // Get the sessions
                                DataSource.getSessions().then(sessions => {
                                    // Save the session information
                                    this._sessionInfo = sessions;

                                    // Set the selected value
                                    let selectedSession = item && item["SessionsLU"];
                                    selectedSession = selectedSession ? selectedSession["Title"] : null;

                                    // Set the default item
                                    props.items = [{
                                        data: [],
                                        text: "Select a Session",
                                        value: ""
                                    }];

                                    // Parse the session information
                                    for (let sessionName in this._sessionInfo) {
                                        // Add A dropdown item for this session
                                        props.items.push({
                                            data: this._sessionInfo[sessionName],
                                            isSelected: sessionName == selectedSession,
                                            text: sessionName,
                                            value: sessionName
                                        });
                                    }

                                    // Get the existing session items
                                    let selectedSessionTimes = this._sessionInfo[selectedSession];
                                    if (selectedSessionTimes) {
                                        // Parse the items
                                        for (let i = 0; i < selectedSessionTimes.length; i++) {
                                            let session = selectedSessionTimes[i];

                                            // Set the selected flag
                                            session.isSelected = session.value == item["SessionsLUId"];
                                        }
                                    }

                                    // Update the items
                                    this._ddlSession.setItems(selectedSessionTimes);

                                    // Resolve the properties
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
                            this._ddlSession = control.get() as any;
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
                        this._btnCancel.click();
                    }
                },
                {
                    text: (item ? "Update" : "Create") + " Registration",
                    type: Components.ButtonTypes.Primary,
                    onClick: () => {
                        // Save the request
                        this.onSave();
                    }
                }
            ]
        });
    }
}