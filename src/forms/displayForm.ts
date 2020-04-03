import { Components, Helper } from "gd-sprest-bs";
import * as DataSource from "./ds"

/**
 * Display Form
 */
export class DisplayForm {
    // Constructor
    constructor(el: HTMLElement) {
        // Hide the default form
        DataSource.hideDefaultForm();

        // Display a loading message
        Helper.SP.ModalDialog.showWaitScreenWithNoClose("Loading the Registration").then(dlg => {
            // Get the session information
            DataSource.getSession(DataSource.getItemID()).then(session => {
                // Render the form
                this.render(el, session);

                // Close the dialog
                dlg.close();
            });
        });
    }

    // Renders the form
    private render(el: HTMLElement, item) {
        // Render a jumbotron
        Components.Jumbotron({
            el,
            title: "Session Information"
        });

        // Set the session information
        let sessionInfo = item["SessionsLU"] || {};

        // Render the form
        Components.Form({
            el,
            rows: [
                {
                    columns: [
                        {
                            control: {
                                label: "Registered Session:",
                                type: Components.FormControlTypes.TextField,
                                isReadonly: true,
                                value: sessionInfo["Title"]
                            }
                        },
                        {
                            control: {
                                label: "Session Information:",
                                type: Components.FormControlTypes.TextField,
                                isReadonly: true,
                                value: sessionInfo["SessionInfo"]
                            }
                        }
                    ]
                }
            ]
        });

        // Render the form buttons
        Components.ButtonGroup({
            el,
            buttons: [
                {
                    className: "mr-2",
                    text: "View Registrations",
                    type: Components.ButtonTypes.Secondary,
                    onClick: () => {
                        // Redirect to the display form
                        document.location.href = document.location.origin + document.location.pathname.replace("/DispForm.aspx", "");
                    }
                },
                {
                    text: "Edit Registration",
                    type: Components.ButtonTypes.Primary,
                    onClick: () => {
                        // Redirect to the edit form
                        document.location.href = "EditForm.aspx?ID=" + item.Id;
                    }
                }
            ]
        });
    }
}