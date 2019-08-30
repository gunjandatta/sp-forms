import { Components, Helper, List } from "gd-sprest-bs";
import * as Common from "./common"

/**
 * Display Form
 */
export const DisplayForm = (el: HTMLElement) => {
    // Display a loading message
    Helper.SP.ModalDialog.showWaitScreenWithNoClose("Loading the Registration").then(dlg => {
        // Get the list form webpart element
        let elWebpart = document.querySelector(".ms-webpart-zone > div[id^='MSOZoneCell']") as HTMLElement;
        if (elWebpart) {
            // Hide the default list form
            elWebpart.style.display = "none";
        }

        // Render a jumbotron
        Components.Jumbotron({
            el,
            title: "Session Registration Form"
        });

        // Get the list item information
        // We are expanding the lookup field to include the "calculated" information field
        // You cannot expand "choice" fields, which is why a "calculated" field was created
        List("Custom Form Demo").Items(Common.getItemID()).query({
            Expand: ["SessionsLU"],
            Select: ["Title", "SessionsLU/Title", "SessionsLU/SessionInfo"]
        }).execute(item => {
            // Get the session day/time information
            let day, time = null;
            let luItem = item["SessionsLU"] || null;
            if (luItem) {
                let sessionInfo = (luItem["SessionInfo"] || "").split(" - ");
                day = sessionInfo[0];
                time = sessionInfo[1];
            }

            // Render the form
            Components.Form({
                el,
                value: {
                    Session: luItem ? luItem["Title"] : "",
                    SessionDay: day || "",
                    SessionTime: time || ""
                },
                rows: [
                    {
                        control: {
                            name: "Session",
                            label: "Registered Session:",
                            type: Components.FormControlTypes.TextField,
                            isReadonly: true
                        }
                    },
                    {
                        control: {
                            name: "SessionDay",
                            label: "Session Day:",
                            type: Components.FormControlTypes.TextField,
                            isReadonly: true
                        }
                    },
                    {
                        control: {
                            name: "SessionTime",
                            label: "Session Time:",
                            type: Components.FormControlTypes.TextField,
                            isReadonly: true
                        }
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
                            document.location.href = document.location.href.replace("/DispForm.aspx", "");
                        }
                    },
                    {
                        text: "Edit Registration",
                        type: Components.ButtonTypes.Primary,
                        onClick: () => {
                            // Redirect to the edit form
                            document.location.href = document.location.href.replace("/DispForm.aspx", "/EditForm.aspx");
                        }
                    }
                ]
            });

            // Close the dialog
            dlg.close();
        });
    });
}