import { Components, Helper, List } from "gd-sprest-bs";

/**
 * Display Form
 */
export const DisplayForm = (el: HTMLElement) => {
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

    // Display a loading message
    Helper.SP.ModalDialog.showWaitScreenWithNoClose("Loading the Registration");

    // Get the item id from the querystring
    let itemId = 0;
    let qs = document.location.search.split('?');
    qs = qs.length > 0 ? qs[1].split('&') : [];
    for (let i = 0; i < qs.length; i++) {
        let qsItem = qs[i].split('=');

        // See if this is the "id" key
        if (qsItem[0] == "ID") {
            // Set the item id
            itemId = parseInt(qsItem[1]);
            break;
        }
    }

    // Get the list item information
    // We are expanding the lookup field to include the "calculated" information field
    // You cannot expand "choice" fields, which is why a "calculated" field was created
    List("Custom Form Demo").Items(itemId).query({
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

        // Close the dialog
        Helper.SP.ModalDialog.commonModalDialogClose();
    });
}