import { Components } from "gd-sprest-bs";
import { ISessionInfo } from "./index.d";

// Method to get the item id from the query string
export function getItemID() {
    let itemId = 0;

    // Get the item id from the querystring
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

    // Return the item id
    return itemId;
}

// Method to intialize the webpart and return the close button element
export function initForm(onSave: () => void): HTMLElement {
    let elCancelButton: HTMLElement = null;

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

        // Set the "Cancel" button
        elCancelButton = elWebpart.querySelector("input[value='Cancel']");
    }

    // Return the button
    return elCancelButton;
}

// Method to update the sessions dropdown
export function updateSessions(ddl: Components.IDropdown, sessions: Array<ISessionInfo> = [], selectedId?: number) {
    let items: Array<Components.IDropdownItem> = [];

    // Ensure the dropdown exists
    if (ddl == null) { return; }

    // Parse the data
    for (let i = 0; i < sessions.length; i++) {
        let session = sessions[i];

        // Add a dropdown item for this session
        items.push({
            isSelected: session.itemId == selectedId,
            text: session.day + " - " + session.time,
            value: session.itemId.toString()
        });
    }

    // Update the dropdown
    ddl.setItems(items);
}
