import { Components, Web, Types } from "gd-sprest-bs";
import { ListNames } from "../cfg";

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

export function getSession(itemId: number): PromiseLike<Types.SP.IListItemQuery> {
    // Return a promise
    return new Promise((resolve, reject) => {
        // Query the main list
        Web().Lists(ListNames.Main).Items(itemId).query({
            Expand: ["SessionsLU"],
            Select: ["Id", "SessionsLUId", "SessionsLU/Title", "SessionsLU/SessionInfo"]
        }).execute(resolve, reject);
    });
}

/**
 * Gets the session information dropdown list items.
 * @returns The sessions information dropdown list items. key - string, value - Array of dropdown items
 */
export function getSessions(): PromiseLike<{ [key: string]: Array<Components.IDropdownItem> }> {
    // Return a promise
    return new Promise((resolve, reject) => {
        // Query the configuration list
        Web().Lists(ListNames.Configuration).Items().query({
            OrderBy: ["Title", "SessionDay", "SessionTime"],
            Select: ["Id", "Title", "SessionInfo"]
        }).execute(items => {
            let sessions: { [key: string]: Array<Components.IDropdownItem> } = {};

            // Parse the items
            for (let i = 0; i < items.results.length; i++) {
                let item = items.results[i];

                // Get the session name
                let session = item["Title"];
                let sessionInfo = item["SessionInfo"];
                if (session && sessionInfo) {
                    // Ensure the session key exists
                    if (sessions[session] == null) {
                        // Create an array for this session
                        sessions[session] = [];
                    }

                    // Append the session information
                    sessions[session].push({
                        data: item,
                        text: sessionInfo,
                        value: item.Id.toString()
                    });
                }
            }

            // Resolve the promise
            resolve(sessions);
        }, reject);
    });
}

/**
 * Hides the default list form.
 * @returns The webpart element
 */
export function hideDefaultForm(): HTMLElement {
    // Get the list form webpart element
    let elWebpart = document.querySelector(".ms-webpart-zone > div[id^='MSOZoneCell']") as HTMLElement;
    if (elWebpart) {
        // Hide the default list form
        elWebpart.style.display = "none";
    }

    // Return the webpart element
    return elWebpart;
}

/**
 * Gets the item information.
 * @param onSave - The event triggered by the 'Save' ribbon button
 * @returns - The form's 'Close' button element
 */
export function initForm(onSave: () => void): HTMLElement {
    let elCancelButton: HTMLElement = null;

    // Hide the list form
    let elWebpart = hideDefaultForm();
    if (elWebpart) {
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