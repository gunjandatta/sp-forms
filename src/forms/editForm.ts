/**
 * Edit Form
 */
export const EditForm = (el: HTMLElement) => {
    // Define the click event for the form
    let onSave = () => {

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
    }

    // TODO
    el.innerHTML = "<h3>TO DO</h3><h5>Create Custom Edit Form</h5>";
}