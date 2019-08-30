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

    // TODO
    el.innerHTML = "<h3>TO DO</h3><h5>Create Custom Display Form</h5>";
}