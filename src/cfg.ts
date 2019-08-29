import { Helper, List, SPTypes } from "gd-sprest-bs";

/** List Name */
export const ListName = "Custom Forms Demo";

/** Configuration */
export const Configuration = Helper.SPConfig({
    /** Lists */
    ListCfg: [
        /** Associated list with session information. */
        {
            ListInformation: {
                Title: "Sessions",
                BaseTemplate: SPTypes.ListTemplateType.GenericList
            },
            CustomFields: [
                {
                    name: "SessionDay",
                    title: "Session Day",
                    type: Helper.SPCfgFieldType.Choice,
                    choices: ["Session 1", "Session 2", "Session 3", "Session 4", "Session 5"]
                } as Helper.IFieldInfoChoice,
                {
                    name: "SessionTime",
                    title: "Session Time",
                    type: Helper.SPCfgFieldType.Choice,
                    choices: ["8:00 AM", "10:00 AM", "12:00 PM", "2:00 PM"]
                } as Helper.IFieldInfoChoice,
                {
                    name: "SessionCapacity",
                    title: "Session Capacity",
                    type: Helper.SPCfgFieldType.Number,
                    numberType: SPTypes.FieldNumberType.Integer
                } as Helper.IFieldInfoNumber
            ],
            ViewInformation: [
                {
                    ViewName: "All Items",
                    ViewFields: ["LinkTitle", "SessionDay", "SessionTime", "SessionCapacity"]
                }
            ]
        },
        /** List with Custom Forms */
        {
            ListInformation: {
                Title: "Custom Form Demo",
                BaseTemplate: SPTypes.ListTemplateType.GenericList
            },
            CustomFields: [
                {
                    name: "SessionsLU",
                    title: "Session",
                    type: Helper.SPCfgFieldType.Lookup,
                    listName: "Sessions",
                    showField: "Title"
                } as Helper.IFieldInfoLookup
            ]
        }
    ]
});

/** Adds sample data to the sessions list */
Configuration["createSampleData"] = () => {
    let classNames = ["SharePoint Development", "SharePoint Branding", "Power Apps", "PowerBI", "Flow"];
    let sessions = ["Session 1", "Session 2", "Session 3", "Session 4", "Session 5"];
    let times = ["8:00 AM", "10:00 AM", "12:00 PM", "2:00 PM"];

    // Generates a random number
    let generateIndex = (min: number, max: number) => { return Math.floor(Math.random() * (max - min + 1) + min); };

    // Create 10 items
    for (var i = 0; i < 10; i++) {
        // Generate a random index for both
        let idx1 = generateIndex(0, classNames.length - 1);
        let idx2 = generateIndex(0, sessions.length - 1);
        let idx3 = generateIndex(0, times.length - 1);

        // Add the list item
        List("Sessions").Items().add({
            Title: classNames[idx1],
            SessionDay: sessions[idx2],
            SessionTime: times[idx3],
            SessionCapacity: 25
        }).execute(
            // Success
            item => { console.log("Created item " + item["Title"], item); },
            // Error
            () => { console.error("Error creating sample item."); },
            true
        );
    }
}