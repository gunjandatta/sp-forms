import { ContextInfo, Helper, List, SPTypes, Types } from "gd-sprest-bs";

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
                } as Helper.IFieldInfoNumber,
                {
                    name: "SessionInfo",
                    title: "Session Information",
                    type: Helper.SPCfgFieldType.Calculated,
                    fieldRefs: ["SessionDay", "SessionTime"],
                    formula: '=[Session Day]&amp;" - "&amp;[Session Time]'
                } as Helper.IFieldInfoCalculated
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
                Title: ListName,
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
            ],
            ViewInformation: [
                {
                    ViewName: "All Items",
                    ViewFields: ["LinkTitle", "SessionsLU"]
                }
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

/** Updates the new/edit/display forms */
Configuration["updateListForms"] = () => {
    // Method to add the webpart to the form
    let addWebPart = (form: Types.SP.Form) => {
        let method: string = "";

        // Determine the type
        switch (form.FormType) {
            case SPTypes.PageType.DisplayForm:
                // Set the method
                method = "DisplayForm";
                break;
            case SPTypes.PageType.EditForm:
                // Set the method
                method = "EditForm";
                break;
            case SPTypes.PageType.NewForm:
                // Set the method
                method = "NewForm";
                break;
            // Skip this form by default
            default:
                return;
        }

        // Add a script editor webpart to the list item form
        Helper.WebPart.addWebPartToPage(form.ServerRelativeUrl, {
            chromeType: "None",
            title: "List Form",
            index: 99, // We want to set the index to render the default list form first
            zone: "Main",
            content: [
                '<div id="list-form"></div>',
                '<script src="' + ContextInfo.webServerRelativeUrl + '/siteassets/sp-forms.js"></script>',
                '<script>',
                '// Wait for the sp-forms script to be loaded',
                'SP.SOD.executeOrDelayUntilScriptLoaded(function() {',
                '\t// Get the element to render the list form to',
                '\tvar el = document.querySelector("#list-form");',
                '',
                '\t// Render the list form',
                '\tFormsDemo.' + method + '(el);',
                '}, "forms-demo");',
                '</script>'
            ].join('\n')
        }).then(() => {
            // Log
            console.log("List form '" + method + "' was updated.");
        });
    }

    // Get the list forms
    List(ListName).Forms().execute(forms => {
        // Parse the forms
        for (let i = 0; i < forms.results.length; i++) {
            // Add the webpart to the form
            addWebPart(forms.results[i]);
        }
    });
}