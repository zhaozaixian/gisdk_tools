Macro "Dbox Sampler"
    // This macro displays a dialog box (dbox_sampler) that illustrates 
    // all of the user interface items that can be shown in a dialog box.  
    // Each item is displayed, along with the GISDK code used to create it.
    // By clicking on each tab in this dialog box and printing it, you
    // can produce a quick reference chart to use when creating dialog boxes.

    // To run this macro using the Maptitude GISDK toolbox:
    // 1. Choose Tools-Add-Ins and open the GIS Developer's Kit toolbox.
    // 2. Compile this macro in test mode using the first button (Compile).
    // 3. Run this macro by clicking the second button (Test), typing the macro 
    //    name "Dbox Sampler", and clicking OK. The type of add-in is macro
    //   (the default).

  RunDbox("dbox_sampler")

endMacro

Dbox "dbox_sampler" Title: "Display All Dialog Box Items"

    init do
        // Create some variables for use in the dialog box items
        introtext = "This macro illustrates all of the user interface " +
            "items that can be shown in a dialog box.  Each item is " +
            "displayed, along with the GISDK code used to create it. " +
            "By clicking on each tab in this dialog box and printing " +
            "it, you can produce a quick reference chart to use when " +
            "creating dialog boxes."
        sampletext = "This is a sample text line."
        testaddress = "1172 Beacon St"
        pw = "apassword"
        integer = 43543
        cost = 43554.98
        button_prompt = "Button"
        samplepoint = SamplePoint("Font character", "Caliper Cartographic|24",
            87, ColorRGB(0, 0, 0), )  // house symbol
	    sampleline = SampleLine(2, null, ColorRGB(65000,0,0),)
        solid_line = LineStyle({{{0, -1, 0}}})
        str3 = "X X X X "
        str4 = " X X X X"
        shaded = FillStyle({str3, str4, str3, str4, str3, str4, str3, str4})
  	    samplearea = SampleArea(3, solid_line, ColorRGB(0, 0, 0), shaded,
  	        colorRGB(0,0,50000),)
        fmt_array = {
            {{1, "L", "02266"}, {8, "L", "BOSTON"}, {21, "L", "MA"}},
            {{1, "L", "02269"}, {8, "L", "QUINCY"}, {21, "L", "MA"}},
            {{1, "L", "02283"}, {8, "L", "BOSTON"}, {21, "L", "MA"}},
            {{1, "L", "02284"}, {8, "L", "BOSTON"}, {21, "L", "MA"}},
            {{1, "L", "02295"}, {8, "L", "BOSTON"}, {21, "L", "MA"}},
            {{1, "L", "02301"}, {8, "L", "BROCKTON"}, {21, "L", "MA"}},
            {{1, "L", "02302"}, {8, "L", "BROCKTON"}, {21, "L", "MA"}},
            {{1, "L", "02303"}, {8, "L", "BROCKTON"}, {21, "L", "MA"}},
            {{1, "L", "02304"}, {8, "L", "BROCKTON"}, {21, "L", "MA"}}}
        list1_idx = 1
        list2 = {"02266", "02269", "02283", "02284", "02295", "02301", "02302",
            "02303", "02304"}
        tree_array = {
                      {"Accounting", {
                          {"John Smith", {
                              {"Computer 432",}, 
                              {"Desk 4454"}, 
                              {"Chair 555"}
                                          }},
                          {"Jane Jones", {
                              {"Computer 432",}, 
                              {"Computer 654",}, 
                              {"Desk 4454"}, 
                              {"Chair 555"}
                                          }}}},
                      {"Marketing", {
                          {"Patricia Jackson", {
                              {"Computer 444",}, 
                              {"Printer 433654"}, 
                              {"Scanner 53355"},
                              {"Chair 232"}}},
                          {"Geoff Bean", {
                              {"Computer 432",}, 
                              {"Computer 654",}, 
                              {"Desk 4454"}, 
                              {"Chair 555"}
                                          }}}}}
                    
        sizes = {0, .25, .5, .75, 1, 1.5, 2, 2.5, 3, 4, 5, 6, 7, 8, 9, 10}
        do_text = "When the button is pressed, the code after the do is executed."
        EndItem

	// This button, by coming before Tab statements is always visible
	// (not on any tab). The Cancel option will let the user also close the 
    // dialog box by pressing Esc or clicking the Close box in the upper right
    // corner of the dialog box.
	Button "Close" 69.75, 26, 10 Cancel do
		Return()
		endItem

    // Tab list definition starts here.
	Tab List 0.5, 0.5, 80, 25 variable: tab_idx

	//First tab ================================================================
	Tab prompt: "Intro" 

        Text 15, 3, 50, 7 Variable: introtext

        Text same, 10, 50 Variable:
            "The following GISDK code is used to create this tab list: " 
        text " " same, after
        Text 20, after, 40 Variable: "Tab List 0.5, 0.5, 60, 25 variable: tab_idx" 
        Text same, after, 40 Variable: 'Tab prompt: "Intro"' 
        Text same, after, 40 Variable: "<dialog box items go here>"
        Text same, after, 40 Variable: "<more tab items go here>"

	//Second tab ===============================================================
	Tab prompt: "Text Items"

    Text 1, 1, 25 Variable: "Item Display"
    Text 1, 2, 65 Variable:
        "--------------------------------------------------------------------"
    Text "Sample Text" 1, 4
    Text 1, 6, 35 Variable: sampletext
    Text 10, 8, 25 Framed Prompt: "Sample: " Variable: sampletext
    Text 1, 11, 35 Disabled Variable: sampletext

    Text 40, 1, 25 Variable: "GISDK Code to Create Item"
    Text 40, after, 38 Variable:
        "--------------------------------------------------------------------"

    Text 40, 4 Variable: 'Text "Sample Text" 1, 3' 
    Text 40, 6 Variable: "Text 1, 5, 35 Variable: sampletext"
    Text 40, 8 Variable: 'Text 10, 8, 35 Framed Prompt: "Sample"'
    Text 40, 9 Variable: "Variable: sampletext"
    Text 40, 11 Variable: "Text 1, 5, 35 Disabled Variable: sampletext"
    Text 40, 13 Variable: "In the Init macro is the assignment statement"
    Text 40, 14 Variable: 'sampletext  = "This is a sample text line."'

	//Third tab ================================================================
	Tab prompt: "Edit Items"

    Text 1, 1, 25 Variable: "Item Display"
    Text 1, 2, 65 Variable:
        "--------------------------------------------------------------------"

    Edit Text "address" 17, 4, 20 Key: Alt_A Prompt: "Address: "
        Variable: testaddress
    Edit Text "address" same, 7, 20 Disabled Prompt: "Address: " Variable: newadd
    Edit Text "pw" same, 11, 20 Key: Alt_P Prompt: "Enter Password: "
        Variable: pw Password
    Edit Int "integer" same, 15, 8 Key: Alt_I Prompt: "Enter Integer: "
        Variable: integer
    Edit Real "cost" same, 18, 14 Key: Alt_C Prompt: "Enter Cost ($): " Variable:
        cost Format: "$###,##0.00"

    Text 40, 1, 25 Variable: "GISDK Code to Create Item"
    Text 40, after, 38 Variable:
         "--------------------------------------------------------------------"

    Text 40, 4 Variable: 'Edit Text "address" 9, 4, 20 '
    Text 40, after Variable: 'Prompt: "Address: " Variable: testaddress'

    Text 40, 7 Variable: 'Edit Text "address" 15, 6, 20 Key: Alt_A'
    Text 40, after Variable: 'Disabled Prompt: "Address: "'
    Text 40, after Variable: 'Prompt: "Address: " Variable: newadd'

    Text 40, 11 Variable: 'Edit Text "pw" 15, 7, 20 Key: Alt_P'
    Text 40, after Variable: 'Prompt: "Enter Password: " Variable: pw'
    Text 40, after Variable: "Password"

    Text 40, 15 Variable: 'Edit Int "integer" 15, 10, 8 Key: Alt_I'
    Text 40, after Variable: 'Prompt: "Enter Integer: " Variable: integer'

    Text 40, 18 Variable: 'Edit Real "cost" 15, 13, 14 Key: Alt_C Prompt:' 
    Text 40, after Variable: '"Enter Cost ($): " Variable: cost'
    Text 40, after Variable: 'Format: "$###,##0.00"'

	//Fourth tab ===============================================================
	Tab prompt: "Buttons"

    Text 1, 1, 25 Variable: "Item Display"
    Text 1, 2, 44 Variable: "--------------------------------------------------"

    Button "Button" 2,4,10 Key: Alt_B
        do ShowMessage(do_text) EndItem

    Button "b1" 2, 6.7, 10 Key: Alt_U Prompt: button_prompt
        do ShowMessage(do_text) EndItem

    Button "Large Button" 2, 9.5, 14, 1.5 Key: Alt_L
        do ShowMessage(do_text) EndItem

    Button "icon_btn1" 2, 13 icons: "bmp\\stop.bmp", "bmp\\stop2.bmp"
        do ShowMessage(do_text) EndItem

    Button "icon_btn2" 2, 15.5 icons: "bmp\\buttons|105.bmp",
        "bmp\\buttons|139.bmp", "bmp\\buttons|173.bmp"
        do ShowMessage(do_text) EndItem
    
    Text 8, 15.5, 23, 3 Variable:
        "Items in buttons.bmp are numbered from 1 in upper-left, going across."

    Sample Button 2, 20, 3.5, 1.5  contents: samplepoint
        do ShowMessage(do_text) EndItem

    Sample Button 10, 20, 3.5, 1.5  contents: sampleline
        do ShowMessage(do_text) EndItem

    Sample Button 18, 20, 3.5, 1.5  contents: samplearea
        do ShowMessage(do_text) EndItem

    Text 31, 1, 25 Variable: "GISDK Code to Create Item"
    Text 31, after, 45 Variable:
        "-------------------------------------------------------------------" +
        "---------------------"

    Text 31, 4 Variable: 'Button "Button" 2,4,10 Key: Alt_B'

    Text 31, 6.7 Variable: 'Button "b1" 2, 6.7, 10 Key: Alt_U Prompt: ' +
        "button_prompt"

    Text 31, 9.5 Variable: 'Button "Large Button" 2, 9.5, 14, 1.5 Key: Alt_L'

    Text 31, 13 Variable: 'Button "icon_btn1" 2, 13 icons: "bmp\\stop.bmp", '
    Text 31, after Variable: '"bmp\\stop2.bmp"'

    Text 31, 15.5 Variable: 'Button "icon_btn2" 2, 15.5 ' +
        'icons: "bmp\\buttons|105.bmp",'
    Text 31, after Variable: '"bmp\\buttons|139.bmp", "bmp\\buttons|173.bmp"'

    Text 31, 20 Variable: "Sample Button 2, 20, 3.5, 1.5  contents: " + 
        "samplepoint"
    Text 35, 21 Variable: "(or sampleline or samplearea)"

	//Fifth tab ================================================================
	Tab prompt: "Lists"

    Text 1, 1, 25 Variable: "Item Display"
    Text 1, 2, 65 Variable:
        "--------------------------------------------------------------------"


    Popdown Menu "pop1" 6, 4, 28, 8  Prompt: "List: " List: fmt_array
        Variable: list1_idx
        do
            showmessage("You chose row number number " + i2s(list1_idx) + ".")
            EndItem

    Popdown Menu "pop2" 6, 7, 9, 8  Prompt: "List: " List: list2
        Variable: list2_idx Editable 
        do
            showmessage('You chose or entered  "' + list2_idx + '".')
            EndItem

    Text 6, 9.5 Variable: "Double-click a row:"
    Scroll List "scroll1" 6, 10.5, 28, 5 Prompt: "List: " List: fmt_array
        Variables: scroll1_idx, click 
        do
            if click = 1 then do  // double click
                zip = fmt_array[scroll1_idx][1][3]
                city = fmt_array[scroll1_idx][2][3]
                state = fmt_array[scroll1_idx][3][3]
                showmessage("You double-clicked on a row: " + zip + " " +
                    city + " " + state)
            end
        EndItem

    Text 6, 16.5 Variable: "Single-click multiple rows:"
    Scroll List "scroll2" 6, 17.5, 28, 5  Prompt: "List: " List: fmt_array
        Multiple Variables: scroll2_idx 
        do
            index_values = i2s(scroll2_idx[1])
            for i = 2 to scroll2_idx.length do
                index_values = index_values + ", " + i2s(scroll2_idx[i])
                end
            ShowMessage("Currently selected rows are: " + index_values)
        EndItem

    Text 17, 6.75, 17, 3 Variable: '"Editable"--pick a value or enter one.'

    Text 40, 1, 25 Variable: "GISDK Code to Create Item"
    Text 40, after, 38 Variable:
        "--------------------------------------------------------------------"

    Text 40, 3.75 Variable: 'Popdown Menu "list1" 6, 4, 28, 8 Prompt: ' +
        '"List: "'
    Text 40, after Variable: "List: fmt_array Variable: list1_idx "

    Text 40, 6.75 Variable: 'Popdown Menu "list2" 6, 6, 9, 8 Prompt: ' +
        '"List: "' 
    Text 40, after Variable: "List: list2 Variable: list2_idx  Editable"

    Text 40, 10.5 Variable: 'Scroll List "scroll1" 6, 10.5, 28, 5'
    Text 40, after Variable: 'Prompt: "List: " List: fmt_array'
    Text 40, after Variable: "Variables: scroll1_idx, click"

    Text 40, 17.5 Variable: 'Scroll List "scroll2" 6, 17.5, 28, 5'
    Text 40, after Variable: 'Prompt: "List: " List: fmt_array'
    Text 40, after Variable: "Multiple Variables: scroll2_idx"

	//Sixth tab ================================================================
	Tab prompt: "Tree Views"

    Text 1, 1, 25 Variable: "Item Display"
    Text 1, 2, 65 Variable:
        "--------------------------------------------------------------------"

    Tree View 6, 4, 28, 10 Prompt: "Tree: " List: tree_array Variables: tree_idx,
        tree_status, tree_states

    Button "Show Variables" 12, 15, 15 do
            ShowArray({tree_idx, tree_status, tree_states})
        EndItem

    Text 6, 16.5 Variable: "Click Show Variables to show the tree variables:"
    Text same, after Variable: "The first array element shows the chosen item"
    Text same, after Variable: "The second array element contains 1 if you double-clicked"
    Text same, after Variable: "   on an item or clicked to expand or collapse an item"
    Text same, after Variable: "The third array element the shows open/close states of"
    Text same, after Variable: "   the tree (1 = open, 0 = closed, null = no subtree)"
  

    Text 40, 1, 25 Variable: "GISDK Code to Create Item"
    Text 40, after, 38 Variable:
        "--------------------------------------------------------------------"

    Text 40, 4 Variable: 'Tree View 6, 4, 28, 10 Prompt: "Tree: "'
    Text 40, after Variable: "List: tree_array "
    Text 40, after Variable: "Variables: tree_idx, tree_status, tree_states"

	//Seventh tab ==============================================================
	Tab prompt: "Other"

    Text 1, 1, 25 Variable: "Item Display"
    Text 1, 2, 65 Variable:
        "--------------------------------------------------------------------"

    Radio List 3, 4, 25, 4.5 Prompt: "A Radio List" Variable: radio_idx
    Radio Button 6, 5.5 Prompt: "Choice 1"
    Radio Button same, 7 Prompt: "Choice 2"

    Frame 3, 10, 28, 8.5 Prompt: "Use a frame to group items"

	Checkbox 5, 13 Prompt: "A Checkbox" variable: opt

    Spinner 16, 16, 8 prompt: "Choose Size" list: sizes  variable: choice 

    Text 40, 1, 25 Variable: "GISDK Code to Create Item"
    Text 40, after Variable:
        "--------------------------------------------------------------------"

    Text 40, 4 Variable: 'Radio List 3, 4, 25, 4.5 Prompt: "A Radio List"' 
    Text 40, after Variable: "     Variable: radio_idx"
    Text 40, after Variable: 'Radio Button 6, 5.5 Prompt: "Choice 1"'
    Text 40, after Variable: 'Radio Button same, 7 Prompt: "Choice 2"'

    Text 40, 10, 39 Variable: "Frame 3, 10, 28, 6.5 "
    Text 40, after, 39 Variable: 'Prompt: "Use a frame to group items"'

    Text 40, 13, 39 Variable: 'Checkbox 5, 12 Prompt: "A Checkbox"'
    Text 40, after, 39 Variable: "variable: opt"

    Text 40, 16, 39 Variable: 'Spinner 16, 14, 6 prompt: "Choose Size"'
    Text 40, after, 39 Variable: "list: sizes  variable: choice"

endDbox
