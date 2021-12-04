macro "CreateIntersectionFigure"
//The options are for a 3-leg intersection.

    shared cc_Colors

    on error do
        ShowMessage(GetLastError())
        return()
        end

    opts = null
    opts.Angles = {
        45.9,  //angle of leg 1, clockwise from North
        211.1, //angle of leg 2, clockwise from North
        90.5   //angle of leg 3, clockwise from North
        }
    opts.Flows = {
        {0,          //flow from leg 1 to leg 1
         1527.8,     //flow from leg 1 to leg 2
         4265.8},    //flow from leg 1 to leg 3

        {2146.8,     //flow from leg 2 to leg 1
         0,          //flow from leg 2 to leg 2
         585.0},     //flow from leg 2 to leg 3

        {3132.1,     //flow from leg 3 to leg 1
         1222.8,     //flow from leg 3 to leg 2
         0}          //flow from leg 3 to leg 3
     }

    opts.Colors = {
        cc_Colors.Black, cc_Colors.Red, cc_Colors.Green, cc_Colors.Orange //...
        }

    opts.[Title text] = "Beacon/Walnut"
    opts.[Title font] = "Arial|Bold|14"
    opts.[Footnote text] = "Am peak hour flows"
    opts.[Footnote font] = "Arial|Bold|12"
    opts.[Flow Labels] = "On" //to display flow values on the diagram
    opts.[Road Labels] = {
        "Road 1",
        "Main St",
        "Walnut St"
        }

    opts.[Road Label Fonts] = "Arial|Bold|10"


    // Add custom error traps here ....
    if opts = null then
        Throw("No inputs provided for intersection diagram")

    CreateFigure("My Intersection Figure", "Intersection", opts)

   EndMacro