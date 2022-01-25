///////////===================
Macro "MyTest"

    //  Show a Startup Bitmap
        opts=null
        opts.StartupClose="True"   opts.Borders="True"   opts.Title="自开发程序 BY Zaixian Zhao" 
        bmpfile="c:\\titlebar.bmp"
        ShowBitmap(bmpfile,opts)

EndMacro

//////////=================自定义的菜单系统 ========================
/////
Menu "My Menu System"
     // This menu item is the top-level title to be added to the main menu system:
     MenuItem "Extra Menu" text: "交通规划程序" key: alt_x   Menu "Extra Dropdown Menu"
EndMenu
 
Menu "Extra Dropdown Menu"
     // This is the body of the drop-down menu to be added
     MenuItem "hello" text: "Hello"   key: alt_h    do    RunMacro("say hello")      endItem
     Separator 
     MenuItem "bye"    text: "Bye"     key: alt_b     do    RunMacro("say bye")        endItem
EndMenu

///////////====================菜单功能实现================ 
/////
Macro "say hello"
     ShowMessage("Hello!")
EndMacro

Macro "say bye"
     ShowMessage("Bye!")
EndMacro