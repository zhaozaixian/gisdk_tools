
//1. Define Global Color
	Color_Red=ColorRGB(65535,0,0)
	Color_Green=ColorRGB(0, 65535,0)
	Color_Blue=ColorRGB(10000, 10000, 65535)
	Color_White=ColorRGB(65535,65535,65535)
	Color_Black=ColorRGB(0,0,0)
	Color_Orange=ColorRGB(65535, 33767, 0)
	Color_Yellow=ColorRGB(65535, 65535, 0)
	Color_Gray= ColorRGB(50000, 50000, 50000)
	Color_Cyan=ColorRGB(0, 58982, 58982)
	Color_Gold=ColorRGB(65535, 54430, 0)
	Color_Brown=ColorRGB(39320, 26215, 13100)
	Color_Purple=ColorRGB(32768,0,65535)

			
//2. **************Show TAZ Color Theme********************
	If Chart_idx=1 then do
		on NotFound goto NEXT
		HideTheme( , "TAZ")
		DestroyTheme("TAZ")
		NEXT:
		Purpose_idx={Purpose_idx[1]}
		Theme_Color = CreateTheme("TAZ",  "view_Join."+(Purpose_list[Purpose_idx[1]]), "Quantiles", 8, )
		Array_colors= { ColorRGB(48400,48400,65025), ColorRGB(49152,65535,53832),ColorRGB(49152,58513,65535),ColorRGB(53832,65535,49152),
									ColorRGB(36864,36864,65025),ColorRGB(65535,58513,49152),ColorRGB(65025,52900,33856),ColorRGB(65535,49152,49152),ColorRGB(65025,34969,65025)}
		SetThemeFillColors(Theme_Color , Array_colors)
		ShowTheme(, Theme_Color)
	end


//3. **************Show TAZ Pie chart********************
	If Chart_idx=2 then do
		on NotFound goto NEXT01
		HideTheme( , "PA Balance")
		DestroyTheme("PA Balance")
		NEXT01:
		My_Array={"view_Join."+(Purpose_list[Purpose_idx[1]])}
		if  Purpose_idx.Length>1then do
				for i=2 to Purpose_idx.Length do
					Fields_Array=InsertArrayElements(My_Array,12,{"view_Join."+(Purpose_list[Purpose_idx[i]])})
					My_Array=Fields_Array
				end
		end
		Theme_PA=CreateChartTheme("PA Balance",  My_Array,"Pie", {"Data Source","Screen"})
		SetThemeFillColors(Theme_PA, fill_colors)
		ShowTheme(, Theme_PA)
	end


//4. **************Show Link Color Theme********************
Theme_Width = CreateContinuousTheme("Vehicle Flows", {view+"."+"AB_"+Flowfield_List[Flowfield_idx]},)             // Creates the line width theme
ShowTheme(, Theme_Width)
Solid = LineStyle({{{1, -1, 0}}})
ls_Array = {Solid}
SetThemeLineStyles(Theme_Width, ls_Array)
SetThemeLineColors(Theme_Width, {Color_Red})
if Flowfield_idx<=Array_Peak[1].length then do
	Theme_Color = CreateTheme("VOC Ratio",view+"."+"AB_"+Left(Flowfield_List[Flowfield_idx],2)+"_VOC", "Manual", 3,{{"values", {{0, "True", 0.8, "False"},
																 {0.8, "True", 1.0,"False"},
																 {1.0, "True", 99,"False"}  }}})
	//Color_Array = GeneratePalette(ColorRGB(0,65535, 0), ColorRGB(65535,0, 0), 1, )		//the number is the middle to insert.
	Color_Array ={Color_Black,Color_Green,Color_Yellow,Color_Red}
	SetThemeLineColors(Theme_Color, Color_Array)
	ShowTheme(,Theme_Color)
	SetLabels(view+"|","AB_"+Flowfield_List[Flowfield_idx],{ {"Rotation", "True"}, {"Font", "Arial|Bold|9"},{"Left/Right","True"},
			 {"Color", Color_Black},{"Visibility","True"}})    //Left/Right to get the BA fields ,{"Format","#####"}
	end
else do
	SetLabels(view+"|",Flowfield_List[Flowfield_idx],{ {"Rotation", "True"}, {"Font", "Arial|Bold|9"},{"Left/Right","True"},
			 {"Color", Color_Black},{"Visibility","True"}})    //Left/Right to get the BA fields ,{"Format","#####"}
end
RunMacro("G30 create legend", "Theme")
