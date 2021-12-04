// Project:
// Project Manager: 
// Project Team members: Jandy,
//GISDK Author: Jandy
//Created Data: 2011-
// Purpose:1. 
//                2. 

Macro "Model"	
	Global Path_Model, Path_Scen,Year,CarOwer_Array,Purpose_Array,Mode_Array

	//Path_Model="D:\\P13 - 汕头市交通模型\\Model\\"
	//Path_Model="D:\\Shantou Model\\"

	if Path_Model=null  then do
		Txt_File_Name="C:\\TC_PowerTool\\UI_Path.txt"
		if  GetFileInfo(Txt_File_Name) then do
			fptr = OpenFile(Txt_File_Name, "r")
			Path_Temp=ReadLine(fptr)
			CloseFile(fptr)
			if  GetFileInfo(Path_Temp+"GISDK\\model.dbd") then do
				Path_Model=Path_Temp
			end
		end
	end

	if Path_Model=null  then Path_Model="D:\\Shantou Model\\"

	if GetFileInfo(Path_Model+"GISDK\\model.dbd")=null then do showMessage("没有找到模型文件系统.") return() end

	CarOwer_Array={"CA","MC","NC"}
	Purpose_Array={"HBW","HBS","HBO","NHB"}    // 1:HBW	 2: HBE	3:HBO	4:NHB	
	Mode_Array={"Bicycle","Moto","Car","PT"}
	
	/*
	RunMacro("Skim_Len_Time_PT")		
	RunMacro("CalMat_Congested")
	RunMacro("Step1 Generation Model")
	RunMacro("Step2 Trip Distribution")
	RunMacro("Step3 Mode Split")
	RunMacro("Step4 PA2OD")
	RunMacro("Step5 Traffic Assignment")
	RunMacro("Step6 Transit Assignment")
	*/
	on NotFound goto NEXT								//Now remove the menu,...
		RemoveMenuItem("Top_Menu")
        NEXT:										
        AddMenuItem("Top_Menu", "Before", "Window")        //Add Menu

	RunDbox("dbox_Interface")	
EndMacro

Menu "Top Menu"       //Inserts a menu item in the same bar as with the standard menu
	MenuItem "Top_Menu" text:"汕头模型" do	 RunDbox("dbox_Interface")	EndItem
EndMenu

//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Dbox "dbox_Interface"Title: "快捷管理模型界面 - " + Project_Name  Toolbox NoKeyboard
init do
	On NotFound goto NEXT
		CloseDbox("dbox_Interface")
	NEXT:

	Project_Name="昆山交通模型"
	Macro_Version="v1.0"

	shared d_matrix_options // default options used for creating matrix editors
	shared Scen_Name
	pt_sz=1
	Scenario_idx=1
	
       About_Us= Project_Name + "  -  " + Macro_Version + 
			"\n\n\n 快捷管理模型系统的功能包括： " +
			"\n 1.小区属性编辑." +
			"\n 2.路网属性编辑." +
			"\n 3.基本的四阶段功能."+
			"\n 4.运行反馈循环模型."

	Array_Scenario_Folder=null
	Array_Scenario_Year=null
	v = OpenTable("v", "FFB", { Path_Model+"scenarios.bin"}, {{"Shared", "True"}})
	rec=GetFirstRecord(v+"|", {{"ID","Ascending"}})
	while rec<>null do 
		Array_Scenario_Folder=Array_Scenario_Folder+{Trim(v.[Name])}
		Array_Scenario_Year=Array_Scenario_Year+{v.[Year]}
		if v.[MyDefault]=1  then do
			Scenario_idx=v.[ID]
		end
		rec=GetNextRecord(v+"|",null, {{"ID","Ascending"}})
	end
	closeview(v)

	Year=Array_Scenario_Year[Scenario_idx]
	Path_Scen=Path_Model+Array_Scenario_Folder[Scenario_idx]+"\\"   //Path为某一情景的文件夹

	if GetDirectoryinfo(Path_Scen+"*","Directory")=null  then do 
		ShowMessage("方案名称有误，或方案文件夹不存在。")   
		Scenario_idx=1 
		Year=Array_Scenario_Year[Scenario_idx]
		Path_Scen=Path_Model+Array_Scenario_Folder[Scenario_idx]+"\\"   //Path为某一情景的文件夹
	end
	Scen_Name=Array_Scenario_Folder[Scenario_idx]
EndItem

Tab List 0.5, 0,83, 25 variable: Tab_idx                         // Tab list definition starts here.  Tab_idx to definiton which tab is selected.
//---First tab -- The following are the left hand icons where a user clicks and a folder opens.
Tab prompt: "运行模型"
	
// The following lines picks the bmp picture used for the front page of  the mian user interface
	
	Checkbox 0, 0 icons: Path_Model+"GISDK\\Front_Logo.bmp" Help:"Model Path: "+Path_Model variable: pt_sz    //the Left pic, Riyadh Traffic Model
	do
		if GetFileInfo(Path_Model+"GISDK\\User's Guide.pdf")  then LaunchDocument(Path_Model+"GISDK\\User's Guide.pdf",)
	endItem

	Text 46,1 variable:"选择方案："
	Popdown Menu same,after, 25 list: Array_Scenario_Folder variable: Scenario_idx
	do
		Year=Array_Scenario_Year[Scenario_idx]
		Path_Scen=Path_Model+Array_Scenario_Folder[Scenario_idx]+"\\"   //Path为某一情景的文件夹
		if GetDirectoryinfo(Path_Scen+"*","Directory")=null  then do ShowMessage("方案名称有误，或方案文件夹不存在。")   Scenario_idx=1 end
		v = OpenTable("v", "FFB", { Path_Model+"scenarios.bin"}, {{"Shared", "True"}})
		rec=GetFirstRecord(v+"|", {{"ID","Ascending"}})
		while rec<>null do 
			v.[MyDefault]=0
			rec=GetNextRecord(v+"|",null, {{"ID","Ascending"}})
		end
		rh2 = LocateRecord("v|", "Name", {Array_Scenario_Folder[Scenario_idx]}, )
		v.[MyDefault]=1
		closeview(v)
		Scen_Name=Array_Scenario_Folder[Scenario_idx]
	EndItem

	Button "查看" after,same,7 Help:"View this Scenario"  //button, View Scenarios
	do
		RunMacro("Close All Maps")
		RunMacro("G30 File Close All")

		//1. Network Map
		Network_Map=Path_Scen+"Network\\Network.map"
		if GetFileInfo(Network_Map) then OpenMap(Network_Map,)

		//2. TAZ Map
		TAZ_Map=Path_Scen+"Zones\\Zones.map"
		if GetFileInfo(TAZ_Map) then OpenMap(TAZ_Map,)
		TileWindows()

		Return()
	EndItem	
   
   	Button "icon_btn1" 44, 6 icons: "bmp\\plantripgen.bmp"  
	do
		TAZ_Map=Path_Scen+"Zones\\Result.map"
		if GetFileInfo(TAZ_Map) then OpenMap(TAZ_Map,)
	enditem
	Button "icon_btn2" same,after icons: "bmp\\plantripdist.bmp" 
	do
		mx=OpenMatrix(Path_Scen+"Matrix\\cgrav.mtx",)
		CreateMatrixEditor("交通分布", mx, d_matrix_options)
	enditem
	Button "icon_btn3" same, after icons: "bmp\\planmodesplit.bmp" 
	do
		mx=OpenMatrix(Path_Scen+"Matrix\\ModeSplit.mtx",)
		CreateMatrixEditor("方式划分", mx, d_matrix_options)
	enditem
	Button "icon_btn4" same, after icons: "bmp\\planmatrix.bmp" 
	do
		mx=OpenMatrix(Path_Scen+"Matrix\\PA2OD.mtx",)
		CreateMatrixEditor("PA2OD", mx, d_matrix_options)
	enditem
	Button "icon_btn5" same, after icons: "bmp\\planassign.bmp" 
	do
		Network_Map=Path_Scen+"Network\\Result.map"
		if GetFileInfo(Network_Map) then OpenMap(Network_Map,)
	enditem
	Button "icon_btn6" same, after icons: "bmp\\planskim.bmp"
	do
		Network_Map=Path_Scen+"Network\\Result-PT.map"
		if GetFileInfo(Network_Map) then OpenMap(Network_Map,)
	enditem
	Button "icon_btn7" same, after icons: "bmp\\plansetup.bmp" 

	 Button "交通生成" after,6,20,1.7					//button, Trip Generation
	   do
		RunMacro("Close All Maps")
		RunMacro("G30 File Close All")
		RunMacro("Step1 Trip Generation")
		TAZ_Map=Path_Scen+"Zones\\Result.map"
		if GetFileInfo(TAZ_Map) then OpenMap(TAZ_Map,)
	EndItem
	   
	 Button "交通分布" same,after,20,1.7					//button, Trip Distribution
	   do
		RunMacro("Close All Maps")
		RunMacro("G30 File Close All")
		RunMacro("Step2 Trip Distribution")
		mx=OpenMatrix(Path_Scen+"Matrix\\cgrav.mtx",)
		CreateMatrixEditor("交通分布", mx, d_matrix_options)
	   EndItem
	   
         Button "方式划分" same, after ,20,1.7					//button,  Mode Split
	   do
		RunMacro("Close All Maps")
		RunMacro("G30 File Close All")
		RunMacro("Step3 Mode Split")
		mx=OpenMatrix(Path_Scen+"Matrix\\ModeSplit.mtx",)
		CreateMatrixEditor("方式划分", mx, d_matrix_options)
	   EndItem
	   
         Button "PA 转 OD" same, after ,20,1.7						//button,  PA to OD
	   do
		RunMacro("Close All Maps")
		RunMacro("G30 File Close All")
		RunMacro("Step4 PA2OD")
		mx=OpenMatrix(Path_Scen+"Matrix\\PA2OD.mtx",)
		CreateMatrixEditor("PA2OD", mx, d_matrix_options)
	   EndItem
	   
         Button "道路分配" same, after ,20,1.7					//button,  Trip Assignment
	   do
		RunMacro("Close All Maps")
		RunMacro("G30 File Close All")
		RunMacro("Step5 Traffic Assignment")
		Network_Map=Path_Scen+"Network\\Result.map"
		if GetFileInfo(Network_Map) then OpenMap(Network_Map,)
	   EndItem
	   
	Button "公交分配" same, after,20,1.7		 		//button,  Select Link Analysis
	do
		RunMacro("Close All Maps")
		RunMacro("G30 File Close All")
		RunMacro("Step6 Transit Assignment")
		Network_Map=Path_Scen+"Network\\Result-PT.map"
		if GetFileInfo(Network_Map) then OpenMap(Network_Map,)
	EndItem
	
	 Button "运行反馈四阶段" same, after,20,1.7					//button,  Sub Area Analysis
	do
		btn = MessageBox("您确定要运行反馈循环的四阶段模型吗？\n 根据您电脑配置不同，模型运行大约需要10分钟。",
		{{"Caption", "PowerTool"},{"Icon","Warning"} ,{"Buttons", "YesNo"}})
		if btn = "No" then Return()
		On notfound goto NEXT
			HideDbox("dbox_Interface")
		NEXT:
		RunMacro("Close All Maps")
		RunMacro("G30 File Close All")
		RunMacro("Run All Steps and Loops")    //feedback loops and all steps
	EndItem

//-- Sencond tab --
Tab prompt: "介绍"								//About us

	Text 1, after
	Text same, after, 42,11 Framed Variable:About_Us

	Text 2, 15 variable:"附加分析功能："

	  Button "更新自由路网" 2,17,20,1.7 Help:"不考虑拥堵的"					//button, Trip Generation
	   do
		RunMacro("Close All Maps")
		RunMacro("G30 File Close All")
		RunMacro("UpdateNetworkFF")
	EndItem
	 Button "更新拥堵路网" after,same,20,1.7 Help:"考虑拥堵的，必须在分配后才行"					//button, Trip Generation
	   do
		RunMacro("Close All Maps")
		RunMacro("G30 File Close All")
		RunMacro("UpdateNetworkCongested")
	EndItem
	Button "获取阻抗矩阵" 2,after,20,1.7 Help:"Len,Time and PT"					//button, Trip Generation
	   do
		RunMacro("Close All Maps")
		RunMacro("G30 File Close All")
		RunMacro("Skim_Len_Time_PT")
	EndItem
	Button "计算综合阻抗" after,same,20,1.7 Help:"GC"					//button, Trip Generation
	   do
		RunMacro("Close All Maps")
		RunMacro("G30 File Close All")
		RunMacro("CalMat_Congested")
		mx=OpenMatrix(Path_Scen+"Matrix\\GC.mtx",)
		CreateMatrixEditor("GC", mx, d_matrix_options)
	EndItem
Close do
	return()
EndItem
EndDbox

Macro "Run All Steps and Loops"

	Shared Mode_Split
	Shared Scen_Name	
	
	Mode_Split={0.309,0.507,0.123,0.061}		//原始调查值
	Max_Loop=5
	Time_Start = GetDateAndTime()
	fptr = OpenFile(Path_Scen+"Mode Share.txt", "a")
	
	WriteLine(fptr, " ")
	WriteLine(fptr, "=====================")
	WriteLine(fptr, "Base Year(2012):  Bicycle:31%  Moto:51%   Car:12%   PT:6%")
	WriteLine(fptr, " ")
	WriteLine(fptr, Time_Start)

	for Feedback_Loop=1 to Max_Loop do    //Max_Loop Define at the start.	
		on Escape goto quit
		 WriteLine(fptr, " ")	
		 WriteLine(fptr, "Loop="+I2S(Feedback_Loop))

		 if Feedback_Loop=1  then do
			SetStatus(1,Scen_Name+":Loop="+I2S(Feedback_Loop)+"(Total="+ I2S(Max_Loop) +").计算交通小区，更新交通生成",)
			RunMacro("Step1 Trip Generation")

			SetStatus(1,Scen_Name+":Loop="+I2S(Feedback_Loop)+"(Total="+ I2S(Max_Loop) +").更新道路网络",)
			RunMacro("UpdateNetworkFF")   //When Feedback_Loop=1, using the FF time, else, using the congested time
		 end
		
		if Feedback_Loop>1  then do
			SetStatus(1,Scen_Name+":Loop="+I2S(Feedback_Loop)+"(Total="+ I2S(Max_Loop) +").更新阻抗时间",)
			RunMacro("UpdateNetworkCongested")
		end
		
		SetStatus(1,Scen_Name+":Loop="+I2S(Feedback_Loop)+"(Total="+ I2S(Max_Loop) +").更新阻抗矩阵",)
		RunMacro("Skim_Len_Time_PT")		//Len Time PT
		RunMacro("CalMat_Congested")		//GC

		SetStatus(1,Scen_Name+":Loop="+I2S(Feedback_Loop)+"(Total="+ I2S(Max_Loop) +"). 交通分布",)
		RunMacro("Step2 Trip Distribution")

		SetStatus(1,Scen_Name+":Loop="+I2S(Feedback_Loop)+"(Total="+ I2S(Max_Loop) +").方式划分，计算分担率",)
		RunMacro("Step3 Mode Split")
		
		//在每次循环时，Mode_Split 应更新
		//打开Mode Split
		File_Mode=Path_Scen+"Matrix\\ModeSplit.mtx"
		mx_Mode=OpenMatrix(File_Mode,)
		Sum_Stat=0
		stat_array = MatrixStatistics(mx_Mode, )
		for i=1 to  Mode_Array.length do
			Sum_Stat = Sum_Stat + stat_array.(Mode_Array[i]).Sum
		end
		str=null
		for i=1 to  Mode_Array.length do
			Mode_Split[i]=stat_array.(Mode_Array[i]).Sum/Sum_Stat	
			str=str+Mode_Array[i]+":"+r2s(round(Mode_Split[i],2)*100)+"%  "
		end		
		WriteLine(fptr, str)
		mx_Mode=null		

		SetStatus(1,Scen_Name+":Loop="+I2S(Feedback_Loop)+"(Total="+ I2S(Max_Loop) +").执行 PA2OD.",)
		RunMacro("Step4 PA2OD")

		SetStatus(1,Scen_Name+":Loop="+I2S(Feedback_Loop)+"(Total="+ I2S(Max_Loop) +").道路分配.",)
		RunMacro("Step5 Traffic Assignment")

		 if Feedback_Loop=Max_Loop  then do
			SetStatus(1,Scen_Name+":Loop="+I2S(Feedback_Loop)+"(Total="+ I2S(Max_Loop) +").公交分配.",)
			RunMacro("Step6 Transit Assignment")
		end
	end

	RunMacro("UpdateNetworkCongested")
	
	CloseFile(fptr)
	//Mesure Time
	Time_End = GetDateAndTime()
	ShowMessage(Scen_Name+":模型开始于: "+Time_Start+" \n\n 成功结束于: "+Time_End)
	SetStatus(2,"@system1",)
	quit:
	return()
EndMacro











//#########################################################################################################################################################
//###################### 第一步 交通生成  ##########################################################################################
//#########################################################################################################################################################

Macro "Step1 Trip Generation"			//Generation Model - Regression Analysis
	
	RunMacro("Close All Maps")
	RunMacro("G30 File Close All")

	//Rates={{1.019 ,0.176 ,0.442 ,0.220 },{1.049 ,0.217 ,0.470 ,0.221 },{0.563 ,0.148 ,0.314 ,0.150 }}
	v = OpenTable("v", "FFB", { Path_Scen+"Rates.bin"}, {{"Shared", "True"}})
	for m=1 to CarOwer_Array.Length do			
			v1 = GetDataVector(v+"|",CarOwer_Array[m], )
			Rates=Rates+{V2A(v1)}			
	end
	closeview(v)

	SUM_P={{1,1,1,1},{1,1,1,1},{1,1,1,1}}
	SUM_A={{1,1,1,1},{1,1,1,1},{1,1,1,1}}

	Coff_HBW={2.373}  //总岗位
	Coff_HBS={0.8675}  //学位
	Coff_HBO={0.378,0.142,0.400} //商业	办公	非农人口
	Coff_NHB={0.1555,0.3758,0.3728,0.0813}   //工业 	商业	办公	非农人口
	Tour_Rate=100 //每万平方米

	Map= RunMacro("G30 new map",Path_Scen+"Zones\\zones.dbd", False)         //Open Map
	SetMap(Map)
	view=GetView()

	EnableProgressBar("Status", 1)     // Allow only a single progress bar
	CreateProgressBar("Calculating...", "True")
	i_rec=0
	rec = GetFirstRecord(view+"|",)			//Update by Purpose, using Regression Analysis - for both Production and Attraction
	While rec<>null do
		//Progress bar
		i_rec=i_rec+1
		stat = UpdateProgressBar("Please wait...", R2I(i_rec/5))
		if stat = "True" then do
			ShowMessage("You Quit! The macro running is stopped.")
			goto quit_loop
		end

		QuWeiFactor=1
		if view.[QuWei]=1  then QuWeiFactor=1.30
		if view.[QuWei]=2  then QuWeiFactor=1.00
		if view.[QuWei]=3  then QuWeiFactor=0.85

		PFactor=1			//特定小区，人口或工业岗位过大，调整？？后无使用
		//if view.[ID]=1703 or view.[ID]=1721 or view.[ID]=1891   then PFactor=0.1

		for m=1 to CarOwer_Array.Length do   //CA, MC, NC
			for n=1 to Purpose_Array.Length do
				//Production
				CoreName=CarOwer_Array[m]+"_"+Purpose_Array[n]
				view.(CoreName+"_P")=view.[Pop]*(1-view.[Arg])*view.(CarOwer_Array[m])*Rates[m][n]*QuWeiFactor*PFactor
				if view.(CoreName+"_P")=null  then view.(CoreName+"_P")=0

				//Attraction
				if Purpose_Array[n]="HBW"  then view.(CoreName+"_A")=
					((view.[办公岗位]+view.[商业岗位]+view.[工业岗位] )* Coff_HBW[1]+view.[旅游资源] *Tour_Rate) * view.(CarOwer_Array[m]+"_Arr")
				if Purpose_Array[n]="HBS"  then view.(CoreName+"_A")=
					(view.[学位] * Coff_HBS[1]+view.[旅游资源] *Tour_Rate) * view.(CarOwer_Array[m]+"_Arr")
				if Purpose_Array[n]="HBO"  then view.(CoreName+"_A")=
					((view.[商业岗位]* Coff_HBO[1]+view.[办公岗位]* Coff_HBO[2]+view.[Pop]*(1-view.[Arg])* Coff_HBO[3])+view.[旅游资源] *Tour_Rate)* view.(CarOwer_Array[m]+"_Arr")
				if Purpose_Array[n]="NHB"  then view.(CoreName+"_A")=
					((view.[工业岗位]* Coff_NHB[1]+ view.[商业岗位]*Coff_NHB[2]+view.[办公岗位]* Coff_NHB[3] +view.[Pop]*(1-view.[Arg])* Coff_NHB[4] )+view.[旅游资源] *Tour_Rate)* view.(CarOwer_Array[m]+"_Arr")
				if view.(CoreName+"_A")=null  then view.(CoreName+"_A")=0
			end
		end		

		if view.[BanMoto]=1 then do   //中心区禁摩，NC各分得禁摩后的100%
			for n=1 to Purpose_Array.Length do
				//view.("CA_"+Purpose_Array[n]+"_P")=view.("CA_"+Purpose_Array[n]+"_P")+view.("MC_"+Purpose_Array[n]+"_P")/2
				view.("NC_"+Purpose_Array[n]+"_P")=view.("NC_"+Purpose_Array[n]+"_P")+view.("MC_"+Purpose_Array[n]+"_P")
				view.("MC_"+Purpose_Array[n]+"_P")=0
				//view.("CA_"+Purpose_Array[n]+"_A")=view.("CA_"+Purpose_Array[n]+"_A")+view.("MC_"+Purpose_Array[n]+"_A")/2
				view.("NC_"+Purpose_Array[n]+"_A")=view.("NC_"+Purpose_Array[n]+"_A")+view.("MC_"+Purpose_Array[n]+"_A")					
				view.("MC_"+Purpose_Array[n]+"_A")=0
			end
		end
		rec=GetNextRecord(view+"|",,)
	End
	quit_loop:
	DestroyProgressBar()

	//Trip Generation - Balancing, 12对
	for m=1 to CarOwer_Array.Length do
		for n=1 to Purpose_Array.Length do
			V1 = GetDataVector(view+"|",CarOwer_Array[m]+"_"+Purpose_Array[n]+"_P", )
			SUM_P[m][n] = VectorStatistic(V1, "Sum", )
			V2 = GetDataVector(view+"|",CarOwer_Array[m]+"_"+Purpose_Array[n]+"_A", )
			SUM_A[m][n] = VectorStatistic(V2, "Sum", )
			SetDataVector(view+"|",CarOwer_Array[m]+"_"+Purpose_Array[n]+"_A", V2*SUM_P[m][n]/SUM_A[m][n], )
		end
	end
	RunMacro("Close All Maps")
	RunMacro("G30 File Close All")
EndMacro


//#########################################################################################################################################################
//###################### 第二步 交通分布  ##########################################################################################
//#########################################################################################################################################################

Macro "Step2 Trip Distribution"

	RunMacro("Close All Maps")
	RunMacro("G30 File Close All")

	//一次成功，准备出行分布的12个，by carower by purpose: Array
	for m=1 to CarOwer_Array.Length do
		for n=1 to Purpose_Array.Length do
			FF_Array=FF_Array+{null}
			Imp_Array=Imp_Array+{{Path_Scen+"Matrix\\GC.mtx", CarOwer_Array[m]+"_"+Purpose_Array[n], "Origin", "Destination"}}
			KF_Array=KF_Array+{{Path_Scen+"Matrix\\K-zones.mtx", "K", "Rows", "Cols"}}
			UseK_Array=UseK_Array+{1}
			//KF_Array=KF_Array+{null}
			Prod_Array=Prod_Array+{"[Zones]."+CarOwer_Array[m]+"_"+Purpose_Array[n]+"_P"}
			Attr_Array=Attr_Array+{"[Zones]."+CarOwer_Array[m]+"_"+Purpose_Array[n]+"_A"}
			Purp_Array=Purp_Array+{CarOwer_Array[m]+"_"+Purpose_Array[n]}
			Iter_Array=Iter_Array+{10}
			Conv_Array=Conv_Array+{0.01}
			Cons_Array=Cons_Array+{"Doubly"}
			Fric_Array=Fric_Array+{"Gamma"}
			
			v = OpenTable("v", "FFB", { Path_Model+"Parameters\\"+CarOwer_Array[m]+"_"+Purpose_Array[n]+".bin"}, {{"Shared", "True"}})
			rec=GetFirstRecord(v+"|",null)
			A_Array=A_Array+{v.[a]}
			B_Array=B_Array+{v.[b]}
			C_Array=C_Array+{v.[c]}
			closeview(v)
		end
	end
	
	RunMacro("TCB Init")
	Opts = null
	Opts.Input.[PA View Set] = {Path_Scen+"Zones\\zones.DBD|Zones", "Zones"}
	Opts.Input.[FF Matrix Currencies] = FF_Array
	Opts.Input.[Imp Matrix Currencies] =Imp_Array
	Opts.Input.[KF Matrix Currencies] = KF_Array
	Opts.Field.[Prod Fields] = Prod_Array
	Opts.Field.[Attr Fields] = Attr_Array
	Opts.Global.[Purpose Names] = Purp_Array
	Opts.Global.Iterations = Iter_Array
	Opts.Global.Convergence = Conv_Array
	Opts.Global.[Constraint Type] = Cons_Array
	Opts.Global.[Fric Factor Type] = Fric_Array
	Opts.Global.[A List] =A_Array
	Opts.Global.[B List] = B_Array
	Opts.Global.[C List] = C_Array
	Opts.Flag.[Use K Factors] = UseK_Array
	Opts.Output.[Output Matrix].Label = "Gravity Matrix"
	Opts.Output.[Output Matrix].Type = "Float"
	Opts.Output.[Output Matrix].[File based] = "FALSE"
	Opts.Output.[Output Matrix].Sparse = "False"
	Opts.Output.[Output Matrix].[Column Major] = "False"
	Opts.Output.[Output Matrix].Compression = 0
	Opts.Output.[Output Matrix].[File Name] = Path_Scen+"Matrix\\cgrav.mtx"
	ret_value = RunMacro("TCB Run Procedure", "Gravity", Opts, &Ret)
	if !ret_value then Return( RunMacro("TCB Closing", ret_value, True ) )


	mx = OpenMatrix(Path_Scen+"Matrix\\cgrav.mtx", )
	AddMatrixCore(mx, "mySum")
	mcs = CreateMatrixCurrencies(mx, , ,  )
	for m=1 to CarOwer_Array.Length do
		for n=1 to Purpose_Array.Length do
			mcs.[mySum]:=nz(mcs.[mySum])+nz(mcs.(CarOwer_Array[m]+"_"+Purpose_Array[n]))
		end
	end


	//operation = {"Copy", null}
	//mc = CreateMatrixCurrency(mx, "mySum", , , )
	//FillMatrix(mc, null,null, operation, {"Diagonal","True1"})
	//CreateMatrixEditor("cgrav", mx, )
endMacro

//#########################################################################################################################################################
//###################### 第三步 Mode Split  ##########################################################################################
//#########################################################################################################################################################

Macro "Step3 Mode Split"

	RunMacro("Close All Maps")
	RunMacro("G30 File Close All")

	//打开cgrav  来自交通分布
	File_Grav=Path_Scen+"Matrix\\cgrav.mtx"
	mx_Grav=OpenMatrix(File_Grav,)
	mcs_Grav=CreateMatrixCurrencies(mx_Grav, , ,)	

	//打开Mode Split - 目标结果
	File_Mode=Path_Scen+"Matrix\\ModeSplit.mtx"
	mx_Mode=OpenMatrix(File_Mode,)
	mcs_Mode=CreateMatrixCurrencies(mx_Mode, , ,)	
	for i=1 to Mode_Array.length  do
		mcs_Mode.(Mode_Array[i]):=0
	end

	Loop_i=0

	RunMacro("TCB Init")
	// STEP: MNL Evaluation
	for m=1 to CarOwer_Array.Length do
		for n=1 to Purpose_Array.Length do
			Loop_i=Loop_i+1			
			Core_Name=CarOwer_Array[m]+"_"+Purpose_Array[n]	

			Opts = null
			Opts.Input.[View Set] = {Path_Scen+"Zones\\ZONES.DBD|Zones", "Zones"}
			Opts.Input.[Destination Set] = {Path_Scen+"Zones\\ZONES.DBD|Zones", "Zones"}
			Opts.Input.[Model Table] = {Path_Model+"Parameters\\MNL_"+Core_Name+".bin"}
			Opts.Input.[Matrix Currencies] = {{Path_Scen+"Matrix\\GC.mtx", "CA_HBW_Bicycle", "Origin", "Destination"}}   //core无所谓
			Opts.Field.[ID Field] = "Zones.ID"
			Opts.Global.[Number of Modes] = 4
			Opts.Global.[Model Name] = "Result"
			Opts.Flag.Aggregate = 1
			Opts.Flag.[Delete Case] = 1
			Opts.Output.[Output Matrix].Label = "Output Matrix"
			Opts.Output.[Output Matrix].Compression = 1
			Opts.Output.[Output Matrix].[File Name] = Path_Model+"GISDK\\MNL\\MNL_EVAL"+i2s(Loop_i)+".mtx"
			ret_value = RunMacro("TCB Run Procedure", "MNL Evaluation", Opts, &Ret)
			if !ret_value then Return( RunMacro("TCB Closing", ret_value, True ) )
			
			mx_Split=OpenMatrix(Path_Model+"GISDK\\MNL\\MNL_EVAL"+i2s(Loop_i)+".mtx",)
			mcs_Split=CreateMatrixCurrencies(mx_Split, , ,)	

			for i=1 to  Mode_Array.length do
				mcs_Mode.(Core_Name+"_"+Mode_Array[i]):=mcs_Grav.(Core_Name)*mcs_Split.(Mode_Array[i])
				mcs_Mode.(Mode_Array[i]):=mcs_Mode.(Mode_Array[i])+mcs_Mode.(Core_Name+"_"+Mode_Array[i])
			end
			mcs_Split=null
			mx_Split=null

		end	//end for
	end
	
	mcs_Mode=null
	mx_Mode=null
	mcs_Grav=null
	mx_Grav=null	
EndMacro



//#########################################################################################################################################################
//###################### 第四步 PA2OD  ##########################################################################################
//#########################################################################################################################################################

Macro "Step4 PA2OD"
	
	RunMacro("Close All Maps")
	RunMacro("G30 File Close All")

	//定义实载率
	Mode_Occ={1, 1.1, 1.3, 35}   //{"Bicycle","Moto","Car","PT"}   //实载率，再转为PCE，moto=1.1, Car=1.3, PT=35，但PT算人次，不转为车次。
	for m=1 to CarOwer_Array.Length do
		for n=1 to Purpose_Array.Length do
			for i=1 to Mode_Array.Length  do				
			Core_Array=Core_Array+{(m-1)*16+(n-1)*4+i+1}
			Adjust_Array=Adjust_Array+{null}
			Peak_Array=Peak_Array+{null}
			HourlyAB_Array=HourlyAB_Array+{"DEP_"+Purpose_Array[n]}
			HourlyBA_Array=HourlyBA_Array+{"RET_"+Purpose_Array[n]}
			Occ_Array=Occ_Array+{Mode_Occ[i]}
			OccAdj_Array=OccAdj_Array+{"No"}
			PeakF_Array=PeakF_Array+{1}
			Conv_Array=(if Mode_Array[i]="PT"  then Conv_Array+{"No"} else Conv_Array+{"Yes"})		//PT不转
			PHF_Array=PHF_Array+{"No"}
			AdjPeak_Array=AdjPeak_Array+{"No"}		
			end
		end
	end

	RunMacro("TCB Init")
	Opts = null
	Opts.Input.[PA Matrix Currency] = {Path_Scen+"Matrix\\ModeSplit.mtx", "CA_HBW_Car", "Origin", "Destination"}
	Opts.Input.[Lookup Set] = {Path_Model+"Parameters\\HOURLY.bin", "HOURLY"}
	Opts.Field.[Matrix Cores] = Core_Array
	Opts.Field.[Adjust Fields] = Adjust_Array
	Opts.Field.[Peak Hour Field] = Peak_Array
	Opts.Field.[Hourly AB Field] = HourlyAB_Array
	Opts.Field.[Hourly BA Field] = HourlyBA_Array
	Opts.Global.[Method Type] = "PA to OD"
	Opts.Global.[Start Hour] = 18
	Opts.Global.[End Hour] = 18
	Opts.Global.[Cache Size] = 500000
	Opts.Global.[Average Occupancies] = Occ_Array
	Opts.Global.[Adjust Occupancies] =OccAdj_Array
	Opts.Global.[Peak Hour Factor] =PeakF_Array
	Opts.Flag.[Separate Matrices] = "No"
	Opts.Flag.[Convert to Vehicles] = Conv_Array
	Opts.Flag.[Include PHF] =PHF_Array
	Opts.Flag.[Adjust Peak Hour] = AdjPeak_Array
	Opts.Output.[Output Matrix].Label = "PA to OD"
	Opts.Output.[Output Matrix].[File Name] = Path_Scen+"Matrix\\PA2OD.mtx"

	ret_value = RunMacro("TCB Run Procedure", "PA2OD", Opts, &Ret)
	if !ret_value then Return( RunMacro("TCB Closing", ret_value, True ) )
endMacro



//#########################################################################################################################################################
//###################### 第五步 道路分配  ##########################################################################################
//#########################################################################################################################################################
Macro "Step5 Traffic Assignment"

	RunMacro("Close All Maps")
	RunMacro("G30 File Close All")

	Mode_Auto={"Moto","Car"}
	Mode_PCU={0.5,1}		//转为PCU
	//打开PA2OD 矩阵结果
	File_PA2OD=Path_Scen+"Matrix\\PA2OD.mtx"
	mx_PA2OD=OpenMatrix(File_PA2OD,)
	mcs_PA2OD=CreateMatrixCurrencies(mx_PA2OD, , ,)	

	//打开外部矩阵
	File_External=Path_Scen+"Matrix\\External.mtx"
	mx_External=OpenMatrix(File_External,)
	mcs_External=CreateMatrixCurrencies(mx_External, , ,)	

	//打开目标矩阵，开始计算
	File_Auto=Path_Scen+"Matrix\\Assignment Auto OD.mtx"
	mx_Auto=OpenMatrix(File_Auto,)
	mcs_Auto=CreateMatrixCurrencies(mx_Auto, , ,)	

	//对外客运交通和货运交通每年5%增长，加载到External.mtx和Assignment Auto OD.mtx中去了
	//*(1+0.05*(Year-2012))
	mcs_Auto.[PM Auto OD]:=nz(mcs_External.[final-pcu pm])+nz(mcs_Auto.[Truck and Other])
	for m=1 to CarOwer_Array.Length do
		for n=1 to Purpose_Array.Length do
			for i=1 to Mode_Auto.Length  do
				mcs_Auto.[PM Auto OD]:=mcs_Auto.[PM Auto OD]+Mode_PCU[i]*mcs_PA2OD.(CarOwer_Array[m]+"_"+Purpose_Array[n]+"_"+Mode_Auto[i]+" (18-19)")
			end
		end
	end
	
	RunMacro("TCB Init")
	Opts = null
	Opts.Input.Database = Path_Scen+"Network\\Streets.DBD"
	Opts.Input.Network = Path_Scen+"Network\\net.net"
	Opts.Input.[OD Matrix Currency] = {Path_Scen+"Matrix\\Assignment Auto OD.mtx", "PM Auto OD","Cols","Rows"}
	Opts.Field.[VDF Fld Names] = {"[AB_FF_Time / BA_FF_Time]", "[AB_Link_Capacity / BA_Link_Capacity]", "Alpha", "Beta", "Preload"}
	Opts.Global.[Load Method] = "UE"
	Opts.Global.[Loading Multiplier] = 1
	Opts.Global.[Alpha Value] = 0.15
	Opts.Global.[Beta Value] = 4
	Opts.Global.Convergence = 0.01
	Opts.Global.Iterations = 20
	Opts.Global.[Proportional Iterations] = 0
	Opts.Global.[Cost Function File] = "bpr.vdf"
	Opts.Global.[VDF Defaults] = {, , 0.15, 4, 0}
	Opts.Output.[Flow Table] = Path_Scen+"Network\\ASN_LinkFlow.bin"
	ret_value = RunMacro("TCB Run Procedure", "Assignment", Opts, &Ret)
	if !ret_value then Return( RunMacro("TCB Closing", ret_value, True ) )

	RunMacro("Close All Maps")
	RunMacro("G30 File Close All")

EndMacro



//#########################################################################################################################################################
//###################### 第六步 公交分配  ##########################################################################################
//#########################################################################################################################################################
Macro "Step6 Transit Assignment"

	//RunMacro("Close All Maps")
	//RunMacro("G30 File Close All")
	
	Mode_PT={"PT"}

	//打开PA2OD 矩阵结果
	File_PA2OD=Path_Scen+"Matrix\\PA2OD.mtx"
	mx_PA2OD=OpenMatrix(File_PA2OD,)
	mcs_PA2OD=CreateMatrixCurrencies(mx_PA2OD, , ,)	

	//打开目标矩阵，开始计算
	File_PT=Path_Scen+"Matrix\\Assignment PT OD.mtx"
	mx_PT=OpenMatrix(File_PT,)
	mcs_PT=CreateMatrixCurrencies(mx_PT, , ,)	

	//先赋值0	
	mcs_PT.[PM PT OD]:=0
	for m=1 to CarOwer_Array.Length do
		for n=1 to Purpose_Array.Length do
			for i=1 to Mode_PT.Length  do
				mcs_PT.[PM PT OD]:=mcs_PT.[PM PT OD]+mcs_PA2OD.(CarOwer_Array[m]+"_"+Purpose_Array[n]+"_"+Mode_PT[i]+" (18-19)")
			end
		end
	end
	
	RunMacro("TCB Init")
	Opts = null
	Opts.Input.[Transit RS] = Path_Scen+"Network\\BUS_ROUTE.rts"
	Opts.Input.Network =Path_Scen+"Network\\tnw.tnw"
	Opts.Input.[OD Matrix Currency] = {Path_Scen+"Matrix\\Assignment PT OD.mtx", "PM PT OD", "Cols", "Rows"}
	Opts.Global.[OD Layer Type] = 2
	Opts.Output.[Flow Table] = Path_Scen+"Network\\TASN_FLW.bin"
	Opts.Output.[Walk Flow Table] = Path_Scen+"Network\\TASN_WFL.bin"
	Opts.Output.[Aggre Table] = Path_Scen+"Network\\TASN_AGG.bin"
	Opts.Output.[OnOff Table] = Path_Scen+"Network\\TASN_ONO.bin"
	ret_value = RunMacro("TCB Run Procedure", "Transit Assignment PF", Opts, &Ret)
	if !ret_value then Return( RunMacro("TCB Closing", ret_value, True ) )

	mx_PA2OD=null
	mx_PT=null
	RunMacro("Close All Maps")
	RunMacro("G30 File Close All")
EndMacro


//----------------------------------------------- 以下为调用的引用函数 --------------------------------------------------------
//#########################################################################################################################################################
//#########################################################################################################################################################
//#########################################################################################################################################################
//#########################################################################################################################################################
//#########################################################################################################################################################
//#########################################################################################################################################################
//#########################################################################################################################################################
//#########################################################################################################################################################
//#########################################################################################################################################################
//#########################################################################################################################################################
//#########################################################################################################################################################
//**********************************************************************************************************************************************************
//功能模块：更新路网属性
//**********************************************************************************************************************************************************
Macro "UpdateNetworkFF"
	SetMapUnits("Kilometers")		//Sets the Unit to Kilometers, for calculating the length in km.	
	VOT_Car=0.6  //   元/分钟。在高速公路上，转为时间后，相当于速度约下降了一半。
	Price_KM=0.4   //   0.4元/公里

	Map= RunMacro("G30 new map",Path_Scen+"Network\\Streets.DBD", False)         //Open Map
	SetMap(Map)
	view=GetView()

	view_set=view+"|"
	rec=getfirstrecord(view_set,null)
	while rec<>null do 
		if view.[Type]=2  then do   //  次支路
			view.[AB_Link_Capacity]=view.[AB_Lanes] * 300
			view.[BA_Link_Capacity]=view.[BA_Lanes] * 300
			view.[AB_FF_Speed]=20								//km/h
			view.[BA_FF_Speed]=20								//km/h
			view.[Alpha]=2										//
			view.[Beta]=3
		end

		if view.[Type]=3  then do   //  支路  乡道
			view.[AB_Link_Capacity]=view.[AB_Lanes] * 400
			view.[BA_Link_Capacity]=view.[BA_Lanes] * 400
			view.[AB_FF_Speed]=30								//km/h
			view.[BA_FF_Speed]=30								//km/h
			view.[Alpha]=2
			view.[Beta]=3
		end

		if view.[Type]=4  then do   //  次干道
			view.[AB_Link_Capacity]=view.[AB_Lanes] * 500
			view.[BA_Link_Capacity]=view.[BA_Lanes] * 500	
			view.[AB_FF_Speed]=35								//km/h
			view.[BA_FF_Speed]=35								//km/h
			view.[Alpha]=2.2
			view.[Beta]=3.5
		end

		if view.[Type]=5  then do   //    一般主干道
			view.[AB_Link_Capacity]=view.[AB_Lanes] * 600
			view.[BA_Link_Capacity]=view.[BA_Lanes] * 600	
			view.[AB_FF_Speed]=40								//km/h
			view.[BA_FF_Speed]=40								//km/h
			view.[Alpha]=2.2
			view.[Beta]=4
		end

		if view.[Type]=6  then do   //  干线主干道、省道
			view.[AB_Link_Capacity]=view.[AB_Lanes] * 800
			view.[BA_Link_Capacity]=view.[BA_Lanes] * 800		
			view.[AB_FF_Speed]=50								//km/h  原来是50
			view.[BA_FF_Speed]=50								//km/h
			view.[Alpha]=2.5
			view.[Beta]=4
		end

		if view.[Type]=7  then do   //   国道
			view.[AB_Link_Capacity]=view.[AB_Lanes] * 1000
			view.[BA_Link_Capacity]=view.[BA_Lanes] * 1000		
			view.[AB_FF_Speed]=60								//km/h
			view.[BA_FF_Speed]=60								//km/h
			view.[Alpha]=2.7
			view.[Beta]=4.5
		end

		if view.[Type]=71  then do   //   快速路
			view.[AB_Link_Capacity]=view.[AB_Lanes] * 1600
			view.[BA_Link_Capacity]=view.[BA_Lanes] * 1600		
			view.[AB_FF_Speed]=80								//km/h
			view.[BA_FF_Speed]=80								//km/h
			view.[Alpha]=2.7
			view.[Beta]=4.5
		end

		if view.[Type]=8  then do   //   高速公路
			view.[AB_Link_Capacity]=view.[AB_Lanes] * 2000
			view.[BA_Link_Capacity]=view.[BA_Lanes] * 2000		
			view.[AB_FF_Speed]=120								//km/h
			view.[BA_FF_Speed]=120								//km/h
			view.[Toll]=view.[Length]*Price_KM						//0.5元/公里，，Toll单位是元
			view.[Alpha]=2.8
			view.[Beta]=5.5
		end

		if view.[Type]=81  then do   //   高速公路---非深汕高速公路
			view.[AB_Link_Capacity]=view.[AB_Lanes] * 2000
			view.[BA_Link_Capacity]=view.[BA_Lanes] * 2000		
			view.[AB_FF_Speed]=100								//km/h
			view.[BA_FF_Speed]=100								//km/h
			view.[Toll]=view.[Length]*Price_KM						//0.5元/公里，，Toll单位是元
			view.[Alpha]=2.8
			view.[Beta]=5.5
		end

		if view.[Type]=9  then do   //   匝道
			view.[AB_Link_Capacity]=view.[AB_Lanes] * 600
			view.[BA_Link_Capacity]=view.[BA_Lanes] * 600	
			view.[AB_FF_Speed]=20								//km/h
			view.[BA_FF_Speed]=20								//km/h
			view.[Alpha]=2.8
			view.[Beta]=5.5
		end

		if view.[Type]=99  then do                 //中心连接线
			view.[AB_Link_Capacity]=view.[AB_Lanes] * 5000
			view.[BA_Link_Capacity]=view.[BA_Lanes] * 5000	
			view.[AB_FF_Speed]=20								//km/h
			view.[BA_FF_Speed]=20								//km/h
			view.[Alpha]=0.15
			view.[Beta]=4
		end

		if view.[Type]=999  then do                 //外部连接线
			view.[AB_Link_Capacity]=view.[AB_Lanes] * 5000
			view.[BA_Link_Capacity]=view.[BA_Lanes] * 5000	
			view.[AB_FF_Speed]=100								//km/h
			view.[BA_FF_Speed]=100								//km/h
			view.[Alpha]=0.15
			view.[Beta]=4
		end

		view.[AB_FF_Time]=view.[Length]*60 / view.[AB_FF_Speed]+nz(view.[AB_ExtraTime])     //unit: min    转为分钟
		view.[BA_FF_Time]=view.[Length]*60 / view.[BA_FF_Speed]+nz(view.[BA_ExtraTime])      //unit: min    转为分钟

		if view.[Type]=8  then do   //   高速公路
			view.[AB_FF_Time]=view.[AB_FF_Time]+ view.[Length]*Price_KM/VOT_Car   //VOT_Car=1    //1元/分钟，即一个小时60元。
			view.[BA_FF_Time]=view.[BA_FF_Time]+ view.[Length]*Price_KM/VOT_Car   //V2013年1月28日OT_Car=1    //1元/分钟，即一个小时60元。
		end

		if view.[Type]=81  then do   //   高速公路---非深汕高速公路
			view.[AB_FF_Time]=view.[AB_FF_Time]+ view.[Length]*Price_KM*1.2/VOT_Car   //VOT_Car=1    //1元/分钟，即一个小时60元。
			view.[BA_FF_Time]=view.[BA_FF_Time]+ view.[Length]*Price_KM*1.2/VOT_Car   //VOT_Car=1    //1元/分钟，即一个小时60元。
		end

		view.[AB_Ped_Time]=(if  view.[Length]<=1 then view.[Length]*60 / 6 else 10)     // 6km/h
		view.[BA_Ped_Time]=view.[AB_Ped_Time]      //6km  unit: min    转为分钟
		view.[AB_Congested]=view.[AB_FF_Time]
		view.[BA_Congested]=view.[BA_FF_Time]
		view.[AB_PT]=view.[AB_FF_Time]*1.2
		view.[BA_PT]=view.[BA_FF_Time]*1.2

		rec=GetNextRecord(view_set,null,null)
	end

	RunMacro("TCB Init")
	Opts = null
	Opts.Input.[Link Set] = {Path_Scen+"Network\\Streets.DBD|Streets", "Streets"}
	Opts.Global.[Network Options].[Node ID] = "Node.ID"
	Opts.Global.[Network Options].[Link ID] = "Streets.ID"
	Opts.Global.[Network Options].[Turn Penalties] = "Yes"
	Opts.Global.[Network Options].[Keep Duplicate Links] = "FALSE"
	Opts.Global.[Network Options].[Ignore Link Direction] = "FALSE"
	Opts.Global.[Network Options].[Time Unit] = "Minutes"
	Opts.Global.[Link Options] = {{"Length", {"Streets.Length", "Streets.Length", , , "False"}}, {"[AB_Link_Capacity / BA_Link_Capacity]", {"Streets.AB_Link_Capacity", "Streets.BA_Link_Capacity", , , "False"}}, {"[AB_FF_Time / BA_FF_Time]", {"Streets.AB_FF_Time", "Streets.BA_FF_Time", , , "False"}}, {"Alpha", {"Streets.Alpha", "Streets.Alpha", , , "False"}}, {"Beta", {"Streets.Beta", "Streets.Beta", , , "False"}}, {"Preload", {"Streets.Preload", "Streets.Preload", , , "False"}}, {"[AB_Ped_Time / BA_Ped_Time]", {"Streets.AB_Ped_Time", "Streets.BA_Ped_Time", , , "False"}}, {"[AB_Congested / BA_Congested]", {"Streets.AB_Congested", "Streets.BA_Congested", , , "False"}}, {"[AB_PT / BA_PT]", {"Streets.AB_PT", "Streets.BA_PT", , , "False"}}}
	Opts.Global.[Length Unit] = "Kilometers"
	Opts.Global.[Time Unit] = "Minutes"
	Opts.Output.[Network File] = Path_Scen+"Network\\net.net"

	ret_value = RunMacro("TCB Run Operation", "Build Highway Network", Opts, &Ret)
	if !ret_value then Return( RunMacro("TCB Closing", ret_value, True ) )

	RunMacro("Close All Maps")
	RunMacro("G30 File Close All")	
EndMacro
//**********************************************************************************************************************************************************
//功能模块：计算阻抗：先Skim Len和时间以及PT Skim后，再调用CalMat_Congested，进行循环计算
//**********************************************************************************************************************************************************
Macro "UpdateNetworkCongested"
	RunMacro("Close All Maps")
	RunMacro("G30 File Close All")	

	Shared PT_Time_Factor
	if PT_Time_Factor=null  then PT_Time_Factor=1.2

	Map= RunMacro("G30 new map",Path_Scen+"Network\\Streets.DBD", False)         //Open Map
	SetMap(Map)
	view=GetView()	 
	MMA_Flow= OpenTable("MMA_Flow", "FFB", {Path_Scen+"Network\\ASN_LinkFlow.bin"})
	view_Join=JoinViews("view_Join", view+".ID", MMA_Flow+".ID1",)
	SetView(view_Join)

	rec=GetFirstRecord(view_Join+"|",null)
	while rec<>null do 
		view_Join.[AB_Congested]=(if view_Join.[AB_Time]<30  then view_Join.[AB_Time] else 30)
		view_Join.[BA_Congested]=(if view_Join.[BA_Time]<30  then view_Join.[BA_Time] else 30)
		view_Join.[AB_PT]=view_Join.[AB_Congested]*PT_Time_Factor
		view_Join.[BA_PT]=view_Join.[BA_Congested]*PT_Time_Factor
		rec=GetNextRecord(view_Join+"|",null,null)
	end
	CloseView(view_Join)
	CloseView(MMA_Flow)	

	//Create Net
	RunMacro("TCB Init")
	Opts = null
	Opts.Input.[Link Set] = {Path_Scen+"Network\\Streets.DBD|Streets", "Streets"}
	Opts.Global.[Network Options].[Node ID] = "Node.ID"
	Opts.Global.[Network Options].[Link ID] = "Streets.ID"
	Opts.Global.[Network Options].[Turn Penalties] = "Yes"
	Opts.Global.[Network Options].[Keep Duplicate Links] = "FALSE"
	Opts.Global.[Network Options].[Ignore Link Direction] = "FALSE"
	Opts.Global.[Network Options].[Time Unit] = "Minutes"
	Opts.Global.[Link Options] = {{"Length", {"Streets.Length", "Streets.Length", , , "False"}}, {"[AB_Link_Capacity / BA_Link_Capacity]", {"Streets.AB_Link_Capacity", "Streets.BA_Link_Capacity", , , "False"}}, {"[AB_FF_Time / BA_FF_Time]", {"Streets.AB_FF_Time", "Streets.BA_FF_Time", , , "False"}}, {"Alpha", {"Streets.Alpha", "Streets.Alpha", , , "False"}}, {"Beta", {"Streets.Beta", "Streets.Beta", , , "False"}}, {"Preload", {"Streets.Preload", "Streets.Preload", , , "False"}}, {"[AB_Ped_Time / BA_Ped_Time]", {"Streets.AB_Ped_Time", "Streets.BA_Ped_Time", , , "False"}}, {"[AB_Congested / BA_Congested]", {"Streets.AB_Congested", "Streets.BA_Congested", , , "False"}}, {"[AB_PT / BA_PT]", {"Streets.AB_PT", "Streets.BA_PT", , , "False"}}}
	Opts.Global.[Length Unit] = "Kilometers"
	Opts.Global.[Time Unit] = "Minutes"
	Opts.Output.[Network File] = Path_Scen+"Network\\net.net"
	ret_value = RunMacro("TCB Run Operation", "Build Highway Network", Opts, &Ret)
	if !ret_value then Return( RunMacro("TCB Closing", ret_value, True ) )

	RunMacro("Close All Maps")
	RunMacro("G30 File Close All")	
EndMacro

Macro "Skim_Len_Time_PT"
	RunMacro("TCB Init")
	//Skim Len
	Opts = null
	Opts.Input.Network = Path_Scen+"Network\\net.net"
	Opts.Input.[Origin Set] = {Path_Scen+"Network\\Streets.DBD|Node", "Node", "Centroids", "Select * where Centroid<>null"}
	Opts.Input.[Destination Set] = {Path_Scen+"Network\\Streets.DBD|Node", "Node", "Centroids"}
	Opts.Input.[Via Set] = {Path_Scen+"Network\\Streets.DBD|Node", "Node"}
	Opts.Field.Minimize = "Length"
	Opts.Field.Nodes = "Node.ID"
	Opts.Output.[Output Matrix].Label = "Shortest Path"
	Opts.Output.[Output Matrix].[File Name] = Path_Scen+"Matrix\\SPMAT-Len.mtx"
	ret_value = RunMacro("TCB Run Procedure", "TCSPMAT", Opts, &Ret)
	if !ret_value then Return( RunMacro("TCB Closing", ret_value, True ) )

	//Skim Time
	Opts = null
	Opts.Input.Network = Path_Scen+"Network\\net.net"
	Opts.Input.[Origin Set] = {Path_Scen+"Network\\Streets.DBD|Node", "Node", "Centroids", "Select * where Centroid<>null"}
	Opts.Input.[Destination Set] = {Path_Scen+"Network\\Streets.DBD|Node", "Node", "Centroids"}
	Opts.Input.[Via Set] = {Path_Scen+"Network\\Streets.DBD|Node", "Node"}
	Opts.Field.Minimize = "[AB_Congested / BA_Congested]"
	Opts.Field.Nodes = "Node.ID"
	Opts.Output.[Output Matrix].Label = "Shortest Path"
	Opts.Output.[Output Matrix].[File Name] = Path_Scen+"Matrix\\SPMAT-Time.mtx"
	ret_value = RunMacro("TCB Run Procedure", "TCSPMAT", Opts, &Ret)
	if !ret_value then Return( RunMacro("TCB Closing", ret_value, True ) )

	//Skim PT
	//Skim PT - 1. Create TNW
	Opts = null
	Opts.Input.[Transit RS] = Path_Scen+"Network\\BUS_ROUTE.rts"
	Opts.Input.[RS Set] = {Path_Scen+"Network\\BUS_ROUTE.rts|Route System", "Route System"}
	Opts.Input.[Walk Set] = {Path_Scen+"Network\\Streets.DBD|Streets", "Streets"}
	Opts.Input.[Stop Set] = {Path_Scen+"Network\\BUS_ROUTES.DBD|Route Stops", "Route Stops"}
	Opts.Global.[Network Label] = "Based on 'Route System' "+GetDateandTime()
	Opts.Global.[Network Options].Walk = "Yes"
	Opts.Global.[Network Options].[Link Attributes] = {{"Length", {"Streets.Length", "Streets.Length"}, "SUMFRAC"}, {"[AB_Ped_Time / BA_Ped_Time]", {"Streets.AB_Ped_Time", "Streets.BA_Ped_Time"}, "SUMFRAC"}, {"[AB_PT / BA_PT]", {"Streets.AB_PT", "Streets.BA_PT"}, "SUMFRAC"}}
	Opts.Global.[Network Options].[Street Attributes] = {{"Length", {"Streets.Length", "Streets.Length"}}, {"[AB_Ped_Time / BA_Ped_Time]", {"Streets.AB_Ped_Time", "Streets.BA_Ped_Time"}}, {"[AB_PT / BA_PT]", {"Streets.AB_Ped_Time", "Streets.BA_Ped_Time"}}}
	Opts.Global.[Network Options].[Route Attributes].Price = {"[Route System].Price"}
	Opts.Global.[Network Options].[Route Attributes].Xfer_Price = {"[Route System].Xfer_Price"}
	Opts.Global.[Network Options].[Route Attributes].Headway = {"[Route System].Headway"}
	Opts.Global.[Network Options].[Route Attributes].Speed = {"[Route System].Speed"}
	Opts.Global.[Network Options].TagField = "NODEID"
	Opts.Global.[Network Options].[Merge Stops] = {"[Route Stops].ID", "Route Stops.NODEID"}
	Opts.Output.[Network File] = Path_Scen+"Network\\tnw.tnw"
	ret_value = RunMacro("TCB Run Operation", "Build Transit Network", Opts, &Ret)
	if !ret_value then Return( RunMacro("TCB Closing", ret_value, True ) )

	//Skim PT - 2. TNW Settings
	Opts = null
	Opts.Input.[Transit RS] =Path_Scen+"Network\\BUS_ROUTE.rts"
	Opts.Input.[Transit Network] = Path_Scen+"Network\\tnw.tnw"

	Opts.Field.[Link Impedance] = "[AB_PT / BA_PT]"
	Opts.Field.[Link Impedance] = "[AB_PT / BA_PT]"
	Opts.Field.[Route Fare] = "Price"
	Opts.Field.[Route Xfer Fare] = "Xfer_Price"
	Opts.Field.[Route Headway] = "Headway"

	Opts.Global.[Global Fare Value] = 2
	Opts.Global.[Global Xfer Fare] = 0.4			//转换到该模式的任何路线的折扣车费（如果有的话） 

	Opts.Global.[Global Fare Weight] = 1
	Opts.Global.[Global Imp Weight] = 1
	Opts.Global.[Global Xfer Weight] = 1
	Opts.Global.[Global IWait Weight] = 1.5
	Opts.Global.[Global XWait Weight] = 1.5
	Opts.Global.[Global Dwell Weight] = 0
	Opts.Global.[Global Dwell Time] = 0
	Opts.Global.[Global Headway] = 10
	Opts.Global.[Global Xfer Time] = 3
	Opts.Global.[Global Max IWait] = 30
	Opts.Global.[Global Min IWait] = 2
	Opts.Global.[Global Max XWait] = 30
	Opts.Global.[Global Min XWait] = 2
	Opts.Global.[Global Layover Time] = 5
	Opts.Global.[Global Max WACC Path] = 4
	Opts.Global.[Global Max Access] = 20
	Opts.Global.[Global Max Egress] = 20
	Opts.Global.[Global Max Transfer] = 5
	Opts.Global.[Global Max Imp] = 240
	Opts.Global.[Path Method] = 3
	Opts.Global.[Value of Time] = 0.1826
	Opts.Global.[Max Xfer Number] = 5
	Opts.Global.[Max Trip Time] = 500
	Opts.Global.[Walk Weight] = 1.5
	Opts.Global.[Zonal Fare Method] = 1
	Opts.Global.[Interarrival Para] = 0.5
	Opts.Global.[Path Threshold] = 1
	Opts.Flag.[Use All Walk Path] = "No"
	Opts.Flag.[Use Stop Access] = "No"
	Opts.Flag.[Use Mode] = "No"
	Opts.Flag.[Use Mode Cost] = "No"
	Opts.Flag.[Combine By Mode] = "Yes"
	Opts.Flag.[Fare By Mode] = "No"
	Opts.Flag.[M2M Fare Method] = 2
	Opts.Flag.[Fare System] = 1
	Opts.Flag.[Use Park and Ride] = "No"
	Opts.Flag.[Use P&R Walk Access] = "No"
	ret_value = RunMacro("TCB Run Operation", "Transit Network Setting PF", Opts, &Ret)
	if !ret_value then Return( RunMacro("TCB Closing", ret_value, True ) )

	//Skim PT - 3. Skim General Cost
	Opts = null
	Opts.Input.Database = Path_Scen+"Network\\Streets.DBD"
	Opts.Input.Network = Path_Scen+"NETWORK\\TNW.TNW"
	Opts.Input.[Origin Set] = {Path_Scen+"Network\\Streets.DBD|Node", "Node", "Centroids", "Select * where Centroid<>null"}
	Opts.Input.[Destination Set] = {Path_Scen+"Network\\Streets.DBD|Node", "Node", "Centroids"}
	Opts.Global.[Skim Var] = {"Generalized Cost"}
	Opts.Global.[OD Layer Type] = 2
	Opts.Output.[Skim Matrix].Label = "Skim Matrix (Pathfinder)"
	Opts.Output.[Skim Matrix].[File Name] = Path_Scen+"Matrix\\SPMAT-PT.mtx"
	ret_value = RunMacro("TCB Run Procedure", "Transit Skim PF", Opts, &Ret)
	if !ret_value then Return( RunMacro("TCB Closing", ret_value, True ) )

EndMacro

//GC 计算，进行循环计算
Macro "CalMat_Congested"

	Shared Mode_Split
	Shared PT_Time_Factor

	if Mode_Split=null  then  Mode_Split={0.309,0.507,0.123,0.061}		//原始调查值
	VOT={{21.63,11.65,16.64,19.14},{10.82,5.82,8.32,9.57},{10.82,5.82,8.32,9.57}}  //来自Excel的Settings.xlsx
	VOT_avg=10.95   // 元/小时

	if GetFileInfo(Path_Scen+"Variables.ini")=null  then do
		ShowMessage("There is not any 'variables.ini' setting file at: "+Path_Scen)
		return()
	end
	Arg=null				//Read the Global Variables to compound array, first clomn is name and second one is value.
	fptr = OpenFile(Path_Scen+"Variables.ini", "r")
	while !FileAtEOF(fptr) do 
		String_Line=Trim(ReadLine(fptr))
		if len(String_Line)>3 and left(String_Line,1)<>";"  and Position(left(String_Line,5),"//")=0 then do
			Array_Row=ParseString(String_Line,"=",)
			Arg=Arg+{Array_Row}		//e.g. { {"Project_Name","Riyadh Traffic Model - RTM"} }   //first clomn is name and second one is value.
		end
	end		
	CloseFile(fptr)

	Bicycle_OVT=S2R(Arg.Bicycle_OVT)
	Moto_Petro=S2R(Arg.Moto_Petro)
	Moto_Parking=S2R(Arg.Moto_Parking)
	Moto_OVT=S2R(Arg.Moto_OVT)
	Moto_Time_Factor=S2R(Arg.Moto_Time_Factor)
	Car_Petro=S2R(Arg.Car_Petro)
	Car_Parking=S2R(Arg.Car_Parking)
	Car_OVT=S2R(Arg.Car_OVT)
	PT_Fare=S2R(Arg.PT_Fare)
	PT_Time_Factor=S2R(Arg.PT_Time_Factor)
	PT_Head=S2R(Arg.PT_Head)
	PT_Wait=S2R(Arg.PT_Wait)
	PT_WalkTime=S2R(Arg.PT_WalkTime)
	IVT_Factor=S2R(Arg.IVT_Factor)
	OVT_Factor=S2R(Arg.OVT_Factor)

	PT_GC_Factor=S2R(Arg.PT_GC_Factor)
	Auto_GC_Factor=S2R(Arg.Auto_GC_Factor)

	mx=OpenMatrix(Path_Scen+"Matrix\\GC.mtx",)
	//Cores=GetMatrixCoreNames(mx)
	//mcs = CreateMatrixCurrencies(mx, , ,)						//mcs is a Array

	//打开Skim Length
	File_Len=Path_Scen+"Matrix\\SPMAT-Len.mtx"
	CoreLen="Shortest Path - Length"
	//RunMacro("Calc_Intrazonal",File_Len,CoreLen)
	mx_Len=OpenMatrix(File_Len,)
	mcs_Len=CreateMatrixCurrencies(mx_Len, , ,)	

	//打开Skim Time
	File_Time=Path_Scen+"Matrix\\SPMAT-Time.mtx"
	CoreTime="Shortest Path - [AB_Congested / BA_Congested]"
	//RunMacro("Calc_Intrazonal",File_Time,CoreTime)
	mx_Time=OpenMatrix(File_Time,)
	mcs_Time=CreateMatrixCurrencies(mx_Time, , ,)	

	//打开PT Skim矩阵
	File_PT=Path_Scen+"Matrix\\SPMAT-PT.mtx"
	CorePT="Generalized Cost"
	//RunMacro("Calc_Intrazonal",File_PT,CorePT)
	mx_PT=OpenMatrix(File_PT,)
	mcs_PT=CreateMatrixCurrencies(mx_PT, , ,)
	//mcs_PT.(CorePT):=mcs_PT.(CorePT)*60/VOT_avg	

	for m=1 to CarOwer_Array.Length do
		for n=1 to Purpose_Array.Length do
				for i=1 to Mode_Array.Length  do
					new_CoreName=CarOwer_Array[m]+"_"+Purpose_Array[n]+"_"+Mode_Array[i]
					//AddMatrixCore(mx, new_CoreName)
					mcs = CreateMatrixCurrencies(mx, , ,)	

					//计算 by CarOwer by Purpose by Mode的阻抗成本
					
					//1. Bicycle   speed=12km/h   乘上60分钟，得到与长度的关系是5倍
					//<30min，广义阻抗为1倍，30-60广义阻抗为1.5倍，大于60min广义阻抗为2倍。
					if i=1  then do
						mcs.(new_CoreName):=(if mcs_Len.(CoreLen)<=6  then mcs_Len.(CoreLen)/12*60+Bicycle_OVT*OVT_Factor
											    else if  mcs_Len.(CoreLen)>=6 and mcs_Len.(CoreLen)<9 then mcs_Len.(CoreLen)/12*60*1.5+Bicycle_OVT*OVT_Factor
											    else mcs_Len.(CoreLen)/12*60*2+Bicycle_OVT*OVT_Factor)
					end

					//2. Moto  GC=IVT+OVT*Factor + Fee/OVT + GC.[BanMoto]  //摩托车比小汽车慢 1.1倍
					//2012年1月1日起，禁止所有悬挂特区以外号牌的摩托车（含澄海区、潮阳区、潮南区和南澳县）进入城市中心区域行驶
					if i=2  then do  
						mcs.(new_CoreName):= IVT_Factor*mcs_Time.(CoreTime)*Moto_Time_Factor+OVT_Factor*Moto_OVT+(mcs_Len.(CoreLen)*Moto_Petro+Moto_Parking)*60/VOT[m][n]
						mcs.(new_CoreName):=mcs.(new_CoreName)*Auto_GC_Factor
					end

					//3. Car  GC=IVT+OVT*Factor + Fee/OVT 
					if i=3  then do  
						mcs.(new_CoreName):= IVT_Factor*mcs_Time.(CoreTime)+OVT_Factor*Car_OVT+(mcs_Len.(CoreLen)*Car_Petro+Car_Parking)*60/VOT[m][n]
						mcs.(new_CoreName):=mcs.(new_CoreName)*Auto_GC_Factor
					end

					//PT要单独算-用PT Skim来得到

					if i=4  then do  
						mcs.(new_CoreName):= mcs_PT.(CorePT)*60/VOT[m][n]
						mcs.(new_CoreName):=mcs.(new_CoreName)*PT_GC_Factor
					end

					RunMacro("Calc_Intrazonal",Path_Scen+"Matrix\\GC.mtx",new_CoreName)

					//如果禁摩，则把禁摩部分清空，在方式划分时，则不考虑这块
					if i=2  then mcs.(new_CoreName):=mcs.(new_CoreName)+mcs.[BanMoto]
				end

				//综合阻抗，按方式划分之权重
				new_CoreName=CarOwer_Array[m]+"_"+Purpose_Array[n]
				//AddMatrixCore(mx, new_CoreName)
				mcs = CreateMatrixCurrencies(mx, , ,)	
				/*
				//按交通小区的比例来划分权重
				mcs.(new_CoreName):=Nz(mcs.(CarOwer_Array[m]+"_"+Purpose_Array[n]+"_"+Mode_Array[1]))*mcs_Split.(Mode_Array[1]+"_Split")+
				Nz(mcs.(CarOwer_Array[m]+"_"+Purpose_Array[n]+"_"+Mode_Array[2]))*mcs_Split.(Mode_Array[2]+"_Split")+
				Nz(mcs.(CarOwer_Array[m]+"_"+Purpose_Array[n]+"_"+Mode_Array[3]))*mcs_Split.(Mode_Array[3]+"_Split")+
				Nz(mcs.(CarOwer_Array[m]+"_"+Purpose_Array[n]+"_"+Mode_Array[4]))*mcs_Split.(Mode_Array[4]+"_Split")
				*/
				//按全市分担率来划分权重
				mcs.(new_CoreName):=Nz(mcs.(CarOwer_Array[m]+"_"+Purpose_Array[n]+"_"+Mode_Array[1]))*Mode_Split[1]+
				Nz(mcs.(CarOwer_Array[m]+"_"+Purpose_Array[n]+"_"+Mode_Array[2]))*Mode_Split[2]+
				Nz(mcs.(CarOwer_Array[m]+"_"+Purpose_Array[n]+"_"+Mode_Array[3]))*Mode_Split[3]+
				Nz(mcs.(CarOwer_Array[m]+"_"+Purpose_Array[n]+"_"+Mode_Array[4]))*Mode_Split[4]

				/*
				//2012年1月1日起，禁止所有悬挂特区以外号牌的摩托车（含澄海区、潮阳区、潮南区和南澳县）进入城市中心区域行驶
				for m=1 to CarOwer_Array.Length do
					for n=1 to Purpose_Array.Length do
						for i=2 to 2  do   //Moto
							new_CoreName=CarOwer_Array[m]+"_"+Purpose_Array[n]+"_"+Mode_Array[i]
							mcs.(new_CoreName):=mcs.(new_CoreName)*mcs.[BanMoto]
						end
					end
				end
				*/
		end
	end

	mx_Len=null
	mx_Time=null
	mx_PT=null
	RunMacro("Close All Maps")
	RunMacro("G30 File Close All")
EndMacro




//------------------------------------------------------------------------------------------------------------------------------
//------------  Calc_Intrazonal ----------------------------------------------------------------------------------------------
//------------------------------------------------------------------------------------------------------------------------------
Macro "Calc_Intrazonal"(File,Core)
	RunMacro("TCB Init")
	Opts = null
	Opts.Input.[Matrix Currency] = {File, Core, ,}
	Opts.Global.Factor = 0.45
	Opts.Global.Neighbors = 3
	Opts.Global.Operation = 1
	Opts.Global.[Treat Missing] = 1
	ret_value = RunMacro("TCB Run Procedure", "Intrazonal", Opts, &Ret)
	if !ret_value then Return( RunMacro("TCB Closing", ret_value, True ) )
EndMacro
//------------------------------------------------------------------------------------------------------------------------------
//------------  Close All Maps ----------------------------------------------------------------------------------------------
//------------------------------------------------------------------------------------------------------------------------------
Macro "Close All Maps"
	on error,Notfound goto NEXT
	maps = GetMapNames()
	if maps<>null then do
		for i = 1 to maps.length do
			CloseMap(maps[i])
		end
	end
	NEXT:
EndMacro

