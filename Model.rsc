// Project:
// Project Manager: 
// Project Team members: Jandy,
//GISDK Author: Jandy
//Created Data: 2011-
// Purpose:1. 
//                2. 

Macro "Model"	
	Global Path_Model, Path_Scen,Year,CarOwer_Array,Purpose_Array,Mode_Array

	//Path_Model="D:\\P13 - ��ͷ�н�ͨģ��\\Model\\"
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

	if GetFileInfo(Path_Model+"GISDK\\model.dbd")=null then do showMessage("û���ҵ�ģ���ļ�ϵͳ.") return() end

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
	MenuItem "Top_Menu" text:"��ͷģ��" do	 RunDbox("dbox_Interface")	EndItem
EndMenu

//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Dbox "dbox_Interface"Title: "��ݹ���ģ�ͽ��� - " + Project_Name  Toolbox NoKeyboard
init do
	On NotFound goto NEXT
		CloseDbox("dbox_Interface")
	NEXT:

	Project_Name="��ɽ��ͨģ��"
	Macro_Version="v1.0"

	shared d_matrix_options // default options used for creating matrix editors
	shared Scen_Name
	pt_sz=1
	Scenario_idx=1
	
       About_Us= Project_Name + "  -  " + Macro_Version + 
			"\n\n\n ��ݹ���ģ��ϵͳ�Ĺ��ܰ����� " +
			"\n 1.С�����Ա༭." +
			"\n 2.·�����Ա༭." +
			"\n 3.�������Ľ׶ι���."+
			"\n 4.���з���ѭ��ģ��."

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
	Path_Scen=Path_Model+Array_Scenario_Folder[Scenario_idx]+"\\"   //PathΪĳһ�龰���ļ���

	if GetDirectoryinfo(Path_Scen+"*","Directory")=null  then do 
		ShowMessage("�����������󣬻򷽰��ļ��в����ڡ�")   
		Scenario_idx=1 
		Year=Array_Scenario_Year[Scenario_idx]
		Path_Scen=Path_Model+Array_Scenario_Folder[Scenario_idx]+"\\"   //PathΪĳһ�龰���ļ���
	end
	Scen_Name=Array_Scenario_Folder[Scenario_idx]
EndItem

Tab List 0.5, 0,83, 25 variable: Tab_idx                         // Tab list definition starts here.  Tab_idx to definiton which tab is selected.
//---First tab -- The following are the left hand icons where a user clicks and a folder opens.
Tab prompt: "����ģ��"
	
// The following lines picks the bmp picture used for the front page of  the mian user interface
	
	Checkbox 0, 0 icons: Path_Model+"GISDK\\Front_Logo.bmp" Help:"Model Path: "+Path_Model variable: pt_sz    //the Left pic, Riyadh Traffic Model
	do
		if GetFileInfo(Path_Model+"GISDK\\User's Guide.pdf")  then LaunchDocument(Path_Model+"GISDK\\User's Guide.pdf",)
	endItem

	Text 46,1 variable:"ѡ�񷽰���"
	Popdown Menu same,after, 25 list: Array_Scenario_Folder variable: Scenario_idx
	do
		Year=Array_Scenario_Year[Scenario_idx]
		Path_Scen=Path_Model+Array_Scenario_Folder[Scenario_idx]+"\\"   //PathΪĳһ�龰���ļ���
		if GetDirectoryinfo(Path_Scen+"*","Directory")=null  then do ShowMessage("�����������󣬻򷽰��ļ��в����ڡ�")   Scenario_idx=1 end
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

	Button "�鿴" after,same,7 Help:"View this Scenario"  //button, View Scenarios
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
		CreateMatrixEditor("��ͨ�ֲ�", mx, d_matrix_options)
	enditem
	Button "icon_btn3" same, after icons: "bmp\\planmodesplit.bmp" 
	do
		mx=OpenMatrix(Path_Scen+"Matrix\\ModeSplit.mtx",)
		CreateMatrixEditor("��ʽ����", mx, d_matrix_options)
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

	 Button "��ͨ����" after,6,20,1.7					//button, Trip Generation
	   do
		RunMacro("Close All Maps")
		RunMacro("G30 File Close All")
		RunMacro("Step1 Trip Generation")
		TAZ_Map=Path_Scen+"Zones\\Result.map"
		if GetFileInfo(TAZ_Map) then OpenMap(TAZ_Map,)
	EndItem
	   
	 Button "��ͨ�ֲ�" same,after,20,1.7					//button, Trip Distribution
	   do
		RunMacro("Close All Maps")
		RunMacro("G30 File Close All")
		RunMacro("Step2 Trip Distribution")
		mx=OpenMatrix(Path_Scen+"Matrix\\cgrav.mtx",)
		CreateMatrixEditor("��ͨ�ֲ�", mx, d_matrix_options)
	   EndItem
	   
         Button "��ʽ����" same, after ,20,1.7					//button,  Mode Split
	   do
		RunMacro("Close All Maps")
		RunMacro("G30 File Close All")
		RunMacro("Step3 Mode Split")
		mx=OpenMatrix(Path_Scen+"Matrix\\ModeSplit.mtx",)
		CreateMatrixEditor("��ʽ����", mx, d_matrix_options)
	   EndItem
	   
         Button "PA ת OD" same, after ,20,1.7						//button,  PA to OD
	   do
		RunMacro("Close All Maps")
		RunMacro("G30 File Close All")
		RunMacro("Step4 PA2OD")
		mx=OpenMatrix(Path_Scen+"Matrix\\PA2OD.mtx",)
		CreateMatrixEditor("PA2OD", mx, d_matrix_options)
	   EndItem
	   
         Button "��·����" same, after ,20,1.7					//button,  Trip Assignment
	   do
		RunMacro("Close All Maps")
		RunMacro("G30 File Close All")
		RunMacro("Step5 Traffic Assignment")
		Network_Map=Path_Scen+"Network\\Result.map"
		if GetFileInfo(Network_Map) then OpenMap(Network_Map,)
	   EndItem
	   
	Button "��������" same, after,20,1.7		 		//button,  Select Link Analysis
	do
		RunMacro("Close All Maps")
		RunMacro("G30 File Close All")
		RunMacro("Step6 Transit Assignment")
		Network_Map=Path_Scen+"Network\\Result-PT.map"
		if GetFileInfo(Network_Map) then OpenMap(Network_Map,)
	EndItem
	
	 Button "���з����Ľ׶�" same, after,20,1.7					//button,  Sub Area Analysis
	do
		btn = MessageBox("��ȷ��Ҫ���з���ѭ�����Ľ׶�ģ����\n �������������ò�ͬ��ģ�����д�Լ��Ҫ10���ӡ�",
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
Tab prompt: "����"								//About us

	Text 1, after
	Text same, after, 42,11 Framed Variable:About_Us

	Text 2, 15 variable:"���ӷ������ܣ�"

	  Button "��������·��" 2,17,20,1.7 Help:"������ӵ�µ�"					//button, Trip Generation
	   do
		RunMacro("Close All Maps")
		RunMacro("G30 File Close All")
		RunMacro("UpdateNetworkFF")
	EndItem
	 Button "����ӵ��·��" after,same,20,1.7 Help:"����ӵ�µģ������ڷ�������"					//button, Trip Generation
	   do
		RunMacro("Close All Maps")
		RunMacro("G30 File Close All")
		RunMacro("UpdateNetworkCongested")
	EndItem
	Button "��ȡ�迹����" 2,after,20,1.7 Help:"Len,Time and PT"					//button, Trip Generation
	   do
		RunMacro("Close All Maps")
		RunMacro("G30 File Close All")
		RunMacro("Skim_Len_Time_PT")
	EndItem
	Button "�����ۺ��迹" after,same,20,1.7 Help:"GC"					//button, Trip Generation
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
	
	Mode_Split={0.309,0.507,0.123,0.061}		//ԭʼ����ֵ
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
			SetStatus(1,Scen_Name+":Loop="+I2S(Feedback_Loop)+"(Total="+ I2S(Max_Loop) +").���㽻ͨС�������½�ͨ����",)
			RunMacro("Step1 Trip Generation")

			SetStatus(1,Scen_Name+":Loop="+I2S(Feedback_Loop)+"(Total="+ I2S(Max_Loop) +").���µ�·����",)
			RunMacro("UpdateNetworkFF")   //When Feedback_Loop=1, using the FF time, else, using the congested time
		 end
		
		if Feedback_Loop>1  then do
			SetStatus(1,Scen_Name+":Loop="+I2S(Feedback_Loop)+"(Total="+ I2S(Max_Loop) +").�����迹ʱ��",)
			RunMacro("UpdateNetworkCongested")
		end
		
		SetStatus(1,Scen_Name+":Loop="+I2S(Feedback_Loop)+"(Total="+ I2S(Max_Loop) +").�����迹����",)
		RunMacro("Skim_Len_Time_PT")		//Len Time PT
		RunMacro("CalMat_Congested")		//GC

		SetStatus(1,Scen_Name+":Loop="+I2S(Feedback_Loop)+"(Total="+ I2S(Max_Loop) +"). ��ͨ�ֲ�",)
		RunMacro("Step2 Trip Distribution")

		SetStatus(1,Scen_Name+":Loop="+I2S(Feedback_Loop)+"(Total="+ I2S(Max_Loop) +").��ʽ���֣�����ֵ���",)
		RunMacro("Step3 Mode Split")
		
		//��ÿ��ѭ��ʱ��Mode_Split Ӧ����
		//��Mode Split
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

		SetStatus(1,Scen_Name+":Loop="+I2S(Feedback_Loop)+"(Total="+ I2S(Max_Loop) +").ִ�� PA2OD.",)
		RunMacro("Step4 PA2OD")

		SetStatus(1,Scen_Name+":Loop="+I2S(Feedback_Loop)+"(Total="+ I2S(Max_Loop) +").��·����.",)
		RunMacro("Step5 Traffic Assignment")

		 if Feedback_Loop=Max_Loop  then do
			SetStatus(1,Scen_Name+":Loop="+I2S(Feedback_Loop)+"(Total="+ I2S(Max_Loop) +").��������.",)
			RunMacro("Step6 Transit Assignment")
		end
	end

	RunMacro("UpdateNetworkCongested")
	
	CloseFile(fptr)
	//Mesure Time
	Time_End = GetDateAndTime()
	ShowMessage(Scen_Name+":ģ�Ϳ�ʼ��: "+Time_Start+" \n\n �ɹ�������: "+Time_End)
	SetStatus(2,"@system1",)
	quit:
	return()
EndMacro











//#########################################################################################################################################################
//###################### ��һ�� ��ͨ����  ##########################################################################################
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

	Coff_HBW={2.373}  //�ܸ�λ
	Coff_HBS={0.8675}  //ѧλ
	Coff_HBO={0.378,0.142,0.400} //��ҵ	�칫	��ũ�˿�
	Coff_NHB={0.1555,0.3758,0.3728,0.0813}   //��ҵ 	��ҵ	�칫	��ũ�˿�
	Tour_Rate=100 //ÿ��ƽ����

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

		PFactor=1			//�ض�С�����˿ڻ�ҵ��λ���󣬵�����������ʹ��
		//if view.[ID]=1703 or view.[ID]=1721 or view.[ID]=1891   then PFactor=0.1

		for m=1 to CarOwer_Array.Length do   //CA, MC, NC
			for n=1 to Purpose_Array.Length do
				//Production
				CoreName=CarOwer_Array[m]+"_"+Purpose_Array[n]
				view.(CoreName+"_P")=view.[Pop]*(1-view.[Arg])*view.(CarOwer_Array[m])*Rates[m][n]*QuWeiFactor*PFactor
				if view.(CoreName+"_P")=null  then view.(CoreName+"_P")=0

				//Attraction
				if Purpose_Array[n]="HBW"  then view.(CoreName+"_A")=
					((view.[�칫��λ]+view.[��ҵ��λ]+view.[��ҵ��λ] )* Coff_HBW[1]+view.[������Դ] *Tour_Rate) * view.(CarOwer_Array[m]+"_Arr")
				if Purpose_Array[n]="HBS"  then view.(CoreName+"_A")=
					(view.[ѧλ] * Coff_HBS[1]+view.[������Դ] *Tour_Rate) * view.(CarOwer_Array[m]+"_Arr")
				if Purpose_Array[n]="HBO"  then view.(CoreName+"_A")=
					((view.[��ҵ��λ]* Coff_HBO[1]+view.[�칫��λ]* Coff_HBO[2]+view.[Pop]*(1-view.[Arg])* Coff_HBO[3])+view.[������Դ] *Tour_Rate)* view.(CarOwer_Array[m]+"_Arr")
				if Purpose_Array[n]="NHB"  then view.(CoreName+"_A")=
					((view.[��ҵ��λ]* Coff_NHB[1]+ view.[��ҵ��λ]*Coff_NHB[2]+view.[�칫��λ]* Coff_NHB[3] +view.[Pop]*(1-view.[Arg])* Coff_NHB[4] )+view.[������Դ] *Tour_Rate)* view.(CarOwer_Array[m]+"_Arr")
				if view.(CoreName+"_A")=null  then view.(CoreName+"_A")=0
			end
		end		

		if view.[BanMoto]=1 then do   //��������Ħ��NC���ֵý�Ħ���100%
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

	//Trip Generation - Balancing, 12��
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
//###################### �ڶ��� ��ͨ�ֲ�  ##########################################################################################
//#########################################################################################################################################################

Macro "Step2 Trip Distribution"

	RunMacro("Close All Maps")
	RunMacro("G30 File Close All")

	//һ�γɹ���׼�����зֲ���12����by carower by purpose: Array
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
//###################### ������ Mode Split  ##########################################################################################
//#########################################################################################################################################################

Macro "Step3 Mode Split"

	RunMacro("Close All Maps")
	RunMacro("G30 File Close All")

	//��cgrav  ���Խ�ͨ�ֲ�
	File_Grav=Path_Scen+"Matrix\\cgrav.mtx"
	mx_Grav=OpenMatrix(File_Grav,)
	mcs_Grav=CreateMatrixCurrencies(mx_Grav, , ,)	

	//��Mode Split - Ŀ����
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
			Opts.Input.[Matrix Currencies] = {{Path_Scen+"Matrix\\GC.mtx", "CA_HBW_Bicycle", "Origin", "Destination"}}   //core����ν
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
//###################### ���Ĳ� PA2OD  ##########################################################################################
//#########################################################################################################################################################

Macro "Step4 PA2OD"
	
	RunMacro("Close All Maps")
	RunMacro("G30 File Close All")

	//����ʵ����
	Mode_Occ={1, 1.1, 1.3, 35}   //{"Bicycle","Moto","Car","PT"}   //ʵ���ʣ���תΪPCE��moto=1.1, Car=1.3, PT=35����PT���˴Σ���תΪ���Ρ�
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
			Conv_Array=(if Mode_Array[i]="PT"  then Conv_Array+{"No"} else Conv_Array+{"Yes"})		//PT��ת
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
//###################### ���岽 ��·����  ##########################################################################################
//#########################################################################################################################################################
Macro "Step5 Traffic Assignment"

	RunMacro("Close All Maps")
	RunMacro("G30 File Close All")

	Mode_Auto={"Moto","Car"}
	Mode_PCU={0.5,1}		//תΪPCU
	//��PA2OD ������
	File_PA2OD=Path_Scen+"Matrix\\PA2OD.mtx"
	mx_PA2OD=OpenMatrix(File_PA2OD,)
	mcs_PA2OD=CreateMatrixCurrencies(mx_PA2OD, , ,)	

	//���ⲿ����
	File_External=Path_Scen+"Matrix\\External.mtx"
	mx_External=OpenMatrix(File_External,)
	mcs_External=CreateMatrixCurrencies(mx_External, , ,)	

	//��Ŀ����󣬿�ʼ����
	File_Auto=Path_Scen+"Matrix\\Assignment Auto OD.mtx"
	mx_Auto=OpenMatrix(File_Auto,)
	mcs_Auto=CreateMatrixCurrencies(mx_Auto, , ,)	

	//������˽�ͨ�ͻ��˽�ͨÿ��5%���������ص�External.mtx��Assignment Auto OD.mtx��ȥ��
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
//###################### ������ ��������  ##########################################################################################
//#########################################################################################################################################################
Macro "Step6 Transit Assignment"

	//RunMacro("Close All Maps")
	//RunMacro("G30 File Close All")
	
	Mode_PT={"PT"}

	//��PA2OD ������
	File_PA2OD=Path_Scen+"Matrix\\PA2OD.mtx"
	mx_PA2OD=OpenMatrix(File_PA2OD,)
	mcs_PA2OD=CreateMatrixCurrencies(mx_PA2OD, , ,)	

	//��Ŀ����󣬿�ʼ����
	File_PT=Path_Scen+"Matrix\\Assignment PT OD.mtx"
	mx_PT=OpenMatrix(File_PT,)
	mcs_PT=CreateMatrixCurrencies(mx_PT, , ,)	

	//�ȸ�ֵ0	
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


//----------------------------------------------- ����Ϊ���õ����ú��� --------------------------------------------------------
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
//����ģ�飺����·������
//**********************************************************************************************************************************************************
Macro "UpdateNetworkFF"
	SetMapUnits("Kilometers")		//Sets the Unit to Kilometers, for calculating the length in km.	
	VOT_Car=0.6  //   Ԫ/���ӡ��ڸ��ٹ�·�ϣ�תΪʱ����൱���ٶ�Լ�½���һ�롣
	Price_KM=0.4   //   0.4Ԫ/����

	Map= RunMacro("G30 new map",Path_Scen+"Network\\Streets.DBD", False)         //Open Map
	SetMap(Map)
	view=GetView()

	view_set=view+"|"
	rec=getfirstrecord(view_set,null)
	while rec<>null do 
		if view.[Type]=2  then do   //  ��֧·
			view.[AB_Link_Capacity]=view.[AB_Lanes] * 300
			view.[BA_Link_Capacity]=view.[BA_Lanes] * 300
			view.[AB_FF_Speed]=20								//km/h
			view.[BA_FF_Speed]=20								//km/h
			view.[Alpha]=2										//
			view.[Beta]=3
		end

		if view.[Type]=3  then do   //  ֧·  ���
			view.[AB_Link_Capacity]=view.[AB_Lanes] * 400
			view.[BA_Link_Capacity]=view.[BA_Lanes] * 400
			view.[AB_FF_Speed]=30								//km/h
			view.[BA_FF_Speed]=30								//km/h
			view.[Alpha]=2
			view.[Beta]=3
		end

		if view.[Type]=4  then do   //  �θɵ�
			view.[AB_Link_Capacity]=view.[AB_Lanes] * 500
			view.[BA_Link_Capacity]=view.[BA_Lanes] * 500	
			view.[AB_FF_Speed]=35								//km/h
			view.[BA_FF_Speed]=35								//km/h
			view.[Alpha]=2.2
			view.[Beta]=3.5
		end

		if view.[Type]=5  then do   //    һ�����ɵ�
			view.[AB_Link_Capacity]=view.[AB_Lanes] * 600
			view.[BA_Link_Capacity]=view.[BA_Lanes] * 600	
			view.[AB_FF_Speed]=40								//km/h
			view.[BA_FF_Speed]=40								//km/h
			view.[Alpha]=2.2
			view.[Beta]=4
		end

		if view.[Type]=6  then do   //  �������ɵ���ʡ��
			view.[AB_Link_Capacity]=view.[AB_Lanes] * 800
			view.[BA_Link_Capacity]=view.[BA_Lanes] * 800		
			view.[AB_FF_Speed]=50								//km/h  ԭ����50
			view.[BA_FF_Speed]=50								//km/h
			view.[Alpha]=2.5
			view.[Beta]=4
		end

		if view.[Type]=7  then do   //   ����
			view.[AB_Link_Capacity]=view.[AB_Lanes] * 1000
			view.[BA_Link_Capacity]=view.[BA_Lanes] * 1000		
			view.[AB_FF_Speed]=60								//km/h
			view.[BA_FF_Speed]=60								//km/h
			view.[Alpha]=2.7
			view.[Beta]=4.5
		end

		if view.[Type]=71  then do   //   ����·
			view.[AB_Link_Capacity]=view.[AB_Lanes] * 1600
			view.[BA_Link_Capacity]=view.[BA_Lanes] * 1600		
			view.[AB_FF_Speed]=80								//km/h
			view.[BA_FF_Speed]=80								//km/h
			view.[Alpha]=2.7
			view.[Beta]=4.5
		end

		if view.[Type]=8  then do   //   ���ٹ�·
			view.[AB_Link_Capacity]=view.[AB_Lanes] * 2000
			view.[BA_Link_Capacity]=view.[BA_Lanes] * 2000		
			view.[AB_FF_Speed]=120								//km/h
			view.[BA_FF_Speed]=120								//km/h
			view.[Toll]=view.[Length]*Price_KM						//0.5Ԫ/�����Toll��λ��Ԫ
			view.[Alpha]=2.8
			view.[Beta]=5.5
		end

		if view.[Type]=81  then do   //   ���ٹ�·---�����Ǹ��ٹ�·
			view.[AB_Link_Capacity]=view.[AB_Lanes] * 2000
			view.[BA_Link_Capacity]=view.[BA_Lanes] * 2000		
			view.[AB_FF_Speed]=100								//km/h
			view.[BA_FF_Speed]=100								//km/h
			view.[Toll]=view.[Length]*Price_KM						//0.5Ԫ/�����Toll��λ��Ԫ
			view.[Alpha]=2.8
			view.[Beta]=5.5
		end

		if view.[Type]=9  then do   //   �ѵ�
			view.[AB_Link_Capacity]=view.[AB_Lanes] * 600
			view.[BA_Link_Capacity]=view.[BA_Lanes] * 600	
			view.[AB_FF_Speed]=20								//km/h
			view.[BA_FF_Speed]=20								//km/h
			view.[Alpha]=2.8
			view.[Beta]=5.5
		end

		if view.[Type]=99  then do                 //����������
			view.[AB_Link_Capacity]=view.[AB_Lanes] * 5000
			view.[BA_Link_Capacity]=view.[BA_Lanes] * 5000	
			view.[AB_FF_Speed]=20								//km/h
			view.[BA_FF_Speed]=20								//km/h
			view.[Alpha]=0.15
			view.[Beta]=4
		end

		if view.[Type]=999  then do                 //�ⲿ������
			view.[AB_Link_Capacity]=view.[AB_Lanes] * 5000
			view.[BA_Link_Capacity]=view.[BA_Lanes] * 5000	
			view.[AB_FF_Speed]=100								//km/h
			view.[BA_FF_Speed]=100								//km/h
			view.[Alpha]=0.15
			view.[Beta]=4
		end

		view.[AB_FF_Time]=view.[Length]*60 / view.[AB_FF_Speed]+nz(view.[AB_ExtraTime])     //unit: min    תΪ����
		view.[BA_FF_Time]=view.[Length]*60 / view.[BA_FF_Speed]+nz(view.[BA_ExtraTime])      //unit: min    תΪ����

		if view.[Type]=8  then do   //   ���ٹ�·
			view.[AB_FF_Time]=view.[AB_FF_Time]+ view.[Length]*Price_KM/VOT_Car   //VOT_Car=1    //1Ԫ/���ӣ���һ��Сʱ60Ԫ��
			view.[BA_FF_Time]=view.[BA_FF_Time]+ view.[Length]*Price_KM/VOT_Car   //V2013��1��28��OT_Car=1    //1Ԫ/���ӣ���һ��Сʱ60Ԫ��
		end

		if view.[Type]=81  then do   //   ���ٹ�·---�����Ǹ��ٹ�·
			view.[AB_FF_Time]=view.[AB_FF_Time]+ view.[Length]*Price_KM*1.2/VOT_Car   //VOT_Car=1    //1Ԫ/���ӣ���һ��Сʱ60Ԫ��
			view.[BA_FF_Time]=view.[BA_FF_Time]+ view.[Length]*Price_KM*1.2/VOT_Car   //VOT_Car=1    //1Ԫ/���ӣ���һ��Сʱ60Ԫ��
		end

		view.[AB_Ped_Time]=(if  view.[Length]<=1 then view.[Length]*60 / 6 else 10)     // 6km/h
		view.[BA_Ped_Time]=view.[AB_Ped_Time]      //6km  unit: min    תΪ����
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
//����ģ�飺�����迹����Skim Len��ʱ���Լ�PT Skim���ٵ���CalMat_Congested������ѭ������
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
	Opts.Global.[Global Xfer Fare] = 0.4			//ת������ģʽ���κ�·�ߵ��ۿ۳��ѣ�����еĻ��� 

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

//GC ���㣬����ѭ������
Macro "CalMat_Congested"

	Shared Mode_Split
	Shared PT_Time_Factor

	if Mode_Split=null  then  Mode_Split={0.309,0.507,0.123,0.061}		//ԭʼ����ֵ
	VOT={{21.63,11.65,16.64,19.14},{10.82,5.82,8.32,9.57},{10.82,5.82,8.32,9.57}}  //����Excel��Settings.xlsx
	VOT_avg=10.95   // Ԫ/Сʱ

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

	//��Skim Length
	File_Len=Path_Scen+"Matrix\\SPMAT-Len.mtx"
	CoreLen="Shortest Path - Length"
	//RunMacro("Calc_Intrazonal",File_Len,CoreLen)
	mx_Len=OpenMatrix(File_Len,)
	mcs_Len=CreateMatrixCurrencies(mx_Len, , ,)	

	//��Skim Time
	File_Time=Path_Scen+"Matrix\\SPMAT-Time.mtx"
	CoreTime="Shortest Path - [AB_Congested / BA_Congested]"
	//RunMacro("Calc_Intrazonal",File_Time,CoreTime)
	mx_Time=OpenMatrix(File_Time,)
	mcs_Time=CreateMatrixCurrencies(mx_Time, , ,)	

	//��PT Skim����
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

					//���� by CarOwer by Purpose by Mode���迹�ɱ�
					
					//1. Bicycle   speed=12km/h   ����60���ӣ��õ��볤�ȵĹ�ϵ��5��
					//<30min�������迹Ϊ1����30-60�����迹Ϊ1.5��������60min�����迹Ϊ2����
					if i=1  then do
						mcs.(new_CoreName):=(if mcs_Len.(CoreLen)<=6  then mcs_Len.(CoreLen)/12*60+Bicycle_OVT*OVT_Factor
											    else if  mcs_Len.(CoreLen)>=6 and mcs_Len.(CoreLen)<9 then mcs_Len.(CoreLen)/12*60*1.5+Bicycle_OVT*OVT_Factor
											    else mcs_Len.(CoreLen)/12*60*2+Bicycle_OVT*OVT_Factor)
					end

					//2. Moto  GC=IVT+OVT*Factor + Fee/OVT + GC.[BanMoto]  //Ħ�г���С������ 1.1��
					//2012��1��1���𣬽�ֹ������������������Ƶ�Ħ�г������κ����������������������ϰ��أ������������������ʻ
					if i=2  then do  
						mcs.(new_CoreName):= IVT_Factor*mcs_Time.(CoreTime)*Moto_Time_Factor+OVT_Factor*Moto_OVT+(mcs_Len.(CoreLen)*Moto_Petro+Moto_Parking)*60/VOT[m][n]
						mcs.(new_CoreName):=mcs.(new_CoreName)*Auto_GC_Factor
					end

					//3. Car  GC=IVT+OVT*Factor + Fee/OVT 
					if i=3  then do  
						mcs.(new_CoreName):= IVT_Factor*mcs_Time.(CoreTime)+OVT_Factor*Car_OVT+(mcs_Len.(CoreLen)*Car_Petro+Car_Parking)*60/VOT[m][n]
						mcs.(new_CoreName):=mcs.(new_CoreName)*Auto_GC_Factor
					end

					//PTҪ������-��PT Skim���õ�

					if i=4  then do  
						mcs.(new_CoreName):= mcs_PT.(CorePT)*60/VOT[m][n]
						mcs.(new_CoreName):=mcs.(new_CoreName)*PT_GC_Factor
					end

					RunMacro("Calc_Intrazonal",Path_Scen+"Matrix\\GC.mtx",new_CoreName)

					//�����Ħ����ѽ�Ħ������գ��ڷ�ʽ����ʱ���򲻿������
					if i=2  then mcs.(new_CoreName):=mcs.(new_CoreName)+mcs.[BanMoto]
				end

				//�ۺ��迹������ʽ����֮Ȩ��
				new_CoreName=CarOwer_Array[m]+"_"+Purpose_Array[n]
				//AddMatrixCore(mx, new_CoreName)
				mcs = CreateMatrixCurrencies(mx, , ,)	
				/*
				//����ͨС���ı���������Ȩ��
				mcs.(new_CoreName):=Nz(mcs.(CarOwer_Array[m]+"_"+Purpose_Array[n]+"_"+Mode_Array[1]))*mcs_Split.(Mode_Array[1]+"_Split")+
				Nz(mcs.(CarOwer_Array[m]+"_"+Purpose_Array[n]+"_"+Mode_Array[2]))*mcs_Split.(Mode_Array[2]+"_Split")+
				Nz(mcs.(CarOwer_Array[m]+"_"+Purpose_Array[n]+"_"+Mode_Array[3]))*mcs_Split.(Mode_Array[3]+"_Split")+
				Nz(mcs.(CarOwer_Array[m]+"_"+Purpose_Array[n]+"_"+Mode_Array[4]))*mcs_Split.(Mode_Array[4]+"_Split")
				*/
				//��ȫ�зֵ���������Ȩ��
				mcs.(new_CoreName):=Nz(mcs.(CarOwer_Array[m]+"_"+Purpose_Array[n]+"_"+Mode_Array[1]))*Mode_Split[1]+
				Nz(mcs.(CarOwer_Array[m]+"_"+Purpose_Array[n]+"_"+Mode_Array[2]))*Mode_Split[2]+
				Nz(mcs.(CarOwer_Array[m]+"_"+Purpose_Array[n]+"_"+Mode_Array[3]))*Mode_Split[3]+
				Nz(mcs.(CarOwer_Array[m]+"_"+Purpose_Array[n]+"_"+Mode_Array[4]))*Mode_Split[4]

				/*
				//2012��1��1���𣬽�ֹ������������������Ƶ�Ħ�г������κ����������������������ϰ��أ������������������ʻ
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

