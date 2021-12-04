
//************************************************************************************
Macro "ToCube"
  RunMacro("ToCubeNodes")
//  RunMacro("ToCubeLinks")
//  RunMacro("ToCubeTransits")
EndMacro

/

//************************************************************************************
Macro "ToCubeNodes"


	
    SetMapUnits("KiloMeters")
    gNodeLayer="Nodes"
    gLinkLayer="Streets"
    
    FileName = ChooseFileName({{"DBF File", "*.dbf"}},"Choose a Nodes DBF File", {,,,,"TRUE"})
    
    NodesDBF = CreateTable("Nodes", FileName, "DBASE", {
	{"N", "Integer", 8, null, "No"},
	{"X", "Real", 15, 2, "No"},
	{"Y", "Real", 15, 2, "No"}
	})
	
    //输出形心
    SetLayer(gNodeLayer)
    qry="Select * where Centroid > 0"
    centroids=SelectByQuery("Centroids", "Several", qry,)


    SetLayer(gNodeLayer)
    crh=GetFirstRecord(gNodeLayer+"|Centroids",)
    While crh<>NULL Do
      cpt=RunMacro("BLToXY",GetPoint(RH2ID(crh)))
    
      cVal=GetRecordValues(gNodeLayer, , {"ID"})
      
      rh = AddRecord(NodesDBF, {
	{"N", cVal[1][2]},
	{"X", cpt[1]},
	{"Y", cpt[2]}
	})
  
      crh=GetNextRecord(gNodeLayer+"|Centroids", , )
    End
    
     //输出节点
    SetLayer(gNodeLayer)

    SetLayer(gNodeLayer)
    nrh=GetFirstRecord(gNodeLayer+"|",)
    While nrh<>NULL Do
    
      nVal=GetRecordValues(gNodeLayer, , {"ID", "Centroid"})
      
      If nVal[2][2]>0 Then Goto NextNode
      
      npt=RunMacro("BLToXY",GetPoint(RH2ID(nrh)))
    
      nVal=GetRecordValues(gNodeLayer, ,{"ID"})
      
      rh = AddRecord(NodesDBF, {
	{"N", nVal[1][2]},
	{"X", npt[1]},
	{"Y", npt[2]}
	})
	
NextNode:  
      nrh=GetNextRecord(gNodeLayer+"|", ,)
    End
    
    CloseView(NodesDBF)

EndMacro


//************************************************************************************
Macro "ToCubeLinks"

    SetMapUnits("KiloMeters")
    gNodeLayer="Nodes"
    gLinkLayer="Streets"
    
    FileName = ChooseFileName({{"DBF File", "*.dbf"}},"Choose a Links DBF File", {,,,,"TRUE"})
    
    LinksDBF = CreateTable("Links", FileName, "DBASE", {
	{"A", "Integer", 8, NULL, "No"},
	{"B", "Integer", 8, NULL, "No"},
	{"Length", "Real", 8, 2, "No"},
	{"Type", "Integer", 4, NULL, "No"},
	{"Lanes", "Real", 5, 2, "No"},
	{"BLane", "Integer", 2, NULL, "No"},
	{"SCRLine", "Integer", 4, NULL, "No"},
	{"LinkCap", "Real", 10, 2, "No"},
        {"InterCap", "Real", 10, 2, "No"},
	{"Speed", "Real", 6, 2, "No"},
        {"SToll_PV", "Real", 10, 2, "No"},
        {"SToll_GV", "Real", 10, 2, "No"},
        {"SToll_VON", "Real", 10, 2, "No"},
        {"DToll_PV", "Real", 10, 2, "No"},
        {"DToll_GV", "Real", 10, 2, "No"},
        {"DToll_VON", "Real", 10, 2, "No"},
        {"XSZPen_PV", "Real", 10, 2, "No"},
        {"XSZPen_PT", "Real", 10, 2, "No"},
        {"XSZPen_GV", "Real", 10, 2, "No"},
        {"XSZPen_VON", "Real", 10, 2, "No"},                                      
	{"LinkID", "Integer", 4, NULL, "No"}
	})    
    
     //输出路段
    SetLayer(gLinkLayer)
    
    EnableProgressBar("Export Links", 1)	//设置进度条
    CreateProgressBar("Export Links...","True")
    Step=0
    Count=GetRecordCount(gLinkLayer,)
    StepII=0
    
    lrh=GetFirstRecord(gLinkLayer+"|",)
    While lrh<>null do
    
      StepII=StepII+1
      temp=RealToInt(StepII*100/Count)	//计算进度
      If temp>Step then do
        Step=temp
        stat=UpdateProgressBar("Export Links..." + String(Step),Step)
        If stat="True" Then Do
          DestroyProgressBar()
          GoTo Quit_Loop
        End
      End

      
      lVal=GetRecordValues(gLinkLayer, ,{"Dir","Length","Mode","Type","Lanes","Blane","SCRLine","LinkCap",
                                         "AB_InterCap","BA_InterCap","Speed","SToll_PV","SToll_GV","SToll_VON",
                                         "DToll_PV","DToll_GV","DToll_VON","XSZPen_PV","XSZPen_PT","XSZPen_GV","XSZPen_VON"})
      For i=1 To lVal.Length Do
        If lVal[i][2]=NULL Then  lVal[i][2]=0
      End
      
      nids=GetEndpoints(RH2ID(lrh))
      
      If lVal[1][2]>=0 then do
        rh = AddRecord(LinksDBF, {
	  {"A",         nids[1]},
	  {"B",         nids[2]},
	  {"Length",    lVal[2][2]},
	  {"Type",      lVal[4][2]},
	  {"Lanes",     lVal[5][2]},
	  {"BLane",     lVal[6][2]},
	  {"SCRLine",   lVal[7][2]},
          {"LinkCap",   lVal[8][2]},
          {"InterCap",  lVal[9][2]},          
	  {"Speed",     lVal[11][2]},
	  {"SToll_PV",  lVal[12][2]},
	  {"SToll_GV",  lVal[13][2]},
	  {"SToll_VON", lVal[14][2]},
	  {"DToll_PV",  lVal[15][2]},
	  {"DToll_GV",  lVal[16][2]},
	  {"DToll_VON", lVal[17][2]},
	  {"XSZPen_PV", lVal[18][2]},
	  {"XSZPen_PT", lVal[19][2]},
	  {"XSZPen_GV", lVal[20][2]},
	  {"XSZPen_VON",lVal[21][2]},	  
	  {"LinkID",    RH2ID(lrh)}	
	  })
      End
      If lVal[1][2]<=0 then do
        rh = AddRecord(LinksDBF, {
	  {"A",         nids[2]},
	  {"B",         nids[1]},
	  {"Length",    lVal[2][2]},
	  {"Type",      lVal[4][2]},
	  {"Lanes",     lVal[5][2]},
	  {"BLane",     lVal[6][2]},
	  {"SCRLine",   lVal[7][2]},
          {"LinkCap",   lVal[8][2]},
          {"InterCap",  lVal[10][2]},
	  {"Speed",     lVal[11][2]},
	  {"SToll_PV",  lVal[12][2]},
	  {"SToll_GV",  lVal[13][2]},
	  {"SToll_VON", lVal[14][2]},
	  {"DToll_PV",  lVal[15][2]},
	  {"DToll_GV",  lVal[16][2]},
	  {"DToll_VON", lVal[17][2]},
	  {"XSZPen_PV", lVal[18][2]},
	  {"XSZPen_PT", lVal[19][2]},
	  {"XSZPen_GV", lVal[20][2]},
	  {"XSZPen_VON",lVal[21][2]},
	  {"LinkID",    RH2ID(lrh)}
	  })
      End
        
      lrh=GetNextRecord(gLinkLayer+"|", ,)
    End
    DestroyProgressBar()

Quit_Loop:
    
    
    CloseView(LinksDBF)

EndMacro


//************************************************************************************
Macro "ToCubeTransits"

    gRouteLayer="Routes"
    gLinkLayer="Streets"
    gNodeLayer="Nodes"
    
    FileName = ChooseFileName({{"LIN File", "*.lin"}},"Choose a Transits LIN File", {,,,,"TRUE"})
        
    rfile=OpenFile(FileName,"w+") 
       
    WriteLine(rfile,";;<<PT>><<LINE>>;;")
       
    SetLayer(gRouteLayer)
    
    rrh=GetFirstRecord(gRouteLayer+"|",)
    
    While rrh<>NULL Do
    
      rVal=GetRecordValues(gRouteLayer, , {"Route_Name", "Vehicle", "Hdwy", "COLOR"})
      
      //s = Substring(rVal[1][2], StringLength(rVal[1][2]), 1)
      //If s = "b" Or s = "B" Then GoTo Continue
      
      For i = 2 To rVal.Length Do
        If rVal[i][2]=NULL THEN rVal[i][2]=0
      End
    
      links = GetRouteLinks(gRouteLayer, rVal[1][2])
      SetLayer(gLinkLayer)
      
      NodeList=NULL
      
      For j=1 To links.Length Do
        nodes = GetEndpoints(links[j][1])
        If j=1 Then Do
          If links[j][2]=1 Then NodeList = InsertArrayElements(NodeList, , {nodes[1], nodes[2]})
          Else NodeList = InsertArrayElements(NodeList, , {nodes[2], nodes[1]})
        End
        Else Do
          If links[j][2]=1 Then NodeList = InsertArrayElements(NodeList, j+1, {nodes[2]})
          Else NodeList = InsertArrayElements(NodeList, j+1, {nodes[1]})
        End
      End 
           
      wline="LINE NAME=\"" + rVal[1][2] + "\", MODE=" + String(rVal[2][2]) + ", COLOR=" + String(rVal[4][2]) +
            ", ONEWAY=T, HEADWAY=" + RealToString(rVal[3][2]) + ","
      WriteLine(rfile, wline)
      
      ///////////////
      //num = NodeList.Length
      //For i = 1 To num-1 Do
      //  NodeList = InsertArrayElements(NodeList, num+1, {NodeList[i]})
      //End
      ////////////////////////
      

      For i = 1 To NodeList.Length Do
        Node = -NodeList[i]
        If i = 1 Or i = NodeList.Length Then Node = -Node
        Else Do
          nVal=GetRecordValues(gNodeLayer, ID2RH(NodeList[i]), {"BusStop"})
          If nVal[1][2] > 0 Then Node = -Node  //暂时定义小区接驳线，轨道接驳线为公交站 （1:常规公交站暂不列为站点)
        End
        
        If i = 1 Then WriteLine(rfile, "     N=" + String(Node) + ",")
        Else Do
          If i = NodeList.Length Then WriteLine(rfile, "     " + String(Node))
          Else WriteLine(rfile, "     " + String(Node) + ",")
        End
      End
      
  Continue:
      rrh=GetNextRecord(gRouteLayer+"|", ,)
      
   End
   
   CloseFile(rfile)
    
    
EndMacro


//************************************************************************************
Macro "ToCubeNTLs"

    global NodeList
    NodeList=NULL
  
    SetMapUnits("KiloMeters")
    gNodeLayer="Nodes"
    gLinkLayer="Streets"
    
    FileName = ChooseFileName({{"NTL File", "*.dbf"}},"Choose a NTL DBF File", {,,,,"TRUE"})
    
    
    NTLsDBF = CreateTable("NTLs", FileName, "DBASE", {
	{"A",     "Integer", 8, null, "No"},
	{"B",     "Integer", 8, null, "No"},
	{"Type",  "Integer", 4, null, "No"}
	})
	
    
    ////////////////////////////////////////////////////////////////////    
    //常规小区辅助公交路段 步行线: type=1
    SetLayer(gNodeLayer)
    qry = "Select * where Centroid > 0"
    Count = SelectByQuery("Centroids", "Several", qry,)
    RecSet = gNodeLayer + "|Centroids"
    
    EnableProgressBar("Export NTLs", 1)	//设置进度条
    CreateProgressBar("Export NTLs...","True")
    Step=0
    StepII=0
    
    nrh=GetFirstRecord(RecSet, )
    While nrh<>null do
    
      StepII=StepII+1
      temp=RealToInt(StepII*100/Count)	//计算进度
      If temp>Step then do
        Step=temp
        stat=UpdateProgressBar("Export NTLs..." + String(Step),Step)
        If stat="True" Then Do
          DestroyProgressBar()
          GoTo Quit_Loop
        End
      End
      
      NodeList = NULL
      RunMacro("FindNodeOfLink", gLinkLayer, RH2ID(nrh), "Type", 1)
      
      For i = 1 To NodeList.Length Do
        AddRecord(NTLsDBF, {
	                    {"A", RH2ID(nrh)},
	                    {"B", NodeList[i]},
	                    {"Type", 101}
	                   })
      End
      nrh=GetNextRecord(RecSet, ,)
    End
    DestroyProgressBar()
    
    ////////////////////////////////////////////////////////////////////
    //轨道小区辅助常规公交路段 步行线: type=23 - > 22
    SetLayer(gNodeLayer)
    qry = "Select * where Centroid > 0"
    Count = SelectByQuery("Centroids", "Several", qry,)
    RecSet = gNodeLayer + "|Centroids"
    
    EnableProgressBar("Export NTLs", 1)	//设置进度条
    CreateProgressBar("Export NTLs...","True")
    Step=0
    StepII=0
    
    nrh=GetFirstRecord(RecSet, )
    While nrh<>null do
    
      StepII=StepII+1
      temp=RealToInt(StepII*100/Count)	//计算进度
      If temp>Step then do
        Step=temp
        stat=UpdateProgressBar("Export NTLs..." + String(Step),Step)
        If stat="True" Then Do
          DestroyProgressBar()
          GoTo Quit_Loop
        End
      End
      
      NodeList = NULL
      RunMacro("FindNodeOfLink", gLinkLayer, RH2ID(nrh), "Type", 23)
      List23 = NodeList
      NodeList = NULL
      
      For i = 1 To List23.Length Do
        RunMacro("FindNodeOfLink", gLinkLayer, List23[i], "Type", 22)
      End  
      
      For i = 1 To NodeList.Length Do
        AddRecord(NTLsDBF, {
	                    {"A", RH2ID(nrh)},
	                    {"B", NodeList[i]},
	                    {"Type", 102}
	                   })
      End
      nrh=GetNextRecord(RecSet, ,)
    End
    DestroyProgressBar()


    ////////////////////////////////////////////////////////////////////
    //轨道小区辅助常规公交路段 步行线: type=23 - > 20 -> 21
    SetLayer(gNodeLayer)
    qry = "Select * where Centroid > 0"
    Count = SelectByQuery("Centroids", "Several", qry,)
    RecSet = gNodeLayer + "|Centroids"
    
    EnableProgressBar("Export NTLs", 1)	//设置进度条
    CreateProgressBar("Export NTLs...","True")
    Step=0
    StepII=0
    
    nrh=GetFirstRecord(RecSet, )
    While nrh<>null do
    
      StepII=StepII+1
      temp=RealToInt(StepII*100/Count)	//计算进度
      If temp>Step then do
        Step=temp
        stat=UpdateProgressBar("Export NTLs..." + String(Step),Step)
        If stat="True" Then Do
          DestroyProgressBar()
          GoTo Quit_Loop
        End
      End
      
      NodeList = NULL
      RunMacro("FindNodeOfLink", gLinkLayer, RH2ID(nrh), "Type", 23)
      List23 = NodeList
      NodeList = NULL
      
      For i = 1 To List23.Length Do
        RunMacro("FindNodeOfLink", gLinkLayer, List23[i], "Type", 20)
      End
      
      List20 = NodeList
      NodeList = NULL
      
      For i = 1 To List20.Length Do
        RunMacro("FindNodeOfLink", gLinkLayer, List20[i], "Type", 21)
      End
      
       
      For i = 1 To NodeList.Length Do
        AddRecord(NTLsDBF, {
	                    {"A", RH2ID(nrh)},
	                    {"B", NodeList[i]},
	                    {"Type", 103}
	                   })
      End
      nrh=GetNextRecord(RecSet, ,)
    End
    DestroyProgressBar()
    
    
    ////////////////////////////////////////////////////////////////////    
    //轨道间的换乘线: type=21
    SetLayer(gLinkLayer)
    qry="Select * where Type = 21"
    Count=SelectByQuery("CLines", "Several", qry,)
    RecSet = gLinkLayer + "|CLines"
    
    EnableProgressBar("Export NTLs", 1)	//设置进度条
    CreateProgressBar("Export NTLs...","True")
    Step=0
    StepII=0
    
    lrh=GetFirstRecord(RecSet, )
    While lrh<>null do
    
      StepII=StepII+1
      temp=RealToInt(StepII*100/Count)	//计算进度
      If temp>Step then do
        Step=temp
        stat=UpdateProgressBar("Export NTLs..." + String(Step),Step)
        If stat="True" Then Do
          DestroyProgressBar()
          GoTo Quit_Loop
        End
      End

      
      nids=GetEndpoints(RH2ID(lrh))
      
      rh = AddRecord(NTLsDBF, {
	{"A", nids[1]},
	{"B", nids[2]},
	{"Type", 104}
	})

      lrh=GetNextRecord(RecSet, ,)
    End
    DestroyProgressBar()
    
Quit_Loop:
    CloseView(NTLsDBF)

EndMacro

//************************************************************************************
Macro "OutLinkPloy"

    gLinkLayer="Streets"
    SetLayer(gLinkLayer)
    
    FileName = ChooseFileName({{"PLY File", "*.ply"}},"Choose a LinkPloy File", {,,,,"TRUE"})
        
    rfile=OpenFile(FileName,"w+") 
       
    WriteLine(rfile,"FROMNODE,TONODE,INDEX,XCOORD,YCOORD")
       
    
    EnableProgressBar("Export Links", 1)	//设置进度条
    CreateProgressBar("Export Links...","True")
    Step=0
    Count=GetRecordCount(gLinkLayer,)
    StepII=0
    
    lrh=GetFirstRecord(gLinkLayer+"|",)
    While lrh<>null do
    
      StepII=StepII+1
      temp=RealToInt(StepII*100/Count)	//计算进度
      If temp>Step then do
        Step=temp
        stat=UpdateProgressBar("Export Links..." + String(Step),Step)
        If stat="True" Then Do
          DestroyProgressBar()
          GoTo Quit_Loop
        End
      End
      
      pts = GetLine(RH2ID(lrh))
      nids=GetEndpoints(RH2ID(lrh))
      if nids[1]<nids[2] then do
        for i=1 to pts.Length do
          wline = string(nids[1])+","+string(nids[2])+","+string(i)+","+string(pts[i].lon)+","+string(pts[i].lat)
          WriteLine(rfile,wline)
        end
      end
      else do
        for i=1 to pts.Length do
          wline = string(nids[2])+","+string(nids[1])+","+string(i)+","+string(pts[pts.Length-i+1].lon)+","+string(pts[pts.Length-i+1].lat)
          WriteLine(rfile,wline)
        end
      end
      
        
      lrh=GetNextRecord(gLinkLayer+"|", ,)
    End
    DestroyProgressBar()

Quit_Loop:
    
    
    CloseFile(rfile)

EndMacro

//********************************************************************************************
//********************************************************************************************
Macro "ToShp"

  FileName = ChooseFileName({{"SHP File", "*.shp"}},"Choose a SHP File", {,,,,"TRUE"})
  field_list = GetFields(, "All")
  ExportArcViewShape(GetView(), FileName, {{"Fields", field_list[1]}, , , {"Transform",{{104000, 32000, 113.944430, 22.656660},{157000, 36000, 114.460910, 22.699530}, {127000, 13000, 114.172050, 22.487220}}}})
EndMacro

//********************************************************************************************
//********************************************************************************************
Macro "ToDxf"

  FileName = ChooseFileName({{"DXF File", "*.dxf"}},"Choose a DXF File", {,,,,"TRUE"})
  field_list = GetFields(, "All")
  ExportDXF(GetView()+"|Selection", FileName, {{"Fields", field_list[1]}, , , {"Transform",{{104000, 32000, 113.944430, 22.656660},{157000, 36000, 114.460910, 22.699530}, {127000, 13000, 114.172050, 22.487220}}}})
EndMacro

//********************************************************************************************
//********************************************************************************************
Macro "InDxf"

  FileName = ChooseFileName({{"DXF File", "*.dxf"}},"Choose a DXF File", {,,,,"TRUE"})
  field_list = GetFields(, "All")
  //ImportDXF(FileName, "c:\\temp\\abc.dbd", "Line", {,,,,,,,,,,,,,{"Transform",{{104000, 32000, 113.944430, 22.656660},{157000, 36000, 114.460910, 22.699530}, {127000, 13000, 114.172050, 22.487220}}}})
  ImportDXF(FileName, "c:\\temp\\abc.dbd", "Line", {
	{"Label", "Street Centerline File"},
	{"Layer Name", "Centerline"},
	{"Layers", "All"},
	{"Fields", "All"},
	{"Table Filename", "abc.bin"},
	{"Optimize", "True"},
	{"Transform",{{104000, 32000, 113.944430, 22.656660},{157000, 36000, 114.460910, 22.699530}, {127000, 13000, 114.172050, 22.487220}}}
	})
EndMacro


//********************************************************************************************
//********************************************************************************************
Macro "Format"(val, dec)
    str=String(val)
    If StringLength(str)<=dec Then Return(LPad(str, dec))
    Else Do
      dot=Position(str,".")
      If dot=0 Or dot-1>dec Then Return(LPad("#", dec))
      Else Return(Left(str, dec))
    End
EndMacro

//********************************************************************************************
//********************************************************************************************
Macro "FindNodeOfLink"(StreetsLayer, Node, LinkAttName, AttValue)

  SetLayer(GetNodeLayer(StreetsLayer))
  links_j = GetNodeLinks(Node)
  SetLayer(StreetsLayer)
  For i=1 To links_j.Length Do
    lVal=GetRecordValues(StreetsLayer, ID2RH(links_j[i]), {LinkAttName})
    If lVal[1][2] = AttValue Then Do
      link_nodes = GetEndpoints(links_j[i])
      If link_nodes[1] =  Node Then Do
        NodeList = InsertArrayElements(NodeList, , {link_nodes[2]})
      End  
      Else Do
        NodeList = InsertArrayElements(NodeList, , {link_nodes[1]})
      End
    End
  End
  
EndMacro


/////**************************************************************************************************
/////**************************************************************************************************
//经纬度坐标换算成深圳本地坐标
Macro "BLToXY"(coor)
  B=coor.lat/1000000.0
  L=coor.lon/1000000.0

//WGS84经纬度坐标到北京54坐标
  PAI=3.1415926535898
  a=6378245.0
  e2=0.00669342162297
  e12=0.00673852541468
  p2=3600.0*180.0/PAI
  P0=PAI/180.0
  
  C0=6367558.49686
  C1=32005.79642
  C2=133.86115
  C3=0.7031
  
  l=(L-114)*3600
  t=tan(B*P0)
  t2=t*t
  t4=t2*t2
  Ita2=e12*cos(B*P0)*cos(B*P0)
  Ita4=Ita2*Ita2
  N=a/sqrt(1-e2*sin(B*P0)*sin(B*P0))
  m=l*cos(B*P0)/p2
  
  m2=m*m
  m4=m2*m2
  SinBf=sin(B*P0)
  SinBf2=SinBf*SinBf
  SinBf4=SinBf2*SinBf2
  
  Temp1=C0*B*P0
  Temp2=cos(B*P0)*SinBf*(C1+C2*SinBf2+C3*SinBf4)
  Temp3=1.0/2.0*N*t*m2
  Temp4=1/24.0*(5.0-t2+9*Ita2+4*Ita4)*N*t*m4
  Temp5=1/720*(61-58*t2+t4)*N*t*(m2*m4)
  
  Temp6=N*m
  Temp7=1/6.0*(1-t2+Ita2)*N*(m*m2)
  Temp8=1/120.0*(5-18*t2+(t2*t2)+14*Ita2*Ita2-58*Ita2*t2)*N*(m*m4)
  X=Temp1-Temp2+Temp3+Temp4+Temp5
  Y=500000+Temp6+Temp7+Temp8
    
//从北京54坐标到本地坐标
  PAI=3.1415926535898
  angle=1.00860457
  s=0.997874528635
  dx=-2460040.63895730
  dy=-433223.52500793
  
  angle=angle/180.0*PAI
  
  lx=X*cos(angle)-Y*sin(angle)
  ly=X*sin(angle)+Y*cos(angle)
  
  lx=lx*s
  ly=ly*s
  
  lx=lx+dx
  ly=ly+dy
  
  return({ly,lx})
  
endMacro


