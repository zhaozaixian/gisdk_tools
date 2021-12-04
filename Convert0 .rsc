Macro "export"

// 取出Centroids, Centroid connectors
	clyr=RunMacro("tsm get layer", "Centroid")    SetLayer(clyr)
	Centroids=null
	rh=GetFirstRecord(clyr+"|",{{"ID", "Ascending"}})
	while rh!=null  do
	      pt.ID=clyr.ID  pt.Coord=GetPoint(clyr.ID)
	      Centroids=Centroids+{CopyArray(pt)}
	      rh=GetNextRecord(clyr+"|",null,{{"ID", "Ascending"}})
	 end
       showarray(Centroids)

	conlyr=RunMacro("tsm get layer", "Centroid Connector")    SetLayer(conlyr)
	Connectors=null
	rh=GetFirstRecord(conlyr+"|",{{"Centroid", "Ascending"}})
        centroid_before=0    link_before=0
	while rh!=null  do
	     if    conlyr.Centroid<>centroid_before  or  conlyr.Link!=link_before  then  do
	           Connectors=Connectors+{{conlyr.Centroid,conlyr.Link}}
                   centroid_before=conlyr.Centroid    link_before=conlyr.Link
	     end
	     rh=GetNextRecord(conlyr+"|",null,{{"Centroid", "Ascending"}})
	end
       showarray(Connectors)

// Export Network DBD
	nlyr=RunMacro("tsm get layer", "Node")
	llyr=RunMacro("tsm get layer", "Link")
	slyr=RunMacro("tsm get layer", "Segment")

       	//建立joinedview: link+segment
	strct = GetTableStructure(slyr)
	aggr=null	    for j=2  to  strct.length  do   aggr=aggr+{{strct[j][1], {{"Max"}}}}  end            //对所有字段进行Aggr
       // jview=  JoinViews( "link+segment", llyr+".ID"   , slyr+".Link",  {{"A",}, {"Fields", aggr}} )

       nflds= GetFields(nlyr , "All")
       lflds = GetFields(llyr , "All")

       opts.[Field Name]=lflds[1]
       opts.[Field Spec]=lflds[2]
       opts.[Node Field Name]=nflds[1]
       opts.[Node Field Spec]=nflds[2]
       showarray(opts)

       ExportGeography(llyr+"|", "c:\\houhai.dbd", opts)

//建立Map，得到Link layer，Node layer
	dbd="c:\\houhai.dbd"      info = GetDBInfo(dbd)   dblyr= GetDBLayers(dbd)
	ThemeMap=CreateMap( ,{{"scope",info[1]}} )
        nlyr =AddLayer(ThemeMap,  , dbd, dblyr[1] ,)        llyr =AddLayer(ThemeMap,  ,dbd, dblyr[2] ,)
	redrawMap()

       //增加LinkType , IsCentroid 字段
       strct = GetTableStructure(llyr)
       for i = 1 to strct.length    do	strct[i] = strct[i] + {strct[i][1]}    end    // Copy the current name to the end of  sub_strct
       strct = InsertArrayElements(strct, 4, {{"LinkType", "Short", 8, , , , , , , , , null}})
       ModifyTable(llyr, strct)

       strct = GetTableStructure(nlyr)
       for i = 1 to strct.length    do	strct[i] = strct[i] + {strct[i][1]}    end    // Copy the current name to the end of  sub_strct
       strct = InsertArrayElements(strct, 4, {{"IsCentroid", "String", 8, , , , , , , , , null}})
       ModifyTable(nlyr, strct)

       //填充LinkType 字段
       SetLayer(llyr)
       rh=GetFirstRecord(llyr+"|", )
       while rh!=null  do
             llyr.LinkType=0
	     if  llyr.Class="Access Road"    then   llyr.LinkType=1
	     if  llyr.Class="Local Street"     then   llyr.LinkType=2
	     if  llyr.Class="Collector"         then   llyr.LinkType=3
	     if  llyr.Class="Minor Arterial"    then   llyr.LinkType=4
	     if  llyr.Class="Major Arterial"    then   llyr.LinkType=5
	     if  llyr.Class="Trunk Arterial"    then   llyr.LinkType=6
	     if  llyr.Class="Expressway"        then   llyr.LinkType=7
	     if  llyr.Class="Highway"            then   llyr.LinkType=8
	     if  llyr.Class="System Ramp"    then   llyr.LinkType=9
	     if  llyr.Class="Ramp"                 then   llyr.LinkType=10
	     if  llyr.Class="Connector"         then   llyr.LinkType=99

	     if    llyr.LinkType=0  then  ShowMessage("Link "+rh+"  没有等级class定义或定义不在范围内！")
	     rh=GetNextRecord(llyr+"|",null, )
       end


// 添加Centroids & Centroid Connectors
SetProgressWindow("Status", 1)     // Allow only a single progress bar
CreateProgressBar("Adding Centroids & Connectors...", "True")

	SetLayer(nlyr)   flds_vals=null
	for i=1 to Centroids.Length   do
	     AddPoint(Centroids[i].Coord, Centroids[i].ID)
	     flds_vals.IsCentroid="Yes"
	     SetRecordValues(nlyr, ID2RH(Centroids[i].ID), flds_vals )
            stat = UpdateProgressBar("Adding Centroids.....    Centroid ID   " + String(Centroids[i].ID), floor(i/Centroids.Length*100))
	    if  stat="True"  then goto quit
	end

	SetLayer(llyr)   opts.[Snap Node]="true"   flds_vals=null
	for i=1 to Connectors.Length  do
	       centroidID=Connectors[i][1]   tolinkID=Connectors[i][2]
	       endpts= GetEndPoints(tolinkID)

	       SetLayer(nlyr)   centroidCOORD= GetPoint(centroidID)
               //如果link一端断头，则断头点为与centroid 连接的节点
	       list1=GetNodeLinks(endpts[1])  list2=GetNodeLinks(endpts[2])
	       if   list1.length=1 then   do  ptsCONNECTED=GetPoint(endpts[1])  goto  next  end
	       if   list2.length=1 then   do  ptsCONNECTED=GetPoint(endpts[2])  goto  next  end
               //如果link两端都有连接，则根据与centroid 连接最短距离判定
	       ptsFROM=GetPoint(endpts[1])    ptsTO=GetPoint(endpts[2])
	       dist1=GetDistance(centroidCOORD,ptsFROM)
	       dist2=GetDistance(centroidCOORD,ptsTO)
	       if  dist1<dist2  then  ptsCONNECTED=ptsFROM    else    ptsCONNECTED=ptsTO

	       next:
               SetLayer(llyr)        	ret=AddLink({centroidCOORD,ptsCONNECTED}, null, opts)

	       flds_vals.LinkType=99
	       SetRecordValues(llyr, ID2RH(ret[1]), flds_vals )
              stat = UpdateProgressBar("Adding Centroid  Connectors.....    Centroid ID   " + String(centroidID), floor(i/Connectors.Length*100))
	      if  stat="True"  then goto quit
	end
quit:
DestroyProgressBar()

// 建立Link Class Themes
        SetLayer(llyr)
	opts=null    opts.Other="False"  opts.Title="道路等级"
        label_arr={"Access Road" ,"Local Street",         "Collector",       "Minor Arterial",   "Major Arterial" ,"Trunk Arterial", "Expressway",  //"Highway" ,
	                     "System Ramp",        "Ramp",             "Connector"}   //没有Highway
	shared cc_Colors
	clr_arr=   {cc_Colors.Gray, cc_Colors.Yellow, cc_Colors.Gold, cc_Colors.Green ,cc_Colors.Cyan, cc_Colors.Blue ,cc_Colors.Red, //cc_Colors.Purple,
	                   cc_Colors.Brown, cc_Colors.Orange,cc_Colors.Black}
	wid_arr=  {1, 1 ,1,2,2,3,3      ,2,2,1}
	ls_arr = RunMacro("G30 setup line styles")
	lsty_arr=  {ls_arr[2],ls_arr[2],ls_arr[2], ls_arr[2], ls_arr[2], ls_arr[69],ls_arr[69],      ls_arr[2], ls_arr[2],ls_arr[5]}

	theme = CreateTheme("Classtheme", llyr+".LinkType", "Categories",512 , opts)
	SetThemeClassLabels(theme, label_arr)
	SetThemeLineColors(theme, clr_arr)
	SetThemeLineStyles(theme,lsty_arr )
	SetThemeLineWidths(theme,wid_arr )
	ShowTheme(,theme)

	Setlayer(nlyr)
	SetIcon(nlyr+"|", "Font Character","Caliper Cartographic|2",36 )
        SelectByQuery("Centroids", "Several", 'Select * where IsCentroid = "Yes" ', )
	SetIcon(nlyr+"|Centroids", "Font Character","Caliper Cartographic|6",63 )
	SetIconColor(nlyr+"|Centroids", cc_Colors.Red)
        SetDisplayStatus(nlyr+"|Centroids", "Active")
	SetLabels(nlyr+"|Centroids", nlyr+".ID",  { {"Font", "Arial|Bold|9"}, {"Color", cc_Colors.Red}})

	RedrawMap()

EndMacro


//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
////////Create Movemt Tabel
Macro "CMT"

Plink=null   Pnode=null    TConns=null

 //从Links Layer 中读出信息， 得到变量Plink
	pi=3.1415926535897932384626433832795
	llyr=RunMacro("tsm get layer", "Link")   SetLayer(llyr)
	Plink=null
	rh=GetFirstRecord(llyr+"|",{{"ID", "Ascending"}})
	while rh!=null  do
	      //读出link.A/Bnode, link.class, link.segments
	      linkIDstr=string(llyr.ID )
	      Plink.(linkIDstr).Anode=llyr.ANode   Plink.(linkIDstr).Bnode=llyr.BNode
	      Plink.(linkIDstr).Dir=llyr.Dir
	      Plink.(linkIDstr).Segments=llyr.Segments

	      //计算link角度：基于拓扑方向
	      coords=Getline(llyr.ID)    node1=coords[1]    node2=coords[coords.length]
	      p1=RunMacro("BLToXY", node1)    p2=RunMacro("BLToXY", node2)
              Plink.(linkIDstr).degree=Atan2(p2.y-p1.y,  p2.x-p1.x)/pi*180

	      SetStatus(2,"Reading LinkLayer,  linkID:"+linkIDstr , )
	      rh=GetNextRecord(llyr+"|",null,{{"ID", "Ascending"}})
	 end

  //从Segment Layer 中读出每个segment每个方向的车道数, 得到变量SegLanes
       SegLanes=null
	slr=RunMacro("tsm get layer", "Segment")  SetLayer(slr)
	rh=GetFirstRecord(slr+"|",{{"ID", "Ascending"}})
	while rh!=null do
	      seg=string(slr.ID)
	      SegLanes.(seg).A=NZ(slr.Lanes_AB)   SegLanes.(seg).B=NZ(slr.Lanes_BA)
	      rh=GetNextRecord(slr+"|",null,{{"ID", "Ascending"}})
	end

 //从Lanes Layer 中读出segment.dir.turn.connectors, 得到变量TConns
        TConns=null
	lanelyr=RunMacro("tsm get layer", "Lane")  SetLayer(lanelyr)
	rh=GetFirstRecord(lanelyr+"|",{{"Segment", "Ascending"}})
	while rh!=null do  //读出每个Lane信息并计算
	      seg=string(lanelyr.Segment)  dir=lanelyr.Dir   Lpos= lanelyr.Position
	      if  dir=1  then  direc="A"  else direc="B"
	      Rpos=SegLanes.(seg).(direc)-Lpos-1   lane_num=R2I(pow(2,Rpos))   //lane的从右边数定位，车道connectors: 1，2，4，8，16

	      p=null
	      turns=lanelyr.Turns   for j=1 to len(turns)  do   t=trim(turns[j])    if t!=null  then p.(t)=lane_num	  end

	      TConns.(seg).(direc).L=NZ(TConns.(seg).(direc).L)+NZ(p.L)
	      TConns.(seg).(direc).T=NZ(TConns.(seg).(direc).T)+NZ(p.T)
	      TConns.(seg).(direc).R=NZ(TConns.(seg).(direc).R)+NZ(p.R)
	      //if  NZ(p.L)=0 and  NZ(p.T)=0  and  NZ(p.R)=0  then TConns.(seg).(direc).T=TConns.(seg).(direc).T+1

	      SetStatus(2,"读取LaneLayer, 当前索引的 segment:"+ seg, )
	      rh=GetNextRecord(lanelyr+"|",null,{{"Segment", "Ascending"}})
	end

        //     shared temp_dir	temp_dir=Project_Info.Dir+Project_Info.Name+"\\TempFile"
       //	fp=OpenFile(temp_dir+"\\seg_turns.txt","w+")
	fp=OpenFile("c:\\seg_turns_num.txt","w+")
	writeline(fp, "segID       1L    1T    1R    |     -1L   -1T    -1R")
	for i=1 to  TConns.length  do
	        idstr=TConns[i][1]
		line=idstr+"  "+string(TConns.(idstr).A.L)+"  "+string(TConns.(idstr).A.T)+"  "+string(TConns.(idstr).A.R)+"   |    "
		                        +string(TConns.(idstr).B.L)+"  "+string(TConns.(idstr).B.T)+"  "+string(TConns.(idstr).B.R)
	       writeline(fp, line)
	end


 //从Segments Layer 中读出信息，并汇总前面结果，得到 Node.link.segment,  Node.link.turn_nums， 得到变量Pnode，
 //目的：得到各个转向connectors，用于计算各转向的饱和流
        Pnode=null
	slyr=RunMacro("tsm get layer", "Segment")  SetLayer(slyr)
	rh=GetFirstRecord(slyr+"|",{{"Link", "Ascending"},{"ID", "Ascending"}})
	while rh!=null  do
	      linkIDstr=string(slyr.Link)    segIDstr=string(slyr.ID)   dir=slyr.Dir    pos=slyr.position
	      if dir>=0  then  do
		  node=string(Plink.(linkIDstr).Bnode)
		  if  pos=Plink.(linkIDstr).Segments-1   then    do  //pos 是从0开始的，故要-1
		      Pnode.(node).(linkIDstr).Segment=slyr.ID
		      Pnode.(node).(linkIDstr).TurnConns=TConns.(segIDstr).A
		  end
	      end
	      if dir<=0  then  do
		  node=string(Plink.(linkIDstr).Anode)
		  if  pos=0   then    do
		      Pnode.(node).(linkIDstr).Segment=slyr.ID
		      Pnode.(node).(linkIDstr).TurnConns=TConns.(segIDstr).B
		  end
	      end

	      SetStatus(2,"Reading SegmentLayer,  Indexed linkID:"+linkIDstr , )
	      rh=GetNextRecord(slyr+"|",null,{{"Link", "Ascending"},{"ID", "Ascending"}})
	end
	Pnode=SortArray(Pnode)   RunMacro ("CMT_Create", Plink, Pnode)
EndMacro

////////////////////////////////////////////////////////////////////////////////////////////////////////
///////////
Macro "CMT_Create"(Plink,Pnode)

 // 打开Movement.bin 文件，修改字段结构，加入相关 字段
      if  GetFileInfo("c:\\Turn_MovementTable.bin") !=null  then  return()

      vm = CreateTable("v_m", "c:\\Turn_MovementTable.bin", "FFB", {
     {"ID", "Integer", 6, null, "Yes"},      {"FromLink", "Integer", 8, null, },    {"Node", "Integer", 6, null, } ,  {"ToLink", "Integer", 8, null, },
     {"Degree_From", "Float", 8, 2, },   {"Degree_To", "Float", 8,2 , },  {"Degree", "Float", 8,2 , },
     {"Turn", "String", 5, 2, } ,   {"Direction", "String", 5, 2, },
     {"Connectors", "Integer", 8, , },  {"Saturation", "Integer", 8, , }
     })

 // 确定Movement.bin 文件中各转向类型：同时计算各转向的fromlink , tolink 的角度及夹角
      recID=0
      nlyr=RunMacro("tsm get layer", "Node")        SetLayer(nlyr)
      rh=GetFirstRecord(nlyr+"|",{{"ID", "Ascending"}})
      while rh!=null  do
           links=GetNodeLinks(nlyr.ID)
	   if nlyr.ID=3492  then
              showmessage(string(nlyr.ID))
	  for i=1  to  links.length  do
		flink=I2S(links[i])
		if   Plink.(flink).Dir>0  and   Plink.(flink).Bnode!=nlyr.ID  then Continue   //单向道路的考虑
		if   Plink.(flink).Dir<0  and   Plink.(flink).Anode!=nlyr.ID  then Continue
		for j=1  to   links.length    do
		      if   i=j  then  Continue
		      tlink=I2S(links[j])
		     if   Plink.(tlink).Dir>0  and   Plink.(tlink).Anode!=nlyr.ID  then Continue
		     if   Plink.(tlink).Dir<0  and   Plink.(tlink).Bnode!=nlyr.ID  then Continue

		    rval=RunMacro("CMT_GetRval",flink, tlink, nlyr.ID, Plink)
		    trn=rval.Turn   if  position("LTR",trn)=0  then  Continue
		    conns=Pnode.(i2s(nlyr.ID)).(flink).TurnConns.(trn)
		    if   conns=0   then   Continue   else   rval.Connectors=conns

		    rval.FromLink=links[i]    rval.ToLink=links[j]   rval.Node=nlyr.ID
		    recID=recID+1   rval.ID=recID
		    AddRecord(vm,rval)
		end
          end
	  SetStatus(2,"Adding Movement  Records,  Indexed nodeID:"+i2s(nlyr.ID) , )
	  rh=GetNextRecord(nlyr+"|",null,{{"ID", "Ascending"}})
      end

      Closeview(vm)
EndMacro

Macro "CMT_GetRval"(flink, tlink, node, Plink)
	      rval=null

              //From  link  degree
	      rval.Degree_From=Plink.(flink).degree
	      if  Plink.(flink).Bnode!=node   then   do  rval.Degree_From=Plink.(flink).degree-180
	                                                                          if    Plink.(flink).degree<0  then  rval.Degree_From=Plink.(flink).degree+180
								            end
              //To    link   degree
	      rval.Degree_To=Plink.(tlink).degree
	      if  Plink.(tlink).Anode!=node   then   do  rval.Degree_To=Plink.(tlink).degree-180
	                                                                          if    Plink.(tlink).degree<0  then  rval.Degree_To=Plink.(tlink).degree+180
									   end

              //From -To  夹角: 转换为0-180度之间
	      rval.Degree=rval.Degree_To-rval.Degree_From
	      if  rval.Degree<0       then   rval.Degree=-rval.Degree
	      if  rval.Degree>180   then   rval.Degree=360-rval.Degree

              //根据角度计算转向类型
	      if   rval.Degree >= 0 and  rval.Degree <=45  then  rval.Turn="T"
	      else  do
	               xc=rval.Degree_To-rval.Degree_From
	               if   (xc >0 and xc<180)  or xc <-180  then  rval.Turn="L"   //左转， 逆时针
	               if   (xc <0 and xc>-180) or xc >180   then  rval.Turn="R"   //右转，顺时针
		       if   xc=180  or xc=-180    then  rval.Turn="U"
	      end

              //判断转向的方向
              d1=RunMacro("Get_Dir",rval.Degree_From)  d2=RunMacro("Get_Dir",rval.Degree_To)
	      rval.Direction=d1[2]+d2[1]

              Return(rval)
EndMacro

Macro "Get_Dir"(degree)

   if degree  between   -45   and    45  then  do   dir1="E"    dir2="W"   end  //  >=-45   and <45
   if degree  between    45   and  135  then  do   dir1="N"    dir2="S"   end
   if degree  between  135   and  181  or  degree  between  -180 and  -135    then    do    dir1="W"    dir2="E"  end
   if degree  between  -135  and   -45  then   do      dir1="S"     dir2="N"   end

   Return ({dir1,dir2})
EndMacro

///////////////////////////////////////////////////////////////////////////////////////////////////////
Macro "Write_Dotnet"(obj)

    fp=openfile("c:\\dotnet.txt","w+")

    for i=1 to obj.length  do
	 name1=obj[i][1]     val1=obj.(name1)
	 writeline(fp,name1)
	 if   !RunMacro("IsCompound",val1)  then  RunMacro("Write_Value", fp,  val1, 1)
	 else  do
        	 for j=1 to val1.length  do
		       name2=val1[j][1]      val2=val1.(name2)
		       writeline(fp,"    "+name2)
		       if   !RunMacro("IsCompound",val2)  then  RunMacro("Write_Value", fp, val2, 4)
		       else   do
				 for k=1 to val2.length  do
				       name3=val2[k][1]      val3=val2.(name3)
				       writeline(fp,"        "+name3)
				       if   !RunMacro("IsCompound",val3) then   RunMacro("Write_Value",fp, val3,8)
				       else  do
				                for m=1 to val3.length  do
						       name4=val3[m][1]      val4=val3.(name4)
						       writeline(fp,"            "+name4)
						       if   !RunMacro("IsCompound",val4)  then  RunMacro("Write_Value",fp, val4,12)
				                end
				       end
				 end
                       end
	         end
         end
    end
    CloseFile(fp)
    LaunchDocument("c:\\dotnet.txt", )
EndMacro

Macro "IsCompound"(val)
   rs=0
   if  (typeof(val)="array")  and  (typeof(val[1])="array")  and (val[1].length=2)   and  (typeof(val[1][1])="string")   then  rs=1
   return (rs)
EndMacro

Macro "Write_Value"(fp, val, nspace)

    type=typeof(val)
    beforespace=null    for i=1 to  nspace+1  do   beforespace=beforespace+" "    end

    if type="string"                           then   writeline(fp, beforespace+val)
    if type="int" or type="double"  then   writeline(fp, beforespace+string(val))
    if type="array"   then  do
	    line=beforespace
	    for i=1 to val.length  do
		     if  typeof(val[i])="string"                                             then  line=line+val[i]+"/"
		     if  typeof(val[i])="int" or  typeof(val[i])="double"    then  line=line+string(val[i])+"/"
		     if  typeof(val[i])="array"   then  do
			    line=line+"{"
			    for j=1 to val[i].length  do
				   if  typeof(val[i][j])="string"    then  line=line+val[i][j]+"/"
				   else  if  typeof(val[i][j])="int" or  typeof(val[i][j])="double"    then  line=line+string(val[i][j])+"/"
					   else    line=line+"未知类型:"+typeof(val[i][j])+"/"
			    end
			    line=line+"}\n"+beforespace
		     end
		     if   typeof(val[i])<>"string"  and   typeof(val[i])<>"int"  and  typeof(val[i])<>"double"  and  typeof(val[i])<>"array"  then   line=line+"未知类型:"+typeof(val[i])+"/"
	    end
            writeline(fp, line)
    end
    if type<>"string"  and  type<>"int"  and   type<>"double"  and  type<>"array"   then   writeline(fp, beforespace+"未知类型:"+type)

EndMacro

//////////////////////////////////////////////////////////////////////////////////////////
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

  p.y=lx+dx
  p.x=ly+dy

  return(p)

endMacro


//RunProgram('cmd /c start /min "" "C:\\Program Files\\TransCAD 6.0\\tcw.exe" -a "D:\\zzx.dbd" -aI test  -N 交通影响评价应用系统 -Q',)