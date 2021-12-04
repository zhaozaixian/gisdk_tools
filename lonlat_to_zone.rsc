Macro "CoordToZone"

 /*To get  how many tables in the AccessFile
   tbs = GetODBCTables("Ms Access Database")
   showarray(tbs)
   */
//  定义文件名称及路径
        shared   access_file, dbd_file, saved_path
        access_file="D:\\重庆项目\\重庆交通调查\\jmcx\\居民出行调查数据——巴南区.accdb"
        dbd_file    ="D:\\重庆项目\\重庆交通小区\\Zones\\zones.dbd"
	saved_path="D:\\重庆项目\\重庆交通调查\\jmcx\\"

//  Detect existing maps and views, and close all of them
        wnds=GetWindows()
	if  wnds<>null then  do
	     for i=1 to wnds[1].length  do
		   if wnds[2][i]="Map"      then   CloseMap(wnds[1][i])
		   if wnds[2][i]="Editor"   then   CloseEditor(wnds[1][i])
	     end
	end
        vws_info = GetViews()
	if vws_info<>null  then    for i=1 to  vws_info[1].length   do  closeview(vws_info[1][i])	end

//  Load the zone map for later tagging
        info = GetDBInfo(dbd_file)   dblyr= GetDBLayers(dbd_file)
	map=CreateMap( ,{{"scope",info[1]},{"Title","重庆市交通小区划分图"}} )  //打开DBD文件
        zone_lyr =AddLayer(map,  , dbd_file, dblyr[1] ,)
	SetLayer(dblyr[1])

//  Load the zone map for later tagging
        EnableProgressBar("坐标到交通小区转换", 3)     // Allow 3 progress bar
	v1=RunMacro("hu")
	v2=RunMacro("ren")
	v3=RunMacro("trips")
        DestroyProgressBar()   DestroyProgressBar()   DestroyProgressBar()

// to display the result  table
      CreateEditor("户记录表-HTable", v1 + "|", ,)     CreateEditor("人记录表-RTable", v2 + "|", ,)     CreateEditor("出行记录表-TTable", v3 + "|", ,)
      CascadeWindows()
EndMacro


//==================================HU=============================================
Macro "hu"
        shared   access_file,dbd_file,saved_path
	on    error   do     ShowMessage("错误！！\n  " + GetLastError())   return()     end

	 view = OpenTable("户记录表", "Access", {access_file,  "户记录","户编号"}, {{"Shared", "True"}})
	start_rec=GetFirstRecord(view+"|", )
	record_count = GetRecordCount(view, null)
	val_array = GetRecordsValues(view+"|", start_rec, null,  {{"户编号","Ascending"}}, record_count, "Row", )

	 strct = GetTableStructure(view)
	 strct[1][2]="integer"  strct[1][3]=8
	 for i=2 to 6    do  strct[i][2]="string"   end    strct[2][3]=16   strct[3][3]=10  strct[4][3]=8  strct[5][3]=40  strct[6][3]=22
	 for i=7 to 28  do  strct[i][2]="integer"   strct[i][3]=6   end     strct[23][2]="Float"   strct[23][3]=6   strct[23][4]=1
         view_zones = CreateTable("户记录表-Zones",saved_path+ "户记录表.bin", "FFB", strct)
	 flds=null  	 for i=1 to strct.length  do  flds=flds+{ strct[i][1]} 	 end
	 AddRecords(view_zones, flds, val_array, )

	for i = 1 to strct.length  do   strct[i] = strct[i] + {strct[i][1]}      end
	fld_zone_home={{"Home_Zone", "string", 6, 0, "False", , , , , , , null}}
	strct = InsertArrayElements(strct, 7, fld_zone_home)
        ModifyTable(view_zones, strct)                                // Modify the table

        CreateProgressBar("户记录转换................", "True")
	rec=GetFirstRecord(view_zones+"|", )   count=0
	while   rec<>null do
		view_zones.[Home_Zone]=RunMacro("lonlat_zone", view_zones.[家庭地址坐标])

                count=count+1   percent=r2i(round(count/record_count,2)*100)
		stat = UpdateProgressBar( "户记录转换................"+String(percent)+"% 完成" , percent)
		if stat = "True" then do     DestroyProgressBar()    return()    end

		rec=GetNextRecord(view_zones+"|",null,null)
	end
	return(view_zones)    Pause(1000)
EndMacro

//=================================REN=============================================
Macro "ren"
        shared   access_file,dbd_file,saved_path
	on    error   do     ShowMessage("错误！！\n  " + GetLastError())   return()     end

	view = OpenTable("人记录表", "Access", {access_file,  "人记录","人编号"}, {{"Shared", "True"}})
	start_rec=GetFirstRecord(view+"|", )
	record_count = GetRecordCount(view, null)
	val_array = GetRecordsValues(view+"|", start_rec, null,  {{"人编号","Ascending"}}, record_count, "Row", )

	 strct = GetTableStructure(view)
	 for i=1 to 10   do  strct[i][2]= "integer"  end      for i=11 to 12   do  strct[i][2]= "string"  end   //type
	 strct[1][3]=10   strct[2][3]=8    for i=3 to 10     do   strct[i][3]=5  end
	 strct[11][3]=40   strct[12][3]=22

         view_zones = CreateTable("人记录表-Zones", saved_path+"人记录表.bin", "FFB", strct)
	 flds=null  	 for i=1 to strct.length  do  flds=flds+{ strct[i][1]} 	 end
	 AddRecords(view_zones, flds, val_array, )

	for i = 1 to strct.length  do   strct[i] = strct[i] + {strct[i][1]}      end
	strct=strct+{{"Work_Zone", "string", 6, 0, "False", , , , , , , null}}
        ModifyTable(view_zones, strct)                                // Modify the table

        CreateProgressBar("人记录转换................", "True")
	rec=GetFirstRecord(view_zones+"|", )   count=0
	while   rec<>null do
		view_zones.[Work_Zone]=RunMacro("lonlat_zone", view_zones.[地址坐标])

                count=count+1   percent=r2i(round(count/record_count,2)*100)
		stat = UpdateProgressBar( "人记录转换................"+String(percent)+"% 完成" , percent)
		if stat = "True" then do     DestroyProgressBar()    return()    end

		rec=GetNextRecord(view_zones+"|",null,null)
	end
       return(view_zones)    Pause(1000)
EndMacro

//===============================TRIPS==========================================
Macro "trips"
        shared   access_file,dbd_file,saved_path
	on    error   do     ShowMessage("错误！！\n  " + GetLastError())   return()     end

	 view = OpenTable("出行记录表", "Access", {access_file,  "出行记录","出行编号"}, {{"Shared", "True"}})
	start_rec=GetFirstRecord(view+"|", )
	record_count = GetRecordCount(view, null)
	val_array = GetRecordsValues(view+"|", start_rec, null,  {{"出行编号","Ascending"}}, record_count, "Row", )

	 strct = GetTableStructure(view)
	 for i=1 to 3    do  strct[i][2]="integer"   end    strct[1][3]=12   strct[2][3]=10   strct[3][3]=4
	 for i=4 to 7    do  strct[i][2]="string"     end    strct[4][3]=20   strct[6][3]=20   strct[5][3]=22   strct[7][3]=22
	 for i=8 to 12  do  strct[i][2]="integer"   strct[i][3]=6   end

         view_zones = CreateTable("出行记录表-Zones",saved_path+ "出行记录表.bin", "FFB", strct)
	 flds=null  	 for i=1 to strct.length  do  flds=flds+{ strct[i][1]} 	 end
	 AddRecords(view_zones, flds, val_array, )

	for i = 1 to strct.length  do   strct[i] = strct[i] + {strct[i][1]}      end
	fld_zone_ori={{"Ori_Zone",   "string", 6, 0, "False", , , , , , , null}}
	fld_zone_des={{"Des_Zone", "string", 6, 0, "False", , , , , , , null}}
	strct = InsertArrayElements(strct, 6,fld_zone_ori )
	strct = InsertArrayElements(strct, 9,fld_zone_des )  //前面插入后，后移一个字段顺序
        ModifyTable(view_zones, strct)                                // Modify the table


        CreateProgressBar("出行记录转换................", "True")
	rec=GetFirstRecord(view_zones+"|", )  count=0
	while   rec<>null do
		view_zones.[Ori_Zone] =RunMacro("lonlat_zone", view_zones.[出发地坐标])
		view_zones.[Des_Zone]=RunMacro("lonlat_zone", view_zones.[到达地坐标])

                count=count+1   percent=r2i(round(count/record_count,2)*100)
		stat = UpdateProgressBar( "出行记录转换................"+String(percent)+"% 完成" , percent)
		if stat = "True" then do     DestroyProgressBar()    return()    end

		rec=GetNextRecord(view_zones+"|",null,null)
	end
       return(view_zones)    Pause(1000)
EndMacro

//=================================================================================
Macro "lonlat_zone" (lonlat)

	     ret=null   coor=null

            subs = ParseString(lonlat, ",，")
	    if   subs.length=2  then do
                 coor.lon=r2i(s2r(subs[1])*1000000)   coor.lat=r2i(s2r(subs[2])*1000000)
                 if   (coor.lon<>0) and (coor.lat<>0)   then  do
		       pt=coord(coor.lon,coor.lat)
		       rh = LocateNearestRecord(pt, 0)
		       if rh=null  then ret="未找到"   else  ret=string(rh2id(rh))
                 end
	     end

             return(ret)
endMacro

Macro "donet_read"
 //       shared   access_file,dbd_file,saved_path
        access_file="D:\\重庆项目\\重庆交通调查\\jmcx\\居民出行调查数据——巴南区.accdb"
        dbd_file    ="D:\\重庆项目\\重庆交通小区\\Zones\\zones.dbd"
	saved_path="D:\\重庆项目\\重庆交通调查\\jmcx\\"


	on    error   do     ShowMessage("错误！！\n  " + GetLastError())   return()     end

	CreateStopwatch("test my macro")

	 view = OpenTable("出行记录表", "Access", {access_file,  "出行记录","出行编号"}, {{"Shared", "True"}})
	start_rec=GetFirstRecord(view+"|", )
	record_count = GetRecordCount(view, null)
	val_array = GetRecordsValues(view+"|", start_rec, null,  {{"出行编号","Ascending"}}, record_count, "Row", )

	rec_data=null

        for i=1 to record_count  do
	  num=val_array[i][1]
	  rec_data.(string(num))=copyarray(subarray(val_array[i],2,11))
        end

time = CheckStopwatch("test my macro")
ShowMessage("The macro took  " + String(time) + " seconds.")
DestroyStopwatch("test my macro")

 endmacro




