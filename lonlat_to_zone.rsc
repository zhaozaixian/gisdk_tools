Macro "CoordToZone"

 /*To get  how many tables in the AccessFile
   tbs = GetODBCTables("Ms Access Database")
   showarray(tbs)
   */
//  �����ļ����Ƽ�·��
        shared   access_file, dbd_file, saved_path
        access_file="D:\\������Ŀ\\���콻ͨ����\\jmcx\\������е������ݡ���������.accdb"
        dbd_file    ="D:\\������Ŀ\\���콻ͨС��\\Zones\\zones.dbd"
	saved_path="D:\\������Ŀ\\���콻ͨ����\\jmcx\\"

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
	map=CreateMap( ,{{"scope",info[1]},{"Title","�����н�ͨС������ͼ"}} )  //��DBD�ļ�
        zone_lyr =AddLayer(map,  , dbd_file, dblyr[1] ,)
	SetLayer(dblyr[1])

//  Load the zone map for later tagging
        EnableProgressBar("���굽��ͨС��ת��", 3)     // Allow 3 progress bar
	v1=RunMacro("hu")
	v2=RunMacro("ren")
	v3=RunMacro("trips")
        DestroyProgressBar()   DestroyProgressBar()   DestroyProgressBar()

// to display the result  table
      CreateEditor("����¼��-HTable", v1 + "|", ,)     CreateEditor("�˼�¼��-RTable", v2 + "|", ,)     CreateEditor("���м�¼��-TTable", v3 + "|", ,)
      CascadeWindows()
EndMacro


//==================================HU=============================================
Macro "hu"
        shared   access_file,dbd_file,saved_path
	on    error   do     ShowMessage("���󣡣�\n  " + GetLastError())   return()     end

	 view = OpenTable("����¼��", "Access", {access_file,  "����¼","�����"}, {{"Shared", "True"}})
	start_rec=GetFirstRecord(view+"|", )
	record_count = GetRecordCount(view, null)
	val_array = GetRecordsValues(view+"|", start_rec, null,  {{"�����","Ascending"}}, record_count, "Row", )

	 strct = GetTableStructure(view)
	 strct[1][2]="integer"  strct[1][3]=8
	 for i=2 to 6    do  strct[i][2]="string"   end    strct[2][3]=16   strct[3][3]=10  strct[4][3]=8  strct[5][3]=40  strct[6][3]=22
	 for i=7 to 28  do  strct[i][2]="integer"   strct[i][3]=6   end     strct[23][2]="Float"   strct[23][3]=6   strct[23][4]=1
         view_zones = CreateTable("����¼��-Zones",saved_path+ "����¼��.bin", "FFB", strct)
	 flds=null  	 for i=1 to strct.length  do  flds=flds+{ strct[i][1]} 	 end
	 AddRecords(view_zones, flds, val_array, )

	for i = 1 to strct.length  do   strct[i] = strct[i] + {strct[i][1]}      end
	fld_zone_home={{"Home_Zone", "string", 6, 0, "False", , , , , , , null}}
	strct = InsertArrayElements(strct, 7, fld_zone_home)
        ModifyTable(view_zones, strct)                                // Modify the table

        CreateProgressBar("����¼ת��................", "True")
	rec=GetFirstRecord(view_zones+"|", )   count=0
	while   rec<>null do
		view_zones.[Home_Zone]=RunMacro("lonlat_zone", view_zones.[��ͥ��ַ����])

                count=count+1   percent=r2i(round(count/record_count,2)*100)
		stat = UpdateProgressBar( "����¼ת��................"+String(percent)+"% ���" , percent)
		if stat = "True" then do     DestroyProgressBar()    return()    end

		rec=GetNextRecord(view_zones+"|",null,null)
	end
	return(view_zones)    Pause(1000)
EndMacro

//=================================REN=============================================
Macro "ren"
        shared   access_file,dbd_file,saved_path
	on    error   do     ShowMessage("���󣡣�\n  " + GetLastError())   return()     end

	view = OpenTable("�˼�¼��", "Access", {access_file,  "�˼�¼","�˱��"}, {{"Shared", "True"}})
	start_rec=GetFirstRecord(view+"|", )
	record_count = GetRecordCount(view, null)
	val_array = GetRecordsValues(view+"|", start_rec, null,  {{"�˱��","Ascending"}}, record_count, "Row", )

	 strct = GetTableStructure(view)
	 for i=1 to 10   do  strct[i][2]= "integer"  end      for i=11 to 12   do  strct[i][2]= "string"  end   //type
	 strct[1][3]=10   strct[2][3]=8    for i=3 to 10     do   strct[i][3]=5  end
	 strct[11][3]=40   strct[12][3]=22

         view_zones = CreateTable("�˼�¼��-Zones", saved_path+"�˼�¼��.bin", "FFB", strct)
	 flds=null  	 for i=1 to strct.length  do  flds=flds+{ strct[i][1]} 	 end
	 AddRecords(view_zones, flds, val_array, )

	for i = 1 to strct.length  do   strct[i] = strct[i] + {strct[i][1]}      end
	strct=strct+{{"Work_Zone", "string", 6, 0, "False", , , , , , , null}}
        ModifyTable(view_zones, strct)                                // Modify the table

        CreateProgressBar("�˼�¼ת��................", "True")
	rec=GetFirstRecord(view_zones+"|", )   count=0
	while   rec<>null do
		view_zones.[Work_Zone]=RunMacro("lonlat_zone", view_zones.[��ַ����])

                count=count+1   percent=r2i(round(count/record_count,2)*100)
		stat = UpdateProgressBar( "�˼�¼ת��................"+String(percent)+"% ���" , percent)
		if stat = "True" then do     DestroyProgressBar()    return()    end

		rec=GetNextRecord(view_zones+"|",null,null)
	end
       return(view_zones)    Pause(1000)
EndMacro

//===============================TRIPS==========================================
Macro "trips"
        shared   access_file,dbd_file,saved_path
	on    error   do     ShowMessage("���󣡣�\n  " + GetLastError())   return()     end

	 view = OpenTable("���м�¼��", "Access", {access_file,  "���м�¼","���б��"}, {{"Shared", "True"}})
	start_rec=GetFirstRecord(view+"|", )
	record_count = GetRecordCount(view, null)
	val_array = GetRecordsValues(view+"|", start_rec, null,  {{"���б��","Ascending"}}, record_count, "Row", )

	 strct = GetTableStructure(view)
	 for i=1 to 3    do  strct[i][2]="integer"   end    strct[1][3]=12   strct[2][3]=10   strct[3][3]=4
	 for i=4 to 7    do  strct[i][2]="string"     end    strct[4][3]=20   strct[6][3]=20   strct[5][3]=22   strct[7][3]=22
	 for i=8 to 12  do  strct[i][2]="integer"   strct[i][3]=6   end

         view_zones = CreateTable("���м�¼��-Zones",saved_path+ "���м�¼��.bin", "FFB", strct)
	 flds=null  	 for i=1 to strct.length  do  flds=flds+{ strct[i][1]} 	 end
	 AddRecords(view_zones, flds, val_array, )

	for i = 1 to strct.length  do   strct[i] = strct[i] + {strct[i][1]}      end
	fld_zone_ori={{"Ori_Zone",   "string", 6, 0, "False", , , , , , , null}}
	fld_zone_des={{"Des_Zone", "string", 6, 0, "False", , , , , , , null}}
	strct = InsertArrayElements(strct, 6,fld_zone_ori )
	strct = InsertArrayElements(strct, 9,fld_zone_des )  //ǰ�����󣬺���һ���ֶ�˳��
        ModifyTable(view_zones, strct)                                // Modify the table


        CreateProgressBar("���м�¼ת��................", "True")
	rec=GetFirstRecord(view_zones+"|", )  count=0
	while   rec<>null do
		view_zones.[Ori_Zone] =RunMacro("lonlat_zone", view_zones.[����������])
		view_zones.[Des_Zone]=RunMacro("lonlat_zone", view_zones.[���������])

                count=count+1   percent=r2i(round(count/record_count,2)*100)
		stat = UpdateProgressBar( "���м�¼ת��................"+String(percent)+"% ���" , percent)
		if stat = "True" then do     DestroyProgressBar()    return()    end

		rec=GetNextRecord(view_zones+"|",null,null)
	end
       return(view_zones)    Pause(1000)
EndMacro

//=================================================================================
Macro "lonlat_zone" (lonlat)

	     ret=null   coor=null

            subs = ParseString(lonlat, ",��")
	    if   subs.length=2  then do
                 coor.lon=r2i(s2r(subs[1])*1000000)   coor.lat=r2i(s2r(subs[2])*1000000)
                 if   (coor.lon<>0) and (coor.lat<>0)   then  do
		       pt=coord(coor.lon,coor.lat)
		       rh = LocateNearestRecord(pt, 0)
		       if rh=null  then ret="δ�ҵ�"   else  ret=string(rh2id(rh))
                 end
	     end

             return(ret)
endMacro

Macro "donet_read"
 //       shared   access_file,dbd_file,saved_path
        access_file="D:\\������Ŀ\\���콻ͨ����\\jmcx\\������е������ݡ���������.accdb"
        dbd_file    ="D:\\������Ŀ\\���콻ͨС��\\Zones\\zones.dbd"
	saved_path="D:\\������Ŀ\\���콻ͨ����\\jmcx\\"


	on    error   do     ShowMessage("���󣡣�\n  " + GetLastError())   return()     end

	CreateStopwatch("test my macro")

	 view = OpenTable("���м�¼��", "Access", {access_file,  "���м�¼","���б��"}, {{"Shared", "True"}})
	start_rec=GetFirstRecord(view+"|", )
	record_count = GetRecordCount(view, null)
	val_array = GetRecordsValues(view+"|", start_rec, null,  {{"���б��","Ascending"}}, record_count, "Row", )

	rec_data=null

        for i=1 to record_count  do
	  num=val_array[i][1]
	  rec_data.(string(num))=copyarray(subarray(val_array[i],2,11))
        end

time = CheckStopwatch("test my macro")
ShowMessage("The macro took  " + String(time) + " seconds.")
DestroyStopwatch("test my macro")

 endmacro




