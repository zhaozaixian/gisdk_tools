// Project:
// Project Manager: 
// Project Team members: Jandy,
//GISDK Author: Jandy
//Created Data: 2011-
// Purpose:1. 公交线路 导出 Nodes 并写到记事本当中。



Macro "Cube_Transit"

LineLayer  ="Streets"
NodeLayer  ="Nodes"
RouteLayer ="Routes"

fptr = OpenFile("C:\\Documents and Settings\\zzx\\桌面\\sz_model\\Routes.lin", "w")
writeline(fptr,";;<<PT>><<LINE>>;;")
writeline(fptr,";*  深圳市公共交通线网文件\n")

  //将公交线路上的路段站点TAG到最近的节点上
n = TagRouteStopsWithNode(RouteLayer, null, "NearstNode", 0.1)
if n<>0  then  ShowMessage("NOTE:Route "+rname+", "+I2S(n) + " stops were not tagged!") 
  	
rec=GetFirstRecord(RouteLayer+"|",null)
while rec<>null do 
	if S2I(rec)>1000  then goto next    //有错误的线路
	
  //得到一条线路的站点列表：StopNodeList
  rname = GetRouteNam(RouteLayer, s2i(rec))      //OR USE RouteLayer.[Route_Name]
	stops = GetRouteStops(RouteLayer, rname, "True")
  StopNodeList=null
	for j=1 to stops.length do   StopNodeList=StopNodeList+{stops[j][6][2]}  end   //[6][2] is the NearstNode number
		
  //得到一条线路的节点列表：RouteNodeList，加入是否站点的信息
	RouteNodeList=null		
	Links=GetRouteLinks(RouteLayer,RouteLayer.[Route_Name])

	SetLayer(LineLayer)
	for i=1 to Links.length  do
		EndPoints=GetEndPoints(Links[i][1])
		if Links[i][2]=1 then  do  fnode=EndPoints[1]  tnode=EndPoints[2]  end
		                 else  do  tnode=EndPoints[1]  fnode=EndPoints[2]  end
		                 	
  	pos_tnode=ArrayPosition(StopNodeList,{tnode},)             //每条LINK只考虑TONODE. 线路起点和终点均假定为站点
  	if  pos_tnode=0 and i<>Links.length  then  tnode=-tnode    //tnode 不是站点情况：tnode不在站点列表中，且不是最后节点，即link不是最后LINK情况
  
    if i=1 then  RouteNodeList=RouteNodeList+{fnode}+{tnode}
    	     else  RouteNodeList=RouteNodeList+{tnode}          //每条LINK只加入TONODE
  end

  //每条线路信息写出到文件
  
  head=GetRecordValues(RouteLayer,null,{"Route_name","Description","Operator","Mode","Hdway",
  	                                      "Vehicle","Seat_Cap","Crush_Cap","FareSystem"})
  writeline(fptr,"LINE NAME="  +"\""+head[1][2]+"\""+","+" LONGNAME="+"\""+head[2][2]+"\""+","+
                 " OPERATOR="  +string(head[3][2])  +","+" Mode="    +string(head[4][2])+",\n    "+
                 " HEADWAY="   +string(head[5][2])  +","+" VEHICLETYPE=" +string(head[6][2])+","+
                 " FARESYSTEM="+string(head[9][2])  +",")
   
  wlines="   N="
  for i=1 to RouteNodeList.length  do 
      wlines=wlines+RunMacro("FormatInt",RouteNodeList[i],6)
      if i<>RouteNodeList.length  then wlines=wlines+","
      if mod(i,8)=0 and  i<>RouteNodeList.length  then wlines=wlines+"\n     "
  end
  WriteLine(fptr, wlines)
 
 	next:
	rec=GetNextRecord(RouteLayer+"|",null,null)
end

CloseFile(fptr)

 
ShowMessage("成功导出到C:\\Documents and Settings\\zzx\\桌面\\sz_model\\Routes.lin")

launchdocument("C:\\Documents and Settings\\zzx\\桌面\\sz_model\\Routes.lin", )
 
EndMacro


MACRO  "FormatInt" (intval,width)
  str=string(intval)
  n=stringlength(str)
  gap=width-n
  if gap>0 then  for i=1 to gap do str=" "+str  end
  return(str)
ENDMACRO
	

//将既有ROUTE 文件转成TRANSCAD NODETABLE 格式，以用于重新生成ROUTE 线网文件
Macro "NodeTable"

LineLayer  ="Streets"
NodeLayer  ="Nodes"
RouteLayer ="Routes"

fptr = OpenFile("c:\\route.txt", "w")


rec=GetFirstRecord(RouteLayer+"|",null)
count=0
while rec<>null do 
	if S2I(rec)>3924 OR rec="3241" then goto next    //有错误的线路
		                             else count=count+1
	
 		
  //得到一条线路的RouteNodeList
	RouteNodeList=null		
	Links=GetRouteLinks(RouteLayer,RouteLayer.[Route_Name])

	SetLayer(LineLayer)
	for i=1 to Links.length  do
		EndPoints=GetEndPoints(Links[i][1])
		if Links[i][2]=1 then  do  fnode=EndPoints[1]  tnode=EndPoints[2]  end
		                 else  do  tnode=EndPoints[1]  fnode=EndPoints[2]  end
		                 	
  
    if i=1 then  RouteNodeList=RouteNodeList+{fnode}+{tnode}
    	     else  RouteNodeList=RouteNodeList+{tnode}          //每条LINK只加入TONODE
  end

  //每条线路信息写出到文件
  for i=1 to RouteNodeList.length  do 
      WriteLine(fptr, i2s(count)+","+RouteLayer.[Route_Name]+","+I2S(RouteNodeList[i]))
  end
   
 	next:
	rec=GetNextRecord(RouteLayer+"|",null,null)
end

CloseFile(fptr)

 
ShowMessage("成功导出到c:\\route.txt")

launchdocument("c:\\route.txt", )
 
EndMacro
