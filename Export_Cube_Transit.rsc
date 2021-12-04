// Project:
// Project Manager: 
// Project Team members: Jandy,
//GISDK Author: Jandy
//Created Data: 2011-
// Purpose:1. ������· ���� Nodes ��д�����±����С�



Macro "Cube_Transit"

LineLayer  ="Streets"
NodeLayer  ="Nodes"
RouteLayer ="Routes"

fptr = OpenFile("C:\\Documents and Settings\\zzx\\����\\sz_model\\Routes.lin", "w")
writeline(fptr,";;<<PT>><<LINE>>;;")
writeline(fptr,";*  �����й�����ͨ�����ļ�\n")

  //��������·�ϵ�·��վ��TAG������Ľڵ���
n = TagRouteStopsWithNode(RouteLayer, null, "NearstNode", 0.1)
if n<>0  then  ShowMessage("NOTE:Route "+rname+", "+I2S(n) + " stops were not tagged!") 
  	
rec=GetFirstRecord(RouteLayer+"|",null)
while rec<>null do 
	if S2I(rec)>1000  then goto next    //�д������·
	
  //�õ�һ����·��վ���б�StopNodeList
  rname = GetRouteNam(RouteLayer, s2i(rec))      //OR USE RouteLayer.[Route_Name]
	stops = GetRouteStops(RouteLayer, rname, "True")
  StopNodeList=null
	for j=1 to stops.length do   StopNodeList=StopNodeList+{stops[j][6][2]}  end   //[6][2] is the NearstNode number
		
  //�õ�һ����·�Ľڵ��б�RouteNodeList�������Ƿ�վ�����Ϣ
	RouteNodeList=null		
	Links=GetRouteLinks(RouteLayer,RouteLayer.[Route_Name])

	SetLayer(LineLayer)
	for i=1 to Links.length  do
		EndPoints=GetEndPoints(Links[i][1])
		if Links[i][2]=1 then  do  fnode=EndPoints[1]  tnode=EndPoints[2]  end
		                 else  do  tnode=EndPoints[1]  fnode=EndPoints[2]  end
		                 	
  	pos_tnode=ArrayPosition(StopNodeList,{tnode},)             //ÿ��LINKֻ����TONODE. ��·�����յ���ٶ�Ϊվ��
  	if  pos_tnode=0 and i<>Links.length  then  tnode=-tnode    //tnode ����վ�������tnode����վ���б��У��Ҳ������ڵ㣬��link�������LINK���
  
    if i=1 then  RouteNodeList=RouteNodeList+{fnode}+{tnode}
    	     else  RouteNodeList=RouteNodeList+{tnode}          //ÿ��LINKֻ����TONODE
  end

  //ÿ����·��Ϣд�����ļ�
  
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

 
ShowMessage("�ɹ�������C:\\Documents and Settings\\zzx\\����\\sz_model\\Routes.lin")

launchdocument("C:\\Documents and Settings\\zzx\\����\\sz_model\\Routes.lin", )
 
EndMacro


MACRO  "FormatInt" (intval,width)
  str=string(intval)
  n=stringlength(str)
  gap=width-n
  if gap>0 then  for i=1 to gap do str=" "+str  end
  return(str)
ENDMACRO
	

//������ROUTE �ļ�ת��TRANSCAD NODETABLE ��ʽ����������������ROUTE �����ļ�
Macro "NodeTable"

LineLayer  ="Streets"
NodeLayer  ="Nodes"
RouteLayer ="Routes"

fptr = OpenFile("c:\\route.txt", "w")


rec=GetFirstRecord(RouteLayer+"|",null)
count=0
while rec<>null do 
	if S2I(rec)>3924 OR rec="3241" then goto next    //�д������·
		                             else count=count+1
	
 		
  //�õ�һ����·��RouteNodeList
	RouteNodeList=null		
	Links=GetRouteLinks(RouteLayer,RouteLayer.[Route_Name])

	SetLayer(LineLayer)
	for i=1 to Links.length  do
		EndPoints=GetEndPoints(Links[i][1])
		if Links[i][2]=1 then  do  fnode=EndPoints[1]  tnode=EndPoints[2]  end
		                 else  do  tnode=EndPoints[1]  fnode=EndPoints[2]  end
		                 	
  
    if i=1 then  RouteNodeList=RouteNodeList+{fnode}+{tnode}
    	     else  RouteNodeList=RouteNodeList+{tnode}          //ÿ��LINKֻ����TONODE
  end

  //ÿ����·��Ϣд�����ļ�
  for i=1 to RouteNodeList.length  do 
      WriteLine(fptr, i2s(count)+","+RouteLayer.[Route_Name]+","+I2S(RouteNodeList[i]))
  end
   
 	next:
	rec=GetNextRecord(RouteLayer+"|",null,null)
end

CloseFile(fptr)

 
ShowMessage("�ɹ�������c:\\route.txt")

launchdocument("c:\\route.txt", )
 
EndMacro
