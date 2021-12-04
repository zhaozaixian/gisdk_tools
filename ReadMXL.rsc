Macro "ReadXML"

xdoc = CreateManagedObject("System.Xml", "System.Xml.XmlDocument",)
xdoc.PreserveWhitespace =true
xdoc.Load("c:\\classinfo.xml")

na=xdoc.CreateNavigator()
ngrs =na.SelectSingleNode("//item[@label='Functional Classes with Non-Linear Speed-Density Functions']/rows")
showmessage(ngrs.OuterXml)
rd=ngrs.ReadSubtree()
showmessage(rd.Name)


/*
 writer= CreateManagedObject("System.Xml", "System.Xml.XmlTextWriter",{"c:\\writer.xml",null})
 table =xdoc.SelectSingleNode("//item[@label='Functional Classes with Non-Linear Speed-Density Functions']/rows")
 table.WriteTo(writer)
 writer.Close()


//�ȼ�ΪFreeway���нڵ�
 row=xdoc.SelectSingleNode("//item[@label='Functional Classes with Non-Linear Speed-Density Functions']/rows/row[v/text()='Freeway']")
showmessage(row.InnerXml)
showmessage(row.InnerText)  //�����ӽڵ�text����

 //���е��ӽڵ㣺Ԫ���ı�ΪFreeway
nodes=xdoc.SelectNodes("//item[@label='Functional Classes with Non-Linear Speed-Density Functions']/rows/row/v[text()='Freeway']")
showmessage(nodes[0].OuterXml) //���ڵ㼰�����ӽڵ�xml
showmessage(nodes[0].InnerText)

 //���е��ı��ڵ�
texts=xdoc.SelectNodes("//item[@label='Functional Classes with Non-Linear Speed-Density Functions']/rows/row/v/text()")
showmessage(texts[1].Value)

 //�ض����е��ı��ڵ�
texts=xdoc.SelectNodes("//item[@label='Functional Classes with Non-Linear Speed-Density Functions']/rows/row[v/text()='Freeway']/v/text()")
for i=0  to  texts.Count-1  do  str=str+texts[i].Value+"\n" end
showmessage(str)


cols=xdoc.SelectNodes("//item[@label='Functional Classes with Non-Linear Speed-Density Functions']/cols/col")

showmessage(string(cols.Count))
for i=0 to cols.Count-1 do arr=arr+{cols[i].GetAttribute("label")}  end
showarray(arr)

rows=xdoc.SelectNodes("//item[@label='Functional Classes with Non-Linear Speed-Density Functions']/rows/row")

for i=0 to  rows.Count-1  do
        arow=null
	for j=0 to rows[i].ChildNodes.Count-1  do
		arow=arow+{rows[i].ChildNodes[j].InnerText}
	end
        allrows=allrows+{arow}
end
showarray(allrows)
*/
EndMacro


