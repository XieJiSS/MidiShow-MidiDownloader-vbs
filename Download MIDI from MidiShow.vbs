'On Error Resume Next

test_connect

'制作：解集SS
'邮箱：c141028@protonmail.com
'@TODO：根据 MidiShow编号/MidiShow网址/MIDI名称 免登录下载MidiShow上的MIDI。
'@Param: 
'日期：2017/05/22
'系统：WinXP可用，Win7+和Vista未测试
'注意：
'	1. 由于脚本联网GET数据并将数据写入磁盘，某免费杀毒软件肯定会报毒，
'	   因此采用开源形式(而不是用专门的软件编译为exe)。
'	2. 免登录下载基于MidiShow的访问权限漏洞（几乎人尽皆知），如果漏洞被修复则脚本作废。
'	3. 本脚本仅供学习、交流使用，下载后请于24小时内删除。
'   4. 本脚本遵照MIT协议开放源代码，在遵照MIT协议的约束条件的前提下，允许自由传播和修改：
'
'	MIT License
'
'	Copyright (c) 2016 - 2017 Ji Jie
'
'	Permission is hereby granted, free of charge, to any person obtaining a copy
'	of this software and associated documentation files (the "Software"), to deal
'	in the Software without restriction, including without limitation the rights
'	to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
'	copies of the Software, and to permit persons to whom the Software is
'	furnished to do so, subject to the following conditions:
'
'	The above copyright notice and this permission notice shall be included in all
'	copies or substantial portions of the Software.
'
'	THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
'	IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
'	FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
'	AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
'	LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
'	OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
'	SOFTWARE.
'
'	Thanks for using this vbscript script file. =)

Function download(midid,ff)
	Dim url
	Const Ms = "Msxml2.ServerXMLHTTP"
	Const Ad = "Adodb.Stream"
	If ff <> vbNullString Then
		url = Replace(Replace(midid,"/midi/","/midi/file/"),".html",".mid")
		midid = Replace(Split(url,"/")(UBound(Split(url,"/"))),".mid","")
	Else
		If isNumeric(midid) Then
			url = "http://www.midishow.com/midi/file/" & midid & ".mid"
		Else
			Dim Reg
			Set Reg = new RegExp
			Reg.Pattern = "[A-Za-z]+://www.midishow.com/midi/[^/s]+"
			Reg.IgnoreCase = True
			If Reg.Test(midid) = True Then
				url = Replace(Replace(midid,"/midi/","/midi/file/"),".html",".mid")
				midid = Replace(Split(url,"/")(UBound(Split(url,"/"))),".mid","")
			Else
				If midid <> vbNullString Then Search midid
				WScript.Quit()
			End If
		End If
	End If
	Const adTypeBinary = 1
	Const adSaveCreateOverWrite = 2
	Dim http,ado,fn
	Set http = CreateObject(Ms)
	http.open "GET",url,False
	http.send()
	Set ado = createobject(Ad)
	fn = name(Replace(Replace(Replace(url,".mid",".html"),"/file/","/"),"htmli","midi")) & " - " & midid & ".mid"
	Dim oHtml, oWindow, nId
	Set oHtml = CreateObject("htmlfile")
	Set oWindow = oHtml.parentWindow
	ado.Type = adTypeBinary
	ado.Open()
	ado.Write http.responseBody
	ado.SaveToFile fn
	ado.Close()
	nId = oWindow.setTimeout(GetRef("on_timeout"), 2000, "VBScript")
	MsgBox "Download Finished. Midi Saved to: " & fn, 32+4, "MidiShow MIDI下载脚本 By 解集SS"
	oWindow.clearTimeout nId
End Function

Sub on_timeout()
    CreateObject("WScript.Shell").SendKeys "{Enter}"
End Sub

Function name(path)
	Set http = CreateObject("Msxml2.ServerXMLHTTP")
	http.open "GET",path,False
	http.send()
	nm = Split(Split(http.responseText,"title>")(1)," - MIDI")(0)
	if nm = "MidiShow" Then
		name = "此Midi不存在。"
	Else
		name = nm
	End If
End Function

Function Search(text)
    On Error Resume Next
	Set http = CreateObject("Msxml2.ServerXMLHTTP")
	http.open "GET","http://www.midishow.com/search/midi?title=" & text & "&search=搜索",False
	http.send
	Dim hrt
	hrt = http.responseText
	If Err Then MsgBox Err.Description, vbCritical + VbMsgBoxSetForeground, "MidiShow MIDI下载脚本 - Error 出错了！"
	On Error Goto 0
	If UBound(Split(hrt,"<span class=" & Chr(34) & "empty" & Chr(34) & ">没有找到数据.</span>")) <> 0 Then
		MsgBox "没有搜索到您输入的MIDI名称。", vbOkCancel + vbInformation, "MidiShow MIDI下载脚本 By 解集SS"
		WScript.Quit()
	Else
		Ahtml = Split(hrt,"<a target=" & Chr(34) & "ms_p" & Chr(34) & " href=")
		If UBound(Ahtml) <> 0 Then
			morew = vbCrlf & vbCrlf & "由于作者(解集SS)懒得调试，所以至多只能显示前 20 条" & vbCrlf & "的搜索结果。"
			For i = 1 To UBound(Ahtml)
				inner = Split(Ahtml(i),"</a></h3>" & vbCrlf & "	<div class=" & Chr(34) & "c" & Chr(34) & ">")
				out = Split(Replace(inner(0),Chr(34),""),">")
				all = Replace(Replace(Split(Split(Split(hrt,"<div class=" & Chr(34) & "summary" & Chr(34) &">")(1),"</div>")(0),"共")(1),"条.","")," ","")
				star = Replace(Mid(Split(inner(1),"small><button class=" & Chr(34) & "ranks star_")(1),1,2),Chr(34),"")
				d = Split(inner(1),"<br />")(0)
				If d = vbCrlf & "	" Then d = vbCrlf & "这个作者太懒，没有留下任何描述→_→"
				d = "------------------" & vbCrlf & d & vbCrlf & vbCrlf & "------------------"
				desc = Split(d & vbCrlf & vbCrlf & Split(Replace(Split(inner(1),"</button>&nbsp;&nbsp;-&nbsp;&nbsp;")(1),"&nbsp;&nbsp;-&nbsp;&nbsp;","  "),"</small>")(0),"标签：")(0) & vbCrlf
				If MsgBox("是否下载MIDI ‘" & out(1) & "’ ?" & vbCrlf & vbCrlf & "作者描述：" & vbCrlf & Replace(desc,":","：") & "MIDI星级： " & star & "星" & vbCrlf & vbCrlf & "第 " & i & " 条，共 " & all & " 条。" & morew, 32+4,"MidiShow MIDI无积分下载 - Confirm") = 6 Then download "http://www.midishow.com" & out(0),out(1)
			Next
		Else
			error = "Error at line 125: " & vbCrlf & "    获取到的搜索结果HTML既不能说明“没有搜索到MIDI”，也不能说明“搜索到了MIDI”。"
			If MsgBox(error & vbCrlf & "可能的原因有：" & vbCrlf & "    1、MidiShow网站升级或改版，导致请求失败或无法解析网页；" & vbCrlf & "    2、网络太差，或连接超时。" & vbCrlf & vbCrlf & "如果可以，请发送此报告到 c141028@protonmail.com 。" & vbCrlf & "点击“确定”键复制报告...(调用TextBox组件，可能失败。)", vbOkCancel + vbInformation, "MidiShow MIDI下载脚本 By 解集SS - Bug Report") = 1 Then copy hrt, True, False
		End If
	End If
End Function

Dim input
input = InputBox("输入 MidiShow MIDI编号 或" & vbCrlf & "MidiShow网址 或" & vbCrlf & "MIDI 名称：" & vbCrlf & vbCrlf & "作者邮箱：c141028@protonmail.com" & vbCrlf & "欢迎大神提建议^_^","MidiShow MIDI免积分下载脚本 By 解集SS")
If input <> "" Then download input, vbNullString

Function copy(txt,mtl,alt)
	'Only availiable on WinXP
	If mtl <> True Then If mtl <> False Then mtl = True
	If alt <> True Then If alt <> False Then alt = False
	Dim Form, TextBox
	Set Form = CreateObject("Forms.Form.1")
	Set TextBox = Form.Controls.Add("Forms.TextBox.1").Object
	TextBox.MultiLine = mtl
	TextBox.Text = txt
	TextBox.SelStart = 0
	TextBox.SelLength = TextBox.TextLength
	TextBox.Copy()
	If alt Then WScript.Echo "Copy Finished."
End Function

Function test_connect()
	Dim wsobj : Set wsobj = CreateObject("WScript.Shell")
	wsobj.run "cmd /c ping /n 2 www.midishow.com || (echo 无法连接到www.midishow.com。 && echo 请检查你的网络连接。 && echo 可能是被墙了。使用VPN以解决这个问题。) | msg * /time:0",0,False
End Function
