<!--#include file="conn.asp"-->
<!--#include file="sp_inc/class_page.asp"-->
<!--#include file="plugIn/Setting.Config.asp"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title><%=config_sitename%></title>
<meta name="keywords" content="<%=config_seokeyword%>">
<meta name="description" content="<%=config_seocontent%>">
<link href="css/public.css" rel="stylesheet" type="text/css">
<style type="text/css">
<!--
.proLi{ width:160px; line-height:30px; border-bottom:#CCCCCC solid 1px; display:block; background:url(images/li.jpg) no-repeat 30px 50%; padding-left:50px; margin-left:32px;}
 -->
</style>
</head>
<body>
<div id="container">
<table id="__01" width="981" height="1145" border="0" cellpadding="0" cellspacing="0">
	<tr>
		<td>
		<!--#include file="head.asp" -->
		</td>
	</tr>
	<tr>
		<td>
		<table id="__01" width="981" height="903" border="0" cellpadding="0" cellspacing="0">
			<tr>
				<td>
				<table id="__01" width="231" height="903" border="0" cellpadding="0" cellspacing="0">
					<tr>
						<td>
							<img src="images/index_hh_01.jpg" width="231" height="48" alt=""></td>
					</tr>
					<tr>
						<td background="images/index_hh_02.jpg" width="231" height="558">
						<!--#include file="left.asp" -->
						</td>
					</tr>
					<tr>
						<td>
							<img src="images/index_hh_03.jpg" width="231" height="25" alt=""></td>
					</tr>
					<tr>
						<td>
							<img src="images/index_hh_04.jpg" width="231" height="45" alt=""></td>
					</tr>
					<tr>
						<td background="images/index_hh_05.jpg" width="231" height="185">
						<div id="index_about">
						<p>�Ͼ�����רҵ���¼ҵ������Ϊ������ȫ��������Ƽ���ָ���ۺ����λ����˾ӵ���ϸ�Ĺ����ƶȣ��Ƚ��ļ���豸�����õ������������������걻�Ͼ��ҵ���ҵЭ��ȼ�������Ϊһ��ά�޵�λ��</p>
						</div>
						</td>
					</tr>
					<tr>
						<td>
							<img src="images/index_hh_06.jpg" width="231" height="25" alt=""></td>
					</tr>
					<tr>
						<td>
							<img src="images/index_hh_07.jpg" width="231" height="17" alt=""></td>
					</tr>
				</table>
				</td>
				<td>
					<table id="__01" width="750" height="903" border="0" cellpadding="0" cellspacing="0">
						<tr>
							<td>
								<img src="images/index_ss_01.jpg" width="750" height="177" alt=""></td>
						</tr>
						<tr>
							<td>
								<img src="images/index_ss_02.jpg" width="750" height="38" alt=""></td>
						</tr>
						<tr>
							<td width="750" height="142">
									<!--���ݿ�ʼ -->
											<script src="JS/MSClass.js"></script>
									<div id="marqueediv6" style=" text-align:center;width:690px;height:170px;margin-left:13px; padding-top:5px;">
								  <table width="100%" border="0" cellspacing="0" cellpadding="0">
									  <tr>
									  
									   <%
			
							set rs = UICon.QUery("select top 20 * from user_case order by hots desc ")
							if rs.recordcount<>0 then
							do while not rs.eof
							'''''''''��ô�ֶ���''''''
							''��ҳ�����DIV+css�����Է���ʵ�������ǳ����㲻��Ҫ��ҳ��
							''ֻ��Ҫ�ı�css�� procontent ��ǩ�µ�li�Ŀ�ȼ���
							''һ�� ֻҪprocontent �Ŀ�Ⱥ�li�Ŀ��һ��
							''��Ҫ���� ֻҪ��li�Ŀ������Ϊ procontent�ļ���֮��
						   %>
										<td width="122"><table width="122" border="0" align="center" cellpadding="0" cellspacing="0"  height="135">
											<tr>
											  <td width="122"><a href="case_in.asp?categoryID=<%=rs("categoryID")%>&amp;itemid=<%=rs("id")%>"><img src="<%=rs("Field_picture")%>"  height="100" ; width="150"   border="0" style="margin-top:5px"/></a>
											  <a href="case_in.asp?categoryID=<%=rs("categoryID")%>&amp;itemid=<%=rs("id")%>" style="display:block; text-align:center; line-height:20px; color:#000; margin-top:5px;"><%=rs("title")%></a>								  </td>
											</tr>
										</table></td>
										<td width="10">&nbsp;</td>
						  <%
							rs.movenext
							loop
							end if
							%>              
									  </tr>
					  </table>
					</div>
									<script>new Marquee("marqueediv6",2,1,690,140,20,0,0)</script>
									<!--���ݽ��� -->
							</td>
						</tr>
						<tr>
							<td>
								<img src="images/index_ss_04.jpg" width="750" height="37" alt=""></td>
						</tr>
						<tr>
							<td width="750" height="165">
							<img src="images/index_ss_05.jpg" width="750" height="165" alt="">
							</td>
						</tr>
						<tr>
							<td>
								<table id="__01" width="750" height="190" border="0" cellpadding="0" cellspacing="0">
									<tr>
										<td>
										<table id="__01" width="505" height="190" border="0" cellpadding="0" cellspacing="0">
											<tr>
												<td>
													<img src="images/index_kk_01.jpg" width="505" height="35" alt=""></td>
											</tr>
											<tr>
												<td width="505" height="155">
														<ul class="index_jdcs">
															<%
																set rs = UICon.Query("select top 12 * from user_sycs order by id desc")
																rs.pagesize=6
																rs.absolutepage=1
																for i=1 to rs.pagesize
																if rs.eof then
																exit for
																end if
															
															%>
																<li>����<a  href="sycs_in.asp?categoryid=<%=rs("categoryID")%>&amp;itemid=<%=rs("id")%>" title="<%=rs("title")%>"  target="_blank" ><%						
									if len(rs("title"))>17 then	
									response.write left(rs("title"),16)&"..."
									else
									response.write rs("title")
									end if%></a></li>	
															<%
																rs.movenext
																next
																rs.close
																set rs=nothing
															%>
														</ul>
														<ul class="index_jdcs">
															<%
																set rs = UICon.Query("select top 12 * from user_sycs order by id desc")
																rs.pagesize=6
																rs.absolutepage=2
																for i=1 to rs.pagesize
																if rs.eof then
																exit for
																end if
															%>
																<li>����<a  href="sycs_in.asp?categoryid=<%=rs("categoryID")%>&amp;itemid=<%=rs("id")%>" title="<%=rs("title")%>"  target="_blank" ><%
									if len(rs("title"))>17 then	
									response.write left(rs("title"),16)&"..."
									else
									response.write rs("title")
									end if%></a></li>	
															<%
																rs.movenext
																next
																rs.close
																set rs=nothing
															%>
														</ul>
												</td>
											</tr>
										</table>
										</td>
										<td>
										<table id="__01" width="245" height="190" border="0" cellpadding="0" cellspacing="0">
											<tr>
												<td>
													<img src="images/index_f_01.jpg" width="245" height="32" alt=""></td>
											</tr>
											<tr>
												<td width="245" height="158">
														<ul id="index_fwcn">
															<li>������ϵ��Ͼ��յ�ά�޷���Ϊ�� --�ͳɱ�����Ҫ���� </li>
															<li>�����Ͼ��յ���װ���ų���Ϊ�� --��Ʒ�������ŵ���  </li>
															<li>�����Ͼ��յ��ƻ����ķ���ΪΪֹ --60���������˻��� </li>
														</ul>
												</td>
											</tr>
										</table>
										</td>
									</tr>
								</table>
							</td>
						</tr>
						<tr>
							<td>
								<img src="images/index_ss_07.jpg" width="750" height="32" alt=""></td>
						</tr>
						<tr>
							<td width="750" height="122">
							<div id="index_qymb">
								<!--���ݿ�ʼ -->
								<% singlename = VerificationUrlParam("signle","string","��ҵĿ��")
									response.Write UserSinglePage(singlename)%>
								<!--���ݽ��� -->
							</div>
							</td>
						</tr>
					</table>
				</td>
			</tr>
		</table>
		</td>
	</tr>
	<tr>
		<td>
		<!--#include file="footer.asp" -->
		</td>
	</tr>
</table>
</div>
</body>
</html>
