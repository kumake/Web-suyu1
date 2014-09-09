							<%
									set rs = UICon.Query("select top 9 * from user_pro order by indexid")
									do while not rs.eof
								
								%>	
							<div  class="left_fwdh">
							
							<table width="160" height="55" border="0" cellpadding="0" cellspacing="0">
							  <tr>
								<td width="70" style="padding-top:3px;padding-left:3px;">
								<a  href="product_in.asp?categoryid=<%=rs("categoryID")%>&amp;itemid=<%=rs("id")%>" title="<%=rs("title")%>"  target="_blank" ><img src="<%=rs("Field_picture")%>" width="60" height="50"/></a>
								</td>
								<td  style="vertical-align:middle">
								<p style="font-size:14px;font-weight:bold;color:#000;"><a  href="product_in.asp?categoryid=<%=rs("categoryID")%>&amp;itemid=<%=rs("id")%>" title="<%=rs("title")%>"  target="_blank" ><%=rs("title")%></a></p>
								</td>
							  </tr>
							</table>
							</div>	
									
								<%
									rs.movenext
									loop
									rs.close
									set rs=nothing
								%>