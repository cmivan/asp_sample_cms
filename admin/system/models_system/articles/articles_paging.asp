<div class="btn-toolbar">
<div class="btn-group">

<%if page<=1 then%>
<span class="btn btn-mini disabled">首页</span>
<span class="btn btn-mini disabled">上一页</span>
<%else%>
<A class="btn btn-mini" href="<%=URI.reUrl("page=1")%>">首页</A>
<A class="btn btn-mini" href="<%=URI.reUrl("page="&(page-1))%>">上一页</A>
<%end if%>
<%
dim pageFrom,pageTo
if (page - pageAdds)<=0 then
   pageFrom = 1
else
   pageFrom = page - pageAdds
end if
if (page + pageAdds)>=pageCounts then
   pageTo = pageCounts
else
   pageTo = page + pageAdds
end if

for pageI = pageFrom to pageTo
    if page = pageI then
%>
<A class="btn btn-mini active" href="<%=URI.reUrl("page="&pageI)%>"><b><%=pageI%></b></A>
<%else%>
<A class="btn btn-mini" href="<%=URI.reUrl("page="&pageI)%>"><%=pageI%></A>
<%
    end if
next
%>
<%if page>=pageCounts then%>
<span class="btn btn-mini disabled">下一页</span>
<span class="btn btn-mini disabled">尾页</span>   
<%else%>   
<A class="btn btn-mini" href="<%=URI.reUrl("page="&(page+1))%>">下一页</A>
<A class="btn btn-mini" href="<%=URI.reUrl("page="&pageCounts)%>">尾页</A>    
<%end if%>
<A class="btn btn-mini disabled">页次：<%=page%>/<%=pageCounts%>页 共 <%=allCounts%> 条记录</A>

</div></div>