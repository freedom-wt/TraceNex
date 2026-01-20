import{p as U}from"./chunk-4BMEZGHF-C7BYB51i.js";import{aw as y,ao as z,a$ as Z,Z as j,H as q,J as H,q as J,r as K,v as Q,t as X,_ as p,E as F,R as Y,w as tt,$ as et,a3 as at,a9 as rt,F as nt}from"./index-rQEwKT5Y.js";import{p as it}from"./radar-MK3ICKWK-J8CPDksL.js";import{d as O}from"./arc-Cy7tnZAO.js";import{o as st}from"./ordinal-Cboi1Yqb.js";import"./semi-ui-wA5uWork.js";import"./react-core-D-iPSUlg.js";import"./tools-BfJF1Swl.js";import"./react-components-BnylIsR_.js";import"./i18n-Bw_U23tu.js";import"./_baseUniq-BGf3rvcq.js";import"./_basePickBy-BfznMFcO.js";import"./clone-159gv5nI.js";import"./init-Gi6I4Gst.js";function ot(t,a){return a<t?-1:a>t?1:a>=t?0:NaN}function lt(t){return t}function ct(){var t=lt,a=ot,m=null,o=y(0),u=y(z),x=y(0);function i(e){var r,l=(e=Z(e)).length,g,w,h=0,c=new Array(l),n=new Array(l),v=+o.apply(this,arguments),A=Math.min(z,Math.max(-z,u.apply(this,arguments)-v)),f,$=Math.min(Math.abs(A)/l,x.apply(this,arguments)),T=$*(A<0?-1:1),d;for(r=0;r<l;++r)(d=n[c[r]=r]=+t(e[r],r,e))>0&&(h+=d);for(a!=null?c.sort(function(S,C){return a(n[S],n[C])}):m!=null&&c.sort(function(S,C){return m(e[S],e[C])}),r=0,w=h?(A-l*T)/h:0;r<l;++r,v=f)g=c[r],d=n[g],f=v+(d>0?d*w:0)+T,n[g]={data:e[g],index:r,value:d,startAngle:v,endAngle:f,padAngle:$};return n}return i.value=function(e){return arguments.length?(t=typeof e=="function"?e:y(+e),i):t},i.sortValues=function(e){return arguments.length?(a=e,m=null,i):a},i.sort=function(e){return arguments.length?(m=e,a=null,i):m},i.startAngle=function(e){return arguments.length?(o=typeof e=="function"?e:y(+e),i):o},i.endAngle=function(e){return arguments.length?(u=typeof e=="function"?e:y(+e),i):u},i.padAngle=function(e){return arguments.length?(x=typeof e=="function"?e:y(+e),i):x},i}var P=j.pie,G={sections:new Map,showData:!1,config:P},M=G.sections,R=G.showData,pt=structuredClone(P),ut=p(()=>structuredClone(pt),"getConfig"),gt=p(()=>{M=new Map,R=G.showData,Y()},"clear"),dt=p(({label:t,value:a})=>{M.has(t)||(M.set(t,a),F.debug(`added new section: ${t}, with value: ${a}`))},"addSection"),ft=p(()=>M,"getSections"),mt=p(t=>{R=t},"setShowData"),ht=p(()=>R,"getShowData"),I={getConfig:ut,clear:gt,setDiagramTitle:q,getDiagramTitle:H,setAccTitle:J,getAccTitle:K,setAccDescription:Q,getAccDescription:X,addSection:dt,getSections:ft,setShowData:mt,getShowData:ht},vt=p((t,a)=>{U(t,a),a.setShowData(t.showData),t.sections.map(a.addSection)},"populateDb"),St={parse:p(async t=>{const a=await it("pie",t);F.debug(a),vt(a,I)},"parse")},yt=p(t=>`
  .pieCircle{
    stroke: ${t.pieStrokeColor};
    stroke-width : ${t.pieStrokeWidth};
    opacity : ${t.pieOpacity};
  }
  .pieOuterCircle{
    stroke: ${t.pieOuterStrokeColor};
    stroke-width: ${t.pieOuterStrokeWidth};
    fill: none;
  }
  .pieTitleText {
    text-anchor: middle;
    font-size: ${t.pieTitleTextSize};
    fill: ${t.pieTitleTextColor};
    font-family: ${t.fontFamily};
  }
  .slice {
    font-family: ${t.fontFamily};
    fill: ${t.pieSectionTextColor};
    font-size:${t.pieSectionTextSize};
    // fill: white;
  }
  .legend text {
    fill: ${t.pieLegendTextColor};
    font-family: ${t.fontFamily};
    font-size: ${t.pieLegendTextSize};
  }
`,"getStyles"),xt=yt,wt=p(t=>{const a=[...t.entries()].map(o=>({label:o[0],value:o[1]})).sort((o,u)=>u.value-o.value);return ct().value(o=>o.value)(a)},"createPieArcs"),At=p((t,a,m,o)=>{F.debug(`rendering pie chart
`+t);const u=o.db,x=tt(),i=et(u.getConfig(),x.pie),e=40,r=18,l=4,g=450,w=g,h=at(a),c=h.append("g");c.attr("transform","translate("+w/2+","+g/2+")");const{themeVariables:n}=x;let[v]=rt(n.pieOuterStrokeWidth);v??(v=2);const A=i.textPosition,f=Math.min(w,g)/2-e,$=O().innerRadius(0).outerRadius(f),T=O().innerRadius(f*A).outerRadius(f*A);c.append("circle").attr("cx",0).attr("cy",0).attr("r",f+v/2).attr("class","pieOuterCircle");const d=u.getSections(),S=wt(d),C=[n.pie1,n.pie2,n.pie3,n.pie4,n.pie5,n.pie6,n.pie7,n.pie8,n.pie9,n.pie10,n.pie11,n.pie12],D=st(C);c.selectAll("mySlices").data(S).enter().append("path").attr("d",$).attr("fill",s=>D(s.data.label)).attr("class","pieCircle");let W=0;d.forEach(s=>{W+=s}),c.selectAll("mySlices").data(S).enter().append("text").text(s=>(s.data.value/W*100).toFixed(0)+"%").attr("transform",s=>"translate("+T.centroid(s)+")").style("text-anchor","middle").attr("class","slice"),c.append("text").text(u.getDiagramTitle()).attr("x",0).attr("y",-400/2).attr("class","pieTitleText");const b=c.selectAll(".legend").data(D.domain()).enter().append("g").attr("class","legend").attr("transform",(s,E)=>{const k=r+l,_=k*D.domain().length/2,B=12*r,V=E*k-_;return"translate("+B+","+V+")"});b.append("rect").attr("width",r).attr("height",r).style("fill",D).style("stroke",D),b.data(S).append("text").attr("x",r+l).attr("y",r-l).text(s=>{const{label:E,value:k}=s.data;return u.getShowData()?`${E} [${k}]`:E});const L=Math.max(...b.selectAll("text").nodes().map(s=>(s==null?void 0:s.getBoundingClientRect().width)??0)),N=w+e+r+l+L;h.attr("viewBox",`0 0 ${N} ${g}`),nt(h,g,N,i.useMaxWidth)},"draw"),Ct={draw:At},Pt={parser:St,db:I,renderer:Ct,styles:xt};export{Pt as diagram};
