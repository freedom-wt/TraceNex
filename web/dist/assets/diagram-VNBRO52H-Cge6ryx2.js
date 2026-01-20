import{p as w}from"./chunk-4BMEZGHF-BqgzsBVs.js";import{Z as B,q as S,r as F,H as z,J as P,t as W,v as T,_ as l,E as x,$ as v,a0 as D,R as E,a3 as _,F as A}from"./index-DJ2dMOmF.js";import{p as N}from"./radar-MK3ICKWK-BvEays-7.js";import"./semi-ui-wA5uWork.js";import"./react-core-D-iPSUlg.js";import"./tools-BfJF1Swl.js";import"./react-components-BnylIsR_.js";import"./i18n-Bw_U23tu.js";import"./_baseUniq-BC5llady.js";import"./_basePickBy-ByhE93s5.js";import"./clone-DU0nv5KA.js";var C={packet:[]},h=structuredClone(C),L=B.packet,Y=l(()=>{const t=v({...L,...D().packet});return t.showBits&&(t.paddingY+=10),t},"getConfig"),H=l(()=>h.packet,"getPacket"),I=l(t=>{t.length>0&&h.packet.push(t)},"pushWord"),M=l(()=>{E(),h=structuredClone(C)},"clear"),m={pushWord:I,getPacket:H,getConfig:Y,clear:M,setAccTitle:S,getAccTitle:F,setDiagramTitle:z,getDiagramTitle:P,getAccDescription:W,setAccDescription:T},O=1e4,R=l(t=>{w(t,m);let e=-1,o=[],s=1;const{bitsPerRow:i}=m.getConfig();for(let{start:a,end:r,label:p}of t.blocks){if(r&&r<a)throw new Error(`Packet block ${a} - ${r} is invalid. End must be greater than start.`);if(a!==e+1)throw new Error(`Packet block ${a} - ${r??a} is not contiguous. It should start from ${e+1}.`);for(e=r??a,x.debug(`Packet block ${a} - ${e} with label ${p}`);o.length<=i+1&&m.getPacket().length<O;){const[b,c]=q({start:a,end:r,label:p},s,i);if(o.push(b),b.end+1===s*i&&(m.pushWord(o),o=[],s++),!c)break;({start:a,end:r,label:p}=c)}}m.pushWord(o)},"populate"),q=l((t,e,o)=>{if(t.end===void 0&&(t.end=t.start),t.start>t.end)throw new Error(`Block start ${t.start} is greater than block end ${t.end}.`);return t.end+1<=e*o?[t,void 0]:[{start:t.start,end:e*o-1,label:t.label},{start:e*o,end:t.end,label:t.label}]},"getNextFittingBlock"),G={parse:l(async t=>{const e=await N("packet",t);x.debug(e),R(e)},"parse")},J=l((t,e,o,s)=>{const i=s.db,a=i.getConfig(),{rowHeight:r,paddingY:p,bitWidth:b,bitsPerRow:c}=a,u=i.getPacket(),n=i.getDiagramTitle(),g=r+p,d=g*(u.length+1)-(n?0:r),k=b*c+2,f=_(e);f.attr("viewbox",`0 0 ${k} ${d}`),A(f,d,k,a.useMaxWidth);for(const[$,y]of u.entries())K(f,y,$,a);f.append("text").text(n).attr("x",k/2).attr("y",d-g/2).attr("dominant-baseline","middle").attr("text-anchor","middle").attr("class","packetTitle")},"draw"),K=l((t,e,o,{rowHeight:s,paddingX:i,paddingY:a,bitWidth:r,bitsPerRow:p,showBits:b})=>{const c=t.append("g"),u=o*(s+a)+a;for(const n of e){const g=n.start%p*r+1,d=(n.end-n.start+1)*r-i;if(c.append("rect").attr("x",g).attr("y",u).attr("width",d).attr("height",s).attr("class","packetBlock"),c.append("text").attr("x",g+d/2).attr("y",u+s/2).attr("class","packetLabel").attr("dominant-baseline","middle").attr("text-anchor","middle").text(n.label),!b)continue;const k=n.end===n.start,f=u-2;c.append("text").attr("x",g+(k?d/2:0)).attr("y",f).attr("class","packetByte start").attr("dominant-baseline","auto").attr("text-anchor",k?"middle":"start").text(n.start),k||c.append("text").attr("x",g+d).attr("y",f).attr("class","packetByte end").attr("dominant-baseline","auto").attr("text-anchor","end").text(n.end)}},"drawWord"),U={draw:J},X={byteFontSize:"10px",startByteColor:"black",endByteColor:"black",labelColor:"black",labelFontSize:"12px",titleColor:"black",titleFontSize:"14px",blockStrokeColor:"black",blockStrokeWidth:"1",blockFillColor:"#efefef"},Z=l(({packet:t}={})=>{const e=v(X,t);return`
	.packetByte {
		font-size: ${e.byteFontSize};
	}
	.packetByte.start {
		fill: ${e.startByteColor};
	}
	.packetByte.end {
		fill: ${e.endByteColor};
	}
	.packetLabel {
		fill: ${e.labelColor};
		font-size: ${e.labelFontSize};
	}
	.packetTitle {
		fill: ${e.titleColor};
		font-size: ${e.titleFontSize};
	}
	.packetBlock {
		stroke: ${e.blockStrokeColor};
		stroke-width: ${e.blockStrokeWidth};
		fill: ${e.blockFillColor};
	}
	`},"styles"),it={parser:G,db:m,renderer:U,styles:Z};export{it as diagram};
