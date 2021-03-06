var woopraTracker=false;

function WoopraKeyValue(_k,_v){
    this.k=_k;
    this.v=_v;
}

function WoopraEvent(_name){

    var entries=new Array();
    this.name=_name;

    this.addProperty=function(key,value){
        entries[entries.length]=new WoopraKeyValue(key,value);
    }

    this.fire=function(){
        var t=woopraTracker;
        var buffer='';

        this.addProperty('name',this.name);

        for (var i=0;i<entries.length;i++){
            buffer+='&'+encodeURIComponent('ce_'+entries[i].k)+'='+encodeURIComponent(entries[i].v);
        }
        buffer+='&'+'alias'+'='+t.site();
        buffer+='&'+''+'cookie'+'='+t.readcookie('wooTracker');
        buffer+='&'+''+'meta'+'='+encodeURIComponent(t.meta());
        buffer+='&'+''+'screen'+'='+encodeURIComponent(t.screeninfo());
        buffer+='&'+''+'language'+'='+encodeURIComponent(t.langinfo());

        if(buffer!=''){
            var _mod = ((document.location.protocol=="https:")?'/woopras/ce.jsp?':'/ce/');
            var _url= t.getEngine() + _mod +'ra='+t.randomstring()+buffer;
            t.request(_url);
        }
    }
}


function WoopraTracker(){

    var pntr=false;
    var chat=false;

    var wx_static=false;
    var wx_engine=false;

    var alias=false;

    var visitor_data=false;
    var idle_timeout=4*60*1000;

    var verified=0;

    this.initialize=function(){

        pntr=this;

        visitor_data=new Array();

	var _c=false;
	_c=pntr.readcookie('wooTracker'); 

        if(!_c){
                _c=pntr.randomstring();
        }
        pntr.createcookie('wooTracker', _c, 10*1000);
	
	_c=pntr.readcookie('sessionCookie');
       	if(!_c){
       	       	_c=pntr.randomstring();
       	}
       	pntr.createcookie('sessionCookie',_c, -1);
 
	if(document.location.protocol=="https:"){
            wx_engine="https://sec1.woopra.com";
	    wx_static="https://static.woopra.com";
        }else{
            wx_engine='http://'+pntr.site()+'.woopra-ns.com';
            wx_static="http://static.woopra.com";
        }
        
        if(document.addEventListener){
            document.addEventListener("mousedown",pntr.clicked,false);
        }
        else{
            document.attachEvent("onmousedown",pntr.clicked);
        }
  
        if(document.addEventListener){
            document.addEventListener("mousemove",pntr.moved,false);
        }
        else{
            document.attachEvent("onmousemove",pntr.moved);
        }

    }
    this.site=function(){
	if(alias){
		return alias;
	}
	return ((location.hostname.indexOf('www.')<0)?location.hostname:location.hostname.substring(4));
    }
    this.addVisitorProperty=function(key,value){
        var cursor=visitor_data.length;
        visitor_data[cursor]=new WoopraKeyValue(key,value);
    }
    this.getStatic=function(){
        return wx_static;
    }
    this.getEngine=function(){
        return wx_engine;
    }
    this.setEngine=function(e){
        wx_engine=e;
    }
    this.setDomain=function(site){
	alias=site;
	if(document.location.protocol=="http:"){
           wx_engine='http://'+site+'.woopra-ns.com';
        }
	var _c=pntr.readcookie('wooTracker');
	if(!_c){
		_c=pntr.randomstring();
	}
        pntr.createcookie('wooTracker', _c, 10*1000);
    }

    this.sleep=function(millis){
        var date = new Date();
        var curDate = new Date();
        while(curDate-date < millis){
            curDate=new Date();
        }
    }

    this.randomstring=function(){
        var chars = "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ";
        var s = '';
        for (i = 0; i < 32; i++) {
            var rnum = Math.floor(Math.random() * chars.length);
            s += chars.substring(rnum, rnum + 1);
        }
        return s;
    }

    this.langinfo=function(){
        return (navigator.browserLanguage || navigator.language || "");
    }

    this.screeninfo=function(){
        return screen.width + 'x' + screen.height;
    }

    this.readcookie=function(k) {
        var c=""+document.cookie;
        var ind=c.indexOf(k);
        if (ind==-1 || k==""){
            return "";
        }
        var ind1=c.indexOf(';',ind);
        if (ind1==-1){
            ind1=c.length;
        }
        return unescape(c.substring(ind+k.length+1,ind1));
    }

    this.createcookie=function(k,v,days){
        var exp='';
        if(days>0){
            var expires = new Date();
            expires.setDate(expires.getDate() + days);
            exp = expires.toGMTString();
        }
        cookieval = k + '=' + v + '; ' + 'expires=' + exp + ';' + 'path=/'+';domain=.'+pntr.site();
        document.cookie = cookieval;
    }

    this.request=function(url){

        var script=document.createElement('script');
        script.type="text/javascript";
        script.src = url;

	var heads=document.getElementsByTagName('head');

	if(heads.length>0){
	        heads[0].appendChild(script);
        }else{
		document.body.appendChild(script);
	}

    }

    this.verify=function(){
        verified=1;
    }

    this.rescue=function(){
       if(verified==0){
       }
    }

    this.meta=function(){
        var meta='';
        if(pntr.readcookie('wooMeta')){
           meta=pntr.readcookie('wooMeta');
        }
        return meta;
    }
   
    this.addParam=function(arr,k,v){
	arr[arr.length]=new WoopraKeyValue(k,v);
    }

    this.track=function(_page,_title){

        var date=new Date();

        var arr=new Array();

	pntr.addParam(arr,'sessioncookie', pntr.readcookie('sessionCookie'));
	pntr.addParam(arr,'cookie', pntr.readcookie('wooTracker'));
	pntr.addParam(arr,'meta', pntr.meta());

	if(alias){
		pntr.addParam(arr,'alias',alias);
	}else{
		pntr.addParam(arr,'alias',pntr.site());
	}

	pntr.addParam(arr,'language',pntr.langinfo());

	if(_page){
		pntr.addParam(arr,'page',_page);
	}else{
		pntr.addParam(arr,'page',window.location.pathname);
	}

	if(_title){
		pntr.addParam(arr,'pagetitle',_title);
	}else{
		pntr.addParam(arr,'pagetitle',document.title);
	}


	pntr.addParam(arr,'referer',document.referrer);
	pntr.addParam(arr,'screen',pntr.screeninfo());
	pntr.addParam(arr,'localtime',date.getHours()+':'+date.getMinutes());

        var c=0;

        for (c=0;c<visitor_data.length;c++){
            pntr.addParam(arr,'cv_'+visitor_data[c].k,visitor_data[c].v);           
        }

        c=0;

        var url='';
        for (c=0;c<arr.length;c++){
            url+="&"+encodeURIComponent(arr[c].k)+"="+encodeURIComponent(arr[c].v);
        }

        var _mod = ((document.location.protocol=="https:")?'/woopras/visit.jsp?':'/visit/');

        pntr.request(wx_engine + _mod +'ra='+pntr.randomstring()+url);
    }


    this.pingServer=function(){
        var _mod = ((document.location.protocol=="https:")?'/woopras/ping.jsp?':'/ping/');
        var _url = wx_engine + _mod + 'cookie='+pntr.readcookie('wooTracker')+'&idle='+parseInt(idle/1000)+'&ra='+pntr.randomstring();
        pntr.request(_url);
    }

    this.clicked=function(e) {

        var cElem = (e.srcElement) ? e.srcElement : e.target;
        if(cElem.tagName == "A"){
            var link=cElem;
            var _download = link.pathname.match(/(?:doc|eps|jpg|png|svg|xls|ppt|pdf|xls|zip|txt|vsd|vxd|js|css|rar|exe|wma|mov|avi|wmv|mp3)($|\&)/);
            var ev=false;
            //
            if(_download && (link.href.toString().indexOf('woopra-ns.com')<0)){
		ev=new WoopraEvent(link.href);
                ev.addProperty('type','download');
                ev.fire();
                pntr.sleep(100);
            }
            if (!_download&&link.hostname != location.host && link.hostname.indexOf('javascript')==-1 && link.hostname!=''){
		ev=new WoopraEvent(link.href);
                ev.addProperty('type','exit');
                ev.fire();
                pntr.sleep(400);
            }
        }
    }

    var last_activity=new Date();
    var idle=0;

    this.moved=function(){
      last_activity=new Date();       
      idle=0;
    }

    this.setIdleTimeout=function(t){
       idle_timeout=t;
    }

    this.ping=function(){

       if(idle>idle_timeout){
 	return;
       }

       pntr.pingServer();
       var now=new Date();
       if(now-last_activity>10000){
         idle=now-last_activity;
       }
    }
    
    this.loadScript=function(src,hook){
	var script=document.createElement('script');
	script.type='text/javascript';
	script.src=src;

	var heads=document.getElementsByTagName('head');

	if(heads.length>0){
		heads[0].appendChild(script);
	}else{
		document.body.appendChild(script);
	}

	script.onload=function(){
		setTimeout(function(){hook.apply();},1000);
	}

	script.onreadystatechange = function() {
		if (this.readyState == 'complete'|| this.readyState=='loaded') {
			setTimeout(hook,1000);
		}
	}

    }
}


woopraTracker=new WoopraTracker();
woopraTracker.initialize();

