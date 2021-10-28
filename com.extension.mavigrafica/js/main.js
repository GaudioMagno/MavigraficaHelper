"object"!=typeof JSON&&(JSON={}),function(){"use strict";var rx_one=/^[\],:{}\s]*$/,rx_two=/\\(?:["\\\/bfnrt]|u[0-9a-fA-F]{4})/g,rx_three=/"[^"\\\n\r]*"|true|false|null|-?\d+(?:\.\d*)?(?:[eE][+\-]?\d+)?/g,rx_four=/(?:^|:|,)(?:\s*\[)+/g,rx_escapable=/[\\"\u0000-\u001f\u007f-\u009f\u00ad\u0600-\u0604\u070f\u17b4\u17b5\u200c-\u200f\u2028-\u202f\u2060-\u206f\ufeff\ufff0-\uffff]/g,rx_dangerous=/[\u0000\u00ad\u0600-\u0604\u070f\u17b4\u17b5\u200c-\u200f\u2028-\u202f\u2060-\u206f\ufeff\ufff0-\uffff]/g,gap,indent,meta,rep;function f(t){return t<10?"0"+t:t}function this_value(){return this.valueOf()}function quote(t){return rx_escapable.lastIndex=0,rx_escapable.test(t)?'"'+t.replace(rx_escapable,function(t){var e=meta[t];return"string"==typeof e?e:"\\u"+("0000"+t.charCodeAt(0).toString(16)).slice(-4)})+'"':'"'+t+'"'}function str(t,e){var r,n,o,u,f,a=gap,i=e[t];switch(i&&"object"==typeof i&&"function"==typeof i.toJSON&&(i=i.toJSON(t)),"function"==typeof rep&&(i=rep.call(e,t,i)),typeof i){case"string":return quote(i);case"number":return isFinite(i)?String(i):"null";case"boolean":case"null":return String(i);case"object":if(!i)return"null";if(gap+=indent,f=[],"[object Array]"===Object.prototype.toString.apply(i)){for(u=i.length,r=0;r<u;r+=1)f[r]=str(r,i)||"null";return o=0===f.length?"[]":gap?"[\n"+gap+f.join(",\n"+gap)+"\n"+a+"]":"["+f.join(",")+"]",gap=a,o}if(rep&&"object"==typeof rep)for(u=rep.length,r=0;r<u;r+=1)"string"==typeof rep[r]&&(o=str(n=rep[r],i))&&f.push(quote(n)+(gap?": ":":")+o);else for(n in i)Object.prototype.hasOwnProperty.call(i,n)&&(o=str(n,i))&&f.push(quote(n)+(gap?": ":":")+o);return o=0===f.length?"{}":gap?"{\n"+gap+f.join(",\n"+gap)+"\n"+a+"}":"{"+f.join(",")+"}",gap=a,o}}"function"!=typeof Date.prototype.toJSON&&(Date.prototype.toJSON=function(){return isFinite(this.valueOf())?this.getUTCFullYear()+"-"+f(this.getUTCMonth()+1)+"-"+f(this.getUTCDate())+"T"+f(this.getUTCHours())+":"+f(this.getUTCMinutes())+":"+f(this.getUTCSeconds())+"Z":null},Boolean.prototype.toJSON=this_value,Number.prototype.toJSON=this_value,String.prototype.toJSON=this_value),"function"!=typeof JSON.stringify&&(meta={"\b":"\\b","\t":"\\t","\n":"\\n","\f":"\\f","\r":"\\r",'"':'\\"',"\\":"\\\\"},JSON.stringify=function(t,e,r){var n;if(indent=gap="","number"==typeof r)for(n=0;n<r;n+=1)indent+=" ";else"string"==typeof r&&(indent=r);if((rep=e)&&"function"!=typeof e&&("object"!=typeof e||"number"!=typeof e.length))throw new Error("JSON.stringify");return str("",{"":t})}),"function"!=typeof JSON.parse&&(JSON.parse=function(text,reviver){var j;function walk(t,e){var r,n,o=t[e];if(o&&"object"==typeof o)for(r in o)Object.prototype.hasOwnProperty.call(o,r)&&(void 0!==(n=walk(o,r))?o[r]=n:delete o[r]);return reviver.call(t,e,o)}if(text=String(text),rx_dangerous.lastIndex=0,rx_dangerous.test(text)&&(text=text.replace(rx_dangerous,function(t){return"\\u"+("0000"+t.charCodeAt(0).toString(16)).slice(-4)})),rx_one.test(text.replace(rx_two,"@").replace(rx_three,"]").replace(rx_four,"")))return j=eval("("+text+")"),"function"==typeof reviver?walk({"":j},""):j;throw new SyntaxError("JSON.parse")})}();

(function () {
    'use strict';

    var path = getPath();
    var bottoniereDir = path +"/Bottoniere";
    var csInterfaceBottoniere = new CSInterface();
    var bottoniere = undefined;
    
    csInterfaceBottoniere.evalScript('getBottoniere("'+bottoniereDir+'")', function(res){
        bottoniere = JSON.parse(res)
    });
    
    setTimeout(function(){
        generaBottoni(bottoniere);
        getDocColor();
    },100);
    
    //Gestione tasto destro del mouse
    var isNS = (navigator.appName == "Netscape") ? 1 : 0;
    
    if(navigator.appName == "Netscape") document.captureEvents(Event.MOUSEDOWN||Event.MOUSEUP);
    
    function mischandler(){
        return false;
    }

    var revisione = document.getElementById("revisione");
    var nRevisioni = 11;
    for(var i=0; i<nRevisioni; i++){
        var lista = document.createElement("option");
        lista.innerHTML = i;
        lista.setAttribute("value",i);
        revisione.appendChild(lista);
    }

    //Inserisci stringa Creato da Gianluca Vitale <3
    function mousehandler(e){
        var myevent = (isNS) ? e : event;
        var eventbutton = (isNS) ? myevent.which : myevent.button;
        if((eventbutton==2)||(eventbutton==3)) return false;
    }
    document.oncontextmenu = mischandler;
    document.onmousedown = mousehandler;
    document.onmouseup = mousehandler;
    
    //Inchiostri tecnici
    
    var cs1 = new CSInterface();
    document.getElementById('logocloudflow').addEventListener('click', function() {
        cs1.openURLInDefaultBrowser("http://192.168.2.240:9090/portal.cgi/JOBS/MAVIGRAFICA/HTML/ordini.html");
    });
    
    var cs2 = new CSInterface();
    document.getElementById('logo3cx').addEventListener('click', function() {
        cs2.openURLInDefaultBrowser("https://157.90.247.97:5001/webclient#people");
    });


}());


function getPath(){
    var path = location.href;
    return path.substring(0, path.length - 11);
}

function generaBottoni(script){
    var buttonholder1 = document.getElementById("scriptHolder1");
    var buttonholder2 = document.getElementById("scriptHolder2");
    var thisButton, thisName, len = script.length;
    for(var i=0; i<len;i++){
        thisName = script[i];
        thisName = thisName.slice(thisName.lastIndexOf("/")+1, thisName.length-4).toUpperCase();
        if (thisName.startsWith("BOTTONIERA ")) thisName = thisName.slice(11, thisName.length);
        thisButton = document.createElement("BUTTON");
        thisButton.innerHTML = thisName;
        thisButton.setAttribute("class", "button-generico");
        thisButton.setAttribute("style", "margin: 5px; width: 90%");
        thisButton.setAttribute("id", script[i].toString());
        thisButton.setAttribute("onclick", "btnClick(this)");
        if (len == 1) {
            thisButton.setAttribute("style", "margin: 5px;");
            buttonholder1.appendChild(thisButton);
            buttonholder2.remove();}
        else {i%2 ? buttonholder1.appendChild(thisButton) : buttonholder2.appendChild(thisButton);}
    }
}

function btnClick(btnElement){
    var interface = new CSInterface();
    interface.evalScript('runscript("' + btnElement.id + '")');
}

function openTab(evt, tab) {
    // Declare all variables
    var i, tabcontent, tablinks;

    // Get all elements with class="tabcontent" and hide them
    tabcontent = document.getElementsByClassName("tabcontent");
    for (i = 0; i < tabcontent.length; i++) {
        tabcontent[i].style.display = "none";
    }

    // Get all elements with class="tablinks" and remove the class "active"
    tablinks = document.getElementsByClassName("tablinks");
    for (i = 0; i < tablinks.length; i++) {
        tablinks[i].className = tablinks[i].className.replace(" active", "");
    }

    // Show the current tab, and add an "active" class to the link that opened the tab
    document.getElementById(tab).style.display = "block";
    evt.currentTarget.className += " active";
}

function getDocColor(){
    var csInterface = new CSInterface();
    csInterface.evalScript('getDocColorName()',function(res){
        var nomeColore = JSON.parse(res);
        var fustella = document.getElementById("fustella");
        var film = document.getElementById("film");
        var tecniciHTML = [fustella, film];
        
        for(var i=0; i<tecniciHTML.length; i++){
            for(var e=0; e<nomeColore.length; e++){
                var lista = document.createElement("option");
                lista.innerHTML = nomeColore[e].toString();
                lista.setAttribute("value",nomeColore[e].toString());
                tecniciHTML[i].appendChild(lista);
            }
        }
    });
}

function leggiScheda(){
    var cs = new CSInterface();
    cs.evalScript('leggiSchedaJSX()', function(res){
        try{
            var elementi = ["commessa", "cliente", "stampatore", "cilindro", "barCode", "mag", "polimero", "grafico", "fascia", "passo", "profilo", "ripFascia", "ripPasso","impBase"];
            var dati = JSON.parse(res);
            if(dati.versione) dati.commessa+=dati.versione;
            var lav = dati.lavorazione.split(" ");
            switch(lav[0]){
                case ("NUOVO"):
                    dati.lavorazione = "new";
                    dati.impBase = "";
                    break;
                case ("RISTAMPA"):
                    dati.lavorazione = "rist";
                    dati.impBase = "";
                    break;
                case ("IMPIANTO"):
                    dati.lavorazione = "abb";
                    dati.impBase = lav[lav.length-1];
                    break;
                case ("VARIANTE"):
                    dati.lavorazione = "var";
                    dati.impBase = lav[lav.length-1];
                    break;
                case ("MODIFICA"):
                    dati.lavorazione = "varrist";
                    dati.impBase = lav[lav.length-1];
                    break;
            }
            if(dati.emulsione == "INTERNA") document.getElementById("interna").checked = true;
            else document.getElementById("esterna").checked = true;
            document.getElementById("lavorazione").value = dati.lavorazione;
            document.getElementById("revisione").value = dati.revisione;
            document.getElementById("fustella").value = dati.fustella;
            document.getElementById("film").value = dati.film;
            
            for(var i=0; i<elementi.length; i++){
                replaceNode(dati, elementi[i]);
            }
            
            for(var i=0; i<11; i++){
                if(i!=10) {replaceNode(dati, "lineatura"+i);
                document.getElementById("varColore"+i).checked = dati["variante"+i];
            }
                else replaceNode(dati, "lineatura", "lineatura0");
            }
            var checkGlobale = false;

            for(var i=0; i<dati.nColori-1; i++){
                if (dati["lineatura"+i] != dati["lineatura"+(++i)]) {
                    checkGlobale = true;
                    break;
                }
            }
            document.getElementById("globale").checked = checkGlobale;

        } catch(e) {
            return
        }
    });
}

function replaceNode(dati, elem, e){
    if (e==undefined || e==null) e=elem;
    var old = document.getElementById(elem);
    if(old != null){
        var newInput = old.cloneNode(true);
        newInput.setAttribute("value", dati[e]);
        old.parentNode.replaceChild(newInput, old)
    }
}

function getDatiLegenda(){
    var radiobutton = document.getElementsByName("emulsione");
    var impBase = document.getElementById("impBase").value;
    var lavorazioneValue = document.getElementById("lavorazione").value;
    var commessa = document.getElementById("commessa").value;
    var fustella = document.getElementById("fustella").value;
    var film = document.getElementById("film").value;
    var globale = document.getElementById("globale").checked;
    var versione = "-";
    
    var path = getPath();

    commessa = commessa.replaceAll(" ","");
    if (commessa.length>7){
        versione = commessa[7];
        commessa = commessa.slice(0,7)
    }

    switch(lavorazioneValue){
        case ("new"):
            var status = "Nuovo Lavoro";
            break;
        case ("rist"):
            var status = "Ristampa con modifica";
            break;
        case ("abb"):
            var status = "Impianto Abbinato "+impBase;
            break;
        case ("var"):
            var status = "Variante di "+impBase;
            break;
        case ("varrist"):
            var status = "Ristampa Variante di "+impBase;
            break;
    }
    var listaDati = ["revisione","profilo","cliente","cilindro","stampatore","barCode","mag","polimero","lineatura","grafico","ripFascia","ripPasso"];
    for(i = 0; i < radiobutton.length; i++) {
        if(radiobutton[i].checked) var checked_emu = radiobutton[i].value;
    }
    // gestione varianti
    var varianti = [];
    for(var i=0; i<10; i++){
        varianti.push(document.getElementById("varColore"+i).checked);
    }

    var lineaturaSingola = [];
    for(var i=0; i<10; i++){
        lineaturaSingola.push(document.getElementById("lineatura"+i).value);
    }
    
    var dati = {path: path, lineaturaSingola:lineaturaSingola, globale:globale, varianti: varianti, commessa:commessa, versione:versione, emulsione: checked_emu, lavorazione:status, fustella:fustella, film:film}
    for(i = 0; i < listaDati.length; i++) {
        if(document.getElementById(listaDati[i]).value!=null)
        dati[listaDati[i]]=document.getElementById(listaDati[i]).value;
        else dati[listaDati[i]]=undefined;
    }

    return dati;
}

function generaLegenda(){
    var dati = getDatiLegenda();
    var datiJSON = JSON.stringify(dati);
    var csInterface = new CSInterface();
    csInterface.evalScript(`legenda('${datiJSON}')`);
}

function bozzaComposita(){
    var csInterface = new CSInterface();
    csInterface.evalScript(`exportComp()`);
}

function setSwatchColor(){
    var csInterface = new CSInterface();
    csInterface.evalScript(`setSwatchColorJSX()`);
}

function checkcrucibel(c1,c2,cr1,cr2){
    var checkbox = document.getElementById(c1);
    var checkbox2 = document.getElementById(c2);
    var crocino = document.getElementById(cr1);
    var crocino2 = document.getElementById(cr2);
    if(checkbox.checked){
        crocino.classList.add("crucibelPressed");
        crocino2.classList.add("crucibelPressed");
        checkbox2.checked=true;
    } else {
        crocino.classList.remove("crucibelPressed");
        crocino2.classList.remove("crucibelPressed");
        checkbox2.checked=false;
    }
}

function crucibel(){
    var legenda = getDatiLegenda();

    var top = document.getElementById('top');
    var left = document.getElementById('left');
    var right = document.getElementById('right');
    var bot = document.getElementById('bot');
    var nameTOP = document.getElementById('nameTOP');
    var nameRIGHT = document.getElementById('nameRIGHT');
    var micropunti = document.getElementById('micropunti');

    var dati = {
        legenda:legenda,
        top:top.checked,
        left:left.checked,
        right:right.checked,
        bot:bot.checked,
        nameTOP:nameTOP.checked,
        nameRIGHT:nameRIGHT.checked,
        micropunti:micropunti.checked
    }

    var datiJSON = JSON.stringify(dati);

    var csInterface = new CSInterface();
    csInterface.evalScript(`crucibelJSX('${datiJSON}')`);
    
}

function quote(){
    var csInterface = new CSInterface();
    csInterface.evalScript(`quoteJSX()`);
}
