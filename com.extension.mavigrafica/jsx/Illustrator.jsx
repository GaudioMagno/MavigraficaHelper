"object"!=typeof JSON&&(JSON={}),function(){"use strict";var rx_one=/^[\],:{}\s]*$/,rx_two=/\\(?:["\\\/bfnrt]|u[0-9a-fA-F]{4})/g,rx_three=/"[^"\\\n\r]*"|true|false|null|-?\d+(?:\.\d*)?(?:[eE][+\-]?\d+)?/g,rx_four=/(?:^|:|,)(?:\s*\[)+/g,rx_escapable=/[\\"\u0000-\u001f\u007f-\u009f\u00ad\u0600-\u0604\u070f\u17b4\u17b5\u200c-\u200f\u2028-\u202f\u2060-\u206f\ufeff\ufff0-\uffff]/g,rx_dangerous=/[\u0000\u00ad\u0600-\u0604\u070f\u17b4\u17b5\u200c-\u200f\u2028-\u202f\u2060-\u206f\ufeff\ufff0-\uffff]/g,gap,indent,meta,rep;function f(t){return t<10?"0"+t:t}function this_value(){return this.valueOf()}function quote(t){return rx_escapable.lastIndex=0,rx_escapable.test(t)?'"'+t.replace(rx_escapable,function(t){var e=meta[t];return"string"==typeof e?e:"\\u"+("0000"+t.charCodeAt(0).toString(16)).slice(-4)})+'"':'"'+t+'"'}function str(t,e){var r,n,o,u,f,a=gap,i=e[t];switch(i&&"object"==typeof i&&"function"==typeof i.toJSON&&(i=i.toJSON(t)),"function"==typeof rep&&(i=rep.call(e,t,i)),typeof i){case"string":return quote(i);case"number":return isFinite(i)?String(i):"null";case"boolean":case"null":return String(i);case"object":if(!i)return"null";if(gap+=indent,f=[],"[object Array]"===Object.prototype.toString.apply(i)){for(u=i.length,r=0;r<u;r+=1)f[r]=str(r,i)||"null";return o=0===f.length?"[]":gap?"[\n"+gap+f.join(",\n"+gap)+"\n"+a+"]":"["+f.join(",")+"]",gap=a,o}if(rep&&"object"==typeof rep)for(u=rep.length,r=0;r<u;r+=1)"string"==typeof rep[r]&&(o=str(n=rep[r],i))&&f.push(quote(n)+(gap?": ":":")+o);else for(n in i)Object.prototype.hasOwnProperty.call(i,n)&&(o=str(n,i))&&f.push(quote(n)+(gap?": ":":")+o);return o=0===f.length?"{}":gap?"{\n"+gap+f.join(",\n"+gap)+"\n"+a+"}":"{"+f.join(",")+"}",gap=a,o}}"function"!=typeof Date.prototype.toJSON&&(Date.prototype.toJSON=function(){return isFinite(this.valueOf())?this.getUTCFullYear()+"-"+f(this.getUTCMonth()+1)+"-"+f(this.getUTCDate())+"T"+f(this.getUTCHours())+":"+f(this.getUTCMinutes())+":"+f(this.getUTCSeconds())+"Z":null},Boolean.prototype.toJSON=this_value,Number.prototype.toJSON=this_value,String.prototype.toJSON=this_value),"function"!=typeof JSON.stringify&&(meta={"\b":"\\b","\t":"\\t","\n":"\\n","\f":"\\f","\r":"\\r",'"':'\\"',"\\":"\\\\"},JSON.stringify=function(t,e,r){var n;if(indent=gap="","number"==typeof r)for(n=0;n<r;n+=1)indent+=" ";else"string"==typeof r&&(indent=r);if((rep=e)&&"function"!=typeof e&&("object"!=typeof e||"number"!=typeof e.length))throw new Error("JSON.stringify");return str("",{"":t})}),"function"!=typeof JSON.parse&&(JSON.parse=function(text,reviver){var j;function walk(t,e){var r,n,o=t[e];if(o&&"object"==typeof o)for(r in o)Object.prototype.hasOwnProperty.call(o,r)&&(void 0!==(n=walk(o,r))?o[r]=n:delete o[r]);return reviver.call(t,e,o)}if(text=String(text),rx_dangerous.lastIndex=0,rx_dangerous.test(text)&&(text=text.replace(rx_dangerous,function(t){return"\\u"+("0000"+t.charCodeAt(0).toString(16)).slice(-4)})),rx_one.test(text.replace(rx_two,"@").replace(rx_three,"]").replace(rx_four,"")))return j=eval("("+text+")"),"function"==typeof reviver?walk({"":j},""):j;throw new SyntaxError("JSON.parse")})}();

function indexOf (lista, target) {
	for (var i = 0, j = lista.length; i < j; i++) {
		if (lista[i] === target) {
			return i;
		}
	}
	return -1;
}

function legenda(datiJSON){
    var doc = app.activeDocument;
    try {
        var lvlLegenda = doc.layers["LEGENDA"];
    } catch (e) {
        var lvlLegenda = doc.layers.add();
        lvlLegenda.name = "LEGENDA";
    }
try{
    var nomeFile = doc.name.slice(0,doc.name.lastIndexOf("."));
    
    var global = JSON.parse(datiJSON);

    global.soggetto = nomeFile;
    global.lvlLegenda = lvlLegenda;

    var tecnici = [global.fustella, global.film];

    try{
        var scheda = doc.groupItems.getByName('Scheda');
        var legenda = doc.groupItems.getByName('LegendaMavi');
        legenda.selected = true;
        app.executeMenuCommand("ungroup");
        scriviScheda(global, tecnici);
        legendaGroup(lvlLegenda,lvlLegenda.pageItems);
        legenda.selected = false;
        app.selection = null;

    } catch (e){
        var legendaFile = app.open(File(global.path+"/template/Legenda_Mavi.ai"));
        app.executeMenuCommand("selectall");
        app.copy();
        app.activeDocument=doc;
        doc.activeLayer = lvlLegenda;
        app.executeMenuCommand("pasteInPlace");
        scriviScheda(global, tecnici);
        var lg = legendaGroup(lvlLegenda,lvlLegenda.pageItems);
        centraLegenda(lg);
        app.selection = null;
        legendaFile.closeNoUI();
    }
    }catch(e){alert(e)}
}

function legendaGroup(livello, oggetti) {
    var doc = app.activeDocument;
    var legenda = livello.groupItems.add();
    for (var i = oggetti.length - 1; i >= 0; i--) {
        if (oggetti[i] != legenda) oggetti[i].move(legenda, ElementPlacement.PLACEATBEGINNING);
    }
    legenda.name = "LegendaMavi";
    return legenda
}

//POSIZIONA CARTIGLIO AL CENTRO
function centraLegenda(legenda) {
    var doc = app.activeDocument;
    var lCartiglio = 885;
    var lPage = doc.artboards[0].artboardRect[2];
    if(lPage<0) lPage*=-1;
    if(lPage>lCartiglio) legenda.translate((lCartiglio-lPage)/2,0);
    else legenda.translate((lPage-lCartiglio)/2,0);
}

function scriviScheda(dati, t){
    try{
        var doc = app.activeDocument;

        var nocolor = new NoColor();
        var rosso = new CMYKColor();
        rosso.magenta = 100;
        rosso.yellow = 100;

        var mm = 2.8346438836889;

        setSwatchColorJSX();
        checkCilindro(dati.stampatore, dati.polimero, dati.cilindro, dati.path);

        dati.nColori = fillScacchiColore(t[0], t[1]);
        dati.nVarianti = 0;

        dati.fascia = (doc.width/mm).toFixed(1);
        dati.passo = (doc.height/mm).toFixed(1);

        var lin=dati.polimero.replace(/[,|.]/,"");
        dati.distorsione = "-";
        if (lin==114) dati.distorsione="-6";
        else if (lin==284) dati.distorsione="-18";
        else if (lin==254) dati.distorsione="-16";

        dati.data = data();

        //inserire qui il controllo della lineatura per singola
        for (var i=0; i<10;i++){
            if (!dati.globale) doc.groupItems.getByName("lineatura"+i).textFrames[0].contents = dati.nColori > i ? String(dati.lineatura).toUpperCase() : "";
            else doc.groupItems.getByName("lineatura"+i).textFrames[0].contents = dati.nColori > i ? String(dati.lineaturaSingola[i]).toUpperCase() : "";
        }

        for (var i=0; i<10; i++){
            var variante = doc.groupItems.getByName('variante'+i);
            variante.pathItems[0].name = dati.varianti[i] ? "on" : "off";
            variante.pathItems[0].fillColor = dati.varianti[i] ? rosso : nocolor;
            variante.textFrames[0].textRange.contents = dati.varianti[i] ? "VAR" : "";
            dati.nVarianti = dati.varianti[i] ? dati.nVarianti+=1 : dati.nVarianti;
        }
        if (dati.nVarianti == 0) dati.nVarianti = "-";

        var elementi = ["cliente", "stampatore", "cilindro", "barCode", "mag", "polimero", "grafico", "fascia", "passo", "profilo","ripFascia","ripPasso","revisione", "emulsione", "commessa", "versione", "lavorazione", "soggetto", "data", "nColori", "nVarianti", "distorsione"];
        
        //Ciclo generale
        for (var i=0; i<elementi.length;i++){
            var elemento = doc.groupItems.getByName(elementi[i]);
            elemento.textFrames[0].contents = String(dati[elementi[i]]).toUpperCase();
        } 

    }catch(e){
        alert("ERRORE! Elemento non trovato\nCancellare livello legenda e reinserirlo da capo")
    }
}


//setta la data del documento
function data() {
    var oggi = new Date();
    var mese = (oggi.getMonth() + 1);
    if (mese < 10)
        return oggi.getDate() + '/0' + mese + '/' + oggi.getFullYear();
    else
        return oggi.getDate() + '/' + mese + '/' + oggi.getFullYear();
}

function newSwatch(colore, tint){
    var doc = app.activeDocument;
    try{
        doc.swatches.getByName(colore);
    }catch(e){
        var swt = doc.swatches.add();
        swt.name = colore;
        swt.color = new CMYKColor();
        swt.color.cyan = tint[0];
        swt.color.magenta = tint[1];
        swt.color.yellow = tint[2];
        swt.color.black = tint[3];
    }
}

function leggiSchedaJSX(){
        var doc = app.activeDocument;
    try {
        var lvlLegenda = doc.layers["LEGENDA"];
        var scheda = doc.groupItems.getByName('Scheda');
        var legenda = doc.groupItems.getByName('LegendaMavi');

        var elementi = ["cliente", "stampatore", "cilindro", "barCode", "mag", "polimero", "grafico", "fascia", "passo", "profilo","ripFascia","ripPasso","revisione", "emulsione", "commessa", "versione", "lavorazione", "soggetto", "data", "nColori", "nVarianti", "distorsione"];

        var fustella = doc.groupItems.getByName('Guide').pathItems[0].name;
        var film = doc.groupItems.getByName('ColoreFilm').pathItems[0].name;

        var result = {fustella:fustella, film:film}

        for (var i=0; i<elementi.length;i++){
            var elemento = legenda.groupItems.getByName(elementi[i]);
            result[elementi[i]] = elemento.textFrames[0].contents ? elemento.textFrames[0].contents : "";
        }

        for (var i=0; i<10;i++){
            var variante = legenda.groupItems.getByName('variante'+i).pathItems[0].name;
            var elemento = legenda.groupItems.getByName("lineatura"+[i]);
            result["variante"+i] = variante == "on" ? true : false;
            result["lineatura"+i] = elemento.textFrames[0].contents ? elemento.textFrames[0].contents : "";
        }

        return JSON.stringify(result);
    } catch (e){
        return
    }
}

function setSwatchColorJSX() {
    var doc = app.activeDocument;

    if (doc.swatchGroups.length > 1){
        for (var x = 0; x <= doc.swatchGroups.length; x++) {
            doc.swatchGroups[1].remove();
        }
    }
	var sw = doc.swatches;
	//Check CMYK
	var ListaInk = doc.inkList;
	var nomiInk = [];
    for (var j = 0; j < ListaInk.length; j++) {
        if (ListaInk[j].inkInfo.printingStatus != InkPrintStatus.DISABLEINK) {
            nomiInk.push(ListaInk[j].name.replace(" quadricromia",""))
            switch (ListaInk[j].inkInfo.kind){
                    case (InkType.BLACKINK): 
                        newSwatch("Nero", [0,0,0,100]);
                        break;
                    case (InkType.CYANINK):
                        newSwatch("Cyan", [100,0,0,0]);
                        break;
                    case (InkType.YELLOWINK): 
                        newSwatch("Giallo", [0,0,100,0]);
                        break;
                    case (InkType.MAGENTAINK): 
                        newSwatch("Magenta", [0,100,0,0]);
                        break;
            }
        }
    }
	for (var x = sw.length-1; x > 1; x--) {
		if (indexOf(nomiInk, sw[x].name) == -1) {
			sw[x].remove();
		}
	}
}

function fillScacchiColore(fustella, film){
    var doc = app.activeDocument;
    var sw = doc.swatches;
    var c=0;
    var nocolor = new NoColor();

    var fustellaColore = doc.groupItems.getByName('Guide');
    fustellaColore.pathItems[0].fillColor=nocolor;

    var filmColore = doc.groupItems.getByName('ColoreFilm');
    filmColore.pathItems[0].fillColor=nocolor;

    for (var i=2; i<sw.length; i++){
        var colore = sw[i].name.replace("PANTONE","P.");
        switch (colore){
            case (fustella):
                fustellaColore.pathItems[0].fillColor=sw[i].color;
                fustellaColore.pathItems[0].name=fustella;
                break;
            case (film):
                filmColore.pathItems[0].fillColor=sw[i].color;
                filmColore.pathItems[0].name=film;
                break;
            default:
                var scacco = doc.groupItems.getByName('Colore'+c);
                var nome = doc.groupItems.getByName('nameColor'+c).textFrames[0].textRange;
                nome.contents = colore;
                for(var j=0;j<6;j++){ scacco.pathItems[j].fillColor=sw[i].color; }
                c++;
                break;
        }
    }
    for (var i=c; i<10; i++){
        var scacco = doc.groupItems.getByName('Colore'+i);
        var nome = doc.groupItems.getByName('nameColor'+i).textFrames[0].textRange;
        nome.contents = "";
        for(var j=0;j<6;j++){scacco.pathItems[j].fillColor=nocolor;}

    }
    return c;
}

//riempi dropdownlist
function getDocColorName(){
    var doc = app.activeDocument;
    var ListaInk = doc.inkList;
    var nomeColore = [];
    for (var j = 0; j < ListaInk.length; j++) {
        if (ListaInk[j].inkInfo.printingStatus != InkPrintStatus.DISABLEINK) {
            for (var i = 0; i < doc.swatches.length; i++) {
                if (ListaInk[j].name == doc.swatches[i].name) {
                    nomeColore.push(doc.swatches[i].name.replace("PANTONE", "P."));
                }
            }
            switch (ListaInk[j].inkInfo.kind){
                    case (InkType.BLACKINK): 
                        nomeColore.push("Nero");
                        break;
                    case (InkType.CYANINK):
                        nomeColore.push("Cyan");
                        break;
                    case (InkType.YELLOWINK): 
                        nomeColore.push("Giallo");
                        break;
                    case (InkType.MAGENTAINK): 
                        nomeColore.push("Magenta");
                        break;
            }
        }
    }
    return JSON.stringify(nomeColore);
}

//
function checkCilindro(cliente, polimero, cilindro, path){
    var arancio = new CMYKColor();
    arancio.yellow = 90;
    arancio.magenta = 40;

    cliente = cliente.toLowerCase();
    cliente = cliente.replace(" ","");
    cliente = cliente.substring(0,12);

    var check = app.activeDocument.pathItems.getByName('CheckCilindro');
    check.fillColor = arancio;

    switch (cliente){
        case "newdimension":
        case "ndp":
            check.fillColor = checkList("ndp", polimero, cilindro, path);
            break;
        case "bluplast":
            check.fillColor = checkList("bluplast", polimero, cilindro, path);
            break;
        case "nuovaerrepla":
        case "nep":
            check.fillColor = checkList("nep", polimero, cilindro, path);
            break;
        case "maca":
        case "macaserino":
        case "macacalvi":
            check.fillColor = checkList("maca", polimero, cilindro, path);
            break;
        case "rossetti":
        case "rossettipack":
            check.fillColor = checkList("rossetti", polimero, cilindro, path);
            break;
        case "pagani":
        case "paganiprint":
            check.fillColor = checkList("pagani", polimero, cilindro, path);
            break;
        case "marpack":
            check.fillColor = checkList("marpack", polimero, cilindro, path);
            break;
        case "gianplast":
            check.fillColor = checkList("gianplast", polimero, cilindro, path);
            break;
        case "bioplast":
            check.fillColor = checkList("bioplast", polimero, cilindro, path);
            break;
        default:
            check.fillColor = arancio;
    }
}

function checkList (nomeCl, polimero, cilindro, path){
    var rosso = new CMYKColor();
    rosso.yellow = 90;
    rosso.magenta = 90;

    var verde = new CMYKColor();
    verde.cyan = 50;
    verde.yellow = 50;

    var xmlfile = new File(path+"/Database/Cilindri.xml");
    xmlfile.open("r");
    var xmlDB = new XML(xmlfile.read());
    xmlfile.close();
    var xPol = xmlDB[nomeCl].Polimero;
    var xCil = xmlDB[nomeCl].Valore;

    var spazio = polimero.indexOf(" ");
    if (spazio > 0) {polimero = polimero.substring(0, spazio)};
    polimero = polimero.replace(",","");

    for (var i=0; i<xPol.length();i++){
        if (xPol[i].children() == polimero && xCil[i].children() == cilindro){
            return verde;
        }
    }
    return rosso;

}

//Crea bozza composita
function exportComp(){
    var doc = app.activeDocument;
    var nomeFile = doc.name.slice(0,doc.name.lastIndexOf("."));
    var activeAB = doc.artboards[doc.artboards.getActiveArtboardIndex()];

    var options = new ImageCaptureOptions();

        options.artBoardClipping = true;
        options.resolution = 300;
        options.antiAliasing = true;
        options.matte = false;
        options.horizontalScale = 100;
        options.verticalScale = 100;
        options.transparency = false;

    var png = new File("~/Desktop/"+nomeFile+".png");

    var dim = activeAB.artboardRect;
    app.executeMenuCommand("Fit Artboard to artwork bounds");
    doc.imageCapture(png, activeAB.artboardRect, options);
    activeAB.artboardRect=dim;

    var pngFile = app.open(png);
    app.executeMenuCommand("Fit Artboard to artwork bounds");

    var dim = activeAB.artboardRect;

        for(var i=0; i<dim.length; i++){
            if(dim[i]<0) dim[i]-=20;
            else dim[i]+=20;
        }
        activeAB.artboardRect=dim;
    var pdf = new File("~/Desktop/"+nomeFile+".pdf");
    var pdfOpz = new PDFSaveOptions();

        pdfOpz.compatibility = PDFCompatibility.ACROBAT5;
        pdfOpz.optimization = true;
        pdfOpz.preserveEditability = false;
        pdfOpz.generateThumbnails = false;
        pdfOpz.preserveEditability = false;
        pdfOpz.printerResolution = 300;

        pngFile.saveAs(pdf, pdfOpz);
        pngFile.close();
        png.remove();
}

//Crucibel
function crucibelJSX(datiJSON){
    var doc = app.activeDocument;
    if (doc.selection.length==1){
        var dati= JSON.parse(datiJSON);
        try {
            var crocLvl = doc.layers["CROCINI"];
        } catch (err) {
            var crocLvl = doc.layers.add();
            crocLvl.name = "CROCINI";
        }

        var obj = doc.pageItems[indexSel(doc.pageItems)];
        var dimObj = obj.controlBounds;
        var hObj = obj.height;
        var wObj = obj.width;
        var k = 19;

        if(dati.top){
            var top = creaCrocino("top", crocLvl, dati.micropunti);
            top.translate(dimObj[2]-(wObj/2),dimObj[1]+k);
        }

        if(dati.bot){
            var bottom = creaCrocino("bottom", crocLvl, dati.micropunti);
            bottom.rotate(180);
            bottom.translate(dimObj[2]-wObj/2,dimObj[3]-k);
        }
        
        if(dati.right){
            var right = creaCrocino("left", crocLvl, dati.micropunti);
            right.rotate(270);
            right.translate(dimObj[2]+k,(dimObj[1]-hObj/2)-2.81);
        }

        if(dati.left){
            var left = creaCrocino("right", crocLvl, dati.micropunti);
            left.rotate(90);
            left.translate(dimObj[0]-k,(dimObj[1]-hObj/2)-2.81);
        }

        if(dati.nameTOP || dati.nameRIGHT){
            try {
                var testi = doc.layers["DICITURE"];
            } catch (err) {
                var testi = doc.layers.add();
                testi.name = "DICITURE";

                var dic = dicituraDX(dati,testi);
                var col = coloreCrucibel(testi);

                if(dati.nameTOP){
                    col.translate((dimObj[2]-(wObj/2))-6,dimObj[1]+k+2);
                    dic.translate((dimObj[2]-(wObj/2))+6,dimObj[1]+k+2);
                } else {
                    dic.rotate(90, true, false, false, false, Transformation.BOTTOMLEFT);
                    col.rotate(90, true, false, false, false, Transformation.BOTTOMLEFT);
                    dic.translate(0,0);
                    col.translate(col.height+1.9937 ,0);
                    dic.translate(dimObj[2]+24.7,dimObj[3]+(hObj/2)+9);
                    col.translate(dimObj[2]+24.7,dimObj[3]+(hObj/2)-col.height-4.5);
                }
            } 
        }

    } else {
        alert("ATTENZIONE\nSelezionare solo un elemento o gruppo\n©Gianluca Vitale - Mavigrafica 2021");
    }
}

function indexSel(obj){
    for (var i=0; i<obj.length; i++){
        if (obj[i].selected){
            return i;
            }
    }
}

function coloreCrucibel(lvl){
    var doc = app.activeDocument;
    var ListaInk = doc.inkList;
    var gruppo = lvl.groupItems.add();
    for (var j=0; j<ListaInk.length; j++){
        if(ListaInk[j].inkInfo.printingStatus!=InkPrintStatus.DISABLEINK){
            for (var i=0;i<doc.swatches.length; i++){
                if(ListaInk[j].name==doc.swatches[i].name){
                    var nomePantone = doc.swatches[i].name.toUpperCase();
                    nomeColore(nomePantone.replace("PANTONE ","P." ),doc.swatches[i].color,gruppo);
                }
            }
            if(ListaInk[j].inkInfo.kind==InkType.BLACKINK){
                var nero = new CMYKColor();
                nero.black=100;
                nomeColore("NERO",nero,gruppo);
            }
            if(ListaInk[j].inkInfo.kind==InkType.CYANINK){
                var ciano = new CMYKColor();
                ciano.cyan=100;
                nomeColore("CIANO",ciano,gruppo);
            }
            if(ListaInk[j].inkInfo.kind==InkType.YELLOWINK){
                var giallo = new CMYKColor();
                giallo.yellow=100;
                nomeColore("GIALLO",giallo,gruppo);
            }
            if(ListaInk[j].inkInfo.kind==InkType.MAGENTAINK){
                var mag = new CMYKColor();
                mag.magenta=100;
                nomeColore("MAGENTA",mag,gruppo);
            }
        }
    }
    return gruppo;
}

function nomeColore(nome, tinta, group){
    var colore = group.textFrames.add();
    var pos = [group.geometricBounds[0]-2,group.geometricBounds[2]];
    colore.contents = nome;
    colore.paragraphs[0].paragraphAttributes.justification = Justification.RIGHT;
    colore.textRange.filled = true;
    colore.textRange.characterAttributes.fillColor = tinta;
    colore.textRange.characterAttributes.size = 6;
    colore.translate(pos[0],pos[1]);
}

function dicituraDX(dati, lvl){
    var doc = app.activeDocument;
    var nomeLavoro = lvl.textFrames.add();
    var nomeFile = doc.name.slice(0,doc.name.lastIndexOf("."));
    
    var mm = 2.8346438836889;
    var fascia = (doc.width/mm).toFixed(1);

    var passo = (doc.height/mm).toFixed(1);

    nomeLavoro.contents = nomeFile.toUpperCase()+" - Pol."+dati.legenda.polimero +" - F."+ fascia+" H."+passo+" Ø."+dati.legenda.cilindro+" "+dati.legenda.emulsione.toUpperCase()+" "+data();
    nomeLavoro.textRange.fillColor = doc.swatches[1].color;
    nomeLavoro.textRange.characterAttributes.size = 6;
    return nomeLavoro;
}

function creaCrocino(name,lvl, mp){
    var doc = app.activeDocument;
    var noColor = new NoColor();
    var CrocinoGr = lvl.groupItems.add();
    CrocinoGr.name = name;
    if(!mp){
        var Crocino = CrocinoGr.compoundPathItems.add();
        var newPath = Crocino.pathItems.add();
        newPath.setEntirePath(Array(Array(0, 0), Array(0, 7)));
        newPath = Crocino.pathItems.add();
        newPath.setEntirePath(Array(Array(-3.5, 3.5), Array(3.5, 3.5)));
        newPath.fillColor = noColor;
        newPath.stroked = true;
        newPath.strokeWidth = 0.3025;
        newPath.strokeColor = doc.swatches[1].color;
    }

    var micropunto = CrocinoGr.pathItems.ellipse(-0.88,-0.25,0.5,0.5);
    micropunto.fillColor = noColor;
    micropunto.stroked = true;
    micropunto.strokeWidth = 0.5;
    micropunto.strokeColor = doc.swatches[1].color;

    return CrocinoGr;
}

//POPOLA TAB BOTTONIERE
function getBottoniere(path){
    var folder = Folder(path);
    var files = folder.getFiles();
    var jsxFiles = [];

    for (var i=0; i<files.length; i++){
        if(files[i].name.indexOf(".jsx")!=-1){
            jsxFiles.push(files[i].fsName);
        }
    }
    return JSON.stringify(jsxFiles);
}

function runscript(script){
    var scriptFile = File(script);
    $.evalFile(scriptFile);

}

//SCRIPT QUOTE
function quoteJSX(){
    try {
        var doc = app.activeDocument;
        // Count selected items
        var selectedItems = parseInt(doc.selection.length, 10) || 0;

        var scaleFactor = app.activeDocument.scaleFactor || 1;

        // Scale
        var setScale = 0;
        var defaultScale = $.getenv("Specify_defaultScale") ? $.getenv("Specify_defaultScale") : setScale;
        // Units
        var setUnits = true;
        var defaultUnits = $.getenv("Specify_defaultUnits") ? convertToBoolean($.getenv("Specify_defaultUnits")) : setUnits;
        // Use Custom Units
        var setUseCustomUnits = false;
        var defaultUseCustomUnits = $.getenv("Specify_defaultUseCustomUnits") ? convertToBoolean($.getenv("Specify_defaultUseCustomUnits")) : setUseCustomUnits;
        // Custom Units
        var setCustomUnits = getRulerUnits();
        var defaultCustomUnits = $.getenv("Specify_defaultCustomUnits") ? $.getenv("Specify_defaultCustomUnits") : setCustomUnits;
        // Decimals
        var setDecimals = 1;
        var defaultDecimals = $.getenv("Specify_defaultDecimals") ? $.getenv("Specify_defaultDecimals") : setDecimals;
        // Font Size
        var setFontSize = 8;
        var defaultFontSize = $.getenv("Specify_defaultFontSize") ? convertToUnits($.getenv("Specify_defaultFontSize")).toFixed(3) : setFontSize;
        // Gap
        var setGap = 4;
        var defaultGap = $.getenv("Specify_defaultGap") ? $.getenv("Specify_defaultGap") : setGap;
        // Stroke width
        var setStrokeWidth = 1;
        var defaultStrokeWidth = $.getenv("Specify_defaultStrokeWidth") ? $.getenv("Specify_defaultStrokeWidth") : setStrokeWidth;
        // Head Tail Size
        var setHeadTailSize = 6;
        var defaultHeadTailSize = $.getenv("Specify_defaultHeadTailSize") ? $.getenv("Specify_defaultHeadTailSize") : setHeadTailSize;

        // SPECIFYDIALOGBOX
        var specifyDialogBox = new Window("dialog", undefined, undefined, { closeButton: false });
        specifyDialogBox.text = "Quote Mavi";
        specifyDialogBox.orientation = "row";
        specifyDialogBox.alignChildren = ["left", "top"];
        specifyDialogBox.spacing = 10;
        specifyDialogBox.margins = 16;

        // DIALOGMAINGROUP
        // ===============
        var dialogMainGroup = specifyDialogBox.add("group", undefined, { name: "dialogMainGroup" });
        dialogMainGroup.orientation = "column";
        dialogMainGroup.alignChildren = ["left", "center"];
        dialogMainGroup.spacing = 10;
        dialogMainGroup.margins = 0;

        // HORIZONTALTABBEDPANEL
        // ===================
        var horizontalTabbedPanel = dialogMainGroup.add("tabbedpanel", undefined, undefined, { name: "horizontalTabbedPanel" });
        horizontalTabbedPanel.alignChildren = "fill";
        horizontalTabbedPanel.preferredSize.width = 363.047;
        horizontalTabbedPanel.margins = 0;
        horizontalTabbedPanel.alignment = ["fill", "center"];

        // TABOPTIONS
        // ==========
        var tabOptions = horizontalTabbedPanel.add("tab", undefined, undefined, { name: "tabOptions" });
        tabOptions.text = "OPZIONI";
        tabOptions.orientation = "row";
        tabOptions.alignChildren = ["fill", "fill"];
        tabOptions.spacing = 10;
        tabOptions.margins = 10;

        // OPTIONSMAINGROUP
        // ================
        var optionsMainGroup = tabOptions.add("group", undefined, { name: "optionsMainGroup" });
        optionsMainGroup.orientation = "column";
        optionsMainGroup.alignChildren = ["fill", "top"];
        optionsMainGroup.spacing = 15;
        optionsMainGroup.margins = 0;

        // DIMENSIONPANEL
        // ==============
        var dimensionPanel = optionsMainGroup.add("panel", undefined, undefined, { name: "dimensionPanel" });
        dimensionPanel.text = "Seleziona i lati da quotare";
        dimensionPanel.orientation = "column";
        dimensionPanel.alignChildren = ["left", "top"];
        dimensionPanel.spacing = 10;
        dimensionPanel.margins = 20;

        var topCheckbox = dimensionPanel.add("checkbox", undefined, undefined, { name: "topCheckbox" });
        topCheckbox.helpTip = "Quota su bordo superiore";
        topCheckbox.text = "Superiore";
        topCheckbox.alignment = ["center", "top"];
        topCheckbox.value = false;
        topCheckbox.onClick = function () {
            topCheckbox.active = true;
            topCheckbox.active = false;

            if (!topCheckbox.value) {
                selectAllCheckbox.value = false;
            }

            activateSpecifyButton();
        };

        // DIMENSIONGROUP
        // ==============
        var dimensionGroup = dimensionPanel.add("group", undefined, { name: "dimensionGroup" });
        dimensionGroup.orientation = "row";
        dimensionGroup.alignChildren = ["center", "top"];
        dimensionGroup.spacing = 20;
        dimensionGroup.margins = 15;
        dimensionGroup.alignment = ["center", "top"];

        var leftCheckbox = dimensionGroup.add("checkbox", undefined, undefined, { name: "leftCheckbox" });
        leftCheckbox.helpTip = "Quota su bordo sinistro";
        leftCheckbox.text = "Sinistra";
        leftCheckbox.value = false;
        leftCheckbox.onClick = function () {
            leftCheckbox.active = true;
            leftCheckbox.active = false;

            if (!leftCheckbox.value) {
                selectAllCheckbox.value = false;
            }

            activateSpecifyButton();
        };

        var selectAllCheckbox = dimensionGroup.add("checkbox", undefined, undefined, { name: "selectAllCheckbox" });
        selectAllCheckbox.helpTip = "Seleziona tutti i bordi.";
        selectAllCheckbox.text = "Tutti i bordi";
        selectAllCheckbox.alignment = ["center", "top"];
        selectAllCheckbox.value = false;
        selectAllCheckbox.onClick = function () {
            selectAllCheckbox.active = true;
            selectAllCheckbox.active = false;

            if (selectAllCheckbox.value) {
                // Select All is checked
                topCheckbox.value = true;
                rightCheckbox.value = true;
                bottomCheckbox.value = true;
                leftCheckbox.value = true;
            } else {
                // Select All is unchecked
                topCheckbox.value = false;
                rightCheckbox.value = false;
                bottomCheckbox.value = false;
                leftCheckbox.value = false;
            }

            activateSpecifyButton();
        };

        var rightCheckbox = dimensionGroup.add("checkbox", undefined, undefined, { name: "rightCheckbox" });
        rightCheckbox.helpTip = "Quota su bordo destro";
        rightCheckbox.text = "Destra";
        rightCheckbox.value = false;
        rightCheckbox.onClick = function () {
            rightCheckbox.active = true;
            rightCheckbox.active = false;

            if (!rightCheckbox.value) {
                selectAllCheckbox.value = false;
            }

            activateSpecifyButton();
        };

        // DIMENSIONPANEL
        // ==============
        var bottomCheckbox = dimensionPanel.add("checkbox", undefined, undefined, { name: "bottomCheckbox" });
        bottomCheckbox.helpTip = "Quota su bordo inferiore";
        bottomCheckbox.text = "Inferiore";
        bottomCheckbox.alignment = ["center", "top"];
        bottomCheckbox.value = false;
        bottomCheckbox.onClick = function () {
            bottomCheckbox.active = true;
            bottomCheckbox.active = false;

            if (!bottomCheckbox.value) {
                selectAllCheckbox.value = false;
            }

            activateSpecifyButton();
        };



        // MULTIPLEOBJECTSPANEL
        // ====================
        var multipleObjectsPanel;
        var betweenCheckbox;

        // If exactly 2 objects are selected, give user option to dimension BETWEEN them
        if (selectedItems == 2) {
            multipleObjectsPanel = optionsMainGroup.add("panel", undefined, undefined, { name: "multipleObjectsPanel" });
            multipleObjectsPanel.text = "Oggetti multipli selezionati";
            //multipleObjectsPanel.preferredSize.height = 65;
            multipleObjectsPanel.orientation = "column";
            multipleObjectsPanel.alignChildren = ["left", "top"];
            multipleObjectsPanel.spacing = 10;
            multipleObjectsPanel.margins = 20;

            betweenCheckbox = multipleObjectsPanel.add("checkbox", undefined, undefined, { name: "betweenCheckbox" });
            betweenCheckbox.text = "Quote tra oggetti selezionati";
            betweenCheckbox.value = false;
            betweenCheckbox.onClick = function () {
                betweenCheckbox.active = true;
                betweenCheckbox.active = false;
            };
        }


        // SCALEPANEL
        // ==========
        var scalePanel = optionsMainGroup.add("panel", undefined, undefined, { name: "scalePanel" });
        scalePanel.text = "Scala";
        scalePanel.orientation = "column";
        scalePanel.alignChildren = ["left", "top"];
        scalePanel.spacing = 10;
        scalePanel.margins = 20;

        var customScaleInfo = scalePanel.add("statictext", undefined, undefined, { name: "customScaleInfo" });
        customScaleInfo.text = "Definisci la scala del documento";

        // CUSTOMSCALEGROUP
        // ================
        var customScaleGroup = scalePanel.add("group", undefined, { name: "customScaleGroup" });
        customScaleGroup.orientation = "row";
        customScaleGroup.alignChildren = ["left", "center"];
        customScaleGroup.spacing = 10;
        customScaleGroup.margins = 0;

        var customScaleLabel = customScaleGroup.add("statictext", undefined, undefined, { name: "customScaleLabel" });
        customScaleLabel.text = "Scala:";

        var customScaleDropdown_array = [];
        for (var n = 1; n <= 30; n++) {
            if (n == 1) {
                customScaleDropdown_array.push("1/" + n + "    (Default)");
                customScaleDropdown_array.push("-");
            } else {
                customScaleDropdown_array.push("1/" + n);
            }
        }

        var customScaleDropdown = customScaleGroup.add("dropdownlist", undefined, undefined, { name: "customScaleDropdown", items: customScaleDropdown_array });
        customScaleDropdown.selection = defaultScale;
        customScaleDropdown.onChange = function () {
            restoreDefaultsButton.enabled = true;
        };

        // SCALEPANEL
        // ==========
        var scaleDivider = scalePanel.add("panel", undefined, undefined, { name: "scaleDivider" });
        scaleDivider.alignment = "fill";

        var customScaleExample = scalePanel.add("statictext", undefined, undefined, { name: "customScaleExample" });
        customScaleExample.text = "Esempio:";
        var customScaleExample2 = scalePanel.add("statictext", undefined, undefined, { name: "customScaleExample" });
        customScaleExample2.text = "250 unità ad 1/4 verrà espresso come 1000";

        // TABSTYLES
        // =========
        var tabStyles = horizontalTabbedPanel.add("tab", undefined, undefined, { name: "tabStyles" });
        tabStyles.text = "STILE";
        tabStyles.orientation = "column";
        tabStyles.alignChildren = ["fill", "fill"];
        tabStyles.spacing = 10;
        tabStyles.margins = 10;

        // LABELSTYLESPANEL
        // ============
        var labelStylesPanel = tabStyles.add("panel", undefined, undefined, { name: "labelStylesPanel" });
        labelStylesPanel.text = "Stile etichetta";
        labelStylesPanel.orientation = "column";
        labelStylesPanel.alignChildren = ["fill", "top"];
        labelStylesPanel.spacing = 10;
        labelStylesPanel.margins = 20;

        var units = labelStylesPanel.add("checkbox", undefined, undefined, { name: "units" });
        units.text = "Includi unità di misura";
        units.value = defaultUnits;
        units.onClick = function () {
            restoreDefaultsButton.enabled = true;

            units.active = true;
            units.active = false;

            if (units.value == false) {
                useCustomUnits.value = false;
                useCustomUnits.enabled = false;
                customUnitsInput.text = getRulerUnits();
                customUnitsInput.enabled = false;
            } else {
                useCustomUnits.enabled = true;
            }
        };

        // CUSTOMIZEUNITSGROUP
        // ===================
        var customizeUnitsGroup = labelStylesPanel.add("group", undefined, { name: "customizeUnitsGroup" });
        customizeUnitsGroup.orientation = "row";
        customizeUnitsGroup.alignChildren = ["left", "center"];
        customizeUnitsGroup.spacing = 10;
        customizeUnitsGroup.margins = 0;

        var useCustomUnits = customizeUnitsGroup.add("checkbox", undefined, undefined, { name: "useCustomUnits" });
        useCustomUnits.text = "Unità del testo";
        useCustomUnits.value = defaultUseCustomUnits;
        if (units.value == false) {
            useCustomUnits.value = false;
            useCustomUnits.enabled = false;
        } else {
            useCustomUnits.enabled = true;
        }
        useCustomUnits.onClick = function () {
            restoreDefaultsButton.enabled = true;
            useCustomUnits.active = true;
            useCustomUnits.active = false;

            if (useCustomUnits.value == true) {
                customUnitsInput.enabled = true;
            } else {
                customUnitsInput.text = getRulerUnits();
                customUnitsInput.enabled = false;
            }
        };

        var customUnitsInput = customizeUnitsGroup.add('edittext {properties: {name: "customUnitsInput"}}');
        customUnitsInput.text = defaultCustomUnits;
        customUnitsInput.enabled = defaultUseCustomUnits;
        customUnitsInput.characters = 20;
        customUnitsInput.preferredSize.width = 120;
        if (useCustomUnits.value == true) {
            customUnitsInput.enabled = true;
        } else {
            customUnitsInput.enabled = false;
        }
        customUnitsInput.onChanging = function () {
            restoreDefaultsButton.enabled = true;
        };
        customUnitsInput.onDeactivate = function () {
            customUnitsInput.text = customUnitsInput.text.replace(/\s+/g, ''); // trim input text
            customUnitsInput.text = customUnitsInput.text.replace(/[^ a-zA-Z]/g, "");
        };


        // DECIMALPLACESGROUP
        // ==================
        var decimalPlacesGroup = labelStylesPanel.add("group", undefined, { name: "decimalPlacesGroup" });
        decimalPlacesGroup.orientation = "row";
        decimalPlacesGroup.alignChildren = ["left", "center"];
        decimalPlacesGroup.spacing = 2;
        decimalPlacesGroup.margins = 0;

        var decimalPlacesLabel = decimalPlacesGroup.add("statictext", undefined, undefined, { name: "decimalPlacesLabel" });
        decimalPlacesLabel.text = "Decimali:";

        var decimalPlacesInput = decimalPlacesGroup.add('edittext {justify: "right", properties: {name: "decimalPlacesInput"}}');
        decimalPlacesInput.characters = 1;
        decimalPlacesInput.preferredSize.width = 40;
        decimalPlacesInput.text = defaultDecimals;
        decimalPlacesInput.onChanging = function () {
            restoreDefaultsButton.enabled = true;
            decimalPlacesInput.text = decimalPlacesInput.text.replace(/[^0-9]/g, "");
        };

        // FONTGROUP
        // =========
        var fontGroup = labelStylesPanel.add("group", undefined, { name: "fontGroup" });
        fontGroup.orientation = "row";
        fontGroup.alignChildren = ["left", "center"];
        fontGroup.spacing = 2;
        fontGroup.margins = 0;

        var fontLabel = fontGroup.add("statictext", undefined, undefined, { name: "fontLabel" });
        fontLabel.text = "Dimensione Font:";

        var fontSizeInput = fontGroup.add('edittext {justify: "right", properties: {name: "fontSizeInput"}}');
        fontSizeInput.text = defaultFontSize;
        fontSizeInput.characters = 5;
        fontSizeInput.preferredSize.width = 60;
        fontSizeInput.onChanging = function () {
            restoreDefaultsButton.enabled = true;
        }
        fontSizeInput.onDeactivate = function () {
            fontSizeInput.text = fontSizeInput.text.replace(/\s+/g, ''); // trim input text
            // If first character is decimal point, don't error, but instead add leading zero to string.
            if (fontSizeInput.text.charAt(0) == ".") {
                fontSizeInput.text = "0" + fontSizeInput.text;
                fontSizeInput.active = true;
            }
        }

        var fontUnitsLabelText = fontGroup.add("statictext", undefined, undefined, { name: "fontUnitsLabelText" });
        fontUnitsLabelText.text = getRulerUnits();

        // LINESTYLESPANEL
        // ============
        var lineStylesPanel = tabStyles.add("panel", undefined, undefined, { name: "lineStylesPanel" });
        lineStylesPanel.text = "Stile linea";
        lineStylesPanel.orientation = "column";
        lineStylesPanel.alignChildren = ["fill", "top"];
        lineStylesPanel.spacing = 10;
        lineStylesPanel.margins = 20;

        // GAPGROUP
        // ========
        var gapGroup = lineStylesPanel.add("group", undefined, { name: "gapGroup" });
        gapGroup.orientation = "row";
        gapGroup.alignChildren = ["left", "center"];
        gapGroup.spacing = 2;
        gapGroup.margins = 0;

        var gapLabel = gapGroup.add("statictext", undefined, undefined, { name: "gapLabel" });
        gapLabel.text = "Gap tra oggetto e quota:";

        var gapInput = gapGroup.add('edittext {justify: "right", properties: {name: "gapInput"}}');
        gapInput.characters = 6;
        gapInput.preferredSize.width = 60;
        gapInput.text = defaultGap;
        gapInput.onChanging = function () {
            restoreDefaultsButton.enabled = true;
        };
        gapInput.onDeactivate = function () {
            gapInput.text = gapInput.text.replace(/\s+/g, ''); // trim input text
            gapInput.text = gapInput.text.replace(/[^0-9\.]/g, "");
            if (gapInput.text.charAt(0) == ".") {
                gapInput.text = "0" + gapInput.text;
            }
        }

        var gapUnitsLabelText = gapGroup.add("statictext", undefined, undefined, { name: "gapUnitsLabelText" });
        gapUnitsLabelText.text = getRulerUnits();

        // STROKEWIDTHGROUP
        // ========
        var strokeWidthGroup = lineStylesPanel.add("group", undefined, { name: "strokeWidthGroup" });
        strokeWidthGroup.orientation = "row";
        strokeWidthGroup.alignChildren = ["left", "center"];
        strokeWidthGroup.spacing = 2;
        strokeWidthGroup.margins = 0;

        var strokeWidthLabel = strokeWidthGroup.add("statictext", undefined, undefined, { name: "strokeWidthLabel" });
        strokeWidthLabel.text = "Spessore Traccia:";

        var strokeWidthInput = strokeWidthGroup.add('edittext {justify: "right", properties: {name: "strokeWidthInput"}}');
        strokeWidthInput.characters = 6;
        strokeWidthInput.preferredSize.width = 60;
        strokeWidthInput.text = defaultStrokeWidth;
        strokeWidthInput.onChanging = function () {
            restoreDefaultsButton.enabled = true;
        };
        strokeWidthInput.onDeactivate = function () {
            strokeWidthInput.text = strokeWidthInput.text.replace(/\s+/g, ''); // trim input text
            strokeWidthInput.text = strokeWidthInput.text.replace(/[^0-9\.]/g, "");
            // If first character is decimal point, don't error, but instead add leading zero to string.
            if (strokeWidthInput.text.charAt(0) == ".") {
                strokeWidthInput.text = "0" + strokeWidthInput.text;
            }
        }

        var strokeWidthUnitsLabelText = strokeWidthGroup.add("statictext", undefined, undefined, { name: "strokeWidthUnitsLabelText" });
        strokeWidthUnitsLabelText.text = getRulerUnits();

        // HEADTAILSIZEGROUP
        // ========
        var headTailSizeGroup = lineStylesPanel.add("group", undefined, { name: "headTailSizeGroup" });
        headTailSizeGroup.orientation = "row";
        headTailSizeGroup.alignChildren = ["left", "center"];
        headTailSizeGroup.spacing = 2;
        headTailSizeGroup.margins = 0;

        var headTailSizeLabel = headTailSizeGroup.add("statictext", undefined, undefined, { name: "headTailSizeLabel" });
        headTailSizeLabel.text = "Lunghezza estremi:";

        var headTailSizeInput = headTailSizeGroup.add('edittext {justify: "right", properties: {name: "headTailSizeInput"}}');
        headTailSizeInput.characters = 6;
        headTailSizeInput.preferredSize.width = 60;
        headTailSizeInput.text = defaultHeadTailSize;
        headTailSizeInput.onChanging = function () {
            restoreDefaultsButton.enabled = true;
        };
        headTailSizeInput.onDeactivate = function () {
            headTailSizeInput.text = headTailSizeInput.text.replace(/\s+/g, ''); // trim input text
            headTailSizeInput.text = headTailSizeInput.text.replace(/[^0-9\.]/g, "");
            // If first character is decimal point, don't error, but instead add leading zero to string.
            if (headTailSizeInput.text.charAt(0) == ".") {
                headTailSizeInput.text = "0" + headTailSizeInput.text;
            }
        }

        var headTailSizeUnitsLabelText = headTailSizeGroup.add("statictext", undefined, undefined, { name: "headTailSizeUnitsLabelText" });
        headTailSizeUnitsLabelText.text = getRulerUnits();


        // HORIZONTALTABBEDPANEL
        // =====================
        horizontalTabbedPanel.selection = tabOptions; // Activate Options tab

        // FOOTERGROUP
        // ===========
        var footerGroup = dialogMainGroup.add("group", undefined, { name: "footerGroup" });
        footerGroup.orientation = "column";
        footerGroup.alignChildren = ["left", "bottom"];
        footerGroup.spacing = 5;
        footerGroup.margins = 0;

        // INNERFOOTERGROUP
        // ===========
        var innerFooterGroup = footerGroup.add("group", undefined, { name: "innerFooterGroup" });
        innerFooterGroup.orientation = "row";
        innerFooterGroup.alignChildren = ["left", "bottom"];
        innerFooterGroup.spacing = 20;
        innerFooterGroup.margins = 0;

        // RESTOREDEFAULTSGROUP
        // ====================
        var restoreDefaultsGroup = innerFooterGroup.add("group", undefined, { name: "restoreDefaultsGroup" });
        restoreDefaultsGroup.orientation = "column";
        restoreDefaultsGroup.alignChildren = ["left", "bottom"];
        restoreDefaultsGroup.spacing = 10;
        restoreDefaultsGroup.margins = 0;
        restoreDefaultsGroup.alignment = ["left", "bottom"];

        var restoreDefaultsButton = restoreDefaultsGroup.add("button", undefined, undefined, { name: "restoreDefaultsButton" });
        restoreDefaultsButton.text = "Reset";
        restoreDefaultsButton.alignment = ["center", "center"];
        restoreDefaultsButton.justify = "left";
        restoreDefaultsButton.enabled = (setFontSize != defaultFontSize || setDecimals != defaultDecimals || setGap != defaultGap || setStrokeWidth != defaultStrokeWidth || setHeadTailSize != defaultHeadTailSize || setScale != defaultScale || setCustomUnits != defaultCustomUnits ? true : false);
        restoreDefaultsButton.onClick = function () {
            restoreDefaults();
        };

        // BUTTONGROUP
        // ===========
        var buttonGroup = innerFooterGroup.add("group", undefined, { name: "buttonGroup" });
        buttonGroup.orientation = "row";
        buttonGroup.alignChildren = ["right", "bottom"];
        buttonGroup.spacing = 10;
        buttonGroup.margins = [70, 0, 0, 0];
        buttonGroup.alignment = ["left", "bottom"];

        var cancelButton = buttonGroup.add("button", undefined, undefined, { name: "cancelButton" });
        cancelButton.text = "Cancella";
        cancelButton.alignment = ["right", "bottom"];
        cancelButton.onClick = function () {
            toggleSpecifyDialog('close');
        };

        var specifyButton = buttonGroup.add("button", undefined, undefined, { name: "specifyButton" });
        specifyButton.text = "Genera quote";
        activateSpecifyButton();
        specifyButton.onClick = function () {
            startSpec();
        };

        var specsLayer;
        var decimals;
        var scale;
        var gap;
        var strokeWidth;
        var headTailSize;
        var newSpot;
        var color;

        function startSpec() {
            try {
                specsLayer = doc.layers["QUOTE"];
            } catch (err) {
                specsLayer = doc.layers.add();
                specsLayer.name = "QUOTE";
            }
            try{
                newSpot = doc.spots["Fustella"];
            } catch(e){
                newSpot = doc.spots.add();
            }
            var newColor = new CMYKColor();
            newColor.cyan = 0;
            newColor.magenta = 100;
            newColor.yellow = 0;
            newColor.black = 0;

            newSpot.name = "Fustella";
            newSpot.colorType = ColorModel.SPOT;
            newSpot.color = newColor;

            color = new SpotColor();
            color.spot = newSpot;
            color.tint = 100;
            specsLayer.blendingMode = BlendModes.DARKEN;

            // Add all selected objects to array
            var objectsToSpec = new Array();
            for (var index = doc.selection.length - 1; index >= 0; index--) {
                objectsToSpec[index] = doc.selection[index];
            }

            // Fetch desired dimensions
            var top = topCheckbox.value;
            var left = leftCheckbox.value;
            var right = rightCheckbox.value;
            var bottom = bottomCheckbox.value;
            // Take focus away from fontSizeInput to validate (numeric)
            fontSizeInput.active = false;

            // Set bool for numeric vars
            var validFontSize = /^[0-9]{1,3}(\.[0-9]{1,3})?$/.test(fontSizeInput.text);

            var validRedColor = /^[0-9]{1,3}$/.test(color.red) && parseInt(color.red) > -1 && parseInt(color.red) < 256;
            var validGreenColor = /^[0-9]{1,3}$/.test(color.green) && parseInt(color.green) > -1 && parseInt(color.green) < 256;
            var validBlueColor = /^[0-9]{1,3}$/.test(color.blue) && parseInt(color.blue) > -1 && parseInt(color.blue) < 256;

            var validDecimalPlaces = /^[0-4]{1}$/.test(decimalPlacesInput.text);
            if (validDecimalPlaces) {
                // Number of decimal places in measurement
                decimals = decimalPlacesInput.text;
                // Set environmental variable
                $.setenv("Specify_defaultDecimals", decimals);
            }

            var validGap = /^(0|[1-9]\d*)(\.\d+)?$/.test(gapInput.text); // Allows for decimals/integers
            if (validGap) {
                // Gap size
                gap = parseFloat(gapInput.text);
                // Set environmental variable
                $.setenv("Specify_defaultGap", gap);
            }

            var validStrokeWidth = /^(0|[1-9]\d*)(\.\d+)?$/.test(strokeWidthInput.text); // Allows for decimals/integers
            if (validStrokeWidth) {
                // Stroke Width
                strokeWidth = parseFloat(strokeWidthInput.text);
                // Set environmental variable
                $.setenv("Specify_defaultStrokeWidth", strokeWidth);
            }

            var validHeadTailSize = /^(0|[1-9]\d*)(\.\d+)?$/.test(headTailSizeInput.text); // Allows for decimals/integers
            if (validHeadTailSize) {
                // Head Tail Size
                headTailSize = parseFloat(headTailSizeInput.text);
                // Set environmental variable
                $.setenv("Specify_defaultHeadTailSize", headTailSize);
            }

            var theScale = parseInt(customScaleDropdown.selection.toString().replace(/1\//g, "").replace(/[^0-9]/g, ""));
            scale = theScale;
            // Set environmental variable
            $.setenv("Specify_defaultScale", customScaleDropdown.selection.index);

            if (selectedItems < 1) {
                beep();
                alert("Selezionare un oggetto.");
                // Close dialog
                toggleSpecifyDialog('close');
            } else if (!top && !left && !right && !bottom) {
                horizontalTabbedPanel.selection = tabOptions; // Activate Options tab
                beep();
                alert("Selezionare almeno un angolo da quotare.");
            } else if (!validFontSize) {
                horizontalTabbedPanel.selection = tabStyles; // Activate Styles tab
                // If fontSizeInput.text does not match regex
                beep();
                alert("Inserire un valore valido. \n0.002 - 999.999");
                fontSizeInput.active = true;
                fontSizeInput.text = setFontSize;
            } else if (parseFloat(fontSizeInput.text, 10) <= 0.001) {
                horizontalTabbedPanel.selection = tabStyles; // Activate Styles tab
                beep();
                alert("Dimensione font deve essere maggiore di 0.001.");
                fontSizeInput.active = true;
            } else if (!validDecimalPlaces) {
                horizontalTabbedPanel.selection = tabStyles; // Activate Styles tab
                // If decimalPlacesInput.text is not numeric
                beep();
                alert("I decimali devono essere tra 0 - 4.");
                decimalPlacesInput.active = true;
                decimalPlacesInput.text = setDecimals;
            } else if (!validGap) {
                horizontalTabbedPanel.selection = tabStyles; // Activate Styles tab
                // If gapInput.text does not match regex decimals/integers
                beep();
                alert("Valore di dimensione gap invalido");
                gapInput.active = true;
                gapInput.text = setGap;
            } else if (!validStrokeWidth) {
                horizontalTabbedPanel.selection = tabStyles; // Activate Styles tab
                // If strokeWidthInput.text does not match regex decimals/integers
                beep();
                alert("Valore di spessore traccia invalido");
                headTailSizeInput.active = true;
                headTailSizeInput.text = setHeadTailSize;
            } else if (!validHeadTailSize) {
                horizontalTabbedPanel.selection = tabStyles; // Activate Styles tab
                // If headTailSizeInput.text does not match regex decimals/integers
                beep();
                alert("Valore degli estremi quota invalido");
                headTailSizeInput.active = true;
                headTailSizeInput.text = setHeadTailSize;
            } else if (selectedItems == 2 && betweenCheckbox.value) {
                if (top) specDouble(objectsToSpec[0], objectsToSpec[1], "top");
                if (left) specDouble(objectsToSpec[0], objectsToSpec[1], "left");
                if (right) specDouble(objectsToSpec[0], objectsToSpec[1], "right");
                if (bottom) specDouble(objectsToSpec[0], objectsToSpec[1], "bottom");
                // Close dialog when finished
                toggleSpecifyDialog('close');
            } else {
                // Iterate over each selected object, creating individual dimensions as you go
                for (var objIndex = objectsToSpec.length - 1; objIndex >= 0; objIndex--) {
                    if (top) specSingle(objectsToSpec[objIndex].geometricBounds, "top");
                    if (left) specSingle(objectsToSpec[objIndex].geometricBounds, "left");
                    if (right) specSingle(objectsToSpec[objIndex].geometricBounds, "right");
                    if (bottom) specSingle(objectsToSpec[objIndex].geometricBounds, "bottom");
                }
                // Close dialog when finished
                toggleSpecifyDialog('close');
            }
        };

        function specSingle(bound, where) {
            // unlock SPECS layer
            specsLayer.locked = false;

            // width and height
            var w = bound[2] - bound[0];
            var h = bound[1] - bound[3];

            var a = bound[0];
            var b = bound[2];
            var c = bound[1];

            var xy = "x";
            var dir = 1;

            switch (where) {
                case "top":
                    a = bound[0];
                    b = bound[2];
                    c = bound[1];
                    xy = "x";
                    dir = 1;
                    break;
                case "right":
                    a = bound[1];
                    b = bound[3];
                    c = bound[2];
                    xy = "y";
                    dir = 1;
                    break;
                case "bottom":
                    a = bound[0];
                    b = bound[2];
                    c = bound[3];
                    xy = "x";
                    dir = -1;
                    break;
                case "left":
                    a = bound[1];
                    b = bound[3];
                    c = bound[0];
                    xy = "y";
                    dir = -1;
                    break;
            }

            var lines = new Array();

            if (xy == "x") {

                // 2 vertical lines
                lines[0] = new Array(new Array(a, c + (gap) * dir));
                lines[0].push(new Array(a, c + (gap + headTailSize) * dir));
                lines[1] = new Array(new Array(b, c + (gap) * dir));
                lines[1].push(new Array(b, c + (gap + headTailSize) * dir));

                // 1 horizontal line
                lines[2] = new Array(new Array(a, c + (gap + headTailSize / 2) * dir));
                lines[2].push(new Array(b, c + (gap + headTailSize / 2) * dir));

                // Create text label
                if (where == "top") {
                    var t = specLabel(w, (a + b) / 2, lines[0][1][1], color);
                    t.top += t.height;
                } else {
                    var t = specLabel(w, (a + b) / 2, lines[0][0][1], color);
                    t.top -= headTailSize;
                }
                t.left -= t.width / 2;

            } else {
                lines[0] = new Array(new Array(c + (gap) * dir, a));
                lines[0].push(new Array(c + (gap + headTailSize) * dir, a));
                lines[1] = new Array(new Array(c + (gap) * dir, b));
                lines[1].push(new Array(c + (gap + headTailSize) * dir, b));

                lines[2] = new Array(new Array(c + (gap + headTailSize / 2) * dir, a));
                lines[2].push(new Array(c + (gap + headTailSize / 2) * dir, b));

                // Create text label
                if (where == "left") {
                    var t = specLabel(h, lines[0][1][0], (a + b) / 2, color);
                    t.left -= t.width;
                    t.rotate(90, true, false, false, false, Transformation.BOTTOMRIGHT);
                    t.top += t.width;
                    t.top += t.height / 2;
                } else {
                    var t = specLabel(h, lines[0][1][0], (a + b) / 2, color);
                    t.rotate(-90, true, false, false, false, Transformation.BOTTOMLEFT);
                    t.top += t.width;
                    t.top += t.height / 2;
                }
            }

            // Draw lines
            var specgroup = new Array(t);

            for (var i = 0; i < lines.length; i++) {
                var p = doc.pathItems.add();
                p.setEntirePath(lines[i]);
                p.strokeDashes = []; // Prevent dashed SPEC lines
                setLineStyle(p, color, parseFloat(strokeWidth));
                specgroup.push(p);
            }
            

            group(specsLayer, specgroup);

        };

        function specDouble(item1, item2, where) {

            var bound = new Array(0, 0, 0, 0);

            var a = item1.geometricBounds;
            var b = item2.geometricBounds;

            if (where == "top" || where == "bottom") {

                if (b[0] > a[0]) { // item 2 on right,

                    if (b[0] > a[2]) { // no overlap
                        bound[0] = a[2];
                        bound[2] = b[0];
                    } else { // overlap
                        bound[0] = b[0];
                        bound[2] = a[2];
                    }
                } else if (a[0] >= b[0]) { // item 1 on right

                    if (a[0] > b[2]) { // no overlap
                        bound[0] = b[2];
                        bound[2] = a[0];
                    } else { // overlap
                        bound[0] = a[0];
                        bound[2] = b[2];
                    }
                }

                bound[1] = Math.max(a[1], b[1]);
                bound[3] = Math.min(a[3], b[3]);

            } else {

                if (b[3] > a[3]) { // item 2 on top
                    if (b[3] > a[1]) { // no overlap
                        bound[3] = a[1];
                        bound[1] = b[3];
                    } else { // overlap
                        bound[3] = b[3];
                        bound[1] = a[1];
                    }
                } else if (a[3] >= b[3]) { // item 1 on top

                    if (a[3] > b[1]) { // no overlap
                        bound[3] = b[1];
                        bound[1] = a[3];
                    } else { // overlap
                        bound[3] = a[3];
                        bound[1] = b[1];
                    }
                }

                bound[0] = Math.min(a[0], b[0]);
                bound[2] = Math.max(a[2], b[2]);
            }
            specSingle(bound, where);
        };

        function specLabel(val, x, y, color) {

            var t = doc.textFrames.add();
            // Get font size from specifyDialogBox.fontSizeInput
            var labelFontSize;
            if (parseFloat(fontSizeInput.text) > 0) {
                labelFontSize = parseFloat(fontSizeInput.text);
            } else {
                labelFontSize = defaultFontSize;
            }

            // Convert font size to RulerUnits
            var labelFontInUnits = convertToPoints(labelFontSize);

            // Set environmental variable
            $.setenv("Specify_defaultFontSize", labelFontInUnits);

            t.textRange.characterAttributes.size = labelFontInUnits;
            t.textRange.characterAttributes.alignment = StyleRunAlignmentType.center;
            t.textRange.characterAttributes.fillColor = color;

            var displayUnitsLabel = units.value;
            // Set environmental variable
            $.setenv("Specify_defaultUnits", displayUnitsLabel);

            var v = (val * scaleFactor) * scale;
            var unitsLabel = "";

            switch (doc.rulerUnits) {
                case RulerUnits.Picas:
                    v = new UnitValue(v, "pt").as("pc");
                    var vd = v - Math.floor(v);
                    vd = 12 * vd;
                    v = Math.floor(v) + "p" + vd.toFixed(decimals);
                    break;
                case RulerUnits.Inches:
                    v = new UnitValue(v, "pt").as("in");
                    v = v.toFixed(decimals);
                    unitsLabel = " in"; // add abbreviation
                    break;
                case RulerUnits.Millimeters:
                    v = new UnitValue(v, "pt").as("mm");
                    v = v.toFixed(decimals);
                    unitsLabel = " mm"; // add abbreviation
                    break;
                case RulerUnits.Centimeters:
                    v = new UnitValue(v, "pt").as("cm");
                    v = v.toFixed(decimals);
                    unitsLabel = " cm"; // add abbreviation
                    break;
                case RulerUnits.Pixels:
                    v = new UnitValue(v, "pt").as("px");
                    v = v.toFixed(decimals);
                    unitsLabel = " px"; // add abbreviation
                    break;
                default:
                    v = new UnitValue(v, "pt").as("pt");
                    v = v.toFixed(decimals);
                    unitsLabel = " pt"; // add abbreviation
            }

            // If custom scale and units label is set
            if (useCustomUnits.value == true && customUnitsInput.enabled && customUnitsInput.text != getRulerUnits()) {
                unitsLabel = customUnitsInput.text;
                $.setenv("Specify_defaultUseCustomUnits", true);
                $.setenv("Specify_defaultCustomUnits", unitsLabel);
            }

            if (displayUnitsLabel) {
                t.contents = v + " " + unitsLabel;
            } else {
                t.contents = v;
            }
            t.top = y;
            t.left = x;

            return t;
        };

        function convertToBoolean(string) {
            switch (string.toLowerCase()) {
                case "true":
                    return true;
                    break;
                case "false":
                    return false;
                    break;
            }
        };

        function setLineStyle(path, color, labelStylesStrokeWidth) {
            path.filled = false;
            path.stroked = true;
            path.strokeColor = color;
            path.strokeWidth = parseFloat(labelStylesStrokeWidth);
            return path;
        };

        // Group items in a layer
        function group(layer, items, isDuplicate) {

            var gg = layer.groupItems.add();

            for (var i = items.length - 1; i >= 0; i--) {

                if (items[i] != gg) { // don't group the group itself
                    if (isDuplicate) {
                        newItem = items[i].duplicate(gg, ElementPlacement.PLACEATBEGINNING);
                    } else {
                        items[i].move(gg, ElementPlacement.PLACEATBEGINNING);
                    }
                }
            }
            return gg;
        };

        function convertToPoints(value) {
            switch (doc.rulerUnits) {
                case RulerUnits.Picas:
                    value = new UnitValue(value, "pc").as("pt");
                    break;
                case RulerUnits.Inches:
                    value = new UnitValue(value, "in").as("pt");
                    break;
                case RulerUnits.Millimeters:
                    value = new UnitValue(value, "mm").as("pt");
                    break;
                case RulerUnits.Centimeters:
                    value = new UnitValue(value, "cm").as("pt");
                    break;
                case RulerUnits.Pixels:
                    value = new UnitValue(value, "px").as("pt");
                    break;
                default:
                    value = new UnitValue(value, "pt").as("pt");
            }
            return value;
        };

        function convertToUnits(value) {
            switch (doc.rulerUnits) {
                case RulerUnits.Picas:
                    value = new UnitValue(value, "pt").as("pc");
                    break;
                case RulerUnits.Inches:
                    value = new UnitValue(value, "pt").as("in");
                    break;
                case RulerUnits.Millimeters:
                    value = new UnitValue(value, "pt").as("mm");
                    break;
                case RulerUnits.Centimeters:
                    value = new UnitValue(value, "pt").as("cm");
                    break;
                case RulerUnits.Pixels:
                    value = new UnitValue(value, "pt").as("px");
                    break;
                default:
                    value = new UnitValue(value, "pt").as("pt");
            }
            return value;
        };

        function getRulerUnits() {
            var rulerUnits;
            switch (doc.rulerUnits) {
                case RulerUnits.Picas:
                    rulerUnits = "pc";
                    break;
                case RulerUnits.Inches:
                    rulerUnits = "in";
                    break;
                case RulerUnits.Millimeters:
                    rulerUnits = "mm";
                    break;
                case RulerUnits.Centimeters:
                    rulerUnits = "cm";
                    break;
                case RulerUnits.Pixels:
                    rulerUnits = "px";
                    break;
                default:
                    rulerUnits = "pt";
            }
            return rulerUnits;
        };

        function activateSpecifyButton() {
            // Update helpTip
            specifyButton.helpTip = (topCheckbox.value || rightCheckbox.value || bottomCheckbox.value || leftCheckbox.value || selectAllCheckbox.value ? "" : "Select at least 1 dimension in the Options tab");
        };

        function restoreDefaults() {
            topCheckbox.value = false;
            rightCheckbox.value = false;
            bottomCheckbox.value = false;
            leftCheckbox.value = false;
            selectAllCheckbox.value = false;
            if (selectedItems == 2) {
                betweenCheckbox.value = false;
            }
            customScaleDropdown.selection = setScale;
            units.value = setUnits;
            useCustomUnits.value = setUseCustomUnits;
            useCustomUnits.enabled = true;
            customUnitsInput.text = setCustomUnits;
            customUnitsInput.enabled = false;
            decimalPlacesInput.text = setDecimals;
            fontSizeInput.text = setFontSize;
            gapInput.text = setGap;
            strokeWidthInput.text = setStrokeWidth;
            headTailSizeInput.text = setHeadTailSize;

            updatePanel(specifyDialogBox);

            // Unset environmental variables
            $.setenv("Specify_defaultUnits", "");
            $.setenv("Specify_defaultFontSize", "");
            $.setenv("Specify_defaultDecimals", "");
            $.setenv("Specify_defaultGap", "");
            $.setenv("Specify_defaultStrokeWidth", "");
            $.setenv("Specify_defaultHeadTailSize", "");
            $.setenv("Specify_defaultScale", "");
            $.setenv("Specify_defaultUseCustomUnits", "");
            $.setenv("Specify_defaultCustomUnits", "");

            restoreDefaultsButton.active = false;
            restoreDefaultsButton.enabled = false;
        };

        function updatePanel(win) {
            specifyDialogBox.layout.layout(true);
        };

        function toggleSpecifyDialog(action) {
            if (!action || (action !== 'open' && action !== 'close')) {
                return
            }
            if (action === 'open') {
                specifyDialogBox.show();
            } else {
                specifyDialogBox.close();
            }
        };

        switch (selectedItems) {
            case 0:
                beep();
                alert("Seleziona almeno un oggetto.");
                break;
            default:
                toggleSpecifyDialog('open');
                break;
        }
    } catch (e) {
        toggleSpecifyDialog('close');
        alert("Error: " + e)
    }

}