// Script per la posizione automatica della bottoniera di Dima per macchina Lombardi
//
// ©Gianluca Vitale, Mavigrafica 2021

try {
    var doc = app.activeDocument;

    if (doc.selection.length == 1) {
        var obj = doc.pageItems[indexSel(doc.pageItems)];
        if (doc.pageItems[indexSel(doc.pageItems)].typename == "PlacedItem") {
            // Definizione Colori

            var registro = doc.swatches.getByName('[Registro]').color;
            var noColor = new NoColor();
            var bianco = new CMYKColor();
            bianco.cyan = 0;
            var nero = new CMYKColor();
            nero.black = 100;
            var ciano = new CMYKColor();
            ciano.cyan = 100;
            var mag = new CMYKColor();
            mag.magenta = 100;
            var giallo = new CMYKColor();
            giallo.yellow = 100;

            var listaColori = [];
            var listaInchiostri = [];
            var listaColoriStringa = [];
            var coloreTecnicoIndex;
            var listaColori = [];
            listaColoriFunc();
            //Variabili Globali

            var stampante = {
                printer: undefined,
                polimero: undefined,
                pi: undefined,
                all: undefined
            };
            var zetaValue = undefined;
            var distFascia;
            var distPasso;
            var bottoniera;
            var gruppoBottoniera;
            var diciture;
            var testi;
            var nomiColore;
            var dataG;
            var righe;
            var etichette;
            var k = 1.4173;
            var mm = 2.834645669;

            //lancio UI
            finestra();
        } else alert("ATTENZIONE\nSelezionare un elemento collegato\n\n© Gianluca Vitale - Mavigrafica 2021");
    } else alert("ATTENZIONE\nSelezionare un elemento\n\n© Gianluca Vitale - Mavigrafica 2021");
} catch (e) {
    alert("ATTENZIONE\nAprire un documento o chiedi a Biagino\n\n© Gianluca Vitale - Mavigrafica 2021");
}



function finestra() {

    // BOTTONIERADIMA
    // ==============
    var UIbottoniera = new Window("dialog");
    UIbottoniera.text = "Bottoniera Dima";
    UIbottoniera.orientation = "row";
    UIbottoniera.alignChildren = ["left", "top"];
    UIbottoniera.spacing = 10;
    UIbottoniera.margins = 16;

    // GRUPPOSINISTRA
    // ==============
    var gruppoSinistra = UIbottoniera.add("group", undefined, {
        name: "gruppoSinistra"
    });
    gruppoSinistra.orientation = "column";
    gruppoSinistra.alignChildren = ["fill", "top"];
    gruppoSinistra.spacing = 10;
    gruppoSinistra.margins = 0;

    // INFOETICHETTAPANEL
    // ==================
    var infoEtichettaPanel = gruppoSinistra.add("panel", undefined, undefined, {
        name: "infoEtichettaPanel"
    });
    infoEtichettaPanel.text = "Info Etichetta";
    infoEtichettaPanel.preferredSize.height = 205;
    infoEtichettaPanel.orientation = "column";
    infoEtichettaPanel.alignChildren = ["left", "top"];
    infoEtichettaPanel.spacing = 10;
    infoEtichettaPanel.margins = 10;

    // DIMENSIONIPANEL
    // ===============
    var dimensioniPanel = infoEtichettaPanel.add("panel", undefined, undefined, {
        name: "dimensioniPanel"
    });
    dimensioniPanel.text = "Dimensione (in mm)";
    dimensioniPanel.orientation = "column";
    dimensioniPanel.alignChildren = ["left", "top"];
    dimensioniPanel.spacing = 10;
    dimensioniPanel.margins = 10;

    // DIMENSIONIGROUP
    // ===============
    var dimensioniGroup = dimensioniPanel.add("group", undefined, {
        name: "dimensioniGroup"
    });
    dimensioniGroup.orientation = "row";
    dimensioniGroup.alignChildren = ["left", "center"];
    dimensioniGroup.spacing = 10;
    dimensioniGroup.margins = 0;

    var baseLabel = dimensioniGroup.add("statictext", undefined, undefined, {
        name: "baseLabel"
    });
    baseLabel.text = "Base";

    var base = dimensioniGroup.add('edittext {justify: "right", properties: {name: "base", readonly: true}}');
    base.text = (obj.width / mm).toFixed(2);
    base.preferredSize.width = 65;

    var altezzaLabel = dimensioniGroup.add("statictext", undefined, undefined, {
        name: "altezzaLabel"
    });
    altezzaLabel.text = "Altezza";

    var altezza = dimensioniGroup.add('edittext {justify: "right", properties: {name: "altezza", readonly: true}}');
    altezza.text = (obj.height / mm).toFixed(2);
    altezza.preferredSize.width = 65;

    // ZETAGROUP
    // =========
    var zetaGroup = dimensioniPanel.add("group", undefined, {
        name: "zetaGroup"
    });
    zetaGroup.orientation = "row";
    zetaGroup.alignChildren = ["center", "center"];
    zetaGroup.spacing = 10;
    zetaGroup.margins = 0;

    var zetaLabel = zetaGroup.add("statictext", undefined, undefined, {
        name: "zetaLabel"
    });
    zetaLabel.text = "Valore Zeta:";

    var zeta = zetaGroup.add('edittext {justify: "right", properties: {name: "zeta"}}');
    zeta.active = true;
    zeta.preferredSize.width = 150;

    // RIPPANEL
    // ========
    var ripPanel = infoEtichettaPanel.add("panel", undefined, undefined, {
        name: "ripPanel"
    });
    ripPanel.text = "Ripetizioni";
    ripPanel.orientation = "column";
    ripPanel.alignChildren = ["left", "top"];
    ripPanel.spacing = 10;
    ripPanel.margins = 10;

    // RIPGROUP
    // ========
    var ripGroup = ripPanel.add("group", undefined, {
        name: "ripGroup"
    });
    ripGroup.orientation = "row";
    ripGroup.alignChildren = ["left", "center"];
    ripGroup.spacing = 14;
    ripGroup.margins = 0;

    var ripFasciaLabel = ripGroup.add("statictext", undefined, undefined, {
        name: "ripFasciaLabel"
    });
    ripFasciaLabel.text = "Fascia";

    var ripFascia = ripGroup.add('edittext {justify: "right", properties: {name: "ripFascia"}}');
    ripFascia.text = "3";
    ripFascia.preferredSize.width = 60;

    var ripPassoLabel = ripGroup.add("statictext", undefined, undefined, {
        name: "ripPassoLabel"
    });
    ripPassoLabel.text = "Passo";

    var ripPasso = ripGroup.add('edittext {justify: "right", properties: {name: "ripPasso"}}');
    ripPasso.text = "3";
    ripPasso.preferredSize.width = 60;

    // INTERSPAZIOPANEL
    // ================
    var interspazioPanel = infoEtichettaPanel.add("panel", undefined, undefined, {
        name: "interspazioPanel"
    });
    interspazioPanel.text = "Interspazio tra etichette";
    interspazioPanel.orientation = "column";
    interspazioPanel.alignChildren = ["left", "top"];
    interspazioPanel.spacing = 10;
    interspazioPanel.margins = 10;

    // INTERSPAZIOGROUP
    // ================
    var interspazioGroup = interspazioPanel.add("group", undefined, {
        name: "interspazioGroup"
    });
    interspazioGroup.orientation = "row";
    interspazioGroup.alignChildren = ["left", "center"];
    interspazioGroup.spacing = 14;
    interspazioGroup.margins = 0;

    var interspazioFasciaLabel = interspazioGroup.add("statictext", undefined, undefined, {
        name: "interspazioFasciaLabel"
    });
    interspazioFasciaLabel.text = "Laterale";

    var interspazioFascia = interspazioGroup.add('edittext {justify: "right", properties: {name: "interspazioFascia"}}');
    interspazioFascia.text = "0";
    interspazioFascia.preferredSize.width = 46.5;

    var interspazioPassoLabel = interspazioGroup.add("statictext", undefined, undefined, {
        name: "interspazioPassoLabel"
    });
    interspazioPassoLabel.text = "Verticale";

    var interspazioPasso = interspazioGroup.add('edittext {justify: "right", properties: {name: "interspazioPasso"}}');
    interspazioPasso.text = "0";
    interspazioPasso.preferredSize.width = 46.5;

    // GRUPPODESTRA
    // ============
    var gruppoDestra = UIbottoniera.add("group", undefined, {
        name: "gruppoDestra"
    });
    gruppoDestra.orientation = "column";
    gruppoDestra.alignChildren = ["right", "center"];
    gruppoDestra.spacing = 10;
    gruppoDestra.margins = 0;

    // FUSTELLA
    // ========
    var fustella = gruppoDestra.add("panel", undefined, undefined, {
        name: "fustella"
    });
    fustella.text = "Fustella";
    fustella.preferredSize.height = 60;
    fustella.orientation = "column";
    fustella.alignChildren = ["left", "top"];
    fustella.spacing = 10;
    fustella.margins = 10;

    // FUSTELLAGROUP
    // =============
    var fustellaGroup = fustella.add("group", undefined, {
        name: "fustellaGroup"
    });
    fustellaGroup.orientation = "column";
    fustellaGroup.alignChildren = ["left", "center"];
    fustellaGroup.spacing = 10;
    fustellaGroup.margins = 0;

    var fustellaLabel = fustellaGroup.add("statictext", undefined, undefined, {
        name: "fustellaLabel"
    });
    fustellaLabel.text = "Seleziona colore fustella";

    var fustellaColore_array = ["[Nessuna Fustella]", ].concat(listaColoriStringa);
    var fustellaColore = fustellaGroup.add("dropdownlist", undefined, undefined, {
        name: "fustellaColore",
        items: fustellaColore_array
    });
    fustellaColore.selection = 0;
    fustellaColore.preferredSize.width = 212;

    // PANNELSTAMPANTE
    // ===============
    var pannelStampante = gruppoDestra.add("panel", undefined, undefined, {
        name: "pannelStampante"
    });
    pannelStampante.text = "Stampante";
    pannelStampante.preferredSize.height = 47;
    pannelStampante.orientation = "column";
    pannelStampante.alignChildren = ["center", "center"];
    pannelStampante.spacing = 10;
    pannelStampante.margins = 10;

    // SELEZIONSTAMPANTE
    // =================
    var selezionStampante = pannelStampante.add("group", undefined, {
        name: "selezionStampante"
    });
    selezionStampante.orientation = "row";
    selezionStampante.alignChildren = ["left", "center"];
    selezionStampante.spacing = 10;
    selezionStampante.margins = 0;

    var stampLomb = selezionStampante.add("radiobutton", undefined, undefined, {
        name: "stampLomb"
    });
    stampLomb.text = "Lombardi";
    stampLomb.value = true;

    var stampEtimac = selezionStampante.add("radiobutton", undefined, undefined, {
        name: "stampEtimac"
    });
    stampEtimac.text = "Etimac";

    var stampErrepi = selezionStampante.add("radiobutton", undefined, undefined, {
        name: "stampErrepi"
    });
    stampErrepi.text = "Errepi";

    // GRUPPODESTRA
    // ============
    var LogoMavi_imgString = "%C2%89PNG%0D%0A%1A%0A%00%00%00%0DIHDR%00%00%00%C3%90%00%00%00(%08%06%00%00%00%C2%8Fet%00%00%00%00%09pHYs%00%00%0B%12%00%00%0B%12%01%C3%92%C3%9D~%C3%BC%00%00%0A%C3%91IDATx%C2%9C%C3%AD%5D%C3%8D%C2%AF%1CG%11%C2%AF%C3%B8%23%04%C2%81%C3%B0%C2%82%7C%C2%88%C2%94%C2%83W%C2%88C%04%07%C2%8F%08G%24%C2%8F%0FH%5C%10%2B%C2%94%7B%26%07%C3%84!%07%C2%96%2B%17%267%24%0E%C3%AC%C2%9F%C2%B0%C2%B9qAY%C3%8E%1C%3C%C3%AF%0F%40%C3%9A'%C2%81P%04H%C2%8B%C2%90C%C3%80%0A%5Ec%13%0C%C2%BC%C2%97A%C3%83%C2%AB%C3%A2%C3%BD%5CS%C3%95%C3%933%3B%C3%BB6%C2%B6%C3%BB'%C2%8D%C3%A6m%7FUuOw%C3%97W%C3%8F%C2%BC%17%C3%AA%C2%BA%C2%A6%C2%84%C2%84%1D%C2%91%13%C3%91%1D%22%3A%C3%A2%C2%BF%C2%9F%1B%5CJ3'!a8%C2%AE%C2%A4%C2%B1K%18%01k%22%C2%BAMD%C3%9B%C3%A7m0%C2%93%0A%C2%97%C2%90%C2%B0%03%1A%09%C2%94%11%C3%91%04%C2%9A%C3%98%C3%B2%C2%8E%12%C3%82%C2%84%C3%AB!%C3%96%11%3B%C3%90%C2%94%2FD%C3%95*u%06%C2%ADKo%C3%B8%0A!%C3%84%C3%97%C3%97%C2%89%C3%A8kD%C3%B41%C2%A8%C2%AE%0BUV%C3%8A%3C%22%C2%A2%C3%8Fr%C2%9A%C3%BC%C3%BDw%22%C3%BA%5Cd%C3%BD%C2%8F%0D%C3%B5X%C3%92N%C2%89%C3%A8%C2%B2%C2%A2%C2%B1%C3%AC%C2%B1%7B%C3%B7%19C%C2%84%3C%C3%A7%13~%C3%AE%C3%82%07B%C3%92%C2%A4%0C%C3%8E%05y%1ER%C3%86z%1E%C2%BB%C3%92%C2%88%C2%81%C2%A6%C3%91%C2%B7%C3%BE%C2%B8%C2%A8%C3%AB%C2%BA%C2%AA%C2%9F%C3%84%C2%B6%C2%AE%C3%ABI%23%C2%99%02WY%C2%B7%C2%91%07%C3%8A%C3%8B%C2%A5i5%C2%98%19%C3%A5%C2%A8U%C3%AA%C2%8C%C2%A6U%0E%C2%AF%C2%BCU%C3%AB%C2%9C%2F%C2%8B6%C3%96%C3%8D%C3%AA%C2%BA~%C3%9C*%11%C3%862%C2%A2%7F1XG%C2%8Cy%C2%88%C2%86U%C3%8E%C2%AB'%C3%B7u%C2%AB%C2%95%C3%B34%2C%C2%AB%C2%9F%C2%87%C2%94%C2%B1%C2%9E%C3%87%C2%AE4%C2%86%C3%B4%C2%A3o%C3%BDQ%2F%C3%8B%C2%89p%C3%8D%C3%98Y%11%C3%8D%C3%AA%C2%9F%C2%B7R%C2%BB%C3%91%C3%AC%1C%C2%B7%C2%8CRC%C3%9A%1A%1B%19%C3%AF%C3%A2%C2%9F%C3%AA%C3%99%C3%AE%1B%2C%3Dv%C3%85M%C2%A6%3F%C3%A9h%C3%87%1B%C3%83%C2%84%03%C3%81s%224%13%C2%A34D4%C3%B1%C2%84%C2%BF%C3%96J%C3%AD%C2%86%C2%B7Pn%C2%B1Jb%C3%91%C2%BA(%C3%9CQ%7Dj%C3%94%C2%ABo9%C2%B4_'%C2%A2%C2%B7%C3%94X%C2%AD%03%C2%9B%C3%8E%C3%ADV%C3%8A9%C3%AE%C3%80%C3%9F%C3%8D%22%C3%BA%26%11%C3%BD%C2%ACU%C3%AA%1C%C3%9E%18%0EA%C2%A3%C3%BA%C2%BC0%C2%A0%C3%9E%C3%9CP%19%C3%85%C2%90%3E%C3%A2%7Be%C2%A8%C3%A0%1A%25%3F%C3%BB%09%C3%94%1F%C3%82%C3%8FA%C3%A1-%20%C3%A2%0E%16*m%C2%A8%C3%B4%C2%99%C3%B2D%C3%B3%60%C3%91%3A%24N%C2%8CI%22%C3%906V%17%C2%BCv%10%C2%BF%22%C2%A2%C3%97%C2%88%C3%A8%C2%83V%C3%8E9%26%1Dc%C2%98p%00%C2%84%16%C2%90%25%C2%85%16%03%C2%A5O%C3%97%C3%A2x%C2%83%17%26%1A%C3%92%7F%24%C2%A2%1B%C3%B0%7B%C3%86%C3%BC%C2%84%60%C3%91%11%C3%BE%C3%BF%C2%AD%C3%92%C3%BF%04%7F%C3%AB%3C%C3%BD%1B%C3%B1n%2B%C3%A5%3C-T%C3%8F%C3%82o%C2%B9%C3%8E%7B%60t%7B%C3%90%1B%C3%97%C2%9F%C2%89%C3%A8%C2%AF%C3%BC%C2%B7%2C%C3%92y%C3%80%C2%A0%C2%BEKD%C3%87D%C3%B4%C2%90%C3%AFw%5B%25%C3%828%C3%A6%C3%9C%C2%B7%C3%B894%C2%AA%C2%AB%C2%A8%C2%AF%C2%927%C3%A7%C2%85%1E%C2%A3M%2C%C2%99o%C3%8B%C3%91%60%01%C3%BB%C3%98%C3%90%C3%B8%06%C3%9F%3F%C2%82%C2%BC%C2%8B%0F%C3%A2%3AF%C2%A9%00%C2%8D%C3%A4i%2B%C3%B7I%C2%84%C2%9C%08%C3%9BV%C3%A96%C2%B4QZ%C2%B4J%C2%B4%C3%8B%C3%A05k%C2%95~%C2%92%7F%C2%AB%C2%9F%C2%92w_%C2%A5%C3%9F7%C3%9A%C3%9F%C2%87%C2%81%2F%C2%90z%C2%A11%C3%9C%C2%A8v-%03%3DT%7FW%C3%A3%5B%C3%93%C3%85g%11%C3%93%C3%97%5D%2FM%C2%A3T%C3%BC%C3%AC%C2%9B%C2%BEy%C2%85%24%10%C2%81%C2%91%5CE%C3%AC%C3%BE%1E%C2%8AH%C2%A9U(%1A%2BC%C3%A2%C3%BD%C2%88%C3%93%C3%B5.%3Bq%C2%8Cy%C2%B4K%1E%C2%AB%C2%BC%C2%A3%40%C2%9E%C3%BE%1D%C2%8B%C2%BE%C3%B5%C2%84%C2%87%C3%9F%C3%B3%C3%BD%C2%A4U%C3%A2%0C%C2%85%C2%92%C3%86%C3%84v%C3%83%C2%91J%C3%B3%C3%AA%13%3C%C3%87%C3%AB%C2%8Af_%5E%1F%C2%B2%C2%A6%C2%80%C3%9A%C2%82%C3%A4%15%C2%AC%C2%AEW%C2%91%C2%AA%C3%AB%10%C3%BAB%C3%A3%3A%C2%A7%3D%3Ch%00%C3%97%C3%995%11U%C2%84%C3%B4%09%C3%AD~%C3%96N%C3%A9%C2%A1Pu%17F%C2%B9%C2%B5Ac%C3%95*%C3%95%C3%9Ea%C2%AD~%3E-%12%C3%88%1AC%2B%C3%8D%C2%AB%C2%8F%C2%97%C3%AC%C3%9C%C3%BB%C2%90%40%C2%95%C2%917%C2%B6%04%0A%C2%B9%C3%8A%C2%ADz%07%C2%95%40%C3%84%C2%9E%C2%92U%2B5%0E9%7B%C2%97bQ(I%C3%92H%C2%90%C3%AF%C2%AB%C2%BA7YR%C2%89%C2%B4jl%C2%A3o%1B%C3%ADk%C2%89%C2%84%C3%B6I%C3%A3e%C3%BB%C2%83%C2%93%C2%A7%7FK%C3%A0%C3%AEoD%C3%B4%05Gg%C2%97%60%C2%ADe%03%C2%89%C3%AD%C3%92%C3%98%1C%C3%B7T%00%C3%B0%18%C3%AA%C2%93%C2%B3%C2%93%C3%AA1%7C%C3%80%C3%A5%C3%AFByq%7F%7F%C2%91%C3%AF%C2%A1%20%C3%A7U%C2%A6%7B%C3%8F%C2%B0%19BAR%C2%91%00%1F%C3%B1%C2%BD%19C%C2%81%C3%B4C%C3%AC%C2%ACGLO4%C2%80%C2%A1%C2%B6%C2%89%04%C3%86%C2%AD%C2%B1%12%09z%0F%C3%8A%C3%A7P%C3%86%1A%C3%8B%C3%B1%C3%A1%C3%AC%C2%9AC%60%C3%AD~%C2%96d%C2%A8%1D%C3%89%22%C3%88T%1B%C3%8BV%C2%893d%1C%7C%C2%B4%C3%AC%C2%AB%C2%8D%C3%81%C2%8B%C3%95%C3%8F%18%09%24%C3%B5%C2%84%C2%8F%C3%90%C3%8Eo%C3%91%10%C2%BC%C3%87w%C2%BD%C2%83%C3%96%06%C2%AF%C2%A1%C3%BE%5B%C2%92C%60I%07%C3%9D%0Fk%C3%A7%C2%8E%C2%A9%C2%AF%C3%9B%C3%A9%C2%92%40%18%C3%94%C2%B6%C3%9A%C2%89%C2%B9r%C3%95%C2%AE%C3%95%C2%96%15%3C%C3%BF1%C2%8F%5B%C2%A8%1F%C2%A3%5CV%20u%2CL%1D%C3%89p%C3%84%C3%92%C3%A3A%2B%C3%A7%0C%C3%9A%C3%9Bd%C3%996%04%5E%20%C3%8B%C2%BE%C3%B2%C3%AA%1C%12%C3%BF%1C%40%C2%BB%C3%8B%C3%BD%C3%BFI%C2%86%1C0%0D%C3%85%C3%81%C3%B6%09K%C2%8A%C2%8E%C2%8E%7D%C2%9E%C3%86%C3%B6%C3%A2Er%C3%AEk%C3%81N%01%0D%C3%AD%C3%92%C2%AEX%7CkU%C3%B0%C2%A6%C2%91F%C2%BC0%C2%BD%C2%A0%C3%A6%10%C2%A0%C3%BAQ%C3%80%C2%BB%2F%C2%B1%08%05%07%25%C2%80X%C2%B1%C2%AA%7C%5B%19%C3%9F%C2%96%5B~%2C%C3%B4%0D%C2%A4%0A%C2%AF%C3%87%C2%AD%1C%1B%C3%9B%3D8%12%2CT%C3%90%0F%C3%A1%C3%B1%C2%83%C2%91%C3%A7%C2%80%C2%8B%7DI%C2%A0%C2%89%C3%B3%C3%B0%1F%C2%80%3D%C2%B5%C3%A8!%C2%85%C3%BA%0C%C3%86%C3%AA%199V%3F4h%C3%BD%C2%B4%C2%A1%C3%A0E0%C3%A6%C2%84%7F%C3%99%C3%B0%C3%94%C3%AE%05%C3%BB%C2%92%403G%C2%B5%C2%A2H%C2%87%C2%84vi%2F%0D%C2%97%C2%B6%07%C3%8F%C3%9D%C2%8E%06%C2%BE%C3%9EECN%04%C2%8DS%C2%A3%C3%BEiD%3D%0B%C3%9A%C3%B8%C3%86%40%C2%AA7%C2%86%2F%C3%B3%C2%84k%C3%AA%C2%BC%C2%A2%C3%BAs%0F%C2%8Cx%C2%8D%C2%B1%02%C2%A9%16%0DQ%C3%93%C2%BEk%C3%A4ua%1A8%C3%9Fw%0A%C2%BC%0A%C3%BD%05%3B%16%1A'%C3%82%C2%97%C2%9Cqx%C2%9F%C3%87o%C3%AF%C2%B0%16%C3%90%11wJ%C3%87%1D%10%C3%AFt%C3%A8%C3%A6%C3%9E%24%C2%BE%16%18%2C%C3%84%0D%C3%87%23g%C2%A9%7C%C2%9AwO%C3%AF%7D%C2%91%C3%AF%C2%A7%C2%86%C3%AA%C3%B7b%C3%87o%C3%84e%C2%A3%C2%BEx%C3%A5B%C3%B5%2CH%3B%5B%C3%BE%1B%C2%BD%7B%C3%9E%18%C2%BE%C3%84ch%C2%A9%C2%B5%C2%92fmR%C2%AFp%C3%9E%16%C3%AECx%C2%B5h%C2%88%C2%AAV%1Ay%5D%C2%90%C2%BA%C3%96s%C2%BB%C2%ACx%26%C3%BE%5B%C3%A6%C2%907%0E_e%C2%8F%C3%A9f%C3%9F%C3%B6%C2%B0%C2%B5%C2%80%C2%88'%C2%AF%C2%A7%C3%A7%1F3S%C3%9E%02%C3%8A%C2%8D%C3%85%C3%B7%C2%8E3%40%C2%9A%C2%A6%3E%C2%BA%C3%93w%01%C2%85%C3%94%00%09r6%0F%C3%A0%C3%97N%C2%9E%C3%B7%1Bq%12%08%60%C2%86%C3%AAY%C3%B0%02%C2%A9Y%60%0C_b%3A%C2%8F%C2%8C%C2%89%C3%BA*%C3%B7%C3%AF%3A%1C%C3%83%C2%921%14%1A%C3%B7%C2%99%C3%AE%C2%98%C2%81TY%C3%AC%C3%92f%C3%8C%C3%A2%C3%8C%C3%B9%12i%C2%85%C3%AF8I%C2%B0%C3%B4C%C3%A0%15%17Z%056%16%C2%8E%C2%83%C3%98%C2%AB%C2%8F%C2%A0%C3%AC~a%C2%B8%5E%2B%C3%A5%C2%AE%C3%94%C3%88%1D%C3%97a%1Ep%2BZ.n%C3%8F%3D%C2%AA%C3%B9%08%C2%B9t%11%C2%96%C3%AB%C3%9Aj%C3%BFC%C3%83%25%1Arc%C3%87%C2%B8MCnl%C2%AB%0D%C3%8F%C3%BD%3B%C3%96%18Z%C3%87uBy%04%C3%8F%C3%95%7B%2FI%60%C2%B9%C2%BAu%3Fb%C3%9C%C3%87%C2%A1%C2%A38%3At%60%C3%8D%05%C3%AB%0A%C3%B1%C2%B8%C2%97%C3%8B%C2%93%40d%C2%A8P%C3%84%C2%86Y%C3%8CQ%C3%B5%7Da%11%C2%90%7C%C2%9E%C3%8A%23%10%C3%BBd%C2%AB%0E%C2%92%12%07I%C3%B1%5D%C2%9C%C2%AB%C3%90%C3%B7%C3%87%C2%BC%C3%AB7n%C3%A8O%13%C3%91%C3%A7%2F%C3%88%06%1A%C2%82%C2%90%C2%9D%13%C3%8A%23%C3%908%C2%B4'P%C3%B3j%C3%99%40%22%C2%9D%C2%84%C3%86U%C2%9E%23V%C3%90Y%C3%92%24%C2%A0%C3%BB%C2%BE!%C2%B1%C2%A4%C2%9D%7F8%C2%BCz%10%1E%7F%C3%87mj%09T%C3%B0%C2%B5%1E%C3%8BA%13Z%40%C2%9BO%C3%A0'%C2%8A%C3%96%C3%BC%C2%B0%C2%B4%1D%C3%B5%C3%80Pg4%C3%84%3E%C2%B9d%C3%A8%C3%8C%C2%AF%C3%B1%C2%A4%C2%91%C3%B4%C3%8F%04%16%C2%AA%C2%85%7D%C3%98%40C%10%C2%B2sBy%04%C2%8B%C3%80S%C2%BFB6%C2%90%C3%8C%C2%93%0A%C3%9A%C3%BE%C2%A1c%C2%9FH%C3%9A%11%C2%94%C3%95%C3%B3Lx%5D%3B%C2%BCz%10Z%C3%96%3BK%C3%94%C3%A1%C2%B0%18%C2%84%C2%A7%C3%B1%C2%B3V%C2%96Q%C3%98%C3%A7%C2%9B%02%C3%96.%2F%0F%C3%B1%2F%C2%AD%C2%9C%C3%A7%07b%C2%93%1C%C3%AE%C3%BB%02%C3%A3%C2%A3%C3%A2%C3%98%C2%90h'K%C2%96%C2%B0%C2%A3%C2%85%07%C3%92Wy%12%C2%9EeH%C2%80%C3%BA%C3%AD%1D%C3%9E%26%08%22-%C2%A0%C2%84%5D!%13%C3%88%C2%B3%C2%9D%C3%BA%40%C2%A4%C3%A0X%C2%AFCd%C3%B0%C2%82%C3%9F%5E%3Cr!%1B(!%C3%A1%C2%A2%C2%91C%C2%A8b%C2%8C%054D%1D%C3%AD%C3%B5v%C3%AB%15%C3%B0%C2%B9%0B%C3%81%15%C3%BF%C3%86%C2%A38%0B%C2%B01%C2%B0a%3C6%C2%8E%C3%9F%2B%C2%93%15%C2%AF%C2%BF%C3%93%C2%A6%C2%8F%C3%9FKY%C3%BD%C2%9D%C2%B35%C3%98%258%C2%90H%5B%C2%BE%07%C2%96q%1B%5BE%0F%C2%BF%17%C2%A6%3FZ%C2%92q%C2%9E%C3%A6%C3%AFB%0E%20%3Ec%C2%90S%08c%C3%98N%C2%95%C2%BA%C2%8F%C2%89X%0F%5C%2F'%C3%83%258%C2%8B%24o%C2%9Df%C2%B0%18%C2%A4%23b%C2%B8%2F%C3%A0%40%25N%C3%AE)%C3%A7I%C2%BAxh2%C2%AE%2B%C3%A9%19_9%C2%A4O%C3%B9.%C3%87wr%C2%9E%C3%983v%C2%AD%C3%A2B%C3%86%C3%B6s%C3%A0I%16A%05e%C3%A6p%C2%9C%C2%A3%C3%A0%3A%13%C2%A8%C2%83%C3%BC%C2%95%40%1B1%C3%A1z%15%5C%C3%92%C2%86%04%C3%B2%C3%96%C3%8A%C2%B1%C2%91%C2%B3Z%C2%83mU%C3%AA%22%18%C2%BB%09%C2%84%07%C3%96%C2%AA%C2%BF%C2%AB%C3%80N%5CBPq%03%7D%C2%AD%C2%81%16~%14q%0B%01H%C2%B1%07%0A._C%1F4%C2%AF%19%C3%8C%05%C3%9C%C3%A8d%C3%9C%2B%C3%98x%09%C3%92%C3%A4%C2%92%C3%BExmhz%13%C3%95g%C3%9D%C2%9E%C3%90%5D%C3%81%06%C2%B8R4%C2%90w%C2%82%C2%BC%5B%11%1F%C2%85y%C2%9B%C2%AF8p%C2%90%C2%AA%C3%A4k%C3%83o%C2%9F%C3%A6%C2%90%C2%BE%C3%A2K%C3%8A%C3%A6*%C2%80%26A%C2%B8%C2%92%C3%AB%C2%96%10%C3%80%C3%8A%C3%B9%7D%C2%9D%C3%8Ax%C3%9B%C3%94%0A%C3%A4a%C3%A0k%C3%83i%3A%40%C2%BA4%C2%BEu%C2%90%1Bmb%7B%25%C3%97%C2%91%0F%18VP%C2%A6%02%C3%9E%C2%AD%C2%A0)%C3%B6k%05%C3%AF%2B%C2%AD9%7F%C3%86%7D%24H%2F%C2%8C%20f%C2%AE%C3%9Eu%C2%AA%C2%807%C3%A1%C2%B3%C3%89%C2%9F%2B%C3%9E%C3%B4%C2%98%C2%93Qo%06%7F%C3%97%C2%8A%17iK%C3%B8%C2%9D%1A%7C%C3%A5%C3%AA%C3%B9%C3%A8g%C2%B1t%C3%86%5D%C3%86s%05c%C2%90C%C3%9F%C2%A6%C3%80%C3%8F%C2%92%C3%93u%1BV0X%3F%03%C2%AB%C2%9C%C2%B4%C2%B7%C2%81q%5D%C3%81Xe%C3%8E%7C%C3%90%C3%B3068m%06e%C3%85%C2%8D-%2B6%07%15f%03jTH%C3%A4!d%C2%97%C3%83%C2%83%7Ck%C2%908%C2%B1(xw)%C3%95%C2%A9d%C2%B9%5B'%C2%BD%C2%BB%C2%B0%C2%84%1DO%C2%90%C2%A9%C2%A3!%1E%C2%B6%C2%B0%7B%C2%8A%04%12)%C3%B9%C2%AE%C2%92t%13.%C2%97%C2%81%14%C2%9A%C2%82%04%0E%01%C2%8F%C3%88%C2%94%C3%B0%C3%9B%C3%B3%20%C2%89%C3%B607%C3%8E%C2%A6%C3%95J%C2%AD%C2%9A%C2%80%C3%84%17%C2%A0%C3%96%C3%A0%1D%C2%BE%C3%8C8o%C3%83%C2%B11%C3%AC%C3%83%C2%92%C3%A9n%14-%C2%91%1AM%C3%BE%C2%9BLS%C2%BE%C2%AA4S%3C%C3%84%C2%9C%C3%84%C3%8E%C2%944%23%C3%AEs%C2%A6h%C2%A3i%C2%A0U%C3%8A%0A%C2%BE%26%C3%94%07%C3%B2%0D%C2%89%C3%97%C2%AD%3AW%40%7D%40la%00%C2%B6%C2%A0%C2%BA%C3%8DUg%C3%9F%04%C3%A6f%C3%B0%10%C2%A5%C2%BD%C2%AD%C3%BA%C2%8D%C3%BAgHg.ap%C2%A6%C3%80%C3%A3O9%10W9%C2%BA%C2%AC%C3%AE%C2%8B%C3%B7r%C2%9E%3C%08%3D%C2%A9%C2%96j%C2%80%C3%97jrj%C2%BBK%3E%23%25%13%C2%B1%C2%84w%C2%996%C2%A0%16%C3%9E%08%2C%C2%82%15%C2%A8%C2%B3%13h%C2%8B%60%C2%ACK%C2%A52c%7F%C2%ACvsx%5E%C3%BAk%C2%A7%18%C3%B4%2C%3A%C2%9E%C2%83%C3%B00%C2%87%C2%8DU%C3%B37e%1Ep%01%C2%AF%C2%95%0A%C2%BB%C2%84%3E%C2%AE%C3%94w%00-%C2%B5Y%C3%92d%C3%BEl%C2%A0%C2%9F%19%C2%A8%C2%8Cs%C3%98%C2%9C%0B.o%C3%99%C3%9C%C2%BB%40%C2%82%C3%A3_%C3%A1S)%C2%AF%C2%A2%C2%9D%C2%9C%C3%9C%C3%98%C2%87C%C2%A6%C2%9C%1CSX%C2%A0%C3%A8%C2%8C!%C3%A3%C2%B7%C2%97Fj%C2%A1e%20%C3%85%2Cg%C3%89T9%C2%90%C3%90%C3%893%C2%81%C3%93(%C3%9A%C2%91S)%C3%BEI9f%C2%A6*Oo%C2%BC%25%2C%22-%11pC%C3%98%18%C2%92g%03%0BT%C3%9A%C2%9C%C3%83%7B%60%C2%B8i.FXD%C3%9F!%C2%A2%C2%9F%C2%AB%C2%B4%C3%A6%7B~_%C2%A6%C2%B4%C2%80%12%12%C2%A2%C3%A0-%C2%92%1F%C2%A4%05%C2%94%C2%90%C3%90%C2%8D%13%C3%A7%C2%9C%C3%A2%C2%BF%C3%92%C2%BFxLH%C3%A8%C3%86%2F%C2%9D%12%C3%9FK%12(!!%0E%C3%A6BI%12(!!%0E%C3%BF!%C2%A2%C2%9F%C3%B0%17%C2%80%C3%BE%C2%BF%C2%98%C2%92%04JH%C3%A8%C2%8F_%C3%BC%C3%AF%C3%BFG%11%C3%BD%C3%A6%C2%BF%00%40qv%04%C2%A9s%3B%00%00%00%00IEND%C2%AEB%60%C2%82";
    var LogoMavi = gruppoDestra.add("iconbutton", undefined, File.decode(LogoMavi_imgString), {
        name: "LogoMavi",
        style: "toolbutton"
    });
    LogoMavi.enabled = false;
    LogoMavi.alignment = ["center", "center"];


    // GROUP1
    // ======
    var group1 = gruppoDestra.add("group", undefined, {
        name: "group1"
    });
    group1.orientation = "row";
    group1.alignChildren = ["left", "top"];
    group1.spacing = 10;
    group1.margins = 0;

    var cancel = group1.add("button", undefined, undefined, {
        name: "cancel"
    });
    cancel.text = "Annulla";
    cancel.onClick = function() {
        UIbottoniera.close();
    }

    var ok = group1.add("button", undefined, undefined, {
        name: "ok"
    });
    ok.text = "Crea";
    ok.onClick = function() {
        if (!zeta.text) alert("Inserire un valore a Z\n©Gianluca Vitale 2021");
        else if (isNaN(zeta.text)) alert("Valore Z non valido\n©Gianluca Vitale 2021");
        else {
            if (stampLomb.value) stampante = {
                printer: "Lombardi",
                polimero: 114,
                pi: 3.175,
                all: 6.06
            };
            else if (stampEtimac.value) stampante = {
                printer: "Etimac",
                polimero: 114,
                pi: 3.141,
                all: 6.06
            };
            else if (stampErrepi.value) stampante = {
                printer: "Errepi",
                polimero: 170,
                pi: 3.141,
                all: 10
            };
            UIbottoniera.close();
            distPasso = (interspazioFascia.text.replace(",", ".")) * mm;
            distFascia = (interspazioPasso.text.replace(",", ".")) * mm;
            zetaValue = zeta.text;
            coloreTecnicoIndex = fustellaColore.selection - 1;
            tabulare(ripFascia.text, ripPasso.text);
            dima();
        }
    }
    UIbottoniera.show();
}

function tabulare(ripFascia, ripPasso) {
    righe = doc.groupItems.add();
    etichette = doc.groupItems.add();
    obj.move(righe, ElementPlacement.PLACEATEND);

    for (var i = 1; i < ripFascia; i++) {
        obj.duplicate();
        obj.translate(obj.width + distPasso, 0);
        obj.move(righe, ElementPlacement.PLACEATEND);
    }
    for (var i = 1; i < ripPasso; i++) {
        righe.move(etichette, ElementPlacement.PLACEATEND);
        righe.duplicate();
        righe.translate(0, obj.height + distFascia);
    }
    if (ripPasso == 1) righe.move(etichette, ElementPlacement.PLACEATEND);
}

function dima() {
    try {
        bottoniera = doc.layers["BOTTONIERA"];
    } catch (e) {
        bottoniera = doc.layers.add();
        bottoniera.name = "BOTTONIERA";
    } finally {
        bottoniera.locked = false;
        gruppoBottoniera = bottoniera.groupItems.add();
        diciture = gruppoBottoniera.groupItems.add();
        testi = diciture.groupItems.add();
        nomiColore = diciture.groupItems.add();
        dataG = diciture.groupItems.add();

        var ldfSX = lineaDiFede();
        ldfSX.translate(etichette.controlBounds[0], etichette.controlBounds[1] - etichette.height / 2 + ldfSX.height / 2);
        var posSX = ldfSX.visibleBounds;

        var ldfDX = lineaDiFede();
        ldfDX.translate(etichette.controlBounds[2] + 10.3412 + k, etichette.controlBounds[1] - etichette.height / 2 + ldfDX.height / 2);
        var posDX = ldfDX.visibleBounds;

        var croSx = creaCrocino();
        croSx.translate(posSX[0] + ldfSX.width / 2, posSX[3] + ldfSX.height / 2);

        var croDx = creaCrocino();
        croDx.translate(posDX[0] + ldfDX.width / 2, posDX[3] + ldfDX.height / 2);

        var crTopSX = crociniEstremi();
        crTopSX.translate(posSX[0] + ldfSX.width / 2, posSX[1]);

        var crBotSX = crociniEstremi();
        crBotSX.rotate(180);
        crBotSX.translate(posSX[0] + ldfSX.width / 2, posSX[3] + 78.8375);

        var crTopDX = crociniEstremi();
        crTopDX.translate(posDX[0] + ldfDX.width / 2, posDX[1]);

        var crBotDX = crociniEstremi();
        crBotDX.rotate(180);
        crBotDX.translate(posDX[0] + ldfDX.width / 2, posDX[3] + 78.8375);

        stringa();
        diciture.translate(posSX[0] - diciture.width / 2 + 6.332, posSX[1]);
        diciture.rotate(270);
        diciture.translate(0, -diciture.height / 2 - 85);

        bottoniInferiori(posSX[3] + 95.67, posSX[0] + 0.7046);
        bottoniInferiori(posDX[3] + 95.67, posDX[0] + 0.7046);

        if (stampante.printer == "Lombardi") markLombardi(croSx.visibleBounds, croSx.height / 2);
        distorsione();
        quote();
        adattaArtboard();
    }
}

function adattaArtboard() {
    var pezzaFinestra = bottoniera.pathItems.rectangle(
        gruppoBottoniera.controlBounds[1] + 17,
        gruppoBottoniera.controlBounds[0] - 17,
        gruppoBottoniera.width + 34,
        gruppoBottoniera.height + 34);
    pezzaFinestra.filled = false;
    pezzaFinestra.stroked = false;
    doc.selectObjectsOnActiveArtboard()
    doc.fitArtboardToSelectedArt(0);
    app.executeMenuCommand("fitall");
    pezzaFinestra.remove();
}

function quote() {
    try {
        var quoteLVL = doc.layers["QUOTE"];
    } catch (e) {
        var quoteLVL = doc.layers.add();
        quoteLVL.name = "QUOTE";
    } finally {
        var quotaLinea = quoteLVL.compoundPathItems.add();
        var lineaOr = quotaLinea.pathItems.add();
        lineaOr.setEntirePath([
            [gruppoBottoniera.controlBounds[0], gruppoBottoniera.controlBounds[1] + 25],
            [gruppoBottoniera.controlBounds[2], gruppoBottoniera.controlBounds[1] + 25]
        ]);
        lineaOr.filled = false;
        lineaOr.strokeColor = coloreTecnicoIndex > 0 ? listaColori[coloreTecnicoIndex] : nero;

        var lineaLatSX = quotaLinea.pathItems.add();
        lineaLatSX.setEntirePath([
            [lineaOr.controlBounds[0], lineaOr.controlBounds[1] + 5],
            [lineaOr.controlBounds[0], lineaOr.controlBounds[1] - 5]
        ]);

        var lineaLatDX = quotaLinea.pathItems.add();
        lineaLatDX.setEntirePath([
            [lineaOr.controlBounds[2], lineaOr.controlBounds[1] + 5],
            [lineaOr.controlBounds[2], lineaOr.controlBounds[1] - 5]
        ]);
    }

    var dimFasciaTot = quoteLVL.textFrames.add();
    dimFasciaTot.translate(lineaOr.controlBounds[0] + lineaOr.width / 2, lineaOr.controlBounds[1] + 2);
    dimFasciaTot.contents = (lineaOr.width / mm).toFixed(2) + " mm";
    dimFasciaTot.fillColor = coloreTecnicoIndex > 0 ? listaColori[coloreTecnicoIndex] : nero;
    dimFasciaTot.paragraphs[0].paragraphAttributes.justification = Justification.CENTER;
}

function distorsione() {
    var cil = zetaValue * stampante.pi;
    var percentuale = (((cil - stampante.all) * 100) / cil).toFixed(3);
    gruppoBottoniera.resize(100, percentuale);
    etichette.resize(100, percentuale);
}

function indexSel(obj) {
    for (var i = 0; i < obj.length; i++) {
        if (obj[i].selected) {
            return i;
        }
    }
}

function markLombardi(pos, alt) {
    var lombMark = gruppoBottoniera.pathItems.rectangle(pos[1] - alt + 7.0866, pos[0] - 14.8962, 14.1732, 14.1732);
    lombMark.fillColor = ciano;
    lombMark.stroked = false;
}

function bottoniInferiori(posX, posY) {
    var bottoni = gruppoBottoniera.groupItems.add();
    var listaRev = [].concat(listaColori).reverse();
    for (var i = 0; i < listaColori.length; i++) {
        var posX = creaCerchi(bottoni, listaRev[i], posX, posY, 1);
    }
    for (var i = 0; i < listaColori.length; i++) {
        var posX = creaCerchi(bottoni, listaColori[i], posX, posY, 0);
    }
}

function creaCerchi(parent, colore, pos1, pos2, elem) {
    var tutto = parent.groupItems.add();
    var cerchio = tutto.pathItems.ellipse(pos1, pos2, 7.51, 7.51);
    cerchio.fillColor = bianco;
    cerchio.stroked = false;
    if (elem) {
        var croceGr = parent.compoundPathItems.add();
        var croce = croceGr.pathItems.add();
        croce.setEntirePath([
            [0, -2.928],
            [0, 2.928]
        ]); //5.856 2.928
        croce = croceGr.pathItems.add();
        croce.setEntirePath([
            [-2.928, 0],
            [2.928, 0]
        ]);
        croce.fillColor = noColor;
        croce.stroked = true;
        croce.strokeWidth = 0.57;
        croce.strokeColor = colore;
        croceGr.translate(cerchio.visibleBounds[0] + 3.76, cerchio.visibleBounds[1] - 3.76);
    } else {
        var cerchioInt = tutto.pathItems.ellipse(pos1 - .5, pos2 + .5, 6.51, 6.51);
        cerchioInt.fillColor = colore;
        cerchioInt.stroked = false;
    }
    return pos1 + 14.17;
}

function crociniEstremi() {
    
    var linBianca = gruppoBottoniera.pathItems.add();
    linBianca.setEntirePath(Array(Array(0, -14.18), Array(0, 0)));
    linBianca.fillColor = noColor;
    linBianca.stroked = true;
    linBianca.strokeWidth = 0.851;
    linBianca.strokeColor = bianco;

    return linBianca;
    // Vecchio codice con crocino grande con fondo bianco

    /*
    var gruppo = gruppoBottoniera.groupItems.add();
    var linBianca = gruppo.pathItems.add();
    linBianca.setEntirePath(Array(Array(0, -4.7357), Array(0, 0)));
    linBianca.fillColor = noColor;
    linBianca.stroked = true;
    linBianca.strokeWidth = 0.851;
    linBianca.strokeColor = bianco;

    var pezzaCrocino = gruppo.pathItems.rectangle(-4.7357, -2.89, 2.89 * 2, 74.1018);
    pezzaCrocino.fillColor = bianco;
    pezzaCrocino.stroked = false;

    var crocioneGR = gruppo.compoundPathItems.add();
    var crocione = crocioneGR.pathItems.add();
    crocione.setEntirePath(Array(Array(0, -78.8375), Array(0, -4.7357)));
    crocione = crocioneGR.pathItems.add();
    crocione.setEntirePath(Array(Array(-2.89, -41.7866), Array(2.89, -41.7866)));
    crocione.fillColor = noColor;
    crocione.stroked = true;
    crocione.strokeWidth = 0.851;
    crocione.strokeColor = registro; 

    return gruppo;
    */
}

function lineaDiFede() {
    var linea = gruppoBottoniera.pathItems.rectangle(0, -k, -8.9239, (zetaValue * stampante.pi) * mm);
    linea.fillColor = registro;
    linea.fillColor.tint = 25.0;
    linea.stroked = false;
    return linea;
}

function creaCrocino() {
    var CrocinoGr = gruppoBottoniera.groupItems.add();

    var pezzaCrocino = CrocinoGr.pathItems.rectangle(25.7996, -3.739, 7.4796, 51.5993);
    pezzaCrocino.fillColor = bianco;
    pezzaCrocino.stroked = false;

    var lineaVert = CrocinoGr.pathItems.add();
    lineaVert.setEntirePath(Array(Array(0, -18.438), Array(0, 18.438)));
    lineaVert.fillColor = noColor;
    lineaVert.stroked = true;
    lineaVert.strokeWidth = 0.851;
    lineaVert.strokeColor = registro;

    var lineaOr = CrocinoGr.pathItems.add();
    lineaOr.setEntirePath(Array(Array(-3.67, 0), Array(3.67, 0)));
    lineaOr.fillColor = noColor;
    lineaOr.stroked = true;
    lineaOr.strokeWidth = 0.567;
    lineaOr.strokeColor = registro;

    var ellisseCrocino = CrocinoGr.pathItems.ellipse(3.39, -3.39, 6.78, 6.78);
    ellisseCrocino.fillColor = noColor;
    ellisseCrocino.stroked = true;
    ellisseCrocino.strokeWidth = 0.567;
    ellisseCrocino.strokeColor = registro;

    return CrocinoGr;
}

function listaColoriFunc() {
    var ListaInk = doc.inkList;
    //CMYK
    for (var j = 0; j < ListaInk.length; j++) {
        for (var i = 0; i < doc.swatches.length; i++) {
            if (ListaInk[j].name == doc.swatches[i].name) {
                var pantone = doc.swatches[i].name.toUpperCase();
                listaColoriStringa.push(pantone.replace("PANTONE", "P."));
                listaColori.push(doc.swatches[i].color);
            }
        }
        if (ListaInk[j].inkInfo.kind == InkType.BLACKINK) {
            listaColoriStringa.push("Nero");
            listaColori.push(nero);
        }
        if (ListaInk[j].inkInfo.kind == InkType.CYANINK) {
            listaColoriStringa.push("Ciano");
            listaColori.push(ciano);
        }
        if (ListaInk[j].inkInfo.kind == InkType.YELLOWINK) {
            listaColoriStringa.push("Giallo");
            listaColori.push(giallo);
        }
        if (ListaInk[j].inkInfo.kind == InkType.MAGENTAINK) {
            listaColoriStringa.push("Magenta");
            listaColori.push(mag);
        }
    }
}

function colore() {
    var ListaInk = doc.inkList;
    //CMYK
    var posizione = [0, 0];

    for (var j = 0; j < ListaInk.length; j++) {
        if (j == coloreTecnicoIndex) continue;
        if (ListaInk[j].inkInfo.printingStatus != InkPrintStatus.DISABLEINK) {
            for (var i = 0; i < doc.swatches.length; i++) {
                if (ListaInk[j].name == doc.swatches[i].name) {
                    var nomePantone = doc.swatches[i].name.toUpperCase();
                    var newpos = nomeColore(nomePantone.replace("PANTONE ", "P."), doc.swatches[i].color, posizione);
                    posizione = posizione.splice();
                    posizione = [newpos[0], newpos[1]];
                    listaColori.push(doc.swatches[i].color);
                }
            }
            if (ListaInk[j].inkInfo.kind == InkType.BLACKINK) {
                var newpos = nomeColore("K", nero, posizione);
                posizione = posizione.splice();
                posizione = [newpos[0], newpos[1]];
                listaColori.push(nero);
            }
            if (ListaInk[j].inkInfo.kind == InkType.CYANINK) {
                var newpos = nomeColore("C", ciano, posizione);
                posizione = posizione.splice();
                posizione = [newpos[0], newpos[1]];
                listaColori.push(ciano);
            }
            if (ListaInk[j].inkInfo.kind == InkType.YELLOWINK) {
                var newpos = nomeColore("G", giallo, posizione);
                posizione = posizione.splice();
                posizione = [newpos[0], newpos[1]];
                listaColori.push(giallo);
            }
            if (ListaInk[j].inkInfo.kind == InkType.MAGENTAINK) {
                var newpos = nomeColore("M", mag, posizione);
                posizione = posizione.splice();
                posizione = [newpos[0], newpos[1]];
                listaColori.push(mag);
            }
        }
    }
}

function nomeColore(nome, tinta, posizione) {
    var gr = nomiColore.groupItems.add();
    var col = new creaTesto(gr, tinta, nome)
    gr.translate(posizione[0] - (gr.controlBounds[3] - 2.5), posizione[1]);
    var nuovaposizione = [(gr.visibleBounds[2] - 3.58), (gr.visibleBounds[3] + 1.25)]
    return nuovaposizione;
}

function stringaNome() {
    var PosInit = doc.name.lastIndexOf(".");
    var NomeEnd = doc.name.substring(0, PosInit).toUpperCase();
    var nome = new creaTesto(testi, registro, NomeEnd + " - Pol." + stampante.polimero + " - Z" + zetaValue);
}

function data() {
    var oggi = new Date();
    var mese = (oggi.getMonth() + 1);
    if (mese < 10)
        var date = oggi.getDate() + '/0' + mese + '/' + oggi.getFullYear();
    else
        var date = oggi.getDate() + '/' + mese + '/' + oggi.getFullYear();
    var dataTxt = new creaTesto(dataG, registro, date);
}

function stringa() {
    colore();
    stringaNome();
    data();
    nomiColore.translate(testi.controlBounds[2] - 7, 0);
    dataG.translate(nomiColore.controlBounds[2] - 1, 0);
    var pos = [diciture.controlBounds[0], diciture.controlBounds[1], diciture.controlBounds[2], diciture.controlBounds[3]];
    var pezza = diciture.pathItems.rectangle(pos[0], pos[0] - 1, diciture.width + 3, -5.7798);
    pezza.fillColor = bianco;
    pezza.stroked = false;
    pezza.zOrder(ZOrderMethod.SENDTOBACK);
    diciture.zOrder(ZOrderMethod.BRINGTOFRONT);
}

function creaTesto(parent, colore, contenuto) {
    this.parent = parent;
    this.instance = this.parent.textFrames.add();
    this.instance.textRange.fillColor = colore;
    this.instance.textRange.characterAttributes.size = 5;
    this.instance.textRange.characterAttributes.textFont = app.textFonts.getByName("MyriadPro-BoldCond");
    this.instance.contents = contenuto;
}