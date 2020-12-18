(function () {
    // "use strict";

    //variable globale nom fichier
    var loadName
    //variable globale nom author
    var loadAuthor
    //variable globale contenant tous les chiffres
    var numberCondtion
    //variable globale position pour ecrire les modifications
    var PositionColon

    var NbrCells

    // TODO réservation avec countcell
    var AdressSaisies
    var UsedRangeval
    var RandomValBegin

    // La fonction d'initialisation doit être exécutée chaque fois qu'une nouvelle page est chargée
    Office.initialize = function (reason) {
        $(document).ready(function () {

            Office.context.document.settings.set("Office.AutoShowTaskpaneWithDocument", true);
            Office.context.document.settings.saveAsync();

            Excel.run(function (context) {

                var worksheet = context.workbook.worksheets.getItem("Data");
                //on ajoute un evenement pour chaque changement dans la feuille Data
                worksheet.onChanged.add(loadEvent);

                var actShWS = context.workbook.worksheets.getItem('CBcapInput').getUsedRange().load("address")

                return context.sync()
                    .then(function () {
                        AdressSaisies = actShWS.address
                        console.log("Event handler successfully registered for onChanged event in the worksheet.");
                    });
            })

        });
    };


    //fonction appelée a chaque nouvelle évenement
    function loadEvent(event) {
        //variable stockant tous les chiffres
        numberCondtion = new RegExp("\\d+", "g");

        Excel.run(function (context) {

            if (event.source == "Local") {

                //on load le fait de pouvoir activer/desactiver l'evenement
                context.runtime.load("enableEvents");
                //on charge le nom du fichier
                var TemploadName = context.workbook.load("name");
                //on charge le nom de l'author
                var TemploadNameloadAuthor = context.workbook.properties.load("author");

                // Feuille CBcapVersion
                var WsVersion = context.workbook.worksheets.getItem("CBcapVersion");
                var CapVersion = WsVersion.getCell(0, 1).load("values");
                var InputActive = WsVersion.getCell(2, 1).load("values");
                var InputSingleCell = WsVersion.getCell(3, 1).load("values");
                var InputLineMax = WsVersion.getCell(4, 1).load("values");

                var TestTemp = event.address
                var InputLineUsed = context.workbook.worksheets.getItem('CBcapInput').getUsedRange().load("rowCount");
                var TempNbrCells = context.workbook.worksheets.getItem('Data').getRange(event.address).load("cellCount")
                var TempUsedRange = context.workbook.worksheets.getItem('CBcapInput').getUsedRange().load("address")
                context.workbook.worksheets.getItem('Data').getUsedRange().format.font.italic = false;

               // context.runtime.load("enableEvents");
             
                return context.sync()

                    .then(function () {

                        NbrCells = TempNbrCells.cellCount;

                        UsedRangeval = TempUsedRange.address;

                        var actShWS = context.workbook.worksheets.getItem('CBcapInput').getUsedRange().getRowsBelow(NbrCells);

                        RandomValBegin = Math.floor(Math.random() * 1000000);
                       
                        actShWS.getCell(NbrCells, 0).values = "2";
                        actShWS.getCell(NbrCells, 1).values = RandomValBegin;

                        return context.sync()

                            .then(function () {

                                    // ATTENTION NOTE : Hard change version bellow to stop all user with previous complement version
                                    if (CapVersion.values == 1) {

                                        if (InputActive.values == 1) {

                                            if (InputLineUsed.rowCount < InputLineMax.values) {
 
                                                loadName = TemploadName.name;
                                                loadAuthor = TemploadNameloadAuthor.author;

                                                //on active l'evenement
                                                context.runtime.enableEvents = false;

                                                //condition pour limiter à 1 cellule
                                                var IsComa = new RegExp(",", "g");
                                                if (!(event.address.match(IsComa))) {

                                                    var IsColon = new RegExp(":", "g");
                                                    if (!(event.address.match(IsColon))) {

                                                        Treatment(event.address);

                                                    } else {

                                                        if (InputSingleCell.values == "") {

                                                            Treatment(event.address);

                                                        } else {

                                                            context.runtime.enableEvents = false;
                                                            var ShData = context.workbook.worksheets.getItem('Data');
                                                            ShData.getRange(event.address).format.fill.color = "red";
                                                            ShData.getRange(event.address).format.font.italic = false;
                                                            ShData.getRange(event.address).format.font.bold = true;
                                                            context.runtime.enableEvents = true;
                                                            TreatmentError("Error Column")
                                                            return context.sync()
                                                        }
                                                    }
                                                } else {
                                                    if (InputSingleCell.values == "") {
                                                        Treatment(event.address);
                                                    } else {
                                                        //  TODO faire une boucle sur tous les éléments d'un tableau éclaté par virgule
                                                        context.runtime.enableEvents = false;
                                                        var ShData = context.workbook.worksheets.getItem('Data');
                                                        ShData.getRange(event.address).format.fill.color = "red";
                                                        ShData.getRange(event.address).format.font.italic = false;
                                                        ShData.getRange(event.address).format.font.bold = true;
                                                        context.runtime.enableEvents = true;
                                                        TreatmentError("Error Column")
                                                        return context.sync();
                                                    }
                                                }
                                            } else {
                                                context.runtime.enableEvents = false;
                                                var ShData = context.workbook.worksheets.getItem('Data');
                                                ShData.getRange(event.address).format.fill.color = "red";
                                                ShData.getRange(event.address).format.font.italic = false;
                                                ShData.getRange(event.address).format.font.bold = true;
                                                context.runtime.enableEvents = true;
                                                TreatmentError("Error Column")
                                                return context.sync();
                                            }
                                        } else {
                                            context.runtime.enableEvents = false;
                                            var ShData = context.workbook.worksheets.getItem('Data');
                                            ShData.getRange(event.address).format.fill.color = "red";
                                            ShData.getRange(event.address).format.font.italic = false;
                                            ShData.getRange(event.address).format.font.bold = true;
                                            context.runtime.enableEvents = true;
                                            TreatmentError("Error Column")
                                            return context.sync();
                                        }
                                    } else {
                                        context.runtime.enableEvents = false;
                                        var ShData = context.workbook.worksheets.getItem('Data');
                                        ShData.getRange(event.address).format.fill.color = "red";
                                        ShData.getRange(event.address).format.font.italic = false;
                                        ShData.getRange(event.address).format.font.bold = true;
                                        context.runtime.enableEvents = true;
                                        TreatmentError("Error Column")
                                        return context.sync();
                                    }
                    })
                 })
            }
        })
    }
    function YearValue(content) {
        return content.getFullYear();
    }
    function MonthValue(content) {
        return content.getMonth();
    }
    function DayValue(content) {
        return content.getDate();
    }
    function HourValue(content) {
        return content.getHours();
    }
    function MinuteValue(content) {
        return content.getMinutes();
    }
    function SecondValue(content) {
        return content.getSeconds();
    }
    function MilliSecondValue(content) {
        return content.getMilliseconds();
    }

    // fonction qui traite les données pour les afficher
    function Treatment(event) {
        Excel.run(async function (ctx) {

            ctx.application.suspendScreenUpdatingUntilNextSync();
            var ShData = ctx.workbook.worksheets.getItem('Data');
            ShData.getRange(event).format.fill.color = "red";

            //on enleve les espaces de l'adresse
            const result = event.split(" ").join("");
            //on remplace les virgules par des espaces
            var ArrayListAdressChanged = result.split(",");

            PositionColon = 0;
            await ctx.sync()

            for (let pas = 0; pas <= ArrayListAdressChanged.length - 1; pas++) {

                var ShData = ctx.workbook.worksheets.getItem('Data');
               
                //event prend la valeur de l'element du tableau actuel
                event = ArrayListAdressChanged[pas]

                //on chercher si il y a un ":"
                var SearchColon = event.indexOf(":") + 1;

                var LengthTemp = event.length;

                //on prend la partie avant les :
                var SliceTempFirst = event.slice(0, SearchColon - 1);
                //si il n'y a pas de : alors SliceTempFirst prend la valeur event
                if (SearchColon == 0) {
                    SliceTempFirst = event;
                }
                
                // Numéro de la 1er colonne de la plage ou de la colonne 
                var NumColonne = ShData.getRange(SliceTempFirst).load('columnIndex');

                // Nombre de colonnes concernée
                var NbrColonne = ShData.getRange(event).load('columnCount');

                await ctx.sync()

                //on coupe en partant de SearColon jusqua la fin de l'adresse, avec D4:D5 -> D5
                var SliceTempEnd = event.slice(SearchColon, LengthTemp);
                //variable qui prend le caractere gauche, avec D4:F6 -> D
                var CharacterLeft = SliceTempFirst.replace(numberCondtion, "");
                //variable qui prend le caractere droit, avec  D4:F6 -> F
                var CharacterWrite = SliceTempEnd.replace(numberCondtion, "");
                //chiffre de la partie de gauche montrant le debut
                var beginCond = SliceTempFirst.slice(CharacterLeft.length, SliceTempFirst.length);
                //chiffre de la partie de droite montrant la fin
                var endCond = SliceTempEnd.slice(CharacterWrite.length, SliceTempEnd.length);

                //on force le numero de la colonne en float
                var tempColumnInd = parseFloat(NumColonne.columnIndex);
                // définit le chiffre de la derniere colonne de la plage
                var conditionMultiRange = tempColumnInd + (parseFloat(NbrColonne.columnCount) - 1);

                //on force beginCond en float
                var valBegin = parseFloat(beginCond);
                var valEnd = parseFloat(endCond);

                //boucle for de la 1er colonne a la derniere
                for (let it = tempColumnInd; it <= conditionMultiRange; it++) {
                    //on se place sur la 1er ligne pour affichier les modifs
 
                    //on reserve une plage pour ecrire 
                    var actShWS = ctx.workbook.worksheets.getItem('CBcapInput').getRange(UsedRangeval).getRowsBelow(NbrCells);
                    //on met l'adresse de la plage réservée dans une variable
                    var RowsbelowAdress = actShWS.load("address");

                    //variable pour manipuler les feuilles Data et Parameter
                    var ShCBcapParameter = ctx.workbook.worksheets.getItem('CBcapParameter');
                    var ShData = ctx.workbook.worksheets.getItem('Data');

                    //on recupere le titre de la colonne modifée
                    var col = ShData.getCell(0, it).load("values");

                    //on recupere les informations de la page Parameter pour la colonne modifiee
                    var IndexNameColNewContent = ShCBcapParameter.getCell(0, it).load("values");
                    var IndexOneOrNot = ShCBcapParameter.getCell(1, it).load("values");
                    var IndexOP = ShCBcapParameter.getCell(2, it).load("values");
                    var IndexTypeMA = ShCBcapParameter.getCell(5, it).load("values");
                    var IndexNameKeyOrignTitle = ShCBcapParameter.getCell(6, it).load("values");
                    var IndexRefKeyOrignTitle = ShCBcapParameter.getCell(7, it).load("values");
                    var IndexValColRef = ShCBcapParameter.getCell(9, it).load("values");
                    var IndexKeyTargetTitle = ShCBcapParameter.getCell(10, it).load("values");
                    var IndexContentTargetTitle = ShCBcapParameter.getCell(11, it).load("values");

                    var IndexOneOrNotBegin = ShCBcapParameter.getCell(3, it).load("values");
                    var IndexWrongColumn = ShCBcapParameter.getCell(2, it).load("values");

                    await ctx.sync();

                    adressTest = RowsbelowAdress.address;

                    var LoadShKeyOrignTitle = "Data!" + IndexRefKeyOrignTitle.values + "1";

                    if (LoadShKeyOrignTitle == "Data!1") {
                        LoadShKeyOrignTitle = "Data!A1";
                    }

                    var IndexLoadNameKeyOrignTitle = ShData.getRange(LoadShKeyOrignTitle).load("values");

                    await ctx.sync();

                    var ShKeyOrignTitle = IndexLoadNameKeyOrignTitle.values.toString();

                    if ((IndexOneOrNotBegin.values != "") && (IndexWrongColumn.values != "")) {

                        if (IndexOneOrNot.values.toString() != "") {

                            if (IndexNameKeyOrignTitle.values == ShKeyOrignTitle) {

                                if (col.values.toString() == IndexNameColNewContent.values.toString()) {

                                    //on enleve l'italique de la colonne
                                    ShData.getUsedRange().getColumn(it).format.font.italic = false;

                                    for (let pas = valBegin; pas <= valEnd; pas++) {

                                        var ShData = ctx.workbook.worksheets.getItem('Data');

                                        var IndexAddress = ShData.getCell(pas - 1, it).load("address");

                                        var IndexvalSelected = ShData.getCell(pas - 1, it).load("values");

                                        var number = pas;
                                        // ex Data!BV3   -> adresse de la clé 

                                        var RefCol = "Data!" + IndexRefKeyOrignTitle.values.toString() + number;
                                        var IndexCellsRef = ShData.getRange(RefCol).load("values");

                                        await ctx.sync();

                                        var CellsRef = IndexCellsRef.values.toString();

                                        await ctx.sync();

                                        if (IndexCellsRef.values != "") {

                                            var ShCBcapInput = ctx.workbook.worksheets.getItem('CBcapInput').getRange(RowsbelowAdress.address);

                                            //cherche si il y a d'autres clés indentiques
                                            if (CellsRef != "") {
                                                var IndexRangeRef = ShData.findAll(CellsRef, { completeMatch: true }).load("address");
                                            }

                                            ctx.runtime.load("enableEvents");

                                            await ctx.sync()

                                                .then(function () {

                                                    ctx.runtime.enableEvents = false;

                                                    //on charge la date et l'heure de la modif
                                                    var ladate = new Date();

                                                    // Input
                                                    ShCBcapInput.getCell(PositionColon, 0).values = "1";
						    ShCBcapInput.getCell(PositionColon, 1).values = RandomValBegin;
                                                    ShCBcapInput.getCell(PositionColon, 3).values = IndexTypeMA.values.toString();

                                                    ShCBcapInput.getCell(PositionColon, 4).values = loadName;

                                                    ShCBcapInput.getCell(PositionColon, 6).values = IndexNameKeyOrignTitle.values.toString();
                                                    ShCBcapInput.getCell(PositionColon, 7).values = col.values.toString();
                                                    ShCBcapInput.getCell(PositionColon, 8).values = IndexKeyTargetTitle.values.toString();
                                                    ShCBcapInput.getCell(PositionColon, 9).values = IndexContentTargetTitle.values.toString();
                                                    ShCBcapInput.getCell(PositionColon, 10).values = IndexCellsRef.values;
                                                    ShCBcapInput.getCell(PositionColon, 11).values = IndexvalSelected.values;
                                                    ShCBcapInput.getCell(PositionColon, 12).values = YearValue(ladate);
                                                    ShCBcapInput.getCell(PositionColon, 13).values = MonthValue(ladate);
                                                    ShCBcapInput.getCell(PositionColon, 14).values = DayValue(ladate);
                                                    ShCBcapInput.getCell(PositionColon, 15).values = HourValue(ladate);
                                                    ShCBcapInput.getCell(PositionColon, 16).values = MinuteValue(ladate);
                                                    ShCBcapInput.getCell(PositionColon, 17).values = SecondValue(ladate);
                                                    ShCBcapInput.getCell(PositionColon, 18).values = MilliSecondValue(ladate);
                                                    ShCBcapInput.getCell(PositionColon, 19).values = NbrCells;

                                                    ShCBcapInput.getCell(PositionColon, 20).values = Math.floor(Math.random() * 1000000);

                                                    if (IndexOP.values.toString() == "InputOnLineByChangeKey") {

                                                        //VerificationRange(IndexRangeRef.address, IndexRefKeyOrignTitle.values.toString(), IndexValColRef.values.toString(), IndexvalSelected.values);
                                                        var tempRangeRef = IndexRangeRef.address;

                                                        //ctx.workbook.worksheets.getItem('Data').getRange(tempRangeRef).format.fill.color = "red";
                                                        //on prend content et on enleve Data! ce qui donnerait avec Data!Bv4  -> Bv4
                                                        const result = tempRangeRef.split("Data!").join("");
                                                        //on remplace les virugles par des espaces , avec Data!B4,Data!Bv5 -> Data!B4 Data!Bv5
                                                        var ArrayAdress = result.split(",");

                                                        //on parcourt notre tableau de valeurs
                                                        for (let pas = 0; pas <= ArrayAdress.length - 1; pas++) {
                                                            //on enleve les chiffres de valeur actuelle
                                                            var tempCaracter = ArrayAdress[pas].replace(numberCondtion, "");

                                                            //si ses lettres de colonnes sont égales à la colonne voulue
                                                            if (tempCaracter == IndexRefKeyOrignTitle.values.toString()) {
                                                                //on remplace ses lettres par celles de la colonne modifiee
                                                                tempCaracter = ArrayAdress[pas].replace(tempCaracter, IndexValColRef.values.toString());

                                                                //on inscrit la valeur
                                                                ShData.getRange(tempCaracter).format.fill.clear();
                                                                ShData.getRange(tempCaracter).values = IndexvalSelected.values;
                                                            }
                                                        }
                                                        ArrayAdress = null;
                                                    }
                                                    //on affiche les informations
                                                    ShData.getRange(IndexAddress.address).format.fill.clear();
                                                    //ShData.getRange(event).format.fill.color = "green";
                                                    ShData.getRange(IndexAddress.address).format.font.italic = true;
                                                    ShData.getRange(IndexAddress.address).format.font.bold = false;

                                                    ctx.runtime.enableEvents = true;
                                                })
                                        } else {
                                            ShData.getRange(event).clear;
                                            ShData.getRange(event).format.fill.color = "red";
                                            ShData.getRange(event).format.font.italic = false;
                                            ShData.getRange(event).format.font.bold = true;
                                            TreatmentError("Error Column")
                                        }
                                        //on prend la ligne suivante
                                        PositionColon = PositionColon + 1;
                                    }
                                } else {
                                    ShData.getRange(event).clear;
                                    ShData.getRange(event).format.fill.color = "red";
                                    ShData.getRange(event).format.font.italic = false;
                                    ShData.getRange(event).format.font.bold = true;
                                    TreatmentError("Error Column")
                                }
                            } else {
                                ShData.getRange(event).clear;
                                ShData.getRange(event).format.fill.color = "red";
                                ShData.getRange(event).format.font.italic = false;
                                ShData.getRange(event).format.font.bold = true;
                                TreatmentError("Error Column")
                            }
                        } else {
                            ShData.getRange(event).clear;
                            ShData.getRange(event).format.fill.color = "red";
                            ShData.getRange(event).format.font.italic = false;
                            ShData.getRange(event).format.font.bold = true;
                            TreatmentError("Error Column")
                        }
                    } else {
                        ShData.getRange(event).clear;
                        ShData.getRange(event).format.fill.color = "red";
                        ShData.getRange(event).format.font.italic = false;
                        ShData.getRange(event).format.font.bold = true;
                        TreatmentError("Error Column")
                    }
                    ShData = null;
                    //on prend la colonne suivante
                    tempColumnInd = tempColumnInd + 1;
                }
            }

            // NOTE ATTENTION PROD DEPLOYEMENT : activer la ligne de dessous lors du déployement
            location.reload();

            return ctx.sync()
                .then(function () {

                    var LoadRandom = ShCBcapInput.getCell(PositionColon , 1).load("values")

                    return ctx.sync()

                        .then(function () {
                            if (LoadRandom.values == RandomValBegin) {
                            } else {
                                ctx.workbook.worksheets.getItem('Data').getRange(event).format.fill.color = "red"
                                ShCBcapInput = null;
                                TreatmentError("Syncro error")
                                return ctx.sync()
                            }
                        })
                })
        })
    } 

    function TreatmentError(txt) {

        Excel.run(async function (ctx) {

                    //on reserve une plage pour ecrire 
                      var actShWS = ctx.workbook.worksheets.getItem('CBcapInput').getUsedRange().getRowsBelow(2);
                    //var actShWS = ctx.workbook.worksheets.getItem('CBcapInput').getRange(UsedRangeval).getRowsBelow(1);
                    //on met l'adresse de la plage réservée dans une variable
                    var RowsbelowAdress = actShWS.load("address");

                    await ctx.sync();

                    var ShCBcapInput = ctx.workbook.worksheets.getItem('CBcapInput').getRange(RowsbelowAdress.address);

                    ctx.runtime.load("enableEvents");

                     await ctx.sync()

                     .then(function () {

                         for (let pas = 0; pas <= 1; pas++) {

                             ctx.runtime.enableEvents = false;

                             // Input
                             ShCBcapInput.getCell(0, 0).values = "-1";
                             ShCBcapInput.getCell(0, 4).values = txt;

                             if (pas == 1) {
                                 ShCBcapInput.getCell(1, 0).values = "2";
                              }
                              ctx.runtime.enableEvents = true;
                          }

                       })

            location.reload();
            return ctx.sync();
        })
    }
})();