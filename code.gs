function onOpen() {
    var ui = SpreadsheetApp.getUi();
    ui.createMenu('Pokemon')
        .addItem('Get pokemon sprite and number', 'getPokemonImageNumber')
        .addItem('Get pokemon shiny sprite and number', 'getPokemonShinyImageNumber')
        .addToUi();
}

function RegExpPokemonSpecies(string) {
    return string.match(/.*(?=-)/)[0];
}

function RegExpPokemonForm(string) {
    return string.match(/(?<=-).+/)[0];
}

function fetchPokemonData(pokemon) {
    var apiUrl = "https://pokeapi.co/api/v2/pokemon/";
    var response = UrlFetchApp.fetch(apiUrl + pokemon, { muteHttpExceptions: true });

    if (response.getResponseCode() === 404) {
        response = UrlFetchApp.fetch("https://pokeapi.co/api/v2/pokemon-species/" + pokemon, { muteHttpExceptions: true });
        if (response.getResponseCode() === 404) {
            var pokemonSpecies = RegExpPokemonSpecies(pokemon).toLowerCase();
            var pokemonRegion = RegExpPokemonForm(pokemon).toLowerCase();
            response = UrlFetchApp.fetch("https://pokeapi.co/api/v2/pokemon-species/" + pokemonSpecies, { muteHttpExceptions: true });
            response = JSON.parse(response);

            var formIndex = 0;
            while (formIndex < response.varieties.length) {
                if (response.varieties[formIndex].pokemon.name.match(pokemonRegion) != null) {
                    response = UrlFetchApp.fetch(apiUrl + response.varieties[formIndex].pokemon.name, { muteHttpExceptions: true });
                    break;
                }
                formIndex++;
            }
        } else {
            response = JSON.parse(response);
            response = UrlFetchApp.fetch(apiUrl + response.varieties[0].pokemon.name, { muteHttpExceptions: true });
        }
    }
    return JSON.parse(response);
}

function getPokemonImageNumber() {
    var activeSheet = SpreadsheetApp.getActiveSheet();
    var selection = activeSheet.getSelection();
    var topCellLocation = [selection.getActiveRange().getRow(), selection.getActiveRange().getColumn()];
    var activeList = selection.getActiveRange().getValues();

    for (var i = 0; i < activeList.length; i++) {
        var pokemon = String(activeList[i]).toLowerCase();
        var response = fetchPokemonData(pokemon);

        activeSheet.getRange(topCellLocation[0] + i, topCellLocation[1] - 1).setValue("=SPLIT(\"" + response.sprites.front_default + "\", \"qwertyuiopasdfghjklzxcvbnmQWERTYUIOPASDFGHJKLZXCVBNM`-=[]:\;',./!@#$%^&*()\")");
        activeSheet.getRange(topCellLocation[0] + i, topCellLocation[1] + 1).setValue("=IMAGE(\"" + response.sprites.front_default + "\", 2)");
    }
}

function getPokemonShinyImageNumber() {
    var activeSheet = SpreadsheetApp.getActiveSheet();
    var selection = activeSheet.getSelection();
    var topCellLocation = [selection.getActiveRange().getRow(), selection.getActiveRange().getColumn()];
    var activeList = selection.getActiveRange().getValues();

    for (var i = 0; i < activeList.length; i++) {
        var pokemon = String(activeList[i]).toLowerCase();
        var response = fetchPokemonData(pokemon);

        activeSheet.getRange(topCellLocation[0] + i, topCellLocation[1] - 1).setValue("=SPLIT(\"" + response.sprites.front_shiny + "\", \"qwertyuiopasdfghjklzxcvbnmQWERTYUIOPASDFGHJKLZXCVBNM`-=[]:\;',./!@#$%^&*()\")");
        activeSheet.getRange(topCellLocation[0] + i, topCellLocation[1] + 1).setValue("=IMAGE(\"" + response.sprites.front_shiny + "\", 2)");
    }
}