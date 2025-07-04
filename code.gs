//More info and example spreadsheet in https://github.com/adriangvg/PokeSpriteAutoFetcher-for-Sheets

//This creates the buttons in the menu. You can then click them to easily get your sprites
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

//Function to fetch Pokémon data, don't change anything here unless you really know what you're doing
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

//Function to fetch Pokémon sprite and number
//Line 63: cell where the number of the Pokémon will be written (one cell left to the active cell)
//Line 64: cell where the sprite of the Pokémon will be written (one cell right to the active cell)
//If you don't need the number, you can just delete line 63 (and 80 for shinies)
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

//Function to fetch Pokémon shiny sprite image
//Same as in the function for non-shinies, lines 80 and 81 are the key ones that affect where data will be written
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