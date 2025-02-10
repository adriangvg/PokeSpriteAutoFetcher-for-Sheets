# PokeSpriteAutoFetcher for Sheets

This is a fork from [Aeiiry's googleSheetsPokemonUtilities](https://github.com/Aeiiry/googleSheetsPokemonUtilities/tree/main). I've been using a modified version of it for a while and decided to update it because it may be of use for someone else (I also wanted to gain some practice with GitHub).

The main difference of my version is that I tried to simplify it and add some options focusing on my actual needs for trade sheets. 

What Aeiiry's version has:
* Button to insert Pokémon sprites (cell to the right)
* Button to insert Pokémon types (cell to the right)
* Button to insert Pokémon weakness/resists (2 cells to the right)

What my version has:
* Button to insert Pokémon sprites (cell to the right) and Pokémon number (cell to the left)
* Button to insert shiny Pokémon sprites (cell to the right) and Pokémon number (cell to the left)

His version also inserts a new column to the right if there's no space. This gave me problems in some situations so my version will override whatever is on the left or right cells.

## The Script

It helps Google Sheets communicate with [PokéAPI](https://pokeapi.co/) to (semi)automatically get Pokémon information (in this case, sprites, shiny sprites and national dex number).

### Why use this?

Other sheets will have a database in a separate sheet, where you have a Pokédex with Pokémon data that is later called in the sheets you use via VLOOKUP. This would need to be updated manually when new Pokémon are added in new games of the franchise. 

Since this script connects to an external database (PokéAPI), this is unnecessary: you just to type the Pokémon's name and click the corresponding button to get the sprite.

Probably the best thing is that it doesn't ask you to change your spreadsheet. It can be added to the one you're already using (unless you don't like the style of the sprites...).

### How to use

Write Pokémon names in cells, select those cells and then click the button of your choice to get the sprites.

Names need to be written with hyphens instead of spaces. Ex: Sandy-Shocks, Deoxys-attack, Wormadam-sandy, etc. 

If you can't seem to get it to recognize a pokemon, head to [PokéAPI](https://pokeapi.co/) and enter "pokemon-species/" followed by the name of the pokemon's species (e.g wormadam), then check under "varieties" for a list of pokemon names the api will recognise.

### Installation

This installation guide is the original one by Aeiiry.

1. If you haven't already, make a new google sheet [here](https://sheets.google.com/)
2. In the toolbar in your sheet, click "Extensions" and then click "Apps Script" - This will open a new window
3. Copy the contents of the [code.gs file](https://github.com/Aeiiry/googleSheetsPokemonUtilities/blob/main/code.gs) to code.gs within the apps script page, making sure to copy it below any existing text
4. Press control+S or click the save icon, go back to your google sheet
5. Wait for a little bit and refresh the page, a button should show up near the top of the page labelled "Pokemon". If it doesn't show up either refresh again and/or wait for ~20 seconds
6. Next we need to grant the script access to your google sheet. Click "Pokemon" in the toolbar and then click any of the buttons that pop up
7. A window will pop up reading "Authorization Required", click "Continue". A new window will pop up
8. Select your google account, click "Advanced", click "Go to [your project name] (unsafe)", then click "Allow"
9. Now the scripts will work!




