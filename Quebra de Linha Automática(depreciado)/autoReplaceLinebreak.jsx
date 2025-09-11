//O objetivo desse script é incluir a funcionalidade de substituir "\n" por uma quebra de linha real "^n", mas possui funcionalidade completa do FindChangeByList para que o usuário possa realizar emendas adicionais se desejar.

//Edições futuras devem ser feitas no arquivo incluso na pasta FindChangeSupport/FindChangeList.txt, que é lido pelo script principal, para serem feitas em texto ou GREP, seguindo os exemplos abaixo:

//text	{findWhat:"<string/RegEx>"}	{changeTo:"<string/RegEx/appliedCharacterStyle>"}	{includeFootnotes:<boolean>, includeMasterPages:<boolean>, includeHiddenLayers:<boolean>, wholeWord:<boolean>}

//Onde findWhat é a propriedade ou caracter a ser identificado em string de texto, changeTo é a propriedade ou caracter a ser alterado em string de texto, e as propriedades seguintes são opções de busca do InDesign.

//Exemplos:
//text	{findWhat:"--"}	{changeTo:"^_"}	{includeFootnotes:true, includeMasterPages:true, includeHiddenLayers:true, wholeWord:false}	Substituição de dois traços por um travessão, em formato string.

//text	{findWhat:"ˆ\d{4}$"} {appliedCharacterStyle:"Auxiliar - Ano"} {includeFootnotes:true, includeMasterPages:true, includeHiddenLayers:true, wholeWord:false}	Aplicação de estilo de caractere em anos com 4 dígitos, em formato GREP, onde \d representa qualquer dígito e {4} representa exatamente 4 ocorrências do dígito, e $ representa o fim da linha, substituindo o estilo da string pelo estilo de caracter Auxiliar para apenas o ano.

//Todas as linhas de referência devem ser organizadas abaixo do arquivo FindChangeList.txt.
//Em caso de mais dúvidas, consultar a documentação de RegEx e ExtendedScript.

main();
function main(){
    var myObject;
    //Make certain that user interaction (display of dialogs, etc.) is turned on.
    app.scriptPreferences.userInteractionLevel = UserInteractionLevels.interactWithAll;
    if(app.documents.length > 0){
        if(app.selection.length > 0){
            switch(app.selection[0].constructor.name){
                case "InsertionPoint":
                case "Character":
                case "Word":
                case "TextStyleRange":
                case "Line":
                case "Paragraph":
                case "TextColumn":
                case "Text":
                case "Cell":
                case "Column":
                case "Row":
                case "Table":
                    myDisplayDialog();
                    break;
                default:
                    //Something was selected, but it wasn't a text object, so search the document.
                    myFindChangeByList(app.documents.item(0));
                    replaceBackslashN(app.documents.item(0)); // <-- Add this line
            }
        }
        else{
            //Nothing was selected, so simply search the document.
            myFindChangeByList(app.documents.item(0));
            replaceBackslashN(app.documents.item(0)); // <-- Add this line
        }
    }
    else{
        alert("No documents are open. Please open a document and try again.");
    }
}
function myDisplayDialog(){
	var myObject;
	var myDialog = app.dialogs.add({name:"FindChangeByList"});
	with(myDialog.dialogColumns.add()){
		with(dialogRows.add()){
			with(dialogColumns.add()){
				staticTexts.add({staticLabel:"Search Range:"});
			}
			var myRangeButtons = radiobuttonGroups.add();
			with(myRangeButtons){
				radiobuttonControls.add({staticLabel:"Document", checkedState:true});
				radiobuttonControls.add({staticLabel:"Selected Story"});
				if(app.selection[0].contents != ""){
					radiobuttonControls.add({staticLabel:"Selection", checkedState:true});
				}
			}			
		}
	}
	var myResult = myDialog.show();
	if(myResult == true){
		switch(myRangeButtons.selectedButton){
			case 0:
				myObject = app.documents.item(0);
				break;
			case 1:
				myObject = app.selection[0].parentStory;
				break;
			case 2:
				myObject = app.selection[0];
				break;
		}
		myDialog.destroy();
		myFindChangeByList(myObject);
	}
	else{
		myDialog.destroy();
	}
}
function myFindChangeByList(myObject){
	var myScriptFileName, myFindChangeFile, myFindChangeFileName, myScriptFile, myResult;
	var myFindChangeArray, myFindPreferences, myChangePreferences, myFindLimit, myStory;
	var myStartCharacter, myEndCharacter;
	var myFindChangeFile = myFindFile("/FindChangeSupport/FindChangeList.txt")
	if(myFindChangeFile != null){
		myFindChangeFile = File(myFindChangeFile);
		var myResult = myFindChangeFile.open("r", undefined, undefined);
		if(myResult == true){
			//Loop through the find/change operations.
			do{
				myLine = myFindChangeFile.readln();
				//Ignore comment lines and blank lines.
				if((myLine.substring(0,4)=="text")||(myLine.substring(0,4)=="grep")||(myLine.substring(0,5)=="glyph")){
					myFindChangeArray = myLine.split("\t");
					//Busca da string de texto, GREP ou glifo.
					myFindType = myFindChangeArray[0];
					//Linha de preferências de busca.
					myFindPreferences = myFindChangeArray[1];
					//String de preferências de alteração.
					myChangePreferences = myFindChangeArray[2];
					//Alcance da busca.
					myFindChangeOptions = myFindChangeArray[3];
					switch(myFindType){
						case "text":
							myFindText(myObject, myFindPreferences, myChangePreferences, myFindChangeOptions);
							break;
						case "grep":
							myFindGrep(myObject, myFindPreferences, myChangePreferences, myFindChangeOptions);
							break;
						case "glyph":
							myFindGlyph(myObject, myFindPreferences, myChangePreferences, myFindChangeOptions);
							break;
					}
				}
			} while(myFindChangeFile.eof == false);
			myFindChangeFile.close();
		}
	}
}
function myFindText(myObject, myFindPreferences, myChangePreferences, myFindChangeOptions){
	//Resetar as preferências de busca/troca de texto antes de cada busca.
	app.changeTextPreferences = NothingEnum.nothing;
	app.findTextPreferences = NothingEnum.nothing;
	var myString = "app.findTextPreferences.properties = "+ myFindPreferences + ";";
	myString += "app.changeTextPreferences.properties = " + myChangePreferences + ";";
	myString += "app.findChangeTextOptions.properties = " + myFindChangeOptions + ";";
	app.doScript(myString, ScriptLanguage.javascript);
	myFoundItems = myObject.changeText();
	//Reset the find/change preferences after each search.
	app.changeTextPreferences = NothingEnum.nothing;
	app.findTextPreferences = NothingEnum.nothing;
}
function myFindGrep(myObject, myFindPreferences, myChangePreferences, myFindChangeOptions){
	//Resetar as preferências de busca/troca de GREP antes de cada busca.
	app.changeGrepPreferences = NothingEnum.nothing;
	app.findGrepPreferences = NothingEnum.nothing;
	var myString = "app.findGrepPreferences.properties = "+ myFindPreferences + ";";
	myString += "app.changeGrepPreferences.properties = " + myChangePreferences + ";";
	myString += "app.findChangeGrepOptions.properties = " + myFindChangeOptions + ";";
	app.doScript(myString, ScriptLanguage.javascript);
	var myFoundItems = myObject.changeGrep();
	app.changeGrepPreferences = NothingEnum.nothing;
	app.findGrepPreferences = NothingEnum.nothing;
}
function myFindGlyph(myObject, myFindPreferences, myChangePreferences, myFindChangeOptions){
	app.changeGlyphPreferences = NothingEnum.nothing;
	app.findGlyphPreferences = NothingEnum.nothing;
	var myString = "app.findGlyphPreferences.properties = "+ myFindPreferences + ";";
	myString += "app.changeGlyphPreferences.properties = " + myChangePreferences + ";";
	myString += "app.findChangeGlyphOptions.properties = " + myFindChangeOptions + ";";
	app.doScript(myString, ScriptLanguage.javascript);
	var myFoundItems = myObject.changeGlyph();
	app.changeGlyphPreferences = NothingEnum.nothing;
	app.findGlyphPreferences = NothingEnum.nothing;
}
function myFindFile(myFilePath){
	var myScriptFile = myGetScriptPath();
	var myScriptFile = File(myScriptFile);
	var myScriptFolder = myScriptFile.path;
	myFilePath = myScriptFolder + myFilePath;
	if(File(myFilePath).exists == false){
		//Em caso de quebra de busca da pasta correta para o arquivo FindChangeList.txt, abrir uma janela de diálogo para localizar o arquivo manualmente.
		myFilePath = File.openDialog("FindChangeList.txt file não encontrado. Por favor, revincule o arquivo.");
	}
	return myFilePath;
}
function myGetScriptPath(){
	try{
		myFile = app.activeScript;
	}
	catch(myError){
		myFile = myError.fileName;
	}
	return myFile;
}

function replaceBackslashN(myObject) {
    app.findTextPreferences = NothingEnum.nothing;
    app.changeTextPreferences = NothingEnum.nothing;
    app.findTextPreferences.findWhat = "\\n";
    app.changeTextPreferences.changeTo = "^n";
    myObject.changeText();
    app.findTextPreferences = NothingEnum.nothing;
    app.changeTextPreferences = NothingEnum.nothing;
}