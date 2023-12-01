(function(window, undefined){
	var flagInit = false;
	var fBtnGetAll = false;
	var fClickLabel = false;
	var fClickBtnCur =  false;

    window.Asc.plugin.init = function(text)
    {
		// this.callCommand(function() {
		// 	var oDocument = Api.GetDocument();
		// 	var oParagraph = Api.CreateParagraph();
		// 	oParagraph.AddText("Hello world!");
		// 	oDocument.InsertContent([oParagraph]);
		// }, true);
		//event "init" for plugin
		// document.getElementById("divS").innerHTML = text.replace(/\n/g,"<br>");

		// document.getElementById("buttonIDPaste").onclick = function() {
		// 	//method for paste text into document
		// 	window.Asc.plugin.executeMethod("PasteText", ["Test paste for document"]);	

		// };		

		// document.getElementById("buttonIDGetAll").onclick = function() {
		// 	//method for get all content controls
		// 	window.Asc.plugin.executeMethod("GetAllContentControls");
		// 	fBtnGetAll = true;					

		// };

		// document.getElementById("buttonIDShowCurrent").onclick = function() {
			
		// 	fClickBtnCur = true;
		// 	//menthod for get current content control (where is the cursor located)
		// 	window.Asc.plugin.executeMethod("GetCurrentContentControl");

		// };
		
		if (!flagInit) {
			flagInit = true;
			//method for get all content controls
			window.Asc.plugin.executeMethod("GetAllContentControls");
			// document.getElementById("buttonIDGetAll").click();
			this.callCommand(() => {
				let oDocument = Api.GetDocument();
				let aContentControls = oDocument.GetAllContentControls();
				let contentControlsArr = [];
				for (let l = 0; l < aContentControls.length; l++) {
					const contentControl = aContentControls[l];
					const contentControlEl = JSON.parse(contentControl.ToJSON())
					let contentControlObjFilter = {
						"Id":contentControl.be.ha.$a,
						"Tag": contentControl.be.ha.TA,
						"Alias": contentControl.be.ha.oE,
						"InternalId": contentControl.be.$a,
					}
					if(contentControlEl.sdtPr.dropDownList || contentControlEl.sdtPr.comboBox){
						//console.log(contentControlEl);
						const contentControlDropDown = contentControl.GetDropdownList();
						const contentControlDropDownItems = contentControlDropDown.GetAllItems();
						let dropdown = [];
						for (let i = 0; i < contentControlDropDownItems.length; i++) {
							const contentControlDropDownItem = contentControlDropDownItems[i];
							dropdown.push({
								text: contentControlDropDownItem.Text,
								index: i,
							});
						}
						contentControlObjFilter["Type"] = "dropdown";
						contentControlObjFilter["dropdown"] = dropdown;
						
					}
					/* contentControl.be.ha.Date */
					else if(contentControlEl.sdtPr.date){
						/* Преобразование даты */
						let fullDate = contentControlEl.sdtPr.date.fullDate;
						let dateObject = new Date(fullDate);
						let day = dateObject.getUTCDate();
						let month = dateObject.getUTCMonth() + 1; // Месяцы в JS начинаются с 0
						let year = dateObject.getUTCFullYear();
						// let formattedDate = (day < 10 ? '0' : '') + day + '-' + (month < 10 ? '0' : '') + month + '-' + year;
						let formattedDate = year + '-' + (month < 10 ? '0' : '') + month + '-' + (day < 10 ? '0' : '') + day;


						let date = {};
						date = contentControlEl.sdtPr.date;
						date.formattedDate = formattedDate;
						contentControlObjFilter["Date"] = date;
						contentControlObjFilter["Type"] = "date";
					}
					else {
						contentControlObjFilter["Type"] = "text"
						let contentControlBlockType = contentControl.GetClassType();
						if(contentControlBlockType == "blockLvlSdt"){
							// console.log(JSON.parse(contentControl.ToJSON(true, true)));
							const text = contentControl.GetContent().GetElement(0).GetText();
							contentControlObjFilter["Content"] = text;
							contentControlObjFilter["ContentControlType"] = "blockLvlSdt";
						}
						else {
							let contentControlObj  = JSON.parse(contentControl.ToJSON(false));
							let contentString = "";

							for (let o = 0; o < contentControlObj.content.length; o++) {
								const contentList = contentControlObj.content[o];
								if(contentList.type == "hyperlink"){
									for (let t = 0; t < contentList.content.length; t++) {
										const mailContent = contentList.content[t];
										for (let a = 0; a < mailContent.content.length; a++) {
											contentString += mailContent.content[a];
										}
									}
								}
								else {
									for (let i = 0; i < contentList.content.length; i++) {
										let contentSecondLvl = contentList.content[i];
	
										contentString = contentString + contentSecondLvl;
									}
								}

							}
							
							contentControlObjFilter["Content"] = contentString;
						}
					}

					contentControlsArr.push(contentControlObjFilter)
				}
				let result = [];
				let duplicates = {};

				contentControlsArr.forEach(item => {
					const { Tag, Alias, InternalId, Id } = item;
					const key = `${Tag}_${Alias}`;
				
					if (!duplicates[key]) {
						// Если это первый раз встречаем элемент с таким ключом
						duplicates[key] = {
							count: 1,
							index: result.length
						};
				
						result.push({ ...item, count: 1 });
					} else {
						// Если элемент с таким ключом уже был встречен
						const { count, index } = duplicates[key];
						result[index][`InternalId_${count}`] = InternalId;
						result[index][`Id_${count}`] = Id;
						result[index].count++;
					}
				});
				contentControlsArr = result;
				
				const groupedContentControls = contentControlsArr.reduce((acc, current) => {
					const existingTagObject = acc.find((item) => item.Tag === current.Tag);
					if (existingTagObject) {
					  // Если объект с таким тегом уже есть, добавляем текущий элемент в его свойство items
					  existingTagObject.items.push({ ...current});
					} else {
					  // Если объекта с таким тегом нет, создаем новый объект и добавляем его в аккумулятор
					  acc.push({
						Tag: current.Tag,
						items: [{...current}],
					  });
					}
				  
					return acc;
				}, []);
				localStorage.setItem("contentControls", JSON.stringify(groupedContentControls));
			})
		
		}
	};

	addLabel = (returnValue, element) => {
		if(returnValue.Type == "dropdown"){
			let checkboxesHtmlblock = '';
			for (let l = 0; l < returnValue.dropdown.length; l++) {
				const checkbox = returnValue.dropdown[l];
				checkboxesHtmlblock = checkboxesHtmlblock + `<br><label style="color: black;"><input type="radio" data-id="${returnValue.InternalId}" value="${checkbox.text}"><div style="display: inline; padding-left: 5px">${checkbox.text}</div></label>`;
			}
			$(element).append($('<div>', {
					class: 'form-group'
				})
				.append($('<label>', {
					text: returnValue.Alias,
				})
				.append(checkboxesHtmlblock)
			))
			let checkboxInputs = document.querySelectorAll(`input[data-id="${returnValue.InternalId}"]`);
			checkboxInputs[0].checked = true;
			for (let o = 0; o < checkboxInputs.length; o++) {
				const input = checkboxInputs[o];
				input.addEventListener("input", (event) => {
					event.preventDefault();
					window.Asc.plugin.executeMethod("SelectContentControl",[returnValue.Id]);
					checkboxInputs.forEach(el => el.checked = false);
					input.checked = true;
					Asc.scope.tag = returnValue.Tag 
					let ids = [];
					for (let t = 0; t < returnValue.count; t++) {
						if(t == 0){
							ids.push(returnValue["Id"]);
						}
						else {
							ids.push(returnValue["Id_"+t]);
						}
					}
					Asc.scope.ids = ids;
					Asc.scope.index = o;
					window.Asc.plugin.callCommand (function () {
						// Api.GetDocument().SearchAndReplace({"searchString": Asc.scope.previousValue, "replaceString": Asc.scope.text});
						let oDocument = Api.GetDocument();
						let allContentControlsByTag = oDocument.GetContentControlsByTag(Asc.scope.tag);
						for (let l = 0; l < allContentControlsByTag.length; l++) {
							const contentControlByTagEl = allContentControlsByTag[l];
							const contentControlByTag = JSON.parse(allContentControlsByTag[l].ToJSON());
							Asc.scope.ids.forEach(id => {
								if(contentControlByTag.sdtPr.id == id){
									let oContentControlList = contentControlByTagEl.GetDropdownList();
									let oItem = oContentControlList.GetItem(Asc.scope.index);
									oItem.Select();
								}
							})
						}
						
				   });
				})
			}
		}
		else if(returnValue.Type == "date"){
			$(element).append(
				$('<div>',{
					class : 'form-group',
				})
				.append($('<label>',{
					text: returnValue.Alias,
				}))
				.append($('<input>',{
					id : returnValue.InternalId,
					class: "form-control micros-convertor-input",
					type: "date",
					value: returnValue.Date.formattedDate,
					min: "1900-01-01",
					on : {
						focus: function(){
							$('.focus').removeClass('focus');
							$(this).addClass('focus');
							// window.Asc.plugin.executeMethod("MoveCursorToContentControl",[this.id, true]);
							window.Asc.plugin.executeMethod("SelectContentControl",[this.id]);
							if(document.activeElement == this ){
							}
						},
						input: function(){
							let originalDate = this.value;

							// Создаем объект Date на основе строки
							let dateObject = new Date(originalDate);
							
							// Получаем компоненты даты
							let day = dateObject.getUTCDate();
							let month = dateObject.getUTCMonth() + 1; // Месяцы в JS начинаются с 0
							let year = dateObject.getUTCFullYear();
							
							let formattedDate = (day < 10 ? '0' : '') + day + '.' + (month < 10 ? '0' : '') + month + '.' + year;

							if(this.value){
								window.Asc.plugin.executeMethod("PasteText", [formattedDate]);
								window.Asc.plugin.executeMethod("SelectContentControl",[this.id]);
							}
						}
					}
				}))
			);	
		}
		else {
			$(element).append(
				$('<div>',{
					// id : returnValue.InternalId,
					class : 'form-group',
					// text : returnValue.InternalId + "	" + (returnValue.Id || 'null'),
					on : {
						click: function(){
							fClickLabel = true;
							// if (element === "#divG") {
							// 	//method for select content control by id
							// 	window.Asc.plugin.executeMethod("SelectContentControl",[this.id]);
							// } else {
							// 	//method for move cursor to content control with specified id
							// 	window.Asc.plugin.executeMethod("MoveCursorToContentControl",[this.id, true]);
							// }
						},
						// mouseover: function(){
						// 	$(this).addClass('label-hovered');
						// },
						// mouseout: function(){
						// 	$(this).removeClass('label-hovered');
						// }
					}
				})
				.append($('<label>',{
					text: returnValue.Alias + (returnValue.count !== 1 ? ` (${returnValue.count})`: ''),
					on: {
						click: function (event) {
							if(returnValue.count !== 1){
							}
							
						}
					}
				}))
				.append($('<input>',{
					id : returnValue.InternalId,
					class: "form-control micros-convertor-input",
					value: returnValue.Content,
					on : {
						focus: function(){
							// $('.focus').removeClass('focus');
							// $(this).addClass('focus');
							// if (element === "#divG") {
							// 	//method for select content control by id
							// 	window.Asc.plugin.executeMethod("SelectContentControl",[this.id]);
							// } else {
							// 	//method for move cursor to content control with specified id
							// 	window.Asc.plugin.executeMethod("MoveCursorToContentControl",[this.id, true]);
							// }
							window.Asc.plugin.executeMethod("SelectContentControl",[this.id]);
							// window.Asc.plugin.executeMethod("SelectContentControl",[this.id]);
							// this.focus()
								// window.Asc.plugin.executeMethod("MoveCursorToContentControl",[this.id, true]);
						},
						input: async function(){
							let previousValue = returnValue.Content;
							returnValue.Content = this.value
							if(returnValue.ContentControlType == "blockLvlSdt"){
								
								// const textArr = this.value.split(" ")
								// console.log(textArr);
								// Asc.scope.textArray = textArr;
								// window.Asc.plugin.executeMethod ("ReplaceTextSmart", [Asc.scope.textArray, String.fromCharCode(9), String.fromCharCode(13)], function (isDone) {
								// 	console.log(isDone);
								// 	if (!isDone)
								// 		window.Asc.plugin.callCommand (function () {
								// 			Api.ReplaceTextSmart (Asc.scope.textArray);
								// 		});
								// });
								// let ids = [];
								// for (let t = 0; t < returnValue.count; t++) {
								// 	if(t == 0){
								// 		ids.push(returnValue["Id"]);
								// 	}
								// 	else {
								// 		ids.push(returnValue["Id_"+t]);
								// 	}
								// }
								// Asc.scope.ids = ids;
								// Asc.scope.text = this.value;
								// Asc.scope.tag = returnValue.Tag 
								// window.Asc.plugin.callCommand (function () {
								// 	let oDocument = Api.GetDocument();
								// 	let allContentControlsByTag = oDocument.GetContentControlsByTag(Asc.scope.tag);
								// 	for (let l = 0; l < allContentControlsByTag.length; l++) {
								// 		const contentControlByTagEl = allContentControlsByTag[l];
								// 		const contentControlByTag = JSON.parse(allContentControlsByTag[l].ToJSON());
								// 		Asc.scope.ids.forEach(id => {
								// 			if(contentControlByTag.sdtPr.id == id){
								// 				console.log(contentControlByTag);
								// 				if(contentControlByTag.type == "blockLvlSdt"){
								// 					console.log(contentControlByTag);
								// 					console.log(contentControlByTagEl.GetContent().GetElement(0).GetText());
								// 				}
								// 				else {

								// 				}
								// 				// console.log(contentControlByTagEl.GetContent().GetElement(0).GetText());
								// 			}
								// 		})
								// 	}
								//  	// var oParagraph = Api.CreateParagraph();
								// 	// oParagraph.AddText(Asc.scope.text);
								// 	// console.log(Asc.scope.text);
								// 	// Api.GetDocument().InsertContent([oParagraph], true, {KeepTextOnly: true});
								// });
								window.Asc.plugin.executeMethod("PasteText", [this.value]);
								window.Asc.plugin.executeMethod("SelectContentControl",[this.id]);
							}
							else {
								// let scriptForObject = `
								// 	let oParagraph = Api.CreateParagraph();
								// 	oParagraph.AddText("${this.value}");
								// 	Api.GetDocument().InsertContent([oParagraph], true, {KeepTextOnly: true});
								// `.replaceAll("\n", "").trim();
								// let arrDocuments = [{
								// 	"Props": {
								// 		"InternalId": returnValue.InternalId,
								// 	},
								// 	"Script": 'var oParagraph = Api.CreateParagraph(); oParagraph.AddText(' + '"' + this.value +'"' +');Api.GetDocument().InsertContent([oParagraph], true, {KeepTextOnly: true});'
								// }]
								// window.Asc.plugin.executeMethod("InsertAndReplaceContentControls", [arrDocuments]);
							// 	Asc.scope.previousValue = previousValue;
							// 	Asc.scope.text = this.value;
							// 	Asc.scope.tag = returnValue.Tag 
							// 	let ids = [];
							// 	for (let t = 0; t < returnValue.count; t++) {
							// 		if(t == 0){
							// 			ids.push(returnValue["Id"]);
							// 		}
							// 		else {
							// 			ids.push(returnValue["Id_"+t]);
							// 		}
							// 	}
							// 	Asc.scope.ids = ids;
							// 	window.Asc.plugin.callCommand (function () {
							// 		Api.GetDocument().SearchAndReplace({"searchString": Asc.scope.previousValue, "replaceString": Asc.scope.text});
							// 		// let oDocument = Api.GetDocument();
							// 		// let allContentControlsByTag = oDocument.GetContentControlsByTag(Asc.scope.tag);
							// 		// for (let l = 0; l < allContentControlsByTag.length; l++) {
							// 		// 	const contentControlByTagEl = allContentControlsByTag[l];
							// 		// 	const contentControlByTag = JSON.parse(allContentControlsByTag[l].ToJSON());
							// 		// 	Asc.scope.ids.forEach(id => {
							// 		// 		if(contentControlByTag.sdtPr.id == id){
							// 		// 		}
							// 		// 	})
							// 		// }
							// 		// var oParagraph = Api.CreateParagraph();
							// 	   	// oParagraph.AddText(Asc.scope.text);
							// 	   	// Api.GetDocument().InsertContent([oParagraph], true, {KeepTextOnly: true});
							//    });
							window.Asc.plugin.executeMethod("PasteText", [this.value]);
							window.Asc.plugin.executeMethod("SelectContentControl",[this.id]);
							}
						}
					}
				}))
			);	
		}
	};
		
	// function  getCurrentControl(id) {
	// 	return new Promise(async (resolve) => {
	// 		let controlContentObject = {};
	
	// 		await window.Asc.plugin.executeMethod("SelectContentControl", [id]);
			
	// 		await window.Asc.plugin.executeMethod("GetCurrentContentControlPr", ["ole"], (obj) => {
	// 			controlContentObject = obj;
	// 			resolve(controlContentObject);
	// 		});
	// 	});
	// }

    window.Asc.plugin.button = function()
    {
		this.executeCommand("close", "");
    };

	window.Asc.plugin.onMethodReturn = async function(returnValue)
	{
		//evend return for completed methods
		var _plugin = window.Asc.plugin;
		if (_plugin.info.methodName == "GetAllContentControls")
		{
			if (fBtnGetAll) {
				// document.getElementById("divP").innerHTML = "";
				// fBtnGetAll = false;
				// for (var i = 0; i < returnValue.length; i++) {	
				// 	addLabel(returnValue[i], "#divP");
				// }
			} else {
				// document.getElementById("divG").innerHTML = "";
				/* Группировка по значению свойства Tag */
				// const groupedArray = returnValue.reduce((acc, current) => {
				// 	const existingTagObject = acc.find((item) => item.Tag === current.Tag);
				// 	if (existingTagObject) {
				// 	  // Если объект с таким тегом уже есть, добавляем текущий элемент в его свойство items
				// 	  existingTagObject.items.push({ ...current});
				// 	} else {
				// 	  // Если объекта с таким тегом нет, создаем новый объект и добавляем его в аккумулятор
				// 	  acc.push({
				// 		Tag: current.Tag,
				// 		items: [{...current}],
				// 	  });
				// 	}
				  
				// 	return acc;
				// }, []);
				// for (let l = 0; l < groupedArray.length; l++) {
				// 	const controlContentItems = groupedArray[l];
				// 	// console.log(controlContentItems);
				// 	$("#divG").append(
				// 		$("<div>", {
				// 			// id: "control-content-"+l,
				// 			class: "card border-0",
				// 		}).append(`<div class="card-header c-header"><a class="card-link">${controlContentItems.Tag}</a></div>`)
				// 		.append(`<div class="card-body" id="control-content-${l}"></div>`)
				// 	)
				// 	for (let o = 0; o < controlContentItems.items.length; o++) {
				// 		const controlContentElement = controlContentItems.items[o];
				// 		await getCurrentControl(controlContentElement.InternalId)
				// 		.then(async (data) => {
				// 			// console.log(data);
				// 			await addLabel(data, "#control-content-"+l);
				// 		})
				// 		.catch((error) => {
				// 			console.error(error);
				// 		})
				// 	}
				// }
				window.Asc.plugin.executeMethod ("MoveCursorToStart", [true]);
				// for (var i = 0; i < returnValue.length; i++) {	
				// 	let element = {};
				// 	getCurrentControl(returnValue[i].InternalId)
				// 	.then((data) => {
				// 		element = data;
				// 		addLabel(element, "#divG");
				// 	})
				// 	.catch((error) => {
				// 		console.error(error);
				// 	})
				// 	// addLabel(returnValue[i], "#divG");
				// }
			}
			

		} 

		// else if (_plugin.info.methodName == "GetCurrentContentControl") {
		// 	if (fClickBtnCur) {
		// 		//method for select content control by id
		// 		window.Asc.plugin.executeMethod("SelectContentControl",[returnValue]);
		// 		fClickBtnCur = false;
		// 	} else if (!($('.label-selected').length && $('.label-selected')[0].id === returnValue) && returnValue) {
		// 		if (document.getElementById(returnValue))
		// 		{
		// 			$('.label-selected').removeClass('label-selected');
		// 			$('#divG #' + returnValue).addClass('label-selected');
		// 			// $('#divP #' + returnValue).addClass('label-selected');


		// 		} else {
		// 			$('.label-selected').removeClass('label-selected');
		// 			addLabel({InternalId: returnValue},"#divG");
		// 			$('#' + returnValue).addClass('label-selected');
		// 		}
		// 	} else if (!returnValue) {
		// 		$('.label-selected').removeClass('label-selected');
		// 	}
		// }
	};

	window.Asc.plugin.onCommandCallback = function() {
		let plugin = window.Asc.plugin;
		const contentControlsJson = localStorage.getItem('contentControls');
		if(contentControlsJson){
			const groupedContentControls = JSON.parse(contentControlsJson);
			localStorage.removeItem("contentControls")
			for (let l = 0; l < groupedContentControls.length; l++) {
				const groupContentControls = groupedContentControls[l];
				$("#divG").append(
					$("<div>", {
						// id: "control-content-"+l,
						class: "card border-0",
					})
					.append(`<div class="card-header c-header"><a class="card-link">${groupContentControls.Tag}</a></div>`)
					.append(`<div class="card-body" id="control-content-${l}"></div>`)
				)
				for (let o = 0; o < groupContentControls.items.length; o++) {
					const controlContentElement = groupContentControls.items[o];
					addLabel(controlContentElement, "#control-content-"+l);
				}
			}
		}
	};

	/* Ивенты */
	window.Asc.plugin.event_onClick = function(isSelectionUse) {
		window.Asc.plugin.executeMethod("GetCurrentContentControlPr", [], function(obj) {
			window.Asc.plugin.currentContentControl = obj;
			var controlTag = obj ? obj.Tag : "";
			if (isSelectionUse)
				controlTag = "";
			console.log(obj);
		});
	};
	
	window.Asc.plugin.event_onFocusContentControl = async function(control){
		let input = document.getElementById(control.InternalId)
		if(input){
			if(document.activeElement !== input){
				setTimeout(() => {
					// Рассчитываем высоту половины видимой области
					let halfViewportHeight = window.innerHeight / 2;
					// Рассчитываем координаты, чтобы элемент был по середине видимой области
					let elementTop = input.getBoundingClientRect().top;
					let scrollToMiddle = elementTop - halfViewportHeight;

					input.scrollIntoView({
						behavior: 'smooth', // добавляет плавность прокрутки
						top: scrollToMiddle,
					});
					input.focus()
				},100)
				
			}
		}
		// let input = document.getElementById(control.InternalId)
		// // Рассчитываем высоту половины видимой области
		// let halfViewportHeight = window.innerHeight / 2;
				
		// // Рассчитываем координаты, чтобы элемент был по середине видимой области
		// let elementTop = input.getBoundingClientRect().top;
		// let scrollToMiddle = elementTop - halfViewportHeight;
		// if(input){
		// 	input.scrollIntoView({
		// 		behavior: 'smooth', // добавляет плавность прокрутки
		// 		top: scrollToMiddle,
		// 	});
		// 	input.focus()
		// }
	};

	window.Asc.plugin.event_onChangeContentControl = (control) => {
		window.Asc.plugin.executeMethod ("GetCurrentContentControlPr", ["ole"], (obj) => {
			let input = document.getElementById(control.InternalId)
			if(input){
				// if(input.value !== obj.content){
					if(input.type == "date"){
						let dateString =  obj.content;

						// Разбиваем строку на компоненты
						let dateComponents = dateString.split('.');

						// Создаем объект Date с использованием компонентов
						let dateObject = new Date(dateComponents[2], dateComponents[1] - 1, dateComponents[0]);
						
						// Получаем компоненты даты
						let year = dateObject.getUTCFullYear();
						let month = dateObject.getUTCMonth() + 1; // Месяцы в JS начинаются с 0, поэтому прибавляем 1
						let day = dateObject.getDate();

						// Форматируем компоненты даты
						let formattedDate = year + '-' + (month < 10 ? '0' : '') + month + '-' + (day < 10 ? '0' : '') + day;

						// Выводим отформатированную дату
						console.log(formattedDate);
						console.log(input.value);
						// input.value = formattedDate;
						
					}
					else {
						input.value = obj.content;
					}
				// }
			}
		})
	}

	// window.Asc.plugin.event_onTargetPositionChanged = function()
	// {
	// 	//event change cursor position
	// 	//all events are specified in the config file in the "events" field
	// 	if (!fClickLabel) {
	// 		//menthod for get current content control (where is the cursor located)
	// 		window.Asc.plugin.executeMethod("GetCurrentContentControl");
	// 	}
	// 	fClickLabel = false;
	// };

	
})(window, undefined);

