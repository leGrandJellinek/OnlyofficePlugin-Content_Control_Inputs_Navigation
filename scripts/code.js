(function(window, undefined){
	var flagInit = false;
	var fBtnGetAll = false;
	var fClickLabel = false;
	var fClickBtnCur =  false;

    window.Asc.plugin.init = function(text){
		if (!flagInit) {
			flagInit = true;
			
			this.callCommand(() => {
				let oDocument = Api.GetDocument();
				// Получаю все элементы управления контентом
				let aContentControls = oDocument.GetAllContentControls();
				let contentControlsArr = [];
				for (let l = 0; l < aContentControls.length; l++) {
					const contentControl = aContentControls[l];
					// Элемент управления контентом в формате JSON
					const contentControlEl = JSON.parse(contentControl.ToJSON())
					// Создаю свой обьект для дальнейшей записи в массив
					let contentControlObjFilter = {
						"Id":contentControl.be.ha.$a,
						"Tag": contentControl.be.ha.TA,
						"Alias": contentControl.be.ha.oE,
						"InternalId": contentControl.be.$a,
					}
					if(contentControlEl.sdtPr.dropDownList || contentControlEl.sdtPr.comboBox){
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
					else if(contentControlEl.sdtPr.date){
						/* Преобразование даты */
						let fullDate = contentControlEl.sdtPr.date.fullDate;
						let dateObject = new Date(fullDate);
						let day = dateObject.getUTCDate();
						let month = dateObject.getUTCMonth() + 1; // Месяцы в JS начинаются с 0
						let year = dateObject.getUTCFullYear();
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
					// Записываю в массив элемент управлением контента
					contentControlsArr.push(contentControlObjFilter)
				}

				/* 
				 * Здесь я перебираю массив для удаления одинаковых элементов управления контента (Одинаковый Tag и Alias)
				 * Я добавляю id и internalId в первый дубликат а остальные удаляю
				 */
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
				
				/* 
				 * Группирую обьекты по значению Tag
				 */
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

				// Записываю финальный массив элементов управления контентом в LocalStorage
				localStorage.setItem("contentControls", JSON.stringify(groupedContentControls));
			})
		
		}
	};

	// Функция  addFields нужна для отрисовки полей внутри плагина
	const addFields = (returnValue, element) => {
		// Для элементов управления типа dropdown & combobox
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
		// Для элементов управления типа "дата"
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
		// Для элементов управления типа "текст"
		else {
			$(element).append(
				$('<div>',{
					class : 'form-group',
					on : {
						click: function(){
							fClickLabel = true;
						},
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
							window.Asc.plugin.executeMethod("SelectContentControl",[this.id]);
						},
						input: async function(){
							window.Asc.plugin.executeMethod("PasteText", [this.value]);
							window.Asc.plugin.executeMethod("SelectContentControl",[this.id]);
						}
					}
				}))
			);	
		}
	};

    window.Asc.plugin.button = function()
    {
		this.executeCommand("close", "");
    };

	/* 
	 * onCommandCallback срабатывает при выполнении callCommand();
	 */
	window.Asc.plugin.onCommandCallback = function() {
		/* В LocalStorage под ключем  contentControls хранятся все элементы управления
		 * Записываются они туда при инициализации моего плагина, дело в том что данные из 
		 * ONLYOFFICE Document Builder передать в plugin нельзя и обходным решением является LocalStorage
		 */
		const contentControlsJson = localStorage.getItem('contentControls');
		if(contentControlsJson){
			const groupedContentControls = JSON.parse(contentControlsJson);
			localStorage.removeItem("contentControls")
			for (let l = 0; l < groupedContentControls.length; l++) {
				const groupContentControls = groupedContentControls[l];
				$("#divG").append(
					$("<div>", {
						class: "card border-0",
					})
					.append(`<div class="card-header c-header"><a class="card-link">${groupContentControls.Tag}</a></div>`)
					.append(`<div class="card-body" id="control-content-${l}"></div>`)
				)
				for (let o = 0; o < groupContentControls.items.length; o++) {
					const controlContentElement = groupContentControls.items[o];
					addFields(controlContentElement, "#control-content-"+l);
				}
			}
		}
	};

	/* Ивенты */

	//Ивент на фокус элемента управления
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
	};
	//Ивент работает при изменения контента внутри элемента управления
	window.Asc.plugin.event_onChangeContentControl = (control) => {
		window.Asc.plugin.executeMethod ("GetCurrentContentControlPr", ["ole"], (obj) => {
			// Инпут, который находится слева в плагине
			let input = document.getElementById(control.InternalId)
			if(input){
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
					
				}
				// Если элемент имеет тип текст
				else {
					input.value = obj.content;
				}
			}
		})
	}
	
})(window, undefined);

