import {IInputs, IOutputs} from "./generated/ManifestTypes";
import DataSetInterfaces = ComponentFramework.PropertyHelper.DataSetApi;
type DataSet = ComponentFramework.PropertyTypes.DataSet;
import * as $ from 'jquery';

// Define const here
const RowRecordId:string = "recordId";

// Style name of Load More Button
const DataSetControl_LoadMoreButton_Hidden_Style = "DataSetControl_LoadMoreButton_Hidden_Style";

export class AFGProjectsEditGrid implements ComponentFramework.StandardControl<IInputs, IOutputs> {

	// Cached context object for the latest updateView
	private contextObj: ComponentFramework.Context<IInputs>;
		
	// Div element created as part of this control's main container
	private mainContainer: HTMLDivElement;

	// Table element created as part of this control's table
	private gridContainer: HTMLDivElement;

	// Button element created as part of this control
	private loadPageButton: HTMLButtonElement;

	private isActive:boolean;
	private isSubmitted:boolean;
	private lockOnInactive:boolean;
	private isEditing:boolean;
	private intervalHandler:any;

	/**
	 * Empty constructor.
	 */
	constructor()
	{
	}

	/**
	 * Used to initialize the control instance. Controls can kick off remote server calls and other initialization actions here.
	 * Data-set values are not initialized here, use updateView.
	 * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to property names defined in the manifest, as well as utility functions.
	 * @param notifyOutputChanged A callback method to alert the framework that the control has new outputs ready to be retrieved asynchronously.
	 * @param state A piece of data that persists in one session for a single user. Can be set at any point in a controls life cycle by calling 'setControlState' in the Mode interface.
	 * @param container If a control is marked control-type='standard', it will receive an empty div element within which it can render its content.
	 */
	public init(context: ComponentFramework.Context<IInputs>, notifyOutputChanged: () => void, state: ComponentFramework.Dictionary, container:HTMLDivElement)
	{
		// Need to track container resize so that control could get the available width. The available height won't be provided even this is true
		context.mode.trackContainerResize(false);

		// Create main table container div. 
		this.mainContainer = document.createElement("div");

		// Create data table container div. 
		this.gridContainer = document.createElement("div");
		this.gridContainer.classList.add("DataSetControl_grid-container");

		// Create data table container div. 
		this.loadPageButton = document.createElement("button");
		this.loadPageButton.setAttribute("type", "button");
		this.loadPageButton.innerText = context.resources.getString("PCF_DataSetControl_LoadMore_ButtonLabel");
		this.loadPageButton.classList.add(DataSetControl_LoadMoreButton_Hidden_Style);
		this.loadPageButton.classList.add("DataSetControl_LoadMoreButton_Style");
		this.loadPageButton.addEventListener("click", this.onLoadMoreButtonClick.bind(this));

		// Adding the main table and loadNextPage button created to the container DIV.
		this.mainContainer.appendChild(this.gridContainer);
		this.mainContainer.appendChild(this.loadPageButton);
		this.mainContainer.classList.add("DataSetControl_main-container");
		container.appendChild(this.mainContainer);
	}


	/**
	 * Called when any value in the property bag has changed. This includes field values, data-sets, global values such as container height and width, offline status, control metadata values such as label, visible, etc.
	 * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to names defined in the manifest, as well as utility functions
	 */
	public updateView(context: ComponentFramework.Context<IInputs>): void
	{
		this.contextObj = context;
		this.toggleLoadMoreButtonWhenNeeded(context.parameters.AFGProjectsDataSet);

		if(!context.parameters.AFGProjectsDataSet.loading){
			this.isActive = false;
			this.isSubmitted = false;
			this.lockOnInactive = false;

			// Get sorted columns on View
			let columnsOnView = this.getSortedColumnsOnView(context);
			let stateColumnName:string = "";

			//context.parameters.AFGProjectsDataSet.records[0].getFormattedValue("");

			if (!columnsOnView || columnsOnView.length === 0) {
				return;
			}

			try {
				let lockInactive:string = context.parameters.lockoninactivestate.raw as string;
				if (lockInactive.toLowerCase().indexOf("true") > -1) this.lockOnInactive = true;

				// Determine if the parent record state is inactice
				//for (let column of columnsOnView){
				//	if(column.name.indexOf("statecode") > -1) {
				//		let state:string = context.parameters.AFGProjectsDataSet.records[context.parameters.AFGProjectsDataSet.sortedRecordIds[0]].getFormattedValue(column.name);

				//		if (state.toLowerCase().indexOf("inactive") > -1) this.isActive = false;
						
				//		break;
				//	}
				//}
			}
			catch(e) {
				var s = e;
			}

			var obj = this;

			context.webAPI.retrieveRecord('caps_submission', (this.contextObj.mode as any).contextInfo.entityId).then(
				function (capsSubmission:ComponentFramework.WebApi.Entity) 
				{
					if (capsSubmission != null) {
						if (capsSubmission.statecode != null && capsSubmission.statecode == 0) obj.isActive = true;
						if (capsSubmission.statuscode != null && capsSubmission.statecode == 2) obj.isSubmitted = true
					}

					while(obj.gridContainer.firstChild)
					{
						obj.gridContainer.removeChild(obj.gridContainer.firstChild);
					}

					obj.gridContainer.appendChild(obj.createGridBody(context, columnsOnView, obj.contextObj.parameters.AFGProjectsDataSet));

					obj.bindEvents();
				},
			function (errorResponse: any) 
			{
				// Error handling code here - record failed to be created
				alert("An error occured while running the query:\r\n\r\n" + errorResponse.message);
			});
		}
		// this is needed to ensure the scroll bar appears automatically when the grid resize happens and all the tiles are not visible on the screen.
		//this.mainContainer.style.maxHeight = window.innerHeight - this.gridContainer.offsetTop - 75 + "px";
	}

	/** 
	 * It is called by the framework prior to a control receiving new data. 
	 * @returns an object based on nomenclature defined in manifest, expecting object[s] for property marked as “bound” or “output”
	 */
	public getOutputs(): IOutputs
	{
		return {};
	}

	/** 
	 * Called when the control is to be removed from the DOM tree. Controls should use this call for cleanup.
	 * i.e. cancelling any pending remote calls, removing listeners, etc.
	 */
	public destroy(): void
	{
		// Add code to cleanup control if necessary
	}

	/**
	 * Get sorted columns on view
	 * @param context 
	 * @return sorted columns object on View
	 */
	private getSortedColumnsOnView(context: ComponentFramework.Context<IInputs>): DataSetInterfaces.Column[]
	{
		if (!context.parameters.AFGProjectsDataSet.columns) {
			return [];
		}
		
		let columns =context.parameters.AFGProjectsDataSet.columns
			.filter(function (columnItem:DataSetInterfaces.Column) { 
				// some column are supplementary and their order is not > 0
				return columnItem.order >= 0 }
			);
		
		// Sort those columns so that they will be rendered in order
		columns.sort(function (a:DataSetInterfaces.Column, b: DataSetInterfaces.Column) {
			return a.order - b.order;
		});
		
		return columns;
	}
	
	/**
	 * funtion that creates the body of the grid and embeds the content onto the tiles.
	 * @param columnsOnView columns on the view whose value needs to be shown on the UI
	 * @param gridParam data of the Grid
	 */
	private createGridBody(context: ComponentFramework.Context<IInputs>, columnsOnView: DataSetInterfaces.Column[], gridParam: DataSet):HTMLDivElement{
		let gridBody:HTMLDivElement = document.createElement("div");
		let rankingFieldName:string = 'bingo'; //context.parameters.rankingfieldname.raw as string;
		let groupingFieldName:string = 'bingo'; //context.parameters.groupingfieldname.raw as string;
		let linkfieldname:string = context.parameters.linkfieldname.raw as string;
		let entitydisplayname:string = context.parameters.entitydisplayname.raw as string;
		let alloweditflaggedforreview_S:string = (context.parameters.alloweditflaggedforreview.raw as string);
		let alloweditflaggedforreview:boolean = false;
		
		if (alloweditflaggedforreview_S.toLowerCase() == 'true') alloweditflaggedforreview = true;

		let spacer:HTMLDivElement = document.createElement("div");
		//spacer.style.height = "20px";
		gridBody.appendChild(spacer);

		let lastCategory:string = "";
		let categoryContainer:HTMLDivElement = document.createElement("div");

		let columnHeader:HTMLDivElement = this.getColumnHeaders(columnsOnView, groupingFieldName);

		let categoryHeader:HTMLDivElement = this.getCategoryHeader('', columnHeader);

		gridBody.appendChild(categoryHeader);

		let width:number = 0;
		columnsOnView.forEach(function(columnItem, index){
			if (!columnItem.isHidden &&  columnItem.name.indexOf(groupingFieldName) == -1) {
				width += 1;
			}
		});

		if (gridParam.sortedRecordIds.length > 0)
		{
			for(let currentRecordId of gridParam.sortedRecordIds){
				let record:DataSetInterfaces.EntityRecord = gridParam.records[currentRecordId];
				let status:string = 'new';
				
				let gridRecord: HTMLDivElement = document.createElement("div");
				gridRecord.classList.add("DataSetControl_grid-dataRow");
				//gridRecord.addEventListener("click", this.onRowClick.bind(this));

				// Set the recordId on the row dom
				gridRecord.setAttribute(RowRecordId, gridParam.records[currentRecordId].getRecordId());
				let linkValue:string = gridParam.records[currentRecordId].getFormattedValue(linkfieldname);
				gridRecord.setAttribute(linkfieldname, linkValue);

				var obj = this;

				columnsOnView.forEach(function(columnItem, index){
					if (columnItem.name.indexOf("statuscode") > -1) {
						status = gridParam.records[currentRecordId].getFormattedValue(columnItem.name).toLowerCase();
					}
					if (!columnItem.isHidden &&  columnItem.name.indexOf(groupingFieldName) == -1) {
						let valuePara:HTMLDivElement = document.createElement("div");
						valuePara.classList.add("DataSetControl_grid-dataCol");
						let priority:number = 0;
						
						if(gridParam.records[currentRecordId].getFormattedValue(columnItem.name) != null && gridParam.records[currentRecordId].getFormattedValue(columnItem.name) != "")
						{
							valuePara.textContent = "";
							if (columnItem.name == "caps_totalprojectcost") valuePara.textContent = "$";
							valuePara.textContent += gridParam.records[currentRecordId].getFormattedValue(columnItem.name);
							//if(columnItem.name.indexOf(rankingFieldName) > -1) priority = gridParam.records[currentRecordId].getValue(columnItem.name) as number;
						}
						else
						{
							valuePara.textContent = "-";
						}

						valuePara.setAttribute('data-datatype', columnItem.dataType);
						valuePara.setAttribute('data-fielddisplayname', columnItem.displayName);
						valuePara.setAttribute('data-value', gridParam.records[currentRecordId].getFormattedValue(columnItem.name));

						let fieldName:string = columnItem.name.toLowerCase();
						if (fieldName.indexOf('caps_projecttype') > -1) {
							valuePara.setAttribute('data-required', 'true');
							valuePara.setAttribute('data-entityname', 'caps_projecttypes');
							let entityRef:ComponentFramework.EntityReference = <ComponentFramework.EntityReference>(gridParam.records[currentRecordId]).getValue(columnItem.name);
							valuePara.setAttribute('data-value', entityRef.id.guid);
						}
						else if (fieldName.indexOf('caps_facility') > -1) {
							valuePara.setAttribute('data-required', 'true');
							valuePara.setAttribute('data-entityname', 'caps_facilities');
							let entityRef:ComponentFramework.EntityReference = <ComponentFramework.EntityReference>(gridParam.records[currentRecordId]).getValue(columnItem.name);
							if (entityRef != null) valuePara.setAttribute('data-value', entityRef.id.guid);
							if (valuePara.textContent == '-') valuePara.textContent = 'Other';
						}
						else if (fieldName.indexOf('caps_projectdescription') > -1) {
							valuePara.setAttribute('data-required', 'true');
						}

						/*
						if(columnItem.name.indexOf(rankingFieldName) > -1) {
							let attr:Attr = document.createAttribute("original_order");
							attr.value = priority.toString();
							valuePara.attributes.setNamedItem(attr);
						}
						*/

						if (columnItem.name == linkfieldname) valuePara.classList.add("click-to-open");

						gridRecord.appendChild(valuePara);
					}
				});

				let allowUnlinkRecord:string = this.contextObj.parameters.allowunlinkrecord.raw as string;
				let allowUnlink:boolean = false;
				if (allowUnlinkRecord != null && allowUnlinkRecord.toLowerCase().indexOf("true") > -1) allowUnlink = true;

				let valuePara:HTMLDivElement = document.createElement("div");
				valuePara.classList.add("DataSetControl_grid-dataColDel");

				if (this.isActive || !this.lockOnInactive) {
					if (status.indexOf('cancelled') == -1) {
						let editLink:HTMLAnchorElement = document.createElement("a");
						editLink.classList.add("edit-button");
						editLink.classList.add("action-button");
						editLink.href = "";
						editLink.innerHTML = "<i class=\"fas fa-edit\"></i>";
						editLink.title = "Edit " + entitydisplayname + " " + linkValue;
						valuePara.appendChild(editLink);

						if (alloweditflaggedforreview == false) {
							let cancelLink:HTMLAnchorElement = document.createElement("a");
							cancelLink = document.createElement("a");
							cancelLink.classList.add("cancelproj-button");
							cancelLink.classList.add("action-button");
							cancelLink.href = "";
							cancelLink.innerHTML = "<i class=\"fas fa-ban\"></i>";
							cancelLink.title = "Cancel " + entitydisplayname + " " + linkValue;
							cancelLink.style.paddingLeft = "5px";
							valuePara.appendChild(cancelLink);
						}
					}

					if (allowUnlink && alloweditflaggedforreview == false) {
						let deleteLink:HTMLAnchorElement = document.createElement("a");
						deleteLink.classList.add("delete-button");
						deleteLink.classList.add("action-button");
						deleteLink.href = "";
						deleteLink.innerHTML = "<i class=\"fas fa-unlink\"></i>";
						deleteLink.title = "Remove " + entitydisplayname + " " + linkValue;
						deleteLink.style.paddingLeft = "5px";
						valuePara.appendChild(deleteLink);
					}
				}

				gridRecord.appendChild(valuePara);
				
				categoryContainer.appendChild(gridRecord);

				//$('.DataSetControl_grid-container').children().first().css("width", width.toString() + "px");
			}
		}
		else
		{
			//let noRecordLabel: HTMLDivElement = document.createElement("div");
			//noRecordLabel.classList.add("DataSetControl_grid-norecords");
			//noRecordLabel.style.width = this.contextObj.mode.allocatedWidth - 25 + "px";
			//noRecordLabel.innerHTML = this.contextObj.resources.getString("PCF_DataSetControl_No_Record_Found");
			//gridBody.appendChild(noRecordLabel);
		}

		if ((this.isActive || !this.lockOnInactive) && alloweditflaggedforreview==false) {
			categoryContainer.appendChild(this.getAddEditRow(context, columnsOnView, entitydisplayname, true, null));
		}
		
		width = (width * 160) + 100;
		gridBody.style.width = width.toString() + "px";

		gridBody.appendChild(categoryContainer);

		return gridBody;
	}

	private getAddEditRow(context:ComponentFramework.Context<IInputs>, columnsOnView: DataSetInterfaces.Column[], entityDisplayName:string, isAddRow:boolean, row:HTMLDivElement|null):HTMLDivElement {
		let alloweditflaggedforreview_S:string = (context.parameters.alloweditflaggedforreview.raw as string);
		let alloweditflaggedforreview:boolean = false;

		if (alloweditflaggedforreview_S.toLowerCase() == 'true') alloweditflaggedforreview = true;

		let gridRecord:HTMLDivElement = document.createElement("div");
		gridRecord.classList.add("DataSetControl_grid-dataRow");

		var obj = this;
		let recordId:string = '';

		if (isAddRow) {
			gridRecord.classList.add("add-dataRow");
		}
		else { //!isAddRow
			var jRow = $(<HTMLDivElement>row);
			var dataFields = jRow.find('div');
			recordId = jRow.attr('recordid')!;
			gridRecord.setAttribute('recordid', recordId);
		}

		let saveSelect:HTMLSelectElement;

		var obj = this;
		var hiddenCount = 0;

		/* Add "Add record" row */
		columnsOnView.forEach(function(columnItem, index){
			if (columnItem.isHidden) hiddenCount++;
			else {
				let valuePara:HTMLDivElement = document.createElement("div");
				valuePara.classList.add("DataSetControl_grid-dataCol");

				if (columnItem.name.indexOf('caps_projectnumber') > -1 || columnItem.name.indexOf('caps_projectcode') > -1) {
					let span:HTMLSpanElement = document.createElement("span");
					if (isAddRow) span.innerText = 'New ' + entityDisplayName;
					else span.innerText = $(dataFields[index-hiddenCount]).data('value');
					valuePara.appendChild(span);
				}
				else if (columnItem.name.indexOf('statuscode') > -1) {
					let span:HTMLSpanElement = document.createElement("span");
					if (isAddRow) span.innerText = 'New';
					else span.innerText = $(dataFields[index-hiddenCount]).data('value');
					valuePara.appendChild(span);
				}
				else if (columnItem.name.indexOf('caps_flaggedforreview') > -1 && alloweditflaggedforreview==false) {
					let span:HTMLSpanElement = document.createElement("span");
					if (isAddRow) span.innerText = 'No';
					else span.innerText = $(dataFields[index-hiddenCount]).data('value');
					valuePara.appendChild(span);
				}
				else if (alloweditflaggedforreview==true && columnItem.name.indexOf('caps_flaggedforreview') == -1) {
					let span:HTMLSpanElement = document.createElement("span");
					if (isAddRow) span.innerText = '';
					else span.innerText = $(dataFields[index-hiddenCount]).data('value');
					valuePara.appendChild(span);
				}
				else {
					let input:HTMLInputElement = document.createElement("input");
					input.id = columnItem.name;
					input.className = 'input-field';

					let attr:Attr = document.createAttribute("autocomplete");
					attr.value = "off";
					input.attributes.setNamedItem(attr);

					switch(columnItem.dataType.toLowerCase())
					{
						case 'lookup.simple':
							var id = '';
							if (!isAddRow) id = $(dataFields[index-hiddenCount]).data('value');	

							let select:HTMLSelectElement = obj.addSelectOptions(context, columnItem.name, isAddRow, id);
							select.setAttribute('data-fielddisplayname', columnItem.displayName);
							if (columnItem.name.indexOf('caps_facility') > -1 || columnItem.name.indexOf('caps_projecttype') > -1) select.setAttribute('data-required', 'true');

							if (columnItem.name.indexOf('caps_facility') > -1) saveSelect = select;
							
							valuePara.appendChild(select);
							break;
						case 'multiple':
							input.maxLength = 2000;
							input.setAttribute('data-fielddisplayname', columnItem.displayName);
							if (columnItem.name.indexOf('caps_projectdescription') > -1) input.setAttribute('data-required', 'true');
							if (!isAddRow) input.value = $(dataFields[index-hiddenCount]).data('value');
							valuePara.appendChild(input);
							break;
						case 'decimal':
							input.maxLength = 20;
							input.setAttribute('data-fielddisplayname', columnItem.displayName);
							input.setAttribute('data-datatype', columnItem.dataType);
							if (!isAddRow) input.value = $(dataFields[index-hiddenCount]).data('value');
							valuePara.appendChild(input);
							break;
						case 'twooptions':
							input.type = 'checkbox';
							input.setAttribute('data-fielddisplayname', columnItem.displayName);
							input.setAttribute('data-datatype', columnItem.dataType);
							if (!isAddRow) {
								let val:string = $(dataFields[index-hiddenCount]).data('value').toLowerCase(); 
								if (val=='true' || val=='yes') input.checked = true;
								else input.checked = false;
							} 
							valuePara.appendChild(input);
							break;
						case 'optionset':
							break;
						default:
							input.maxLength = 200;
							input.setAttribute('data-fielddisplayname', columnItem.displayName);
							if (!isAddRow) input.value = $(dataFields[index-hiddenCount]).data('value');
							if (columnItem.name.indexOf('caps_otherfacility') > -1) {
								input.setAttribute('data-required', 'true');
								if (input.value.length === 0) input.disabled = true;
								else {
									var inputId = saveSelect.id;
									obj.intervalHandler = window.setInterval(function(){
										var sel = $('#' + inputId);
										if (sel.length > 0 && sel.children().length > 1) {
											window.clearInterval(obj.intervalHandler)
											$('#' + inputId).val('0');
										}
									},50);
								};
							}
							valuePara.appendChild(input);
							break;
					}
				}

				gridRecord.appendChild(valuePara);
			}
		});

		let valuePara:HTMLDivElement = document.createElement("div");
		valuePara.classList.add("DataSetControl_grid-dataColDel");

		if (this.isActive || !this.lockOnInactive) {
			let editLink:HTMLAnchorElement = document.createElement("a");
			editLink.classList.add("add-button");
			editLink.classList.add("action-button");
			editLink.href = "";
			editLink.innerHTML = "<i class=\"fas fa-save\"></i>";
			editLink.title = "Save " + entityDisplayName;
			valuePara.appendChild(editLink);

			if (!isAddRow) {
				editLink = document.createElement("a");
				editLink.classList.add("cancel-button");
				editLink.classList.add("action-button");
				editLink.href = "";
				editLink.innerHTML = "<i class=\"fas fa-ban\"></i>";
				editLink.title = "Cancel edit";
				editLink.style.paddingLeft = "5px";
				valuePara.appendChild(editLink);
			}
		}
		
		gridRecord.appendChild(valuePara);

		return gridRecord;
	}

	private addSelectOptions(context: ComponentFramework.Context<IInputs>, columnName:string, isAddRow:boolean, recordId:string):HTMLSelectElement {
		let select: HTMLSelectElement = document.createElement("select");
		let valueField:string = '';
		let textField:string = '';

		if (isAddRow) select.id = columnName;
		else select.id = columnName + '_' + recordId;
		select.className = 'input-field';

		let query:string = '';

		switch (columnName.toLowerCase())
		{
			case 'caps_facility':
				select.setAttribute('data-entityname', 'caps_facilities')
				valueField = 'caps_facilityid';
				textField = 'caps_name';
				query = '?$select=' + valueField + ',' + textField + '&$orderby=' + textField + ' asc';
				break;
			case 'caps_projecttype':
				select.setAttribute('data-entityname', 'caps_projecttypes')
				valueField = 'caps_projecttypeid';
				textField = 'caps_type';
				query = '?fetchXml=<fetch top="50" distinct="true" >' +
					'<entity name="caps_projecttype" >' +
					'<attribute name="' + valueField + '" />' +
					'<attribute name="' + textField + '" />' +
					'<order attribute="caps_type" />' +
					'<link-entity name="caps_submissioncategory" from="caps_submissioncategoryid" to="caps_submissioncategory" link-type="inner">' +
						'<filter type="and" >' +
							'<condition attribute="caps_type" operator="eq" value="200870002" />' +
						'</filter>' +
					'</link-entity>' +
					'</entity>' +
				'</fetch>';
				break;
			default:
				break;
		}

		if (query.length > 0) {
			var selectId = select.id;
			var obj=this;
			try{
				context.webAPI.retrieveMultipleRecords(columnName.toLowerCase(), query).then(
					function (response) 
					{
						if (response.entities != null && response.entities.length > 0) {
							var select = $('#' + selectId);
							select.children().remove();

							let option:HTMLOptionElement = document.createElement('option');
							option.value = '';
							option.text = '-- SELECT --';
							select.append(option);

							if (columnName.toLowerCase() == 'caps_facility') {
								option = document.createElement('option');
								option.value = '0';
								option.text = 'Other (Please Specify)';
								select.append(option);

								select.on('change', obj.facilityChanged);
							}

							for (var i = 0; i < response.entities.length; i++) {
								let entity:ComponentFramework.WebApi.Entity = response.entities[i];
								option = document.createElement('option');
								option.value = entity[valueField];
								option.text = entity[textField];

								select.append(option);
							}

							select.val(recordId);
						}
					},
					function (errorResponse: any) 
					{
						// Error handling code here - record failed to be created
						alert("An error occured while running the query:\r\n\r\n" + errorResponse.message);
					}
				);
			}
			catch (ex) { alert (ex); }
		}

		return select;
	}

	private facilityChanged(event:Event) {
		let slct:HTMLSelectElement = event.target as HTMLSelectElement;

		var otherFacility = $(slct).parent().parent().find('#caps_otherfacility');

		if (slct.value === '0') {
			otherFacility.prop('disabled', false);
		}
		else {
			$(otherFacility).val('');
			otherFacility.prop('disabled', true);
			otherFacility.css('border-width', '1px');
			otherFacility.css('border-color', '#C3C3C3');
		}
	}
	
	private getCategoryHeader(categoryName:string, columnHeader:HTMLDivElement):HTMLDivElement {
		let categoryHeader:HTMLDivElement = document.createElement("div");
		categoryHeader.classList.add("DataSetControl_category-header");

		let container:HTMLDivElement = document.createElement("div");

		/*
		container.classList.add("DataSetControl_category-title");

		let plusIcon = document.createElement("i"); //<i class="far fa-plus-square"></i> <i class="far fa-minus-square"></i>
		plusIcon.classList.add("far", "fa-minus-square", "icon_collapser");
		container.appendChild(plusIcon);

		let input:HTMLInputElement = document.createElement("input") as HTMLInputElement;
		input.type = "checkbox";
		input.checked = true;
		input.title = categoryName;
		input.classList.add("DataSetControl_collapser");
		container.appendChild(input);

		let span:HTMLSpanElement = document.createElement("span");
		span.innerHTML = "&nbsp;" + categoryName;
		container.appendChild(span);

		categoryHeader.appendChild(container);

		container = document.createElement("div");
		*/

		container.innerHTML = columnHeader.innerHTML;
		categoryHeader.appendChild(container);

		return categoryHeader;
	}

	private getColumnHeaders(columnsOnView: DataSetInterfaces.Column[], groupingFieldName:string):HTMLDivElement {
		let columnHeader:HTMLDivElement = document.createElement("div");
		columnHeader.classList.add("DataSetControl_grid-header");

		columnsOnView.forEach(function(columnItem, index){
			if (!columnItem.isHidden && columnItem.name != groupingFieldName) {
				let labelPara:HTMLDivElement = document.createElement("div");
				labelPara.classList.add("DataSetControl_grid-headerCol");
				labelPara.textContent = columnItem.displayName;
				columnHeader.appendChild(labelPara);
			}
		});

		let allowUnlinkRecord:string = this.contextObj.parameters.allowunlinkrecord.raw as string;
		let allowUnlink:boolean = false;
		if (allowUnlinkRecord != null && allowUnlinkRecord.toLowerCase().indexOf("true") > -1) allowUnlink = true;

		let labelPara:HTMLDivElement = document.createElement("div");
		labelPara.classList.add("DataSetControl_grid-headerColDel");
		if (allowUnlink == true) labelPara.textContent = "Actions";
		columnHeader.appendChild(labelPara);

		return columnHeader;
	}

	private getColumnItem(columnsOnView: DataSetInterfaces.Column[], columnName:string): string {
		columnsOnView.forEach(function(columnItem, index){
			if (columnItem.name == columnName) return columnItem.displayName;
		});

		return columnName;
	}

	/**
	 * Row Click Event handler for the associated row when being clicked
	 * @param event
	 */
	private onRowClick(event: Event): void {
		let rowRecordId = (event.currentTarget! as HTMLTableRowElement).parentElement!.getAttribute(RowRecordId);

		if(rowRecordId)
		{
			let entityReference = this.contextObj.parameters.AFGProjectsDataSet.records[rowRecordId].getNamedReference();
			let entityFormOptions = {
				entityName: (entityReference as any).LogicalName,
				entityId: rowRecordId,
				target: 2
			}
			this.contextObj.navigation.openForm(entityFormOptions);
		}
	}

	/**
	 * Toggle 'LoadMore' button when needed
	 */
	private toggleLoadMoreButtonWhenNeeded(gridParam: DataSet): void{
		
		if(gridParam.paging.hasNextPage && this.loadPageButton.classList.contains(DataSetControl_LoadMoreButton_Hidden_Style))
		{
			this.loadPageButton.classList.remove(DataSetControl_LoadMoreButton_Hidden_Style);
		}
		else if(!gridParam.paging.hasNextPage && !this.loadPageButton.classList.contains(DataSetControl_LoadMoreButton_Hidden_Style))
		{
			this.loadPageButton.classList.add(DataSetControl_LoadMoreButton_Hidden_Style);
		}

	}

	/**
	 * 'LoadMore' Button Event handler when load more button clicks
	 * @param event
	 */
	private onLoadMoreButtonClick(event: Event): void {
		this.contextObj.parameters.AFGProjectsDataSet.paging.loadNextPage();
		this.toggleLoadMoreButtonWhenNeeded(this.contextObj.parameters.AFGProjectsDataSet);
	}

	private onRowEditClick(event: Event): void {
		event.preventDefault();

		$(event.target as HTMLElement).parent().find('.edit-button').click();
	}

	private bindEvents() {
		$(".DataSetControl_grid-dataRow").on("click", ".click-to-open", this.onRowClick.bind(this));
		//$(".DataSetControl_grid-dataRow").on('click', this.onRowEditClick.bind(this));

		$('.action-button').off('click');

		var obj = this;

		$('.cancel-button').on('click', function(event) {
			event.preventDefault();

			obj.isEditing = false;

			$(this).parent().parent().prev().show();
			$(this).parent().parent().remove();
			$('.add-dataRow').show();

			obj.bindEvents();

			return false;
		});

		$('.add-button').on('click', function(event) {
			event.preventDefault();

			if ($(this).prop('disabled')===true) return false;

			$(this).prop('disabled',true);

			var div = $(this).parent().parent();
			var fields = div.find('.input-field');
			let entity:ComponentFramework.WebApi.Entity = {};

			var recordId = div.attr('recordId');

			// UNSUPPORTED
			entity['caps_Submission@odata.bind'] = 'caps_submissions(' + (obj.contextObj.mode as any).contextInfo.entityId + ')';

			obj.contextObj.webAPI.retrieveMultipleRecords('caps_submissioncategory', '?$select=caps_submissioncategoryid&$filter=caps_type eq 200870002&$top=1').then(
				function (response) 
				{
					let valid:boolean = true;
					let validationErrors:string = '';

					if (response.entities != null && response.entities.length > 0) {						
						entity['caps_SubmissionCategory@odata.bind'] = 'caps_submissioncategories(' + response.entities[0]['caps_submissioncategoryid'] + ')';

						fields.each(function() {
							let dataType:string =$(this).data('datatype');
							let entityName:string = $(this).data('entityname');
							let req:string = <string>$(this).data('required') + '';
							let required:boolean = false;
							if (req != null && req == 'true') required = true;

							if (required && $(this).val() == '' && $(this).prop('disabled') === false) {
								let fieldDisplayName:string = $(this).data('fielddisplayname');
								validationErrors += '"' + fieldDisplayName + '" is required.\r\n';
								valid = false;
								this.style.borderColor = '#ff0000';
								this.style.borderWidth = '1px';
								$('.add-button').prop('disabled',false);
								return;
							}
			
							if (entityName != null && entityName.length > 0) {
								if (this.id.indexOf('caps_projecttype') > -1)
									entity['caps_ProjectType@odata.bind'] = entityName + '(' + $(this).val() + ')';
								else if (this.id.indexOf('caps_facility') > -1)
								{ 
									if ($(this).val() != '0')
										entity['caps_Facility@odata.bind'] = entityName + '(' + $(this).val() + ')';
									else 
										entity['_caps_facility_value'] = null;
								}
							}
							else if (dataType == null || dataType.length == 0) {
								entity[this.id] = $(this).val();
							}
							else if (dataType.toLowerCase().indexOf('decimal') > -1) {
								let s:string = (<string>$(this).val()).replace(',','').replace('$','');
								let x:number = parseFloat(s);
								if (s != null && s.length > 0) entity[this.id] = x;
								if (s.length > 0 && isNaN(x)) {
									validationErrors += '"' + $(this).data('fielddisplayname') + '" is not a valid number.\r\n';
									valid = false;
									this.style.borderColor = '#ff0000';
									this.style.borderWidth = '1px';
									$('.add-button').prop('disabled',false);
									return;
								}
							}
							else if (dataType.toLowerCase().indexOf('twooptions') > -1) {
								entity[this.id] = $(this).prop("checked");
							}
						});
			
						if (valid==true) {
							//$('.DataSetControl_category-header').siblings().last().hide("slow");
							if (recordId != null && recordId.length > 0) {
								obj.contextObj.webAPI.updateRecord('caps_project', recordId, entity).then(
									function (response: ComponentFramework.EntityReference) {
										obj.isEditing = false;
										obj.contextObj.parameters.AFGProjectsDataSet.refresh();
									},
									function (errorResponse: any) 
									{
										//Error handling code here - record delete failed
										alert("An error occured while saving the record:\r\n\r\n" + errorResponse.message);
										//$('.DataSetControl_category-header').siblings().last().show("slow");

										$('.add-button').prop('disabled',false);
									}
								);
							}
							else {
								obj.contextObj.webAPI.createRecord('caps_project', entity).then(
									function (response: ComponentFramework.EntityReference) {
										obj.isEditing = false;
										obj.contextObj.parameters.AFGProjectsDataSet.refresh();
									},
									function (errorResponse: any) 
									{
										//Error handling code here - record delete failed
										alert("An error occured while creating the record:\r\n\r\n" + errorResponse.message);
										$('.DataSetControl_category-header').siblings().last().show("slow");
										$('.add-button').prop('disabled',false);
									}
								);
							}
						}
						else {
							var alertStrings = { confirmButtonLabel: "Ok", text: validationErrors, title: "Validation Errors:" };
							obj.contextObj.navigation.openAlertDialog(alertStrings);
							$('.add-button').prop('disabled',false);
						}
					}
					else {
						alert('Error: AFG submission category was not found.');
						$('.add-button').prop('disabled',false);
					}
				},
				function (errorResponse: any) 
				{
					// Error handling code here - record failed to be created
					alert("An error occured while querying AFG submission category:\r\n\r\n" + errorResponse.message);
					$('.add-button').prop('disabled',false);
				}
			);

			return false;
		});

		$(".edit-button").on('click', function(event) {
			event.preventDefault();

			if (obj.isEditing===true) {
				if (confirm('You\'re currently editing another row. Do you want to cancel editing the other row and edit this row instead? Unsaved changes will be lost.\r\n\r\nClick \'Cancel\' to continue editing the other row.')) {
					$('.cancel-button').click();
				}
				else return false;
			};

			obj.isEditing = true;
			$('.add-dataRow').hide();
			var row = $(this).parent().parent();
			let linkField:string = row.attr(obj.contextObj.parameters.linkfieldname.raw as string)!;
			let entityDisplayName:string = obj.contextObj.parameters.entitydisplayname.raw as string;
			let parentEntityDisplayName:string = obj.contextObj.parameters.parententitydisplayname.raw as string;

			row.hide();
			let div:HTMLDivElement = obj.getAddEditRow(obj.contextObj, obj.getSortedColumnsOnView(obj.contextObj), entityDisplayName, false, <HTMLDivElement>row[0]);
			$(div).insertAfter(row);

			obj.bindEvents();

			return false;
		});

		$(".cancelproj-button").on("click", function(event) {
			event.preventDefault();

			var row = $(this).parent().parent();
			let linkField:string = row.attr(obj.contextObj.parameters.linkfieldname.raw as string)!;
			let entityDisplayName:string = obj.contextObj.parameters.entitydisplayname.raw as string;

			obj.contextObj.navigation.openConfirmDialog({title:'Cancel ' + entityDisplayName + ' "' + linkField + '"?', text:'Are you sure you want to cancel ' + entityDisplayName + ' "' + linkField + '"?'}, {}).then(
				function(success) {
					if (success.confirmed) {
						let entityName:string = obj.contextObj.parameters.AFGProjectsDataSet.getTargetEntityType();
						//let parentEntityFieldName:string = obj.contextObj.parameters.parententityfieldname.raw as string;
						let recordId:string = row.attr(RowRecordId)!;

						let entity:ComponentFramework.WebApi.Entity = {};
						entity['statecode'] = 1;
						entity['statuscode'] = 100000010;

						obj.contextObj.webAPI.updateRecord(entityName, recordId, entity).then(
							function (response: ComponentFramework.EntityReference) {
								obj.contextObj.parameters.AFGProjectsDataSet.refresh();
								//var group = row.parent();
								//row.remove();
								//var items = group.children();
								//obj.updateItem(items, 0);
							},
							function (errorResponse: any) 
							{
								//Error handling code here - record delete failed
								alert("An error occured while removing the record:\r\n\r\n" + errorResponse.message);
							}
						);
					}
				}
			);

			return false;
		});

		$(".delete-button").on("click", function(event) {
			event.preventDefault();

			var row = $(this).parent().parent();
			let linkField:string = row.attr(obj.contextObj.parameters.linkfieldname.raw as string)!;
			let entityDisplayName:string = obj.contextObj.parameters.entitydisplayname.raw as string;
			let parentEntityDisplayName:string = obj.contextObj.parameters.parententitydisplayname.raw as string;

			obj.contextObj.navigation.openConfirmDialog({title:'Remove ' + entityDisplayName + ' "' + linkField + '"?', text:'Are you sure you want to remove ' + entityDisplayName + ' "' + linkField + '"?'}, {}).then(
				function(success) {
					if (success.confirmed) {
						let entityName:string = obj.contextObj.parameters.AFGProjectsDataSet.getTargetEntityType();
						let parentEntityFieldName:string = obj.contextObj.parameters.parententityfieldname.raw as string;
						let recordId:string = row.attr(RowRecordId)!;

						var data:any = {};
						data["_" + parentEntityFieldName + "_value"] = null;

						//obj.contextObj.webAPI.deleteRecord(entityName, recordId).then(
						obj.contextObj.webAPI.updateRecord(entityName, recordId, data).then(
						//obj.contextObj.webAPI.deleteRecord(entityName, recordId).then(
							function (response: ComponentFramework.EntityReference) {
								obj.contextObj.parameters.AFGProjectsDataSet.refresh();
								var group = row.parent();
								row.remove();
								var items = group.children();
								//obj.updateItem(items, 0);
							},
							function (errorResponse: any) 
							{
								//Error handling code here - record delete failed
								alert("An error occured while removing the record:\r\n\r\n" + errorResponse.message);
							}
						);
					}
				}
			);

			return false;
		});
	}
}
