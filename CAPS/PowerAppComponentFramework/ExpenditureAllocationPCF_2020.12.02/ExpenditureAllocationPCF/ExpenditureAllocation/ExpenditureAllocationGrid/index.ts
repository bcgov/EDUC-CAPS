import {IInputs, IOutputs} from "./generated/ManifestTypes";
import DataSetInterfaces = ComponentFramework.PropertyHelper.DataSetApi;
import * as $ from 'jquery';
type DataSet = ComponentFramework.PropertyTypes.DataSet;

// Define const here
const RowRecordId:string = "recordId";

// Style name of Load More Button
const DataSetControl_LoadMoreButton_Hidden_Style = "DataSetControl_LoadMoreButton_Hidden_Style";

export class ExpenditureAllocationGrid implements ComponentFramework.StandardControl<IInputs, IOutputs> {
	// Cached context object for the latest updateView
	private contextObj: ComponentFramework.Context<IInputs>;
		
	// Div element created as part of this control's main container
	private mainContainer: HTMLDivElement;

	// Table element created as part of this control's table
	private gridContainer: HTMLDivElement;

	// Button element created as part of this control
	private loadPageButton: HTMLButtonElement;

	private isActive:boolean;
	private lockOnInactive:boolean;

	private parentEntityName:string;
	private parentEntityId:string;

	private tabIndex:number = 0;

	private _notifyOutputChanged: () => void;

	//private _totalCost:number = 0;
	//private _totalAllocated:number = 0;
	//private _totalUnallocated:number = 0;

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
		context.mode.trackContainerResize(true);

		this._notifyOutputChanged = notifyOutputChanged;

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
		this.toggleLoadMoreButtonWhenNeeded(context.parameters.ExpenditureAllocationDataSet);

		if(!context.parameters.ExpenditureAllocationDataSet.loading){
			this.isActive = true;
			this.lockOnInactive = false;

			// Get sorted columns on View
			let columnsOnView = this.getSortedColumnsOnView(context);
			let stateColumnName:string = "";

			//context.parameters.SortableDataSet.records[0].getFormattedValue("");

			if (!columnsOnView || columnsOnView.length === 0) {
				return;
			}

			/*if (this._totalCost == null || this._totalCost == 0) {
				this.GetParentEntityData(context);
				return;
			}*/

			while(this.gridContainer.firstChild)
			{
				this.gridContainer.removeChild(this.gridContainer.firstChild);
			}
			

			//Deanna - Removing Total Cost, Allocated and Unallocated fields.
			//this.gridContainer.appendChild(this.createAllocationHeader(context, columnsOnView));

			try {
				let lockInactive:string = context.parameters.lockOnInactiveState.raw as string;
				if (lockInactive.toLowerCase().indexOf("true") > -1) this.lockOnInactive = true;

				// Determine if the parent record state is inactice
				for (let column of columnsOnView){
					if(column.name.indexOf("statecode") > -1) {
						let state:string = context.parameters.ExpenditureAllocationDataSet.records[context.parameters.ExpenditureAllocationDataSet.sortedRecordIds[0]].getFormattedValue(column.name);

						if (state.toLowerCase().indexOf("inactive") > -1) this.isActive = false;
						
						break;
					}
				}
			}
			catch(e) {
				var s = e;
			}

			this.gridContainer.appendChild(this.createGridBody(context, columnsOnView, context.parameters.ExpenditureAllocationDataSet));

			//this.calculateAllocation(false);
		}

		// this is needed to ensure the scroll bar appears automatically when the grid resize happens and all the tiles are not visible on the screen.
		this.mainContainer.style.maxHeight = window.innerHeight - this.gridContainer.offsetTop - 75 + "px";

		this.bindEvents();
	}

	/*
	private GetParentEntityData(context: ComponentFramework.Context<IInputs>):void {
		let parentRecordFieldName:string = context.parameters.parentRecordFieldName.raw as string;
		let totalCostFieldName:string = context.parameters.totalCostFieldName.raw as string;

		this.parentEntityName = context.parameters.parentRecordEntityName.raw as string;

		let recordId:string = context.parameters.ExpenditureAllocationDataSet.sortedRecordIds[0];

		var obj = this;

		context.webAPI.retrieveRecord(context.parameters.ExpenditureAllocationDataSet.getTargetEntityType(),recordId,"?$select=_" + parentRecordFieldName + "_value").then
		(
			function (response) 
			{
				obj.parentEntityId = response["_" + parentRecordFieldName + "_value"];

				context.webAPI.retrieveRecord(obj.parentEntityName,obj.parentEntityId,"?$select=" + totalCostFieldName).then(
					function (response) 
					{
						obj._totalCost = response[totalCostFieldName];
						obj.updateView(obj.contextObj);
					},
					function (errorResponse: any) 
					{
						// Error handling code here - record failed to be created
						alert("An error occured while retrieving the parent record (1): " + errorResponse.toString());
					}
				);
			},
			function (errorResponse: any) 
			{
				// Error handling code here - record failed to be created
				alert("An error occured while retrieving the record (1): " + errorResponse.toString());
			}
		);

		//
	}*/

	/*
	private createAllocationHeader(context: ComponentFramework.Context<IInputs>, columnsOnView: DataSetInterfaces.Column[]):HTMLDivElement{
		let allocationHeader:HTMLDivElement = document.createElement("div");
		allocationHeader.classList.add("allocation-header");

		let totalAllocated:number = 0; 
		//if (context.parameters.totalAllocated != null) totalAllocated = context.parameters.totalAllocated.raw as number;
		let totalUnallocated:number = 0;
		//if (context.parameters.totalUnallocated != null) totalUnallocated = context.parameters.totalUnallocated.raw as number;

		let div:HTMLDivElement = document.createElement("div");
		div.innerText = "Total Cost";
		allocationHeader.appendChild(div);

		var formatter = new Intl.NumberFormat('en-US', {
			style: 'currency',
			currency: 'USD',
		  });

		let totalCost:string = formatter.format(this._totalCost).replace(/\$/g, "");

		let input:HTMLInputElement = document.createElement("input") as HTMLInputElement;
		input.type = "Text";
		input.id = "total_cost";
		input.value = totalCost;
		div = document.createElement("div");
		div.appendChild(input);
		allocationHeader.appendChild(div);

		div = document.createElement("div");
		div.innerText = "Total Allocated";
		div.classList.add("DataSetControl_AllocHeadr");
		allocationHeader.appendChild(div);

		div = document.createElement("div");
		div.id = "total_allocated";
		div.innerText = totalAllocated.toString();
		allocationHeader.appendChild(div);

		div = document.createElement("div");
		div.innerText = "Total Unallocated";
		div.classList.add("DataSetControl_AllocHeadr");
		allocationHeader.appendChild(div);

		div = document.createElement("div");
		div.id = "total_unallocated";
		div.innerText = totalUnallocated.toString();
		allocationHeader.appendChild(div);

		return allocationHeader;
	}
	*/

	/**
	 * funtion that creates the body of the grid and embeds the content onto the tiles.
	 * @param columnsOnView columns on the view whose value needs to be shown on the UI
	 * @param gridParam data of the Grid
	 */
	private createGridBody(context: ComponentFramework.Context<IInputs>, columnsOnView: DataSetInterfaces.Column[], gridParam: DataSet):HTMLDivElement{
		let gridBody:HTMLDivElement = document.createElement("div");
		let yearlyExpenditureFieldName:string = context.parameters.yearlyExpenditureFieldName.raw as string;

		if(gridParam.sortedRecordIds.length > 0)
		{
			let gridHeader:HTMLDivElement = this.getGridHeader(columnsOnView);
			gridBody.appendChild(gridHeader);

			//DataSetControl_grid-headerSec
			let gridColSec:HTMLDivElement = document.createElement("div");;
			let quotient:number = Math.floor(gridParam.sortedRecordIds.length / 2);
			let remainder:number = gridParam.sortedRecordIds.length % 2;
			let count:number = 1;

			for(let currentRecordId of gridParam.sortedRecordIds){
				if (count==1 || count == (quotient + remainder + 1)) {
					gridColSec = document.createElement("div");
					gridColSec.classList.add("DataSetControl_grid-headerSec");
				}

				let record:DataSetInterfaces.EntityRecord = gridParam.records[currentRecordId];
				
				let gridRecord: HTMLDivElement = document.createElement("div");
				gridRecord.classList.add("DataSetControl_grid-dataRow");

				// Set the recordId on the row dom
				gridRecord.setAttribute(RowRecordId, gridParam.records[currentRecordId].getRecordId());

				var obj = this;

				columnsOnView.forEach(function(columnItem, index){
					if (!columnItem.isHidden && columnItem.name.indexOf(".statecode") == -1) { //&& columnItem.name.indexOf(groupingFieldName) == -1
						let valuePara:HTMLDivElement = document.createElement("div");
						valuePara.classList.add("DataSetControl_grid-dataCol");
						let val:string = "";
						
						if(gridParam.records[currentRecordId].getFormattedValue(columnItem.name) != null && gridParam.records[currentRecordId].getFormattedValue(columnItem.name) != "")
						{
							valuePara.textContent = "";
							val = gridParam.records[currentRecordId].getFormattedValue(columnItem.name);
							if (val.indexOf(".") > -1) val = val.substring(0, val.indexOf("."));

							//if (columnItem.name == yearlyExpenditureFieldName) val = "$" + val;
						}
						else
						{
							val = "$0";
						}

						if (columnItem.name == yearlyExpenditureFieldName && (obj.isActive || !obj.lockOnInactive)) {
							let textBox:HTMLInputElement = document.createElement("input");
							textBox.type = "text";
							textBox.id = currentRecordId;
							textBox.tabIndex = count;
							textBox.value = val;
							textBox.classList.add("expenditure-allocation");

							valuePara.appendChild(textBox);
						} 
						else {
							valuePara.textContent = val;
						}

						gridRecord.appendChild(valuePara);
					}
				});

				gridColSec.appendChild(gridRecord);

				if (count == (quotient + remainder) || count == gridParam.sortedRecordIds.length) {
					gridBody.appendChild(gridColSec);

					gridColSec = document.createElement("div");
					gridColSec.classList.add("DataSetControl_grid-headerSec_Spcr");
					gridColSec.innerHTML = "&nbsp;";
					gridBody.appendChild(gridColSec);
				}

				//gridBody.appendChild(gridRecord);

				count++;
			}
		}
		else
		{
			let noRecordLabel: HTMLDivElement = document.createElement("div");
			noRecordLabel.classList.add("DataSetControl_grid-norecords");
			noRecordLabel.style.width = this.contextObj.mode.allocatedWidth - 25 + "px";
			noRecordLabel.innerHTML = this.contextObj.resources.getString("PCF_DataSetControl_No_Record_Found");
			gridBody.appendChild(noRecordLabel);
		}

		return gridBody;
	}

	/**
	 * funtion that creates the header of the grid.
	 * @param columnsOnView columns on the view whose value needs to be shown on the UI
	 */
	private getGridHeader(columnsOnView: DataSetInterfaces.Column[]):HTMLDivElement {
		let gridHeader:HTMLDivElement = document.createElement("div");
		gridHeader.classList.add("DataSetControl_grid-header");
		//let groupingFieldName:string = this.contextObj.parameters.groupingfieldname.raw as string;

		for (let i=1;i<=2;i++)
		{
			let gridCol:HTMLDivElement = document.createElement("div");
			gridCol.classList.add("DataSetControl_grid-headerSec");

			columnsOnView.forEach(function(columnItem, index){
				if (!columnItem.isHidden && columnItem.name.indexOf(".statecode") == -1) { //&& columnItem.name != groupingFieldName
					let labelPara:HTMLDivElement = document.createElement("div");
					labelPara.classList.add("DataSetControl_grid-headerCol");
					labelPara.textContent = columnItem.displayName;
					gridCol.appendChild(labelPara);
				}
			});

			gridHeader.appendChild(gridCol);

			gridCol = document.createElement("div");
			gridCol.classList.add("DataSetControl_grid-headerSec_Spcr");
			gridCol.innerHTML = "&nbsp;";
			gridHeader.appendChild(gridCol);
		}

		return gridHeader;
	}

	/**
	 * Get sorted columns on view
	 * @param context 
	 * @return sorted columns object on View
	 */
	private getSortedColumnsOnView(context: ComponentFramework.Context<IInputs>): DataSetInterfaces.Column[]
	{
		if (!context.parameters.ExpenditureAllocationDataSet.columns) {
			return [];
		}
		
		let columns =context.parameters.ExpenditureAllocationDataSet.columns
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
	 * It is called by the framework prior to a control receiving new data. 
	 * @returns an object based on nomenclature defined in manifest, expecting object[s] for property marked as “bound” or “output”
	 */
	public getOutputs(): IOutputs
	{
		return {
			//totalAllocated: this._totalAllocated,
			//totalUnallocated: this._totalUnallocated
		};
	}

	private formatInputValueAsNumber(input:string):number {
		input = input.replace(/\$/g, "").replace(/,/g,"");
		if (input.indexOf(".") > -1) input = input.substring(0, input.indexOf("."));

		let parseNum:number = Number(input);
		if (!isNaN(parseNum)) return parseNum;
		return 0;
	}

	private formatNumberAsString(input:number):string {
		var formatter = new Intl.NumberFormat('en-US', {
			style: 'currency',
			currency: 'USD',
		  });

		let formatted:string = formatter.format(input).replace(/\$/g, "");
		if (formatted.indexOf(".") > -1) formatted = formatted.substring(0, formatted.indexOf("."));

		return formatted;
	}

	/**
	 * 'LoadMore' Button Event handler when load more button clicks
	 * @param event
	 */
	private onChangeAllocation(event: Event): void {
		this.enableFields(false);
		let input:HTMLInputElement = event.target as HTMLInputElement;
		this.tabIndex = input.tabIndex;

		let val:number = 0;
		let tempVal:string = input.value;
		tempVal = tempVal.replace(/\$/g, "").replace(/,/g,"");
		
		let parseNum:number = Number(tempVal);
		if (!isNaN(parseNum))
		{
			 val = this.formatInputValueAsNumber(input.value);
			 $(input).val(this.formatNumberAsString(parseNum));
		}
		else {
			input.classList.add("input-error");
			return;
		}

		let id:string = input.id;
		let yearlyExpenditureFieldName:string = this.contextObj.parameters.yearlyExpenditureFieldName.raw as string;

		var data:any = {};
		data[yearlyExpenditureFieldName] = val;

		try 
		{
			let entityName:string = this.contextObj.parameters.ExpenditureAllocationDataSet.getTargetEntityType();
			var obj = this;
			this.contextObj.webAPI.updateRecord(entityName, id, data).then
			(
				function (response: ComponentFramework.EntityReference) 
				{ 
					// Callback method for successful update of record

					// Get the ID of the new record created
					//let id: string = response.id.guid;

					//obj.contextObj.parameters.ExpenditureAllocationDataSet.refresh();
					obj.contextObj.parameters.ExpenditureAllocationDataSet.refresh();
				},
				function (errorResponse: any) 
				{
					// Error handling code here - record failed to be created
					alert("An error occured while saving the updated expenditure (1): " + errorResponse.toString());
				}
			);
		}
		catch(ex)
		{
			alert("An error occured while saving the updated expenditure: " + ex.toString());
		}

		//this.calculateAllocation(true);
		//this.contextObj.parameters.ExpenditureAllocationDataSet.refresh();
	}
/*
	private onChangeTotal(event: Event): void {
		this.enableFields(false);
		let input:HTMLInputElement = event.target as HTMLInputElement;

		let tempVal:string = input.value;
		tempVal = tempVal.replace(/\$/g, "").replace(/,/g,"");
		
		let parseNum:number = Number(tempVal);
		if (!isNaN(parseNum))
		{
			 $(input).val(this.formatNumberAsString(parseNum));
			 $(input).removeClass("input-error");
		}
		else {
			input.classList.add("input-error");
			return;
		}

		//this.calculateAllocation(true);
		
	}
	*/
/*
	private calculateAllocation(refreshDataset:boolean) {
		
		let total:number = 0;
		let allocated:number = 0;
		let unallocated:number = 0;

		//if (this.contextObj.parameters.totalCost != null && this.contextObj.parameters.totalCost.raw != null) total = this.contextObj.parameters.totalCost.raw as number;

		let tempVal:string = $("#total_cost").val() as string;
		tempVal = tempVal.replace(/\$/g, "").replace(/,/g,"");

		let parseNum:number = Number(tempVal);
		if (parseNum != NaN) total = parseNum;
		var obj = this;

		$(".expenditure-allocation").each(function(){
			$(this).removeClass("input-error");
			let val:number = obj.formatInputValueAsNumber($(this).val() as string);

			allocated += val;
			unallocated = total - allocated;
		});

		let allocatedS:string = this.formatNumberAsString(allocated);
		let unallocatedS:string = this.formatNumberAsString(unallocated);

		$("#total_allocated").text(allocatedS);
		$("#total_unallocated").text(unallocatedS);
		$("#total_cost").val(this.formatNumberAsString(total));

		// Set bound columns and notify parent form of change
		// TODO: MS Bug? Do bound parameters work on dataset control?
		this._totalAllocated = allocated;
		this._totalUnallocated = unallocated;
		this._notifyOutputChanged();

		var data:any = {};
		data[this.contextObj.parameters.totalCostFieldName.raw as string] = total;
		data[this.contextObj.parameters.totalAllocatedFieldName.raw as string] = allocated;
		data[this.contextObj.parameters.totalUnallocatedFieldName.raw as string] = unallocated;
		var obj = this;
		this.contextObj.webAPI.updateRecord(this.parentEntityName, this.parentEntityId, data).then(
			function (response: ComponentFramework.EntityReference) 
				{ 
					//obj.enableFields(true);
					// Callback method for successful update of record

					// Get the ID of the new record created
					//let id: string = response.id.guid;

					if (refreshDataset==true) {
						obj._totalCost = 0;
						obj.contextObj.parameters.ExpenditureAllocationDataSet.refresh();
					}
				},
				function (errorResponse: any) 
				{
					// Error handling code here - record failed to be created
					alert("An error occured while saving the updated parent record (1): " + errorResponse.toString());
					obj.enableFields(true);
				}
		);
		
	}*/

	private bindEvents() {
		$(".expenditure-allocation").on("change", this.onChangeAllocation.bind(this));
		//$("#total_cost").on("change", this.onChangeTotal.bind(this));

		$(".DataSetControl_grid-dataCol input[type=text]")
			.eq(this.tabIndex).trigger("focus");
	}

	private enableFields(enable:boolean): void {
		$(".expenditure-allocation").prop("disabled", !enable);
		//$("#total_cost").prop("disabled", !enable);
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
		this.contextObj.parameters.ExpenditureAllocationDataSet.paging.loadNextPage();
		this.toggleLoadMoreButtonWhenNeeded(this.contextObj.parameters.ExpenditureAllocationDataSet);
	}

	/** 
	 * Called when the control is to be removed from the DOM tree. Controls should use this call for cleanup.
	 * i.e. cancelling any pending remote calls, removing listeners, etc.
	 */
	public destroy(): void
	{
		// Add code to cleanup control if necessary
	}
}