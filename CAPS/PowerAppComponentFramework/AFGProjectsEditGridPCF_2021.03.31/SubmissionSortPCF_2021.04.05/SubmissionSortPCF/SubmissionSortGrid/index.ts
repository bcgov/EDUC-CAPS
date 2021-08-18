import {IInputs, IOutputs} from "./generated/ManifestTypes";
import DataSetInterfaces = ComponentFramework.PropertyHelper.DataSetApi;
import * as $ from 'jquery';
import 'jqueryui';

type DataSet = ComponentFramework.PropertyTypes.DataSet;

// Define const here
const RowRecordId:string = "recordId";

// Style name of Load More Button
const DataSetControl_LoadMoreButton_Hidden_Style = "DataSetControl_LoadMoreButton_Hidden_Style";

class EntityReference {
	id: string;
	entityName: string;
	constructor(entityName: string, id: string) {
		this.id = id;
		this.entityName = entityName;
	}
}

class Column {
	/**
	 * Name of the column, unique in this dataset
	 */
	name: string;

	/**
	 * Localized display name for the column
	 */
	displayName: string;

	/**
	 * The manifest type of this column's values.
	 */
	dataType: string;

	/**
	 * The alias of this column.
	 */
	alias: string;

	/**
	 * The column order for the layout
	 */
	order: number;

	/**
	 * Customized column width ratios
	 */
	visualSizeFactor: number;

	/**
	 * The column visibility state.
	 */
	isHidden?: boolean;

	/**
	 * Is specific column the primary attrribute of the view's entity
	 */
	isPrimary?: boolean;

	/**
	 * Prevents the UI from making the column sortable.
	 */
	disableSorting?: boolean;

	constructor(name: string, displayName: string, dataType: string, order:number, isHidden?: boolean, isPrimary?: boolean) {
		this.name = name;
		this.displayName = displayName;
		this.dataType = dataType;
		this.order = order;
		this.isHidden = isHidden;
		this.isPrimary = isPrimary;
	}
}

export class SubmissionSortGrid implements ComponentFramework.StandardControl<IInputs, IOutputs> {

	// Cached context object for the latest updateView
	private contextObj: ComponentFramework.Context<IInputs>;
		
	// Div element created as part of this control's main container
	private mainContainer: HTMLDivElement;

	// Table element created as part of this control's table
	private gridContainer: HTMLDivElement;

	// Button element created as part of this control
	private loadPageButton: HTMLButtonElement;

	// The PO/PR Teams for the drop down list
	private POPRTeams: ComponentFramework.WebApi.Entity[];

	private submissionCategoryId:string;

	private callforSubmissionReference: EntityReference;

	private columnsOnView: DataSetInterfaces.Column[];

	private isActive:boolean;
	private lockOnInactive:boolean;

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

		this.callforSubmissionReference = new EntityReference(
			(<any>context).page.entityTypeName,
			(<any>context).page.entityId
		);

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

		let div:HTMLDivElement = document.createElement("div");
		div.className = "teamSelector";

		let lbl:HTMLLabelElement = document.createElement("label");
		lbl.textContent = "PO/PR Team: ";
		div.append(lbl);

		let ddl:HTMLSelectElement = document.createElement("select");
		ddl.id = "ddlTeam";
		ddl.className = "ddl";

		let opt:HTMLOptionElement = document.createElement("option");
		opt.value = "";
		opt.text = "-- Select PO/PR Team --";
		ddl.append(opt);
		div.append(ddl);

		this.mainContainer.append(div);

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
		this.toggleLoadMoreButtonWhenNeeded(context.parameters.SortableDataSet);

		if(!context.parameters.SortableDataSet.loading){
			this.isActive = true;
			this.lockOnInactive = false;

			// TODO: Need to run custom query to render PO/RD Team drop down
			// TODO: Need to run custom query to retrieve and render all project requests for selected team/call for submission

			// Get sorted columns on View
			this.columnsOnView = this.getSortedColumnsOnView(context);
			let stateColumnName:string = "";

			//context.parameters.SortableDataSet.records[0].getFormattedValue("");

			if (!this.columnsOnView || this.columnsOnView.length === 0) {
				return;
			}

			while(this.gridContainer.firstChild)
			{
				this.gridContainer.removeChild(this.gridContainer.firstChild);
			}

			if (this.POPRTeams == null)
				this.loadPOPRTeams(context);
			else
				this.renderPOPRTeamsDDL(context);

			//let projectsChanged:boolean = this.assignMissingDistrictPriorities(context, columnsOnView, context.parameters.SortableDataSet);
			//if (projectsChanged==true) context.parameters.SortableDataSet.refresh();

			try {
				let lockInactive:string = context.parameters.lockoninactivestate.raw as string;
				if (lockInactive.toLowerCase().indexOf("true") > -1) this.lockOnInactive = true;

				// Determine if the parent record state is inactice
				for (let column of this.columnsOnView){
					if(column.name.indexOf("statecode") > -1) {
						let state:string = context.parameters.SortableDataSet.records[context.parameters.SortableDataSet.sortedRecordIds[0]].getFormattedValue(column.name);

						if (state.toLowerCase().indexOf("inactive") > -1) this.isActive = false;
						
						break;
					}
				}
			}
			catch(e) {
				var s = e;
			}

			//let groupingFieldName:string = context.parameters.groupingfieldname.raw as string;
			//let gridHeader:HTMLDivElement = this.getGridHeader(columnsOnView, groupingFieldName);
			//this.mainContainer.insertBefore(gridHeader, this.gridContainer);

			let teamId:string = $(".ddl").val() as string;
			if(teamId != null && teamId != undefined && teamId.length > 0)
			{
				this.gridContainer.appendChild(this.createSubmissionHeaders(context, teamId));
			}

			this.createGridBody(context, this.columnsOnView, context.parameters.SortableDataSet);
		}
		// this is needed to ensure the scroll bar appears automatically when the grid resize happens and all the tiles are not visible on the screen.
		//this.mainContainer.style.maxHeight = window.innerHeight - this.gridContainer.offsetTop - 75 + "px";

		var binder = this.bindEvents.bind(this);
		binder();
	}

	//<fetch distinct="true" >
  //<entity name="caps_project" >
  //<attribute name="caps_ministrypriority" />
  //<attribute name="caps_ministryportfoliopriority" />
  //<attribute name="caps_districtpriority" />
  //<attribute name="statuscode" />
  //<attribute name="caps_projectid" />
  //<attribute name="caps_submissioncategory" />
  //<attribute name="caps_projectdescription" />
  //<attribute name="statecode" />
  //<link-entity name="caps_submission" from="caps_submissionid" to="caps_submission" link-type="inner" >
  //<filter>
	  //<condition attribute="caps_callforsubmission" operator="eq" value="13d6b3e3-b509-eb11-a813-000d3af42496" />
  //</filter>
  //</link-entity>
  //<link-entity name="edu_schooldistrict" from="edu_schooldistrictid" to="caps_schooldistrict" link-type="inner" >
  //<attribute name="owningteam" />
	//<filter>
	  //<condition attribute="owningteam" operator="eq" value="714c6301-f25f-ea11-a811-000d3af42c56" />
	//</filter>
  //</link-entity>
//</entity>
//</fetch>

	public createSubmissionHeaders(context: ComponentFramework.Context<IInputs>, teamId: string): HTMLDivElement
	{
		let categoriesDiv:HTMLDivElement = document.createElement("div");
		categoriesDiv.id = "divCategories";

		// this.callforSubmissionReference.id;

		let fetch:string = "?fetchXml=<fetch distinct=\"true\"><entity name=\"caps_project\"><attribute name=\"caps_submissioncategory\" /> " +
			"<link-entity name=\"caps_submission\" from=\"caps_submissionid\" to=\"caps_submission\" link-type=\"inner\"><filter>" +
			"<condition attribute=\"caps_callforsubmission\" operator=\"eq\" value=\"" + this.callforSubmissionReference.id + "\" />" +
			"</filter></link-entity><link-entity name=\"edu_schooldistrict\" from=\"edu_schooldistrictid\" to=\"caps_schooldistrict\" link-type=\"inner\">" +
			"<filter><condition attribute=\"owningteam\" operator=\"eq\" value=\"" + teamId + "\" /></filter></link-entity>" +
		  	"<link-entity name=\"caps_submissioncategory\" from=\"caps_submissioncategoryid\" to=\"caps_submissioncategory\" link-type=\"inner\">" +
			"<attribute name=\"caps_name\" /><order attribute=\"caps_name\" /></link-entity></entity></fetch>";
		var obj = this;

		context.webAPI.retrieveMultipleRecords("caps_project", fetch).then(
			function (response: ComponentFramework.WebApi.RetrieveMultipleResponse) {
				let isFirstLoad:boolean = false;
				let isFirstCat:boolean = true;
				let cssClass:string = "catLink";
				var categoriesDiv = $("#divCategories");
				for(var project of response.entities)
				{
					cssClass = "catLink";

					if (isFirstCat===true && (obj.submissionCategoryId === undefined || obj.submissionCategoryId === null))
					{
						isFirstLoad = true;
						obj.submissionCategoryId = project["_caps_submissioncategory_value"];
						cssClass = "catLink catLink_selected";
					}
					else if (obj.submissionCategoryId !== undefined && obj.submissionCategoryId !== null && obj.submissionCategoryId === project["_caps_submissioncategory_value"])
						cssClass = "catLink catLink_selected";

					categoriesDiv.append("<a href=\"javascript:;\" class=\"" + cssClass + "\" data-id=\"" + project["_caps_submissioncategory_value"] + "\">" + project["caps_submissioncategory3.caps_name"] + "</a>&nbsp;");

					isFirstCat = false;
				}

				$(".catLink")
					.on("click", function(event) {		
						event.preventDefault();
						$(".catLink").removeClass("catLink_selected");

						var link = $(this);
						link.addClass("catLink_selected");

						obj.submissionCategoryId = link.data("id");

						obj.updateView(obj.contextObj);
						return false;
					});

				if (isFirstLoad===true) obj.updateView(obj.contextObj);
			},
			function (errorResponse: any) 
			{
				//Error handling code here - record delete failed
				alert("An error occured while submission categories: " + errorResponse.message);
			}
		);

		return categoriesDiv;
	}

	public loadPOPRTeams(context: ComponentFramework.Context<IInputs>): void
	{
		let fetch:string = "?fetchXml=<fetch distinct=\"true\"><entity name=\"team\"><attribute name=\"teamid\" /><attribute name=\"name\" />" +
			"<filter><condition attribute=\"caps_category\" operator=\"eq\" value=\"200870000\" /></filter>" +
			"<order attribute=\"name\" /></entity></fetch>";
		var obj = this;

		context.webAPI.retrieveMultipleRecords("team", fetch).then(
			function (response: ComponentFramework.WebApi.RetrieveMultipleResponse) {
				obj.POPRTeams = response.entities;

				obj.renderPOPRTeamsDDL(obj.contextObj);
			},
			function (errorResponse: any) 
			{
				//Error handling code here - record delete failed
				alert("An error occured while retrieving PO/PR Teams: " + errorResponse.message);
			}
		);
	}

	public renderPOPRTeamsDDL(context: ComponentFramework.Context<IInputs>): void
	{
		var ddl = $("#ddlTeam");
 
		if (ddl.children().length == 1)
		{
			for (var team of this.POPRTeams)
			{
				let option:string = "<option value=\"" + team.teamid + "\">" + team.name + "</option>";

				ddl.append(option);
			}
		}
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
		let cols:DataSetInterfaces.Column[] = [];
		let col:DataSetInterfaces.Column = new Column("caps_projectid", "Project Request", "Guid", 0, true, true);
		cols.push(col);

		
		//col = new Column("caps_ministryprogrampriority", "Portfolio Rank", "Whole.Number", 1, false, false);
		col = new Column("caps_ministryportfoliopriority", "Portfolio Rank", "Whole.Number", 1, false, false);
		cols.push(col);

		col = new Column("caps_ministrypriority", "PO Category Rank", "Whole.Number", 2, false, false);
		cols.push(col);

		col = new Column("caps_districtpriority", "SD Category Rank", "Whole.Number", 3, false, false);
		cols.push(col);

		col = new Column("caps_ministrycomments", "MinistryComments", "SingleLine.Text", 4, false, false);
		cols.push(col);

		col = new Column("caps_projectcode", "Name", "SingleLine.Text", 5, false, false);
		cols.push(col);

		col = new Column("caps_projectdescription", "Description", "Multiple", 6, false, false);
		cols.push(col);

		col = new Column("caps_schooldistrict", "School District", "Lookup.Simple", 7, false, false);
		cols.push(col);

		col = new Column("statuscode", "Status", "Optionset", 8, false, false);
		cols.push(col);

		col = new Column("caps_submissioncategory", "Submission Category", "Lookup.Simple", 9, true, false);
		cols.push(col);

		if (!cols) { //context.parameters.SortableDataSet.columns) {
			return [];
		}
		
		let columns = cols //context.parameters.SortableDataSet.columns
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

	private assignMissingDistrictPriorities(context: ComponentFramework.Context<IInputs>, columnsOnView: DataSetInterfaces.Column[], gridParam: DataSet): boolean {
		let projectsChanged = false;
		let rankingFieldName:string = context.parameters.rankingfieldname.raw as string;
		let groupingFieldName:string = ""; //context.parameters.groupingfieldname.raw as string;
		
		if(gridParam.sortedRecordIds.length > 0)
		{
			let lastCategory:string = "";
			let count:number = 1;
			let maxPriority:number = 0;

			for(let currentRecordId of gridParam.sortedRecordIds){

				let record:DataSetInterfaces.EntityRecord = gridParam.records[currentRecordId];

				let category = record.getFormattedValue(groupingFieldName);

				// Render a header for each category
				if (category != lastCategory) {
					count = 1;
					maxPriority = 0;
					lastCategory = category;
				}

				var obj = this;

				columnsOnView.forEach(function(columnItem, index){
					if(columnItem.name.indexOf(rankingFieldName) > -1) {
						let priority:number = 0;
						
						if(gridParam.records[currentRecordId].getFormattedValue(columnItem.name) != null && gridParam.records[currentRecordId].getFormattedValue(columnItem.name) != "")
						{
							if(columnItem.name.indexOf(rankingFieldName) > -1) priority = gridParam.records[currentRecordId].getValue(columnItem.name) as number;
						}
						else
						{
							projectsChanged = true;

							priority = obj.getMaxPriority(gridParam, category, rankingFieldName, groupingFieldName) + 1;
							if (priority <= maxPriority) priority = maxPriority + 1;

							var data:any = {};
							data[rankingFieldName] = priority;

							try {
								let entityName:string = context.parameters.SortableDataSet.getTargetEntityType();
								context.webAPI.updateRecord(entityName, record.getRecordId(), data).then
								(
									function (response: ComponentFramework.EntityReference) 
									{ 
										// Callback method for successful update of record
					
										// Get the ID of the new record created
										//let id: string = response.id.guid;
									},
									function (errorResponse: any) 
									{
										// Error handling code here - record failed to be created
										alert("An error occured while saving the updated priority (1): " + errorResponse.toString());
									}
								);
							}
							catch(ex)
							{
								alert("An error occured while saving the updated priority (2): " + ex.toString());
							}
						}

						if (maxPriority < priority) maxPriority = priority;
					}
				});

				count++;
			}
		}

		return projectsChanged;
	}

	private createSortContainer(isRanked:boolean): HTMLDivElement
	{
		let sortContainer:HTMLDivElement = document.createElement("div");

		if (isRanked === true) sortContainer.id = "sortRanked";
		else sortContainer.id = "sortUnranked";

		if (this.isActive || !this.lockOnInactive) sortContainer.classList.add("can-sort");

		let div:HTMLDivElement = document.createElement("div");
		div.className = "dropHereToRank";

		if (isRanked===true) {
			div.id = "dropHereToRank";
			div.innerHTML = "Drop here to rank.";
		}
		else {
			div.id = "dropHereToUnrank";
			div.innerHTML = "Drop here to unrank.";
		}

		sortContainer.appendChild(div);

		return sortContainer;
	}
	
	/**
	 * funtion that creates the body of the grid and embeds the content onto the tiles.
	 * @param columnsOnView columns on the view whose value needs to be shown on the UI
	 * @param gridParam data of the Grid
	 */
	private createGridBody(context: ComponentFramework.Context<IInputs>, columnsOnView: DataSetInterfaces.Column[], gridParam: DataSet):void{
		let gridBody:HTMLDivElement = document.createElement("div");
		let rankingFieldName:string = context.parameters.rankingfieldname.raw as string;
		let groupingFieldName:string = context.parameters.groupingfieldname.raw as string;
		let linkfieldname:string = context.parameters.linkfieldname.raw as string;
		let entitydisplayname:string = context.parameters.entitydisplayname.raw as string;

		let teamId:string = $(".ddl").val() as string;
		let submissionCategoryId:string = this.submissionCategoryId;

		//let gridScroll:HTMLDivElement = document.createElement("div");
		//gridScroll.classList.add("DataSetControl_grid-dataScroll");

		let gridHeader:HTMLDivElement = this.getGridHeader(columnsOnView, groupingFieldName);
		gridBody.appendChild(gridHeader);

		if(teamId !== undefined && teamId !== null && teamId.length > 0 &&
			submissionCategoryId !== undefined && submissionCategoryId !== null && submissionCategoryId.length > 0 &&
			gridParam.sortedRecordIds.length > 0)
		{
			let fetch:string = "?fetchXml=<fetch distinct=\"true\"><entity name=\"caps_project\"><attribute name=\"caps_projectid\" />" +
			"<attribute name=\"caps_ministryportfoliopriority\" /><attribute name=\"caps_ministrypriority\" />" +
			"<attribute name=\"caps_districtpriority\" /><attribute name=\"caps_projectcode\" />" +
			"<attribute name=\"caps_projectdescription\" /><attribute name=\"statuscode\" /><attribute name=\"statecode\" />" +
			"<attribute name=\"caps_submissioncategory\" /><attribute name=\"caps_schooldistrict\" />" +
			"<attribute name=\"caps_ministrycomments\" />" +
			"<filter><condition attribute=\"caps_submissioncategory\" operator=\"eq\" value=\"" + submissionCategoryId + "\" /></filter>" +
			"<link-entity name=\"caps_submission\" from=\"caps_submissionid\" to=\"caps_submission\" link-type=\"inner\"><filter>" +
			"<condition attribute=\"caps_callforsubmission\" operator=\"eq\" value=\"" + this.callforSubmissionReference.id + "\" />" +
			"</filter></link-entity><link-entity name=\"edu_schooldistrict\" from=\"edu_schooldistrictid\" to=\"caps_schooldistrict\" link-type=\"inner\">" +
			"<filter><condition attribute=\"owningteam\" operator=\"eq\" value=\"" + teamId + "\" /></filter></link-entity>" +
		  	"<link-entity name=\"caps_submissioncategory\" from=\"caps_submissioncategoryid\" to=\"caps_submissioncategory\" link-type=\"inner\">" +
			"<attribute name=\"caps_name\" /><order attribute=\"caps_name\" /></link-entity><order attribute=\"caps_ministryportfoliopriority\" />" +
			"<order attribute=\"caps_ministrypriority\" /><order attribute=\"caps_districtpriority\" /></entity></fetch>";
			var obj = this;

			context.webAPI.retrieveMultipleRecords("caps_project", fetch).then(
				function (response: ComponentFramework.WebApi.RetrieveMultipleResponse) {
					if (response.entities !== undefined && response.entities !== null && response.entities.length > 0)
					{
						let isRanked:boolean = true;
						let spacer:HTMLDivElement = document.createElement("div");
						spacer.style.height = "20px";
						gridBody.appendChild(spacer);
			
						let lastCategory:string = "";
						let categoryContainer:HTMLDivElement = document.createElement("div");
			
						let columnHeader:HTMLDivElement = obj.getColumnHeaders(columnsOnView, groupingFieldName);

						let rankedContainer:HTMLDivElement = obj.createSortContainer(true);
						let unrankedContainer:HTMLDivElement = obj.createSortContainer(false);
			
						//for(let currentRecordId of gridParam.sortedRecordIds){
						for(let record of response.entities){
							let programPriority:number = record[rankingFieldName];

							if (programPriority !== undefined && programPriority !== null && programPriority > 0)
								isRanked = true;
							else
								isRanked = false;
							
							let gridRecord: HTMLDivElement = document.createElement("div");
							gridRecord.classList.add("DataSetControl_grid-dataRow");
			
							// Set the recordId on the row dom
							gridRecord.setAttribute(RowRecordId, record["caps_projectid"]);
							let linkValue:string = record[linkfieldname];
							gridRecord.setAttribute(linkfieldname, linkValue);
			
							//var obj = this;
							let width:number = 0;
			
							columnsOnView.forEach(function(columnItem, index){
								if (!columnItem.isHidden &&  columnItem.name.indexOf(groupingFieldName) == -1) {
									let valuePara:HTMLDivElement = document.createElement("div");
									valuePara.classList.add("DataSetControl_grid-dataCol");
									let priority:number = 0;
									width += 1;

									var colName = columnItem.name;
									if (colName === "caps_schooldistrict") colName = "_caps_schooldistrict_value"
									
									if(record[colName] != null && record[colName] != "")
									{
										valuePara.textContent = "";

										if (colName == "caps_totalprojectcost") valuePara.textContent = "$" + record[colName];
										else if (colName === "statuscode") valuePara.textContent += record["statuscode@OData.Community.Display.V1.FormattedValue"];
										else if (colName === "_caps_schooldistrict_value") valuePara.textContent += record["_caps_schooldistrict_value@OData.Community.Display.V1.FormattedValue"];
										else valuePara.textContent += record[colName];

										if(columnItem.name.indexOf(rankingFieldName) > -1) priority = record[colName] as number;
									}
									else
									{
										valuePara.textContent = "-";	
									}
			
									if(columnItem.name.indexOf(rankingFieldName) > -1) {
										let attr:Attr = document.createAttribute("original_order");
										attr.value = priority.toString();
										valuePara.attributes.setNamedItem(attr);
									}

									let attr:Attr = document.createAttribute("title");
									attr.value = valuePara.textContent;
									valuePara.attributes.setNamedItem(attr);
			
									if (columnItem.name == "caps_projectcode") valuePara.classList.add("click-to-open");
			
									gridRecord.appendChild(valuePara);
								}
							});
			
							let allowUnlinkRecord:string = "false"; //this.contextObj.parameters.allowunlinkrecord.raw as string;
							let allowUnlink:boolean = false;
							if (allowUnlinkRecord != null && allowUnlinkRecord.toLowerCase().indexOf("true") > -1) allowUnlink = true;
			
							let valuePara:HTMLDivElement = document.createElement("div");
							valuePara.classList.add("DataSetControl_grid-dataColDel");
							if (allowUnlink && (obj.isActive || !obj.lockOnInactive)) {
								let deleteLink:HTMLAnchorElement = document.createElement("a");
								deleteLink.classList.add("delete-button");
								deleteLink.href = "";
								deleteLink.innerHTML = "<i class=\"fas fa-unlink\"></i>";
								deleteLink.title = "Remove " + entitydisplayname + " " + linkValue;
								valuePara.appendChild(deleteLink);
							}
			
							gridRecord.appendChild(valuePara);
							
							if (isRanked === true)
								rankedContainer.appendChild(gridRecord);
							else
								unrankedContainer.appendChild(gridRecord);

							width = (width * 160) + 100;
							gridBody.style.width = width.toString() + "px";
						}

						gridBody.appendChild(obj.getCategoryHeader("Ranked", columnHeader));
						gridBody.appendChild(rankedContainer);

						gridBody.appendChild(obj.getCategoryHeader("Unranked", columnHeader));
						gridBody.appendChild(unrankedContainer);
					}
					else
					{
						let noRecordLabel: HTMLDivElement = document.createElement("div");
						noRecordLabel.classList.add("DataSetControl_grid-norecords");
						noRecordLabel.style.width = obj.contextObj.mode.allocatedWidth - 25 + "px";
						noRecordLabel.innerHTML = obj.contextObj.resources.getString("PCF_DataSetControl_No_Record_Found");
						gridBody.appendChild(noRecordLabel);
					}

					obj.gridContainer.appendChild(gridBody);

					var dragDrop = obj.bindDragDrop.bind(obj);
					dragDrop();
				},
				function (errorResponse: any) 
				{
					//Error handling code here
					alert("An error occured while retrieving projects: " + errorResponse.message);
				}
			);
		}
	}

	private getMaxPriority(gridParam: DataSet, category: string, rankingFieldName: string, groupingFieldName: string): number {
		let maxPriority:number = 0;
		let priority:number = 1;

		for(let currentRecordId of gridParam.sortedRecordIds){
			if (gridParam.records[currentRecordId].getFormattedValue(groupingFieldName) == category) {
				if (gridParam.records[currentRecordId].getFormattedValue(rankingFieldName) != null && gridParam.records[currentRecordId].getFormattedValue(rankingFieldName) != "") {
					priority = gridParam.records[currentRecordId].getValue(rankingFieldName) as number;
					if (priority > maxPriority) maxPriority = priority;
				}
			}
		}

		return maxPriority;
	}

	private getCategoryHeader(categoryName:string, columnHeader:HTMLDivElement):HTMLDivElement {
		let categoryHeader:HTMLDivElement = document.createElement("div");
		categoryHeader.classList.add("DataSetControl_category-header");

		let container:HTMLDivElement = document.createElement("div");
		container.classList.add("DataSetControl_category-title");

		//let plusIcon = document.createElement("i"); //<i class="far fa-plus-square"></i> <i class="far fa-minus-square"></i>
		//plusIcon.classList.add("far", "fa-minus-square", "icon_collapser");
		//container.appendChild(plusIcon);

		//let input:HTMLInputElement = document.createElement("input") as HTMLInputElement;
		//input.type = "checkbox";
		//input.checked = true;
		//input.title = categoryName;
		//input.classList.add("DataSetControl_collapser");
		//container.appendChild(input);

		let span:HTMLSpanElement = document.createElement("span");
		span.innerHTML = "&nbsp;" + categoryName;
		container.appendChild(span);

		categoryHeader.appendChild(container);

		container = document.createElement("div");
		container.innerHTML = columnHeader.innerHTML;
		categoryHeader.appendChild(container);

		return categoryHeader;
	}

	private getGridHeader(columnsOnView: DataSetInterfaces.Column[], groupingFieldName:string):HTMLDivElement {
		let gridHeader:HTMLDivElement = document.createElement("div");
		gridHeader.classList.add("DataSetControl_grid-FullHeader");

		//if (groupingFieldName != null && groupingFieldName.length > 0) {
		//	let gridShowHide:HTMLDivElement = document.createElement("div");
		//	gridShowHide.classList.add("DataSetControl_grid-showhide");

		//	let input:HTMLInputElement = document.createElement("input");
		//	input.checked = true;
		//	input.classList.add("show-hide-check");
		//	gridShowHide.appendChild(input);

		//	let showHideLink:HTMLAnchorElement = document.createElement("a");
		//	showHideLink.classList.add("show-hide-all");
		//	showHideLink.text = "Expand/Collapse all Categories";
		//	showHideLink.href = "#";
		//	gridShowHide.appendChild(showHideLink);
			
		//	gridHeader.appendChild(gridShowHide);
		//}

		return gridHeader;
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

		let allowUnlinkRecord:string = "false"; //this.contextObj.parameters.allowunlinkrecord.raw as string;
		let allowUnlink:boolean = false;
		if (allowUnlinkRecord != null && allowUnlinkRecord.toLowerCase().indexOf("true") > -1) allowUnlink = true;

		let labelPara:HTMLDivElement = document.createElement("div");
		labelPara.classList.add("DataSetControl_grid-headerColDel");
		if (allowUnlink == true) labelPara.textContent = "Remove";
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
			//let entityReference = this.contextObj.parameters.SortableDataSet.records[rowRecordId].getNamedReference();
			let entityFormOptions = {
				entityName: "caps_project",
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
		this.contextObj.parameters.SortableDataSet.paging.loadNextPage();
		this.toggleLoadMoreButtonWhenNeeded(this.contextObj.parameters.SortableDataSet);
	}

	private bindEvents() {
		var obj = this;

		$(".ddl")
		.off("change")
		.on("change", function (event){
			let teamId:string = $(".ddl").val() as string;
			obj.updateView(obj.contextObj);
		});
	}

	private bindDragDrop() {
		try{
			$("#sortRanked, #sortUnranked")
				.sortable({placeholder: "ui-state-highlight",
				connectWith: ".can-sort",
				start: function sortStart(event, ui) {
						ui.item.data("start-pos", ui.item.index());
					},
				update: this.sortUpdate.bind(this)
			}).disableSelection();

			$(".click-to-open").on("click", this.onRowClick.bind(this));
		}
		catch (e) {
			var s = "";
		}

		var obj = this;

		//$(".DataSetControl_collapser").on("change", function (event){
		//	let input:HTMLInputElement = event.target as HTMLInputElement;
		//	var icon = $(input).siblings("i");
			
		//	if (input.checked) {
		//		$(event.target).parent().parent().next().show();
		//		$(event.target).parent().next().show();
		//		icon.addClass("fa-minus-square");
		//		icon.removeClass("fa-plus-square");
		//	}
		//	else {
		//		$(event.target).parent().parent().next().hide();
		//		$(event.target).parent().next().hide();
		//		icon.addClass("fa-plus-square");
		//		icon.removeClass("fa-minus-square");
		//	}
		//});

		//$(".icon_collapser").on("click", function(event) {
		//	var checkbox = $(this).siblings("input");
			//if (checkbox.prop("checked") == true) {
				//$(this).addClass("fa-plus-square");
				//$(this).removeClass("fa-minus-square");
			//}
			//else {
				//$(this).addClass("fa-minus-square");
				//$(this).removeClass("fa-plus-square");
			//}
		//	checkbox.click();
		//});

		//$(".show-hide-all").on("click", function(event){
		//	event.preventDefault();
		//	let checked:boolean = !$(".show-hide-check").prop("checked");

		//	$(".show-hide-check").prop("checked", checked);
		//	$(".DataSetControl_collapser").each(function(){
		//		var input = $(this);
		//		if (input.prop("checked") != checked) input.click();
		//	});

		//	return false;
		//});

		//$(".delete-button").on("click", function(event) {
		//	event.preventDefault();

		//	var row = $(this).parent().parent();
		//	let linkField:string = row.attr(obj.contextObj.parameters.linkfieldname.raw as string)!;
		//	let entityDisplayName:string = obj.contextObj.parameters.entitydisplayname.raw as string;
		//	let parentEntityDisplayName:string = ""; //obj.contextObj.parameters.parententitydisplayname.raw as string;

		//	obj.contextObj.navigation.openConfirmDialog({title:"Remove " + entityDisplayName + " " + linkField + "?", text:"Are you sure you want to remove " + entityDisplayName + " " + linkField + " from this " + parentEntityDisplayName + "?"}, {}).then(
		//		function(success) {
		//			if (success.confirmed) {
		//				let entityName:string = obj.contextObj.parameters.SortableDataSet.getTargetEntityType();
		//				let parentEntityFieldName:string = ""; //obj.contextObj.parameters.parententityfieldname.raw as string;
		//				let recordId:string = row.attr(RowRecordId)!;

		//				var data:any = {};
		//				data["_" + parentEntityFieldName + "_value"] = null;

						//obj.contextObj.webAPI.deleteRecord(entityName, recordId).then(
		//				obj.contextObj.webAPI.updateRecord(entityName, recordId, data).then(
		//					function (response: ComponentFramework.EntityReference) {
		//						obj.contextObj.parameters.SortableDataSet.refresh();
		//						var group = row.parent();
		//						row.remove();
		//						var items = group.children();
		//						obj.updateItem(items, 0);
		//					},
		//					function (errorResponse: any) 
		//					{
								//Error handling code here - record delete failed
		//						alert("An error occured while removing the record: " + errorResponse.toString());
		//					}
		//				);
		//			}
		//		}
		//	);

		//	return false;
		//});

		//$(document).ready(function(){
		//	var container = $(".DataSetControl_main-container").closest(".webkitScroll");
		//	let maxY:number = container[0].offsetHeight;

		//	$(".DataSetControl_main-container").css("height", (maxY - 250));
		//	$(".DataSetControl_grid-dataScroll").css("height", (maxY - 340));
		//});
	}

	private updateItem(items:JQuery<HTMLElement>, index:number, level:number):void {		
		let isRanked:boolean = true;

		if (items === null || items === undefined || items.length === 0) {
			this.contextObj.parameters.SortableDataSet.refresh();
			return;
		}

		var parent = $(items[0]).parent();
		if (parent[0].id==="sortRanked") {
			isRanked = true;
			items.remove("#dropHereToRank");
			items = items.not("#dropHereToRank");
		} 
		else 
		{
			isRanked = false;
			items.remove("#dropHereToUnrank");
			items = items.not("#dropHereToUnrank");
		}

		console.log("updateItem(), index:" + index + ", level:" + level + ", isRanked:" + isRanked);

		if (items.length === 0 && level===0)
		{
			index = 0;
			level = 1;
			items = $("#sortUnranked").children();

			isRanked = false;
			items.remove("#dropHereToUnrank");
			items = items.not("#dropHereToUnrank");
		}

		let entityName:string = "caps_project";

		let recordId:string = items[index].getAttribute(RowRecordId) as string;

		var data:any = {};
		
		if (isRanked === true)
			data[this.contextObj.parameters.rankingfieldname.raw as string] = index + 1;
		else
			data[this.contextObj.parameters.rankingfieldname.raw as string] = null;

		var item = $(items[index]);
		item.children().first().text(index + 1);
		item.fadeTo(150, .3);

		var obj = this;
	
		try {
			this.contextObj.webAPI.updateRecord(entityName, recordId, data).then
			(
				function (response: ComponentFramework.EntityReference) {
					item.fadeTo(150, 1);

					if (level===0 && index === (items.length - 1)) {
						index = 0;
						level = 1;
						items = $("#sortUnranked").children();
					}
					else if (level===1 && index === (items.length - 1)) {
						obj.contextObj.parameters.SortableDataSet.refresh();
						return;
					}
					else index++;
					
					obj.updateItem(items, index, level);

					/*
					if (index < (items.length - 1) && level===0) { 
						obj.updateItem(items, index+1, 0);
					}
					else {
						if (items.length > 0 && level===0) {
							if (isRanked === true) items = $("#sortUnranked").children();
							else items = $("#sortRanked").children();

							obj.updateItem(items, 0, 1);
						}
						else {
							obj.isSorting = false;
							obj.contextObj.parameters.SortableDataSet.refresh();
						}
					}
					*/
				},
				function (errorResponse: any) 
				{
					// Error handling code here - record failed to be created
					alert("An error occured while saving the updated priority (1): " + errorResponse.toString());
				}
			);
		}
		catch(ex)
		{
			alert("An error occured while saving the updated priority (2): " + ex.toString());
		}
	}

	private sortUpdate(event:Event, ui:JQueryUI.SortableUIParams):void {
		var target = $(event.target as any);

		$("#sortRanked, #sortUnranked").sortable( "option", "disabled", false );
		$("#sortRanked, #sortUnranked").disableSelection();

		let start_pos:number = ui.item.data("start-pos");
		let end_pos:number = ui.item.index();

		var items = $("#sortRanked").children(); //ui.item.parent().children();

		this.updateItem(items, 0, 0);
	}
}