import {IInputs, IOutputs} from "./generated/ManifestTypes";
import DataSetInterfaces = ComponentFramework.PropertyHelper.DataSetApi;
import { array } from "prop-types";
type DataSet = ComponentFramework.PropertyTypes.DataSet;
const RowRecordId:string = "rowRecId";
let readOnly: string[] ;
readOnly= ['createdby', 'createdonbehalfby', 'createdbyexternalparty','createdon','processid','statecode','statuscode'];

export class ActivitySummary implements ComponentFramework.StandardControl<IInputs, IOutputs> {

	private contextObj: ComponentFramework.Context<IInputs>;
		
	// Div element created as part of this control's main container
	private mainContainer: HTMLDivElement;

	// Table element created as part of this control's table
	private dataTable: HTMLTableElement;

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

		// Create main table container div. 
		this.mainContainer = document.createElement("div");
		this.mainContainer.classList.add("SimpleTable_MainContainer_Style");

		// Create data table container div. 
		this.dataTable = document.createElement("table");
		this.dataTable.classList.add("SimpleTable_Table_Style");

	
	
		this.mainContainer.appendChild(this.dataTable);

		container.appendChild(this.mainContainer);
	}


	/**
	 * Called when any value in the property bag has changed. This includes field values, data-sets, global values such as container height and width, offline status, control metadata values such as label, visible, etc.
	 * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to names defined in the manifest, as well as utility functions
	 */
	public updateView(context: ComponentFramework.Context<IInputs>): void
	{
		this.contextObj = context;
	

			if(!context.parameters.smartGridDataSet.loading){
				
				// Get sorted columns on View
				let columnsOnView = this.getSortedColumnsOnView(context);

				if (!columnsOnView || columnsOnView.length === 0) {
                    return;
				}

				let columnWidthDistribution = this.getColumnWidthDistribution(context, columnsOnView);


				while(this.dataTable.firstChild)
				{
					this.dataTable.removeChild(this.dataTable.firstChild);
				}
				let isTypeColumn:boolean=false
				columnsOnView.forEach(function (columnItem) {
				if(columnItem.name=="activitytypecode")
				{
					isTypeColumn=true;
				}
				});
				if(isTypeColumn)
				{
				this.dataTable.appendChild(this.createTableHeader(columnsOnView, columnWidthDistribution));		
				this.dataTable.appendChild(this.createTableBody(columnsOnView, columnWidthDistribution, context.parameters.smartGridDataSet));
				this.dataTable.parentElement!.style.height = window.innerHeight - this.dataTable.offsetTop - 70 + "px";
				this.updateCount( context.parameters.smartGridDataSet);
				}
				else
				{
					this.mainContainer.innerHTML="Activity Type Column Missing"
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
			if (!context.parameters.smartGridDataSet.columns) {
				return [];
			}
			
			let columns =context.parameters.smartGridDataSet.columns
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
		 * Get column width distribution
		 * @param context context object of this cycle
		 * @param columnsOnView columns array on the configured view
		 * @returns column width distribution
		 */
		private getColumnWidthDistribution(context: ComponentFramework.Context<IInputs>, columnsOnView: DataSetInterfaces.Column[]): string[]{

			let widthDistribution: string[] = [];
			
			// Considering need to remove border & padding length
			let totalWidth:number = context.mode.allocatedWidth - 250;
			let widthSum = 0;
			
			columnsOnView.forEach(function (columnItem) {
				widthSum += columnItem.visualSizeFactor;
			});

			let remainWidth:number = totalWidth;
			
			columnsOnView.forEach(function (item, index) {
				let widthPerCell = "";
				if (index !== columnsOnView.length - 1) {
					let cellWidth = Math.round((item.visualSizeFactor / widthSum) * totalWidth);
					remainWidth = remainWidth - cellWidth;
					widthPerCell = cellWidth + "px";
				}
				else {
					widthPerCell = remainWidth + "px";
				}
				widthDistribution.push(widthPerCell);
			});

			return widthDistribution;

		}

		private createTableHeader(columnsOnView: DataSetInterfaces.Column[], widthDistribution: string[]):HTMLTableSectionElement{

			let tableHeader:HTMLTableSectionElement = document.createElement("thead");
			let tableHeaderRow: HTMLTableRowElement = document.createElement("tr");
			tableHeaderRow.classList.add("SimpleTable_TableRow_Style");

			columnsOnView.forEach(function(columnItem, index){
				if(columnItem.name=="activitytypecode")
				{
				
				let tableHeaderCell = document.createElement("th");
				tableHeaderCell.id="typeHeader";
				tableHeaderCell.classList.add("SimpleTable_TableHeader_Style");
				let innerDiv = document.createElement("div");
				innerDiv.classList.add("SimpleTable_TableCellInnerDiv_Style");
				innerDiv.style.maxWidth = widthDistribution[200];
				innerDiv.innerText = columnItem.displayName;
				tableHeaderCell.appendChild(innerDiv);
				tableHeaderRow.appendChild(tableHeaderCell);

				let tableHeaderCellount = document.createElement("th");
				tableHeaderCellount.id="countHeader";
				tableHeaderCellount.classList.add("SimpleTable_TableHeader_Style");
				let innerDivCounts = document.createElement("div");
				innerDivCounts.classList.add("SimpleTable_TableCellInnerDiv_Style");
				innerDivCounts.style.maxWidth = widthDistribution[200];
				innerDivCounts.innerText = "Count";
				tableHeaderCellount.appendChild(innerDivCounts);
				tableHeaderRow.appendChild(tableHeaderCellount);
				
				}
			});
		


			tableHeader.appendChild(tableHeaderRow);
			return tableHeader;
		}




//
		private createTableBody(columnsOnView: DataSetInterfaces.Column[], widthDistribution: string[], gridParam: DataSet):HTMLTableSectionElement{

			let tableBody:HTMLTableSectionElement = document.createElement("tbody");

			if(gridParam.sortedRecordIds.length > 0)
			{
				let activityTypeName: string[];
				activityTypeName=[''];
				for(let currentRecordId of gridParam.sortedRecordIds){
					
					

						
						{
						let activityCode:string=gridParam.records[currentRecordId].getValue("activitytypecode").toString();
						if(!(activityTypeName.includes(activityCode)))
						{
							let tableRecordRow: HTMLTableRowElement = document.createElement("tr");
							tableRecordRow.classList.add("SimpleTable_TableRow_Style");
					

								activityTypeName.push(activityCode);
								tableRecordRow.id=activityCode;
								let tableRecordCell = document.createElement("td");
								tableRecordCell.classList.add("SimpleTable_TableCell_Style");
								let innerDiv = document.createElement("div");
								innerDiv.classList.add("SimpleTable_TableCellInnerDiv_Style");
								innerDiv.style.maxWidth = widthDistribution[200];
								innerDiv.innerText = gridParam.records[currentRecordId].getFormattedValue("activitytypecode");
								tableRecordCell.appendChild(innerDiv);
								tableRecordRow.appendChild(tableRecordCell);

								let tableRecordCellCount = document.createElement("td");
								tableRecordCellCount.classList.add("SimpleTable_TableCell_Style");
								let innerDivCount = document.createElement("div");
								innerDivCount.classList.add("SimpleTable_TableCellInnerDiv_Style");
								innerDivCount.style.maxWidth = widthDistribution[200];
								innerDivCount.id="count_"+activityCode;
								innerDivCount.innerText = "0";
								tableRecordCellCount.appendChild(innerDivCount);
								tableRecordRow.appendChild(tableRecordCellCount);
								tableBody.appendChild(tableRecordRow);
						}
						}
					

					
				}
				



			}
			else
			{
				let tableRecordRow: HTMLTableRowElement = document.createElement("tr");
				let tableRecordCell: HTMLTableCellElement = document.createElement("td");
				tableRecordCell.classList.add("No_Record_Style");
				tableRecordCell.colSpan = columnsOnView.length;
				tableRecordCell.innerText = this.contextObj.resources.getString("PCF_TSTableGrid_No_Record_Found");
				tableRecordRow.appendChild(tableRecordCell)
				tableBody.appendChild(tableRecordRow);
			}

			return tableBody;
		}


	private updateCount( gridParam: DataSet):void{

		for(let currentRecordIdCount of gridParam.sortedRecordIds){

			let activityCodeVal:string=gridParam.records[currentRecordIdCount].getValue("activitytypecode").toString();
			let fieldID:string="count_"+activityCodeVal;
			var currentCountElem=document.getElementById(fieldID) as HTMLDivElement;
			let currentCount=Number(currentCountElem.innerText);
			currentCountElem.innerText=(currentCount+1).toString();

		}
	}

}