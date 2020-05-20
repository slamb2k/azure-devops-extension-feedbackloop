import "./Pivot.scss";

import * as React from "react";
import * as SDK from "azure-devops-extension-sdk";

import { showRootComponent } from "../../Common";

import { getClient, IProjectPageService, CommonServiceIds, INavigationElement, IPageRoute, IHostNavigationService } from "azure-devops-extension-api";
import { TeamContext } from "azure-devops-extension-api/Core";

import { WorkRestClient, TimeFrame, BacklogType } from "azure-devops-extension-api/Work";
import { WorkItemTrackingRestClient, WorkItem, Wiql } from "azure-devops-extension-api/WorkItemTracking";

import { Table, ITableColumn, renderSimpleCell, renderSimpleCellValue } from "azure-devops-ui/Table";
import { ArrayItemProvider } from "azure-devops-ui/Utilities/Provider";
import * as _ from "lodash";

interface IPivotContentState {
    userName?: string;
    projectName?: string;
    teamName?: string;
    iframeUrl?: string;
    extensionData?: string;
    extensionContext?: SDK.IExtensionContext;
    host?: SDK.IHostContext;
    navElements?: INavigationElement[];
    route?: IPageRoute;
    columns: ITableColumn<any>[];
    scopeChanges?: ArrayItemProvider<ScopeChangeWorkItem>;
}

enum ScopeChangeType {
    CreatedInSprint = 1,
    Deleted,
    AddedToSprint,
    RemovedFromSprint,
    StoryPointsIncreased,
    StoryPointsDecreased,
}

class ScopeChangeWorkItem implements WorkItem {
    commentVersionRef!: import("azure-devops-extension-api/WorkItemTracking").WorkItemCommentVersionRef;
    fields!: {
        [key: string]: any;
    };
    id!: number;
    relations!: import("azure-devops-extension-api/WorkItemTracking").WorkItemRelation[];
    rev!: number;
    _links: any;
    url!: string;

    scopeChangeType!: ScopeChangeType;
}

class PivotContent extends React.Component<{}, IPivotContentState> {

    constructor(props: {}) {
        super(props);

        this.state = {
            iframeUrl: window.location.href,
            columns: [{
                id: "id",
                name: "Work Item Id",
                renderCell: renderSimpleCell,
                width: 100
            },
            {
                id: "title",
                name: "Title",
                renderCell: (rowIndex: number, columnIndex: number, tableColumn: ITableColumn<WorkItem>, tableItem: WorkItem): JSX.Element => {
                    return renderSimpleCellValue<any>(columnIndex, tableColumn, tableItem.fields.title);
                },
                width: 400
            },
            {
                id: "changeddate",
                name: "Changed Date",
                renderCell: (rowIndex: number, columnIndex: number, tableColumn: ITableColumn<WorkItem>, tableItem: WorkItem): JSX.Element => {
                    return renderSimpleCellValue<any>(columnIndex, tableColumn, tableItem.fields["System.ChangedDate"]);
                },
                width: 100
            },
            {
                id: "changedby",
                name: "Changed By",
                renderCell: (rowIndex: number, columnIndex: number, tableColumn: ITableColumn<WorkItem>, tableItem: WorkItem): JSX.Element => {
                    return renderSimpleCellValue<any>(columnIndex, tableColumn, tableItem.fields["System.ChangedBy"]);
                },
                width: 100
            },
            {
                id: "scopechangetype",
                name: "Change Type",
                renderCell: (rowIndex: number, columnIndex: number, tableColumn: ITableColumn<ScopeChangeWorkItem>, tableItem: ScopeChangeWorkItem): JSX.Element => {
                    return renderSimpleCellValue<any>(columnIndex, tableColumn, ScopeChangeType[tableItem.scopeChangeType]);
                },
                width: 100
            }]
        };
    }

    public componentDidMount() {
        SDK.init();
        //this.initializeComponent();

        console.debug("ININTIALIZING STATE");
        this.initializeState();
    }

    private async initializeState(): Promise<void> {
        await SDK.ready();
        
        const userName = SDK.getUser().displayName;
        this.setState({
            userName,
            extensionContext: SDK.getExtensionContext(),
            host: SDK.getHost()
         });

        const navService = await SDK.getService<IHostNavigationService>(CommonServiceIds.HostNavigationService);
        const navElements = await navService.getPageNavigationElements();
        this.setState({ navElements });

        console.debug("GOT NAV ITEMS");

        const route = await navService.getPageRoute();        
        const projectName = route.routeValues.project;
        const teamName = route.routeValues.teamName;
        this.setState({ 
                route: route,
                projectName: projectName,
                teamName: teamName
             });

        console.debug("PAGE ROUTE, PROJECT AND TEAM");

        if (!this.state.projectName || !this.state.teamName)
            throw ("Missing project and team. Exiting...");

        const teamContext: TeamContext = { projectId: "", teamId: "", project: this.state.projectName, team: this.state.teamName };

        console.debug("TEAM CONTEXT CONFIGURED");

        // Get the current iteration for the specified team    
        const iterationPath = route.routeValues.iteration;    
        const iterations = await getClient(WorkRestClient).getTeamIterations(teamContext);
        
        if (!iterations)
            throw("Unable to retrieve iterations.");

        console.debug(`RETRIEVED ${ iterations.length } ITERATIONS`);

        // Find the iteration by name and return the object so we have the id
        const selectedIteration = _.find(iterations, { 'path': iterationPath.replace("/", "\\") });

        if (!selectedIteration)
            throw("Unable to retrieve selected iteration.");        
        
        console.debug(`SELECTED ITERATION: ${ selectedIteration.name }`);

        // Get the backlog level configurations for the specified team    
        const backlogs = await getClient(WorkRestClient).getBacklogs(teamContext);

        if (!backlogs)
            throw("No backlog configurations found.");
        
        console.debug(`RETRIEVED ${ backlogs.length } BACKLOG CONFIGURATIONS`);

        // Get the requirement back log
        const requirementBacklog = _.find(backlogs, { 'type': BacklogType.Requirement });

        if (!requirementBacklog)
            throw("No Requirement Backlog");

        console.debug(`SELECTED BACKLOG: ${ requirementBacklog.name }`);

        // Create a filter for a wiql query
        const backlogTypesFilter = _.map(requirementBacklog.workItemTypes, wit => { return `'${ wit.name }'` }).join(",");

        // Get the end of the first day of the iteration
        const plannedDate = selectedIteration.attributes.startDate;
        plannedDate.setHours(23,59,59,999);

        const plannedWorkWiql = {
            query: `Select [System.Id], [System.Title], [System.State]
                    from WorkItems
                    where [System.IterationPath] = '${ selectedIteration.path }' and
                          [System.WorkItemType] in (${ backlogTypesFilter })
                    asof '${ plannedDate.toISOString() }'
                    order by [System.Id] asc` };

        console.debug(`PLANNED WORK WIQL: ${ plannedWorkWiql.query }`);

        const plannedWorkItemRefs = await getClient(WorkItemTrackingRestClient).queryByWiql(plannedWorkWiql, this.state.projectName);
        const plannedWorkItemIds = plannedWorkItemRefs.workItems.map(wi => wi.id);

        const plannedWorkItemsCompare = [] as WorkItem[];

        if (plannedWorkItemIds.length > 0)
        {
            const plannedWorkItems = await getClient(WorkItemTrackingRestClient).getWorkItems(plannedWorkItemIds);

            console.debug(`RETRIEVED ${ plannedWorkItems.length } PLANNED WORKITEMS`);

            _.merge(plannedWorkItemsCompare, plannedWorkItems);
        }



        const currentWorkWiql = {
            query: `Select [System.Id], [System.Title], [System.State]
                    from WorkItems
                    where [System.IterationPath] = '${ selectedIteration.path }' and
                          [System.WorkItemType] in (${ backlogTypesFilter })
                    order by [System.Id] asc` };

        console.debug(`CURRENT WORK WIQL: ${ currentWorkWiql.query }`);

        const currentWorkItemRefs = await getClient(WorkItemTrackingRestClient).queryByWiql(currentWorkWiql, this.state.projectName);
        const currentWorkItemIds = currentWorkItemRefs.workItems.map(wi => wi.id);

        const currentWorkItemsCompare = [] as WorkItem[];

        if (currentWorkItemIds.length > 0)
        {
            const currentWorkItems = await getClient(WorkItemTrackingRestClient).getWorkItems(currentWorkItemIds);

            console.debug(`RETRIEVED ${ currentWorkItems.length } CURRENT WORKITEMS`);

            = _.merge(currentWorkItemsCompare, currentWorkItems);
        }



        const removedWorkItems = _.differenceWith(plannedWorkItems, currentWorkItems, (pwi, cwi) => {
            return pwi.id === cwi.id;
         });    

        console.debug(`${ removedWorkItems.length } PLANNED WORKITEMS NOT FOUND IN CURRENT SET - REMOVED/MOVED:`);

        const removedChanges = _.map(removedWorkItems, function (v) {
            const sci = v as ScopeChangeWorkItem;
            sci.scopeChangeType = ScopeChangeType.RemovedFromSprint;
            return sci;
        });

        console.debug(JSON.stringify(removedChanges));

        const addedWorkItems = _.differenceWith(plannedWorkItems, currentWorkItems, (pwi, cwi) => {
            return pwi.id === cwi.id;
         });    

        console.debug(`${ removedWorkItems.length } CURRENT WORKITEMS NOT FOUND IN PLANNED SET - ADDED:`);

        const addedChanges = _.map(addedWorkItems, function (v) {
            const sci = v as ScopeChangeWorkItem;
            sci.scopeChangeType = ScopeChangeType.AddedToSprint;
            return sci;
        });

        console.debug(JSON.stringify(addedChanges));

        this.setState({
            scopeChanges: new ArrayItemProvider(_.merge(removedChanges, addedChanges))
        });
    }

    public render(): JSX.Element {

        const { userName, projectName, teamName, host, iframeUrl, extensionContext, route, navElements } = this.state;

        return (
            <div className="sample-pivot">
                {
                    !this.state.scopeChanges &&
                    <p>Loading...

                    <div>Hello, {userName}!</div>
                    {
                        (projectName && teamName) &&
                        <div>
                            <div>Project: {projectName}</div>
                            <div>Team: {teamName}</div>
                        </div>

                    }
                    <div>iframe URL: {iframeUrl}</div>
                    {
                        extensionContext &&
                        <div>
                            <div>Extension id: {extensionContext.id}</div>
                            <div>Extension version: {extensionContext.version}</div>
                        </div>
                    }
                    {
                        host &&
                        <div>
                            <div>Host id: {host.id}</div>
                            <div>Host name: {host.name}</div>
                            <div>Host service version: {host.serviceVersion}</div>
                        </div>
                    }
                    {
                        navElements && <div>Nav elements: {JSON.stringify(navElements)}</div>
                    }
                    {
                        route && <div>Route: {JSON.stringify(route)}</div>
                    } 
                    </p>
                }
                {
                    this.state.scopeChanges &&
                    <Table
                        columns={this.state.columns}
                        itemProvider={this.state.scopeChanges}
                    />
                }
            </div>
        );
    }
}

showRootComponent(<PivotContent />);