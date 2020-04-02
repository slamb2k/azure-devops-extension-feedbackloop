import "./Pivot.scss";

import * as React from "react";
import * as SDK from "azure-devops-extension-sdk";

import { showRootComponent } from "../../Common";

import { getClient, IProjectPageService, CommonServiceIds } from "azure-devops-extension-api";
import { WorkItemTrackingRestClient, WorkItem } from "azure-devops-extension-api/WorkItemTracking";

import { Table, ITableColumn, renderSimpleCell, renderSimpleCellValue } from "azure-devops-ui/Table";
import { ArrayItemProvider } from "azure-devops-ui/Utilities/Provider";

interface IPivotContentState {
    revisions?: ArrayItemProvider<WorkItem>;
    columns: ITableColumn<any>[];
}

class PivotContent extends React.Component<{}, IPivotContentState> {

    constructor(props: {}) {
        super(props);

        this.state = {
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
                    return renderSimpleCellValue<any>(columnIndex, tableColumn, tableItem.fields["System.Title"]);
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
            }]
        };
    }

    public componentDidMount() {
        SDK.init();
        this.initializeComponent();
    }

    private async initializeComponent() {
        const projectService = await SDK.getService<IProjectPageService>(CommonServiceIds.ProjectPageService);
        const project = await projectService.getProject();

        if (project)
        {
            console.log("PROJECT: " + project.name);

            const revisionHistory = await getClient(WorkItemTrackingRestClient).readReportingRevisionsGet(project.name);
            const revisions = revisionHistory.values;
            revisions.filter()
            console.log("REVISION COUNT: " + revisions.length);

            //const revisions = await getClient(WorkItemTrackingRestClient).getWorkItems([1]);

            const _ = require('lodash');

            const plannedWorkItems = revisions;
            const currentWorkItems = revisions;

            let difference = plannedWorkItems.filter(x => !(currentWorkItems.indexOf(x) !== -1));
 
            this.setState({
                revisions: new ArrayItemProvider(revisions)
            });
        }
    }

    public render(): JSX.Element {
        return (
            <div className="sample-pivot">
                {
                    !this.state.revisions &&
                     <p>Loading...</p>
                }
                {
                    this.state.revisions &&
                    <Table
                        columns={this.state.columns}
                        itemProvider={this.state.revisions}
                    />
                }
            </div>
        );
    }
}

showRootComponent(<PivotContent />);