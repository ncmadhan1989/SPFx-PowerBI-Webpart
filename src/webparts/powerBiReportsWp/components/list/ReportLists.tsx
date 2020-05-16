import * as React from 'react';
import { groupBy } from '@microsoft/sp-lodash-subset';
import { } from '@microsoft/sp-lodash-subset';
import { IReport } from '../models/IReport';
import { IReportListsState } from './IReportListsState';
import { ReportDataProvider } from '../dataprovider/ReportDataProvider';
import IFrameContainer from '../frame/IFrameContainer';
import { Fabric } from 'office-ui-fabric-react/lib/index';
import { GroupedList, IGroup, IGroupDividerProps }
    from 'office-ui-fabric-react/lib/components/GroupedList';
import { GroupHeader } from 'office-ui-fabric-react/lib/components/GroupedList/GroupHeader';
import { findIndex } from 'office-ui-fabric-react/lib/Utilities';
import { FocusZone } from 'office-ui-fabric-react/lib/FocusZone';
import { SelectionMode, SelectionZone, Selection } from 'office-ui-fabric-react/lib/Selection';
import { IColumn, DetailsRow, IDetailsGroupRenderProps } from 'office-ui-fabric-react/lib/DetailsList';
import { Icon } from 'office-ui-fabric-react/lib/Icon';
import { mergeStyleSets, getTheme } from 'office-ui-fabric-react/lib/Styling';

const theme = getTheme();
const classNames = mergeStyleSets({
    controlWrapper: {
        width: '100%',
        background: 'rgb(244, 244, 244) !important',
        marginTop: '10px',
    },
    selectionDetails: {
        marginBottom: '20px',
    },
    detailRow: {
        width: '100%',
        background: 'rgb(244, 244, 244) !important',
        borderBottom: '1px solid rgb(255, 255, 255) !important',
    },
    groupHeader: {
        background: 'gainsboro !important',
        display: 'flex',
        alignItems: 'center',
        height: '42px',
    },
    groupHeaderTitle: {
        paddingLeft: '12px',
        fontSize: '21px',
        fontWeight: '300',
        cursor: 'pointer',
        whiteSpace: 'nowrap',
        textOverflow: 'ellipsis',
    },
    groupHeaderButton: {
        position: 'relative',
        padding: '0px',
        display: 'flex',
        alignItems: 'center',
        justifyContent: 'center',
        fontSize: '18px',
        width: '32px',
        height: '42px',
        color: 'rgb(102, 102, 102)',
        outline: 'transparent',
        border: 'none !important',
        background: 'none transparent !important',
    },
    iconClass: {
        fontSize: 20
    },
    iconCollapsed: {
        transform: 'rotate(0deg) !important'
    },
    iconExpand: {
        transform: 'rotate(90deg) !important'
    }
});

export interface IReportListsProps {
    siteurl: string;
    listname: string;
}

export default class CustomerList extends React.Component<IReportListsProps, IReportListsState> {
    private _reportDataProvider: ReportDataProvider;
    private _selection: Selection = null;
    private _columns: IColumn[] = [{
        key: "ReportName",
        name: "Report Name",
        fieldName: "ReportName",
        minWidth: 400
    }];

    constructor(props) {
        super(props);
        this._reportDataProvider = new ReportDataProvider(props);
        this.state = {
            listItemsGroupedByCategory: [],
            groups: [],
            selection: this._selection,
            columns: [],
            iframesrc: ""
        };
    }

    public componentDidMount(): void {
        this._getAndGroupCustomerItems();
    }

    private _getAndGroupCustomerItems() {
        this._reportDataProvider.getItems()
            .then((res: IReport[]) => {
                //const _sortedReports = res.sort((a, b) => (a.CategoryName < b.CategoryName) ? -1 :
                //    (a.CategoryName > b.CategoryName) ? 1 : 0);                
                //const _groups = this._generateIGroups(res, 2, "CategoryName");
                const _groups = this.__generateIGroups(res, "CategoryName", 0, 2, 0, 0, true);
                this._selection = new Selection({
                    onSelectionChanged: () => {
                        const _selectedReport: IReport = this._getSelectionDetails();
                        if (_selectedReport) {
                            this.setState({
                                iframesrc: _selectedReport.ReportURL
                            });
                        }
                    },
                });
                this._selection.setItems(res);
                this.setState({
                    listItemsGroupedByCategory: res,
                    groups: _groups,
                    selection: this._selection,
                    columns: this._columns
                });
            });
    }

    private _generateIGroups(sortedCustomerItems: IReport[], groupDepth, groupColumnName: string): IGroup[] {
        let _groups: IGroup[] = [];
        let _groupsNew: IGroup[] = [];
        const _groupByType: _.Dictionary<IReport[]> = groupBy(sortedCustomerItems, (i: IReport) => i.CategoryName);

        Object.keys(_groupByType).forEach((group, index) => {
            _groups.push({
                name: group,
                key: group + index,
                startIndex: findIndex(sortedCustomerItems, (i: IReport) => i.CategoryName === group),
                count: _groupByType[group].length,
                isCollapsed: true
            });
        });

        _groupsNew.push(
            {
                name: "MRH", key: "MRH0", startIndex: 0, count: 3, isCollapsed: true,
                children: [
                    { name: "Compliance", key: "Compliance0", startIndex: 0, count: 2, level: 1, isCollapsed: true },
                    { name: "Crew", key: "Crew0", startIndex: 2, count: 1, level: 1, isCollapsed: true }
                ]
            },
            {
                name: "Technical", key: "Technical0", startIndex: 3, count: 1, isCollapsed: true,
                children: [
                    { name: "Cost", key: "Cost1", startIndex: 3, count: 1, level: 1, isCollapsed: true }
                ]
            }
        );

        return _groups;
    }

    private __generateIGroups(sortedCustomerItems: IReport[],
        groupColumnName: string,
        groupCount: number,
        groupDepth: number,
        startIndex: number,
        level: number = 0,
        isCollapsed: boolean = true
    ): IGroup[] {
        let _groupsNew: IGroup[] = [];
        const _groupByType: _.Dictionary<IReport[]> = groupBy(sortedCustomerItems, (i: IReport) => {
            if (groupColumnName === "CategoryName")
                return i["CategoryName"];
            return i["SubCategory"];
        });

        Object.entries(_groupByType).map((group, index) => {
            const _group = group[0];
            const _count = group[1].length;
            const _items = group[1];
            _groupsNew.push({
                count: _count,
                key: _group + index,
                name: _group,
                startIndex: startIndex,
                level: level,
                isCollapsed: true,
                children:
                    (groupDepth > 1 && _items.length > 0)
                        ? this.__generateIGroups(_items, "SubCategory", _count, groupDepth - 1, startIndex, 1, isCollapsed)
                        : []
            });
            startIndex = startIndex + _count;
        });
        return _groupsNew;
    }

    public render() {
        const { listItemsGroupedByCategory, groups, selection, columns, iframesrc } = this.state;
        return (
            (this.props.siteurl && this.props.listname) ?
                (
                    <div className="container-fluid">
                        <div className="row">
                            <div className="col-lg-9">
                                <div className={classNames.controlWrapper}>
                                    <IFrameContainer iframesrc={iframesrc} />
                                </div>
                            </div>
                            <div className="col-lg-3">
                                <Fabric>
                                    <div className={classNames.controlWrapper}>
                                        {
                                            <FocusZone>
                                                <SelectionZone selection={selection} selectionMode={SelectionMode.single}>
                                                    <GroupedList
                                                        items={listItemsGroupedByCategory}
                                                        groupProps={{
                                                            onRenderHeader: this._onRenderHeader
                                                        }}
                                                        selection={selection}
                                                        groups={groups}
                                                        onRenderCell={this._onRenderCell}
                                                        selectionMode={SelectionMode.single}
                                                    />
                                                </SelectionZone>
                                            </FocusZone>
                                        }
                                    </div>
                                </Fabric>
                            </div>
                        </div>
                    </div>
                )
                :
                (
                    <div className="alert alert-danger text-center" role="alert">
                        Please provide with the Base site url and Reports list name webpart properties
                    </div>
                )
        );
    }

    private _onRenderCell = (nestingDepth: number, item: IReport, itemIndex: number): JSX.Element => {
        return (
            <DetailsRow
                className={classNames.detailRow}
                columns={this.state.columns}
                groupNestingDepth={nestingDepth}
                item={item}
                itemIndex={itemIndex}
                selection={this.state.selection}
                selectionMode={SelectionMode.single}
                compact={false}
            />
        );
    }

    private _onRenderHeader(props: IGroupDividerProps): JSX.Element {
        const onToggleSelectGroup = () => {
            props.onToggleCollapse(props.group);
        };
        const onToggleCollapse = () => {
            props.onToggleCollapse(props.group);
            const _collaspedGroup = props.groups.filter(g => g.key != props.group.key);
            _collaspedGroup.forEach((_group, index) => {
                if (!_group.isCollapsed)
                    props.onToggleCollapse(_group);
            });
        };
        return (
            <GroupHeader {...props}
                className={classNames.groupHeader}
                onToggleCollapse={onToggleCollapse}
                onToggleSelectGroup={onToggleSelectGroup} />
        );
    }

    private _getSelectionDetails(): IReport {
        const selectionCount = this._selection.getSelectedCount();
        const _item: IReport = this._selection.getSelection()[0] as IReport;
        return _item;
    }

}
