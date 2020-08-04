import * as React from 'react';
import { styles, classNames, icons } from '../globalStyles';
import { groupBy } from '@microsoft/sp-lodash-subset';
import { } from '@microsoft/sp-lodash-subset';
import { IReport } from '../models/IReport';
import { IReportListsState } from './IReportListsState';
import { ReportDataProvider } from '../dataprovider/ReportDataProvider';
import IFrameContainer from '../frame/IFrameContainer';
import { Fabric, Spinner } from 'office-ui-fabric-react/lib/index';
import { PrimaryButton } from 'office-ui-fabric-react/lib/Button';
import { Stack } from 'office-ui-fabric-react/lib/Stack';
import { Label } from 'office-ui-fabric-react/lib/Label';
import { GroupedList, IGroup }
    from 'office-ui-fabric-react/lib/components/GroupedList';
import { GroupHeader, IGroupHeaderProps } from 'office-ui-fabric-react/lib/components/GroupedList/GroupHeader';
import { Icon, IconButton } from 'office-ui-fabric-react';
import { Toggle } from 'office-ui-fabric-react/lib/Toggle';
import { findIndex, IRenderFunction } from 'office-ui-fabric-react/lib/Utilities';
import { FocusZone } from 'office-ui-fabric-react/lib/FocusZone';
import { SelectionMode, SelectionZone, Selection } from 'office-ui-fabric-react/lib/Selection';
import {
    IColumn, DetailsRow, CheckboxVisibility
} from 'office-ui-fabric-react/lib/DetailsList';
import { LayerHost } from 'office-ui-fabric-react/lib/Layer';
import { Panel, IPanelProps, IPanelHeaderRenderer, PanelType } from 'office-ui-fabric-react/lib/Panel';
import { SearchBox } from 'office-ui-fabric-react/lib/SearchBox';
import { IFocusTrapZoneProps } from 'office-ui-fabric-react/lib/FocusTrapZone';

const focusTrapZoneProps: IFocusTrapZoneProps = {
    isClickableOutsideFocusTrap: false,
    forceFocusInsideTrap: false,
};

export interface IReportListsProps {
    siteurl: string;
    listname: string;
    iframeheight: number;
    reportsmenutitle: string;
    menuposition: string;
    webparttitle: string;
    paneltype: string;
    panelwidth: number;
    shownavigationpane: boolean;
    showfilterpane: boolean;
    openpropertypane(): void;
}

export default class CustomerList extends React.Component<IReportListsProps, IReportListsState> {
    private _reportDataProvider: ReportDataProvider;
    private _results: IReport[];
    private _selection: Selection = null;
    private _columns: IColumn[] = [{
        key: "ReportName",
        name: "Report Name",
        fieldName: "ReportName",
        minWidth: 400
    }];

    constructor(props) {
        super(props);
        this._reportDataProvider = ReportDataProvider.getInstance();
        this.state = {
            isOpen: true,
            isLoading: true,
            isAllGroupsCollapsed: false,
            listItemsGroupedByCategory: [],
            groups: [],
            menuPosition: "right",
            selection: this._selection,
            columns: [],
            iframesrc: ""
        };
    }

    private openPanel = () => {
        this.setState({
            isOpen: true
        });
    }

    private dismissPanel = () => {
        this.setState({
            isOpen: false
        });
    }

    public componentDidMount(): void {
        if (this.props.listname) {
            this._getAndGroupCustomerItems();
        }
    }

    public componentDidUpdate(prevProps: IReportListsProps, prevState: IReportListsState): void {
        if (this.props.listname !== prevProps.listname) {
            if (this.props.listname) {
                this._getAndGroupCustomerItems();
            }
        }
    }

    private _getAndGroupCustomerItems() {
        this._reportDataProvider.getItems(this.props.listname)
            .then((res: IReport[]) => {
                this._results = res;
                this._setResults(res);
            });
    }

    private _setResults(res: IReport[]) {
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
            isLoading: false,
            groups: _groups,
            selection: this._selection,
            columns: this._columns
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
                isCollapsed: isCollapsed,
                children:
                    (groupDepth > 1 && _items.length > 0)
                        ? this.__generateIGroups(_items, "SubCategory", _count, groupDepth - 1, startIndex, 1, isCollapsed)
                        : []
            });
            startIndex = startIndex + _count;
        });
        return _groupsNew;
    }

    private _onSearchReport = (text: string) => {
        const _prevResults: IReport[] = this._results;
        const _filter: IReport[] = text ?
            _prevResults.filter(res => res.ReportName.toLocaleLowerCase().indexOf(text) > -1)
            :
            this._results;
        this._setResults(_filter);
    }

    private _onExpandCollapseAll = (ev: React.MouseEvent<HTMLElement>, checked: boolean) => {
        const _groups = this.__generateIGroups(this._results, "CategoryName", 0, 2, 0, 0, !checked);
        this.setState({
            groups: _groups,
        });
    }

    private _onSearchCleared = (ev: React.FormEvent<HTMLElement | HTMLTextAreaElement>) => {
        this._setResults(this._results);
    }

    private _selectPanelType(type: string): PanelType {
        switch (type) {
            case 'custom': return PanelType.custom;
            case 'small': return PanelType.smallFixedFar;
            case 'medium': return PanelType.medium;
            default: return PanelType.smallFixedFar;
        }
    }

    public render() {
        const { listItemsGroupedByCategory, isAllGroupsCollapsed, groups, selection, columns, iframesrc, isLoading } = this.state;
        const panelType: PanelType = this._selectPanelType(this.props.paneltype);
        return (
            <div className="container-fluid">
                <div className="row">
                    <div className="col-lg-12">
                        <div className={classNames.controlWrapper}>
                            <div className={classNames.menuAppbar}>
                                <div className={classNames.menuToolbar}>
                                    {
                                        this.props.menuposition == 'right' ?
                                            <h6 className={classNames.menuHeading}>{this.props.webparttitle}</h6> :
                                            null
                                    }
                                    {
                                        (this.props.listname) ?
                                            (<IconButton iconProps={icons.menuIcon}
                                                styles={styles.menuIconStyles} title="Open Menu"
                                                onClick={this.openPanel} />)
                                            : null
                                    }
                                    {
                                        this.props.menuposition == 'left' ?
                                            <h6 className={classNames.menuHeading}>{this.props.webparttitle}</h6> :
                                            null
                                    }
                                </div>
                            </div>
                            <h3 className={classNames.centerbg}>
                                {
                                    (iframesrc) ? null :
                                        (<div><Icon iconName="PowerBILogo" className={classNames.powerbilogobg}></Icon></div>)
                                }
                                {(this.props.listname) ? null :
                                    (<div>
                                        <PrimaryButton text="Configure Webpart"
                                            className={classNames.buttonconfigure}
                                            onClick={this._openPropertyPane.bind(this)} />
                                    </div>)
                                }
                            </h3>
                            <IFrameContainer iframesrc={iframesrc} {...this.props} />
                            <LayerHost id="layerHostMenu"
                                className={this.props.menuposition == 'right' ?
                                    classNames.layerHostClassRight :
                                    classNames.layerHostClassLeft
                                }
                            />
                        </div>
                    </div>
                    {
                        (this.props.listname) ?
                            (
                                <Panel
                                    isOpen={this.state.isOpen}
                                    hasCloseButton
                                    type={panelType}
                                    customWidth={this.props.panelwidth + 'px'}
                                    closeButtonAriaLabel="Close"
                                    onRenderHeader={this._onRenderPanelHeader}
                                    onRenderNavigationContent={this._onRenderPanelNavigationContent}
                                    styles={this.props.menuposition == 'right' ?
                                        styles.panelRightStyles :
                                        styles.panelLeftStyles}
                                    focusTrapZoneProps={focusTrapZoneProps}
                                    layerProps={{ hostId: 'layerHostMenu' }}
                                    onDismiss={this.dismissPanel}
                                >
                                    <Fabric>
                                        <div className={classNames.controlWrapper}>
                                            {
                                                isLoading ?
                                                    (
                                                        <div className={classNames.spinnerWrapper}>
                                                            <Spinner label="Working on it..." labelPosition="bottom"></Spinner>
                                                        </div>
                                                    )
                                                    :
                                                    (
                                                        <FocusZone>
                                                            <SelectionZone selection={selection} selectionMode={SelectionMode.single}>
                                                                <GroupedList
                                                                    items={listItemsGroupedByCategory}
                                                                    groupProps={{
                                                                        onRenderHeader: this._onRenderHeader
                                                                    }}
                                                                    usePageCache={true}
                                                                    selection={selection}
                                                                    groups={groups}
                                                                    onRenderCell={this._onRenderCell}
                                                                    selectionMode={SelectionMode.single}
                                                                />
                                                            </SelectionZone>
                                                        </FocusZone>
                                                    )
                                            }
                                        </div>
                                    </Fabric>

                                </Panel>
                            )
                            : null
                    }
                </div>
            </div>
        );
    }

    private _openPropertyPane(e) {
        if (e)
            e.preventDefault();

        this.props.openpropertypane();
    }

    private _onRenderPanelNavigationContent: IRenderFunction<IPanelProps> = (props, defaultRender) => {
        return (
            <>
                <SearchBox
                    placeholder="Search report..."
                    ariaLabel="Search the report."
                    iconProps={icons.filterIcon}
                    styles={styles.searchBoxStyles}
                    onClear={this._onSearchCleared}
                    onEscape={this._onSearchCleared}
                    onSearch={this._onSearchReport}
                />
                <IconButton iconProps={icons.closeIcon}
                    styles={styles.closeIconButtonStyles}
                    onClick={this.dismissPanel} title="Close" ariaLabel="Close" />
            </>
        );
    }

    private _onRenderPanelHeader: IPanelHeaderRenderer = (props, defaultRender) => {
        return (
            <>
                <Stack horizontal tokens={{ childrenGap: 10 }} className={classNames.panelHeader}>
                    <Label styles={{ root: { fontSize: '18px' } }}>{this.props.reportsmenutitle}</Label>
                    <Toggle onText="Expand" offText="Collapse"
                        onChange={this._onExpandCollapseAll} styles={styles.toggleExpandCollapse} />
                </Stack>
            </>
        );
    }

    private _onRenderCell = (nestingDepth: number, item: IReport, itemIndex: number): JSX.Element => {
        return (
            <DetailsRow
                styles={styles.detailsRowStyles}
                columns={this.state.columns}
                groupNestingDepth={nestingDepth}
                item={item}
                checkboxVisibility={CheckboxVisibility.always}
                indentWidth={10}
                itemIndex={itemIndex}
                selection={this.state.selection}
                selectionMode={SelectionMode.single}
                compact={false}
            />
        );
    }

    private _onRenderHeader(props: IGroupHeaderProps): JSX.Element {
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
                styles={styles.groupHeaderStyles}
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
