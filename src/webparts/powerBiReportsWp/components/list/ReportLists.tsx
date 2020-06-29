import * as React from 'react';
import { groupBy } from '@microsoft/sp-lodash-subset';
import { } from '@microsoft/sp-lodash-subset';
import { IReport } from '../models/IReport';
import { IReportListsState } from './IReportListsState';
import { ReportDataProvider } from '../dataprovider/ReportDataProvider';
import IFrameContainer from '../frame/IFrameContainer';
import { Fabric, IIconStyles } from 'office-ui-fabric-react/lib/index';
import { PrimaryButton, IButtonProps, IButtonStyles } from 'office-ui-fabric-react/lib/Button';
import { GroupedList, IGroup, IGroupDividerProps, IGroupHeaderStyles }
    from 'office-ui-fabric-react/lib/components/GroupedList';
import { GroupHeader, IGroupHeaderProps } from 'office-ui-fabric-react/lib/components/GroupedList/GroupHeader';
import { Icon, IconButton, IIconProps } from 'office-ui-fabric-react';
import { findIndex, IRenderFunction } from 'office-ui-fabric-react/lib/Utilities';
import { FocusZone } from 'office-ui-fabric-react/lib/FocusZone';
import { SelectionMode, SelectionZone, Selection } from 'office-ui-fabric-react/lib/Selection';
import {
    IColumn, DetailsRow,
    IDetailsRowStyles, IDetailsRowStyleProps, CheckboxVisibility
} from 'office-ui-fabric-react/lib/DetailsList';
import { mergeStyleSets, getTheme, mergeStyles } from 'office-ui-fabric-react/lib/Styling';
import { LayerHost } from 'office-ui-fabric-react/lib/Layer';
import { Panel, IPanelProps, IPanelStyleProps, IPanelStyles } from 'office-ui-fabric-react/lib/Panel';
import { IFocusTrapZoneProps } from 'office-ui-fabric-react/lib/FocusTrapZone';
import * as util from '../Util';

const theme = getTheme();
const menuIcon: IIconProps = { iconName: 'GlobalNavButton' };
const closeIcon: IIconProps = { iconName: 'Cancel' };
const classNames = mergeStyleSets({
    controlWrapper: {
        width: '100%',
        marginTop: '10px',
    },
    powerbilogobg: {
        fontSize: '42px',
        cursor: 'pointer',
    },
    centerbg: {
        position: 'absolute',
        top: '50%',
        left: '0',
        right: '0',
        textAlign: 'center',
        margin: '0 auto'
    },
    buttonconfigure: {
        zIndex: 99999,
    },
    menuAppbar: {
        width: '100%',
        top: '0',
        left: 'auto',
        right: '0',
        position: 'absolute',
        display: 'flex',
        zIndex: 1100,
        flexDirection: 'column',
        backgroundColor: 'rgb(243, 242, 241)',
        boxShadow: '0px 2px 4px -1px rgba(0,0,0,0.2), 0px 4px 5px 0px rgba(0,0,0,0.14), 0px 1px 10px 0px rgba(0,0,0,0.12)'
    },
    menuToolbar: {
        display: 'flex',
        position: 'relative',
        height: '32px',
        padding: '0 15px 0 15px',
    },
    menuHeading: {
        flexGrow: 1,
        margin: '0px',
        lineHeight: '1.75',
        textAlign: 'center',
    }

});
const layerHostClass = mergeStyles({
    position: 'absolute',
    width: 'auto',
    height: '100%',
    top: '32px',
    right: 0,
    zIndex: 1200
});
const panelStyle = (props: IPanelStyleProps): Partial<IPanelStyles> => ({
    ...({
        header: {
            marginTop: '0px !important',
            marginBottom: '0px !important',
        },
        headerText:{
            fontSize: '16px',
        }
    })
});
const detailRowStyle = (props: IDetailsRowStyleProps): Partial<IDetailsRowStyles> => ({
    ...({
        root: {
            width: '100%',
            background: 'rgb(244, 244, 244)',
            borderBottom: '1px solid rgb(255, 255, 255) !important'
        }
    })
});
const groupHeaderStyle = (props: IGroupHeaderProps): Partial<IGroupHeaderStyles> => ({
    ...({
        root: {
            background: 'gainsboro !important',
            display: 'flex',
            alignItems: 'center',
            height: '42px',
        },
        title: {
            fontSize: '16px',
            fontWeight: '400'
        }
    })
});
const closeIconButtonStyle: IButtonStyles = {
    root: {
        fontSize: '14px',
        fontWeight: '600',
        float: 'right',
    }
};
const focusTrapZoneProps: IFocusTrapZoneProps = {
    isClickableOutsideFocusTrap: true,
    forceFocusInsideTrap: false,
};

export interface IReportListsProps {
    siteurl: string;
    listname: string;
    iframeheight: number;
    reportsmenutitle: string;
    webparttitle: string;
    openpropertypane(): void;
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
        this._reportDataProvider = ReportDataProvider.getInstance();
        this.state = {
            isOpen: false,
            listItemsGroupedByCategory: [],
            groups: [],
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
            <div className="container-fluid">
                <div className="row">
                    <div className="col-lg-12">
                        <div className={classNames.controlWrapper}>
                            <div className={classNames.menuAppbar}>
                                <div className={classNames.menuToolbar}>
                                    <h6 className={classNames.menuHeading}>{this.props.webparttitle}</h6>
                                    {
                                        (this.props.listname) ?
                                            (<IconButton iconProps={menuIcon} title="Open Menu" onClick={this.openPanel} />)
                                            : null
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
                            <LayerHost id="layerHostMenu" className={layerHostClass} />
                        </div>
                    </div>
                    {
                        (this.props.listname) ?
                            (
                                <Panel
                                    isOpen={this.state.isOpen}
                                    hasCloseButton
                                    closeButtonAriaLabel="Close"
                                    onRenderNavigation={this._onRenderPanelNavigation}
                                    styles={panelStyle}
                                    headerText={this.props.reportsmenutitle}
                                    focusTrapZoneProps={focusTrapZoneProps}
                                    layerProps={{ hostId: 'layerHostMenu' }}
                                    onDismiss={this.dismissPanel}
                                >
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
                                                            usePageCache={true}
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

    private _onRenderPanelNavigation: IRenderFunction<IPanelProps> = (props, defaultRender) => {
        return (
            <>
                <IconButton iconProps={closeIcon} 
                    styles={closeIconButtonStyle} 
                    onClick={this.dismissPanel} title="Close" ariaLabel="Close" />
            </>
        );
    }
    
    private _onRenderCell = (nestingDepth: number, item: IReport, itemIndex: number): JSX.Element => {
        return (
            <DetailsRow
                styles={detailRowStyle}
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
                styles={groupHeaderStyle}
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
