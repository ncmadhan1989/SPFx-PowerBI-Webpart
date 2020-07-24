import { IReport } from '../models/IReport';
import { IGroup } from 'office-ui-fabric-react/lib/components/GroupedList/index';
import { Selection } from 'office-ui-fabric-react/lib/Selection';
import { IColumn } from 'office-ui-fabric-react/lib/DetailsList';
export interface IReportListsState {
    isOpen: boolean;
    isAllGroupsCollapsed: boolean;
    listItemsGroupedByCategory: IReport[];
    groups: IGroup[];
    selection: Selection;
    columns: IColumn[];
    iframesrc: string;
}