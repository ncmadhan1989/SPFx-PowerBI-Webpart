import { mergeStyleSets, getTheme } from 'office-ui-fabric-react/lib/Styling';
import { IDocumentCardStyles } from 'office-ui-fabric-react/lib/DocumentCard';
import { IIconProps, IGroupHeaderStyles, IGroupHeaderStyleProps, IToggleStyles, IToggleStyleProps } from 'office-ui-fabric-react/lib/index';
import { IButtonStyles } from 'office-ui-fabric-react/lib/Button';
import { IPanelStyles, IPanelStyleProps } from 'office-ui-fabric-react/lib/Panel';
import { ISearchBoxStyles, ISearchBoxStyleProps } from 'office-ui-fabric-react/lib/SearchBox';
import { IDetailsRowStyles, IDetailsRowStyleProps } from 'office-ui-fabric-react/lib/DetailsList';

const theme = getTheme();

const cardStyles: IDocumentCardStyles = {
    root: {
        width: '100%',
        maxWidth: 'none',
        minWidth: '300px',
        marginTop: '10px',
        backgroundColor: '#f3f2f1',
    }
};
const closeIconButtonStyles: IButtonStyles = {
    root: {
        fontSize: '14px',
        fontWeight: '600',        
    }
};
const expandCollapseIconStyles: IButtonStyles ={
    root:{
        fontSize: '16px',
        lineHeight: '2',
    }
};
const menuIconStyles: IButtonStyles = {
    root:{
        color: '#000000'
    }
};
const panelLeftStyles = (props: IPanelStyleProps): Partial<IPanelStyles> => ({
    ...({
        main:{
            left: 0,
            right: 'auto'
        },
        header: {
            marginTop: '0px !important',
            marginBottom: '0px !important',
        },
        headerText: {
            fontSize: '16px',
        },
        content:{
            paddingLeft: '15px !important',
            paddingRight: '15px !important'
        }
    })
});
const panelRightStyles = (props: IPanelStyleProps): Partial<IPanelStyles> => ({
    ...({
        main:{
            right: 0,
            left: 'auto'
        },
        header: {
            marginTop: '0px !important',
            marginBottom: '0px !important',
        },
        headerText: {
            fontSize: '16px',
        },
        content:{
            paddingLeft: '15px !important',
            paddingRight: '15px !important'
        }
    })
});
const detailRowStyles = (props: IDetailsRowStyleProps): Partial<IDetailsRowStyles> => ({
    ...({
        root: {
            width: '100%',
            background: 'rgb(244, 244, 244)',
            borderBottom: '1px solid rgb(255, 255, 255) !important'
        }
    })
});
const groupHeaderStyles = (props: IGroupHeaderStyleProps): Partial<IGroupHeaderStyles> => ({
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
const searchBoxStyles = (props: ISearchBoxStyleProps): Partial<ISearchBoxStyles> => ({
    ...({
        root: {
            width: '100%',
            margin: '5px',
        }
    })
});
const toggleExpandCollapse = (props: IToggleStyleProps): Partial<IToggleStyles> => ({
    ...({
        root:{
            marginTop: '10px'
        },
        text:{
            fontSize: '12px'
        }
    })
});

export const styles = {
    cardStyles: cardStyles,
    closeIconButtonStyles: closeIconButtonStyles,
    menuIconStyles: menuIconStyles,
    expandCollapseIconStyles: expandCollapseIconStyles,
    panelLeftStyles: panelLeftStyles,
    panelRightStyles: panelRightStyles,
    detailsRowStyles: detailRowStyles,
    groupHeaderStyles: groupHeaderStyles,
    searchBoxStyles: searchBoxStyles,
    toggleExpandCollapse: toggleExpandCollapse,
};

const menuIcon: IIconProps = { iconName: 'GlobalNavButton' };
const closeIcon: IIconProps = { iconName: 'Cancel' };
const filterIcon: IIconProps = { iconName: 'Filter' };
const expandIcon: IIconProps = { iconName: 'ExploreContent' };
const collapseIcon: IIconProps = { iconName: 'CollapseContent' };

export const icons = {
    menuIcon: menuIcon,
    closeIcon: closeIcon,
    filterIcon: filterIcon,
    expandIcon: expandIcon,
    collapseIcon: collapseIcon,
};

export const classNames = mergeStyleSets({
    formWrapper: {
        display: 'flex',
        alignItems: 'center',
        flexDirection: 'column',
        width: '100%',
        minHeight: '100%',
        padding: '20px',
    },
    controlWrapper: {
        width: '100%',
        marginTop: '5px',
    },
    layerHostClassRight: {
        position: 'absolute',
        width: 'auto',
        height: '100%',
        top: '37px',
        right: '15px',
        zIndex: 1200
    },
    layerHostClassLeft: {
        position: 'absolute',
        width: 'auto',
        height: '100%',
        top: '37px',
        left: '15px',
        right: 'auto',
        zIndex: 1200
    },
    powerbilogobg: {
        fontSize: '42px',
        cursor: 'pointer',
        color: '#000000',
    },
    centerbg: {
        position: 'absolute',
        top: '50%',
        left: '0',
        right: '0',
        textAlign: 'center',
        margin: '0 auto'
    },
    panelHeader:{
        paddingLeft: '40px',
        paddingRight: '40px',
    },
    buttonconfigure: {
        zIndex: 99999,        
    },
    menuAppbar: {
        width: '100%',
        top: '0',
        left: 'auto',
        right: '0',
        position: 'sticky',
        display: 'flex',
        zIndex: 1100,
        flexDirection: 'column',
        backgroundColor: 'rgb(243, 242, 241)',
        color: '#000000',
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