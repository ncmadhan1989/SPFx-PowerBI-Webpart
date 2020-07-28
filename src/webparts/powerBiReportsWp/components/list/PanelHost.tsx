import * as React from 'react';
import { useEffect } from 'react';
import { classNames } from '../globalStyles';
import { LayerHost } from 'office-ui-fabric-react/lib/Layer';

export default function PanelHost(props) {

    useEffect(() => {
        console.log('loaded host.');
    }, [props.menuposition]);

    return (
        <LayerHost id="layerHostMenu"
            className={props.menuposition == 'right' ?
                classNames.layerHostClassRight :
                classNames.layerHostClassLeft
            }
        />
    );
}
