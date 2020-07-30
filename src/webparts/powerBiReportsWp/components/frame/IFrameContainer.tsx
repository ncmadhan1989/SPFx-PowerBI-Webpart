import * as React from 'react';
import { useEffect } from 'react';

export default function IFrameContainer(props) {

    useEffect(() => {
        console.log('loaded iframe.');
    }, [props.iframesrc, props.iframeheight]);

    let src: string = "";
    if(props.iframesrc)
        src = `${props.iframesrc}&filterPaneEnabled=${props.showfilterpane}&navContentPaneEnabled=${props.shownavigationpane}`;

    return (
        <div className="container-fluid" style={{ padding: '0px', marginTop: '5px' }}>
            <div className="row">
                <div className="col-lg-12" style={{ padding: '1px' }}>
                    <iframe src={src} width="100%" height={props.iframeheight} frameBorder="0"></iframe>
                </div>
            </div>
        </div>
    );
}
