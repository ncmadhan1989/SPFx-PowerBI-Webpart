import * as React from 'react';
import { useEffect } from 'react';

export default function IFrameContainer(props) {

    useEffect(() => {
        console.log('loaded iframe.');
    }, [props.iframesrc, props.iframeheight]);

    return (
        <div className="container-fluid" style={{ padding: '0px', marginTop: '32px' }}>
            <div className="row">
                <div className="col-lg-12" style={{ padding: '0px' }}>
                    <iframe src={props.iframesrc} width="100%" height={props.iframeheight} frameBorder="0"></iframe>
                </div>
            </div>
        </div>
    );
}
