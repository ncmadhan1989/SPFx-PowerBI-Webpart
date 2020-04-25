import * as React from 'react';
import { useEffect } from 'react';

export default function IFrameContainer(props) {

    useEffect(() => {
        console.log('loaded iframe.');
    }, [props.iframesrc]);

    return (
        <div className="container-fluid">
            <div className="row">
                <div className="col-lg-12">
                    <iframe src={props.iframesrc} width="100%" height="650px" frameBorder="0"></iframe>
                </div>
            </div>
        </div>
    );
}
