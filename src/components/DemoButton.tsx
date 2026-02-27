import * as React from 'react';
import { FunctionComponent, MouseEventHandler } from 'react';

interface IDemoButtonProps {
    onClick: MouseEventHandler<HTMLButtonElement>
    text: string
}

const DemoButton: FunctionComponent<IDemoButtonProps> = props => {
    return (
        <button
            className="b-button b-green"
            onClick={props.onClick}
            style={{ width : '100%' }}
        >
            {props.text}
        </button>
    );
};

export default DemoButton;
