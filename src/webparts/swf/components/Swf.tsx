import * as React from 'react';
import MainLayout from './Layout/MainLayout';
import ToDoModule from './Modules/ToDo/ToDoModule';
import { ISwfProps } from './ISwfProps';

const Swf: React.FC<ISwfProps> = ({ context }) => {
    const [activeModule, setActiveModule] = React.useState<string>('ToDo');

    const renderModule = () => {
        switch (activeModule) {
            case 'ToDo':
                return <ToDoModule context={context} />;
            default:
                return (
                    <div style={{
                        display: 'flex',
                        alignItems: 'center',
                        justifyContent: 'center',
                        height: '100%',
                        flexDirection: 'column',
                        gap: 12,
                        color: '#666',
                        fontSize: 16
                    }}>
                        <span style={{ fontSize: 40 }}>🚧</span>
                        <span>This module is under development</span>
                    </div>
                );
        }
    };

    return (
        <MainLayout
            activeModule={activeModule}
            onModuleChange={setActiveModule}
        >
            {renderModule()}
        </MainLayout>
    );
};

export default Swf;
