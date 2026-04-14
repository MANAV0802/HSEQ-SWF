import * as React from 'react';
import {
    Stack,
    Label,
    Icon,
    IconButton,
    Text,
    Spinner,
    SpinnerSize
} from '@fluentui/react';
import { IAttachment } from '../../models/IToDoItem';
import styles from './AttachmentControl.module.scss';

export interface IAttachmentControlProps {
    attachments: IAttachment[];
    onUpload: (file: File) => Promise<void>;
    onDelete: (fileName: string) => Promise<void>;
}

const AttachmentControl: React.FC<IAttachmentControlProps> = ({ attachments, onUpload, onDelete }) => {
    const [uploading, setUploading] = React.useState(false);
    const fileInputRef = React.useRef<HTMLInputElement>(null);

    const handleFileChange = async (event: React.ChangeEvent<HTMLInputElement>) => {
        const file = event.target.files?.[0];
        if (file) {
            setUploading(true);
            try {
                await onUpload(file);
            } finally {
                setUploading(false);
                if (fileInputRef.current) fileInputRef.current.value = '';
            }
        }
    };

    return (
        <div className={styles.attachmentControl}>
            <Label>Attachments</Label>
            <Stack tokens={{ childrenGap: 10 }}>
                {attachments.map((file, idx) => (
                    <Stack horizontal verticalAlign="center" key={idx} className={styles.attachmentItem}>
                        <Icon iconName="Page" className={styles.fileIcon} />
                        <a href={file.ServerRelativeUrl} target="_blank" rel="noreferrer" className={styles.fileName}>
                            {file.FileName}
                        </a>
                        <IconButton 
                            iconProps={{ iconName: 'Cancel' }} 
                            title="Delete" 
                            onClick={() => onDelete(file.FileName)} 
                        />
                    </Stack>
                ))}
                
                <div 
                    className={styles.uploadArea} 
                    onClick={() => fileInputRef.current?.click()}
                >
                    {uploading ? (
                        <Spinner size={SpinnerSize.medium} label="Uploading..." />
                    ) : (
                        <>
                            <Icon iconName="Upload" className={styles.uploadIcon} />
                            <Text>Click to upload files</Text>
                        </>
                    )}
                </div>
                <input 
                    type="file" 
                    ref={fileInputRef} 
                    style={{ display: 'none' }} 
                    onChange={handleFileChange} 
                />
            </Stack>
        </div>
    );
};

export default AttachmentControl;
