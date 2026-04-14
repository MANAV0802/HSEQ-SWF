import * as React from 'react';
import { IToDoItem } from '../../../models/IToDoItem';
import { SPService } from '../../../services/SPService';
import GenericGrid from '../../Shared/GenericGrid';
import { IColumn, Panel, PanelType, Icon } from '@fluentui/react';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import ToDoForm from './ToDoForm';

export interface IToDoModuleProps {
    context: WebPartContext;
}

// ── Comprehensive Export Definitions ──────────────────────────────────────────
const formatDate = (d: string | undefined) => {
    if (!d) return '';
    const date = new Date(d);
    const day = String(date.getDate()).padStart(2, '0');
    const month = String(date.getMonth() + 1).padStart(2, '0');
    const year = date.getFullYear();
    return `${day}/${month}/${year}`;
};

const stripHtml = (html: string | undefined): string => {
    if (!html) return '';
    return html.replace(/(<([^>]+)>)/gi, "").trim();
};

const ToDoModule: React.FC<IToDoModuleProps> = ({ context }) => {
    const [items, setItems] = React.useState<IToDoItem[]>([]);
    const [loading, setLoading] = React.useState(true);
    const [selectedItem, setSelectedItem] = React.useState<IToDoItem | null>(null);
    const [isPanelOpen, setIsPanelOpen] = React.useState(false);

    const spService = React.useMemo(() => new SPService(context), [context]);

    const fetchItems = async () => {
        setLoading(true);
        try {
            const data = await spService.getToDoItems();
            setItems(data);
        } catch (e) {
            console.error('Fetch failed', e);
        } finally {
            setLoading(false);
        }
    };

    React.useEffect(() => { fetchItems(); }, []);

    // ── Grid Column Definitions ───────────────────────────────────────────────
    const columns: IColumn[] = [
        {
            key: 'Id', name: 'ID', fieldName: 'Id',
            minWidth: 40, maxWidth: 55, isResizable: true
        },
        {
            key: 'Regarding', name: 'Regarding', fieldName: 'Regarding',
            minWidth: 120, maxWidth: 180, isResizable: true
        },
        {
            key: 'TaskOwner', name: 'Task Owner', fieldName: 'TaskOwner',
            minWidth: 120, maxWidth: 160, isResizable: true,
            onRender: (item: IToDoItem) => {
                const title = item.TaskOwner?.Title || '';
                if (!title) return <span>—</span>;
                return (
                    <div style={{ display: 'flex', alignItems: 'center', gap: 6 }}>
                        <Icon iconName="Contact" style={{ fontSize: 12, color: '#0078d4' }} />
                        <span>{title}</span>
                    </div>
                );
            }
        },
        {
            key: 'StartDate', name: 'Start Date', fieldName: 'StartDate',
            minWidth: 90, maxWidth: 110, isResizable: true,
            onRender: (item: IToDoItem) => <span>{formatDate(item.StartDate)}</span>
        },
        {
            key: 'CompletionDate', name: 'Completion Date', fieldName: 'CompletionDate',
            minWidth: 110, maxWidth: 130, isResizable: true,
            onRender: (item: IToDoItem) => <span>{formatDate(item.CompletionDate)}</span>
        },
        {
            key: 'CompletedPercent', name: '% Done', fieldName: 'CompletedPercent',
            minWidth: 65, maxWidth: 80, isResizable: true,
            onRender: (item: IToDoItem) => {
                const pct = item.CompletedPercent ?? 0;
                return (
                    <div style={{ display: 'flex', alignItems: 'center', gap: 5 }}>
                        <div style={{ flex: 1, height: 6, background: '#e0e0e0', borderRadius: 3, minWidth: 36 }}>
                            <div style={{
                                width: `${Math.min(pct, 100)}%`, height: '100%',
                                background: pct >= 100 ? '#107C10' : pct >= 50 ? '#0078D4' : '#F7901E',
                                borderRadius: 3, transition: 'width 0.3s'
                            }} />
                        </div>
                        <span style={{ fontSize: 11, color: '#555', whiteSpace: 'nowrap' }}>{pct}%</span>
                    </div>
                );
            }
        }
    ];

    // ── Flatten items for export (FULL DATA) ──────────────────────────────────
    const flattenItems = (rawItems: any[]) => rawItems.map(item => ({
        'ID': item.Id,
        'Subject': item.Title || '',
        'Regarding': item.Regarding || '',
        'Status': item.Status?.Title || '',
        'Category': item.Category?.Title || '',
        'Classification': item.Classification?.Title || '',
        'Priority': item.Priority?.Title || '',
        'Task Owner': item.TaskOwner?.Title || '',
        'Assignee Internal': item.AssigneeInternal?.Title || '',
        'Assignee External': item.AssigneeExternal?.Title || '',
        'Due Date': formatDate(item.DueDate),
        'Start Date': formatDate(item.StartDate),
        'Completion Date': formatDate(item.CompletionDate),
        '% Done': `${item.CompletedPercent ?? 0}%`,
        'Resolution': stripHtml(item.Resolution),
        'Description': stripHtml(item.Description),
        'Email Notifications': item.EmailNotifications ? 'Yes' : 'No',
        'Created By': item.Author?.Title || '',
        'Created On': formatDate(item.Created),
        'Modified By': item.Editor?.Title || '',
        'Modified On': formatDate(item.Modified)
    }));

    // ── Export Handlers ───────────────────────────────────────────────────────
    const handleExportExcel = (data: any[]) => {
        import('../../../services/ExportService').then(({ ExportService }) => {
            ExportService.exportToExcel(flattenItems(data), 'ASP_Assist_Group_Tasks');
        });
    };

    const handleExportCSV = (data: any[]) => {
        import('../../../services/ExportService').then(({ ExportService }) => {
            ExportService.exportToCSV(flattenItems(data), 'ASP_Assist_Group_Tasks');
        });
    };

    const handleExportPDF = (data: any[]) => {
        import('../../../services/ExportService').then(({ ExportService }) => {
            const flattened = flattenItems(data);
            // For PDF, we need flat keys and headers separately
            const headers = Object.keys(flattened[0] || {});
            ExportService.exportToPDF(
                flattened,
                'ASP Assist Group Tasks',
                headers,
                headers // Use headers as keys directly since objects are now keyed by display names
            );
        });
    };

    const handleExportZip = (data: any[]) => {
        import('../../../services/ExportService').then(({ ExportService }) => {
            const flattened = flattenItems(data);
            const headers = Object.keys(flattened[0] || {});
            ExportService.exportToZip(
                flattened,
                'ASP_Assist_Group_Tasks',
                headers,
                headers
            ).catch(e => console.error('ZIP export failed', e));
        });
    };

    // ── CRUD Handlers ─────────────────────────────────────────────────────────
    const handleDelete = async (selectedItems: any[]) => {
        if (confirm(`Delete ${selectedItems.length} item(s)?`)) {
            for (const item of selectedItems) await spService.deleteToDoItem(item.Id);
            await fetchItems();
        }
    };

    const handleNew = () => { setSelectedItem(null); setIsPanelOpen(true); };
    const handleEdit = (item: IToDoItem) => { setSelectedItem(item); setIsPanelOpen(true); };

    const handleSave = async (payload: any, mode: 'stay' | 'close' | 'new') => {
        try {
            let resultItemId = selectedItem?.Id;
            if (selectedItem) {
                await spService.updateToDoItem(selectedItem.Id!, payload);
            } else {
                const result = await spService.addToDoItem(payload);
                resultItemId = result.data?.Id || result.Id;
            }
            const allItems = await spService.getToDoItems();
            setItems(allItems);

            if (mode === 'close') {
                setIsPanelOpen(false);
                setSelectedItem(null);
            } else if (mode === 'new') {
                setSelectedItem(null);
            } else if (mode === 'stay' && resultItemId) {
                const freshItem = allItems.find(i => i.Id === resultItemId);
                if (freshItem) setSelectedItem(freshItem);
            }
        } catch (e) {
            console.error('Save failed', e);
            alert('Save failed. Check browser console for details.');
        }
    };

    return (
        <React.Fragment>
            <GenericGrid
                items={items}
                columns={columns}
                loading={loading}
                onNew={handleNew}
                onEdit={handleEdit}
                onDelete={handleDelete}
                onRefresh={fetchItems}
                onExportExcel={handleExportExcel}
                onExportCSV={handleExportCSV}
                onExportPDF={handleExportPDF}
                onExportZip={handleExportZip}
            />

            <Panel
                isOpen={isPanelOpen}
                onDismiss={() => setIsPanelOpen(false)}
                type={PanelType.custom}
                customWidth="1100px"
                onRenderHeader={() => null}
                isLightDismiss={false}
                styles={{
                    content: { padding: 0 },
                    scrollableContent: { overflow: 'hidden' },
                    commands: { display: 'none' }
                }}
            >
                <ToDoForm
                    item={selectedItem}
                    spService={spService}
                    context={context}
                    onClose={() => setIsPanelOpen(false)}
                    onSave={handleSave}
                />
            </Panel>
        </React.Fragment>
    );
};

export default ToDoModule;
