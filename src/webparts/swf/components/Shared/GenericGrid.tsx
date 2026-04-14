import * as React from 'react';
import {
    DetailsListLayoutMode,
    Selection,
    IColumn,
    SelectionMode,
    CommandBar,
    ICommandBarItemProps,
    ShimmeredDetailsList,
    ConstrainMode,
    Icon
} from '@fluentui/react';
import styles from './GenericGrid.module.scss';

export interface IGenericGridProps {
    items: any[];
    columns: IColumn[];
    onNew?: () => void;
    onDelete?: (selectedItems: any[]) => void;
    onExportExcel?: (items: any[]) => void;
    onExportPDF?: (items: any[]) => void;
    onExportCSV?: (items: any[]) => void;
    onExportZip?: (items: any[]) => void;
    onRefresh?: () => void;
    onEdit?: (item: any) => void;
    loading?: boolean;
    title?: string;
}

const GenericGrid: React.FC<IGenericGridProps> = (props) => {
    const [filteredItems, setFilteredItems] = React.useState<any[]>(props.items);
    const [selectionCount, setSelectionCount] = React.useState(0);

    const [selection] = React.useState<Selection>(
        new Selection({
            onSelectionChanged: () => {
                setSelectionCount(selection.getSelectedCount());
            }
        })
    );

    React.useEffect(() => {
        setFilteredItems(props.items);
    }, [props.items]);

    const getExportTarget = (): any[] =>
        selectionCount > 0 ? selection.getSelection() : filteredItems;

    const commandItems: ICommandBarItemProps[] = [
        {
            key: 'new',
            text: 'New',
            iconProps: { iconName: 'Add' },
            className: styles.cmdNew,
            onClick: props.onNew
        },
        {
            key: 'edit',
            text: 'Edit',
            iconProps: { iconName: 'Edit' },
            disabled: selectionCount !== 1,
            onClick: () => {
                if (selectionCount === 1) props.onEdit?.(selection.getSelection()[0]);
            }
        },
        {
            key: 'delete',
            text: 'Delete',
            iconProps: { iconName: 'Delete' },
            disabled: selectionCount === 0,
            onClick: () => props.onDelete?.(selection.getSelection())
        },
        {
            key: 'export',
            text: 'Export',
            iconProps: { iconName: 'Download' },
            subMenuProps: {
                items: [
                    {
                        key: 'excel',
                        text: 'Excel (.xlsx)',
                        iconProps: { iconName: 'ExcelDocument' },
                        onClick: () => props.onExportExcel?.(getExportTarget())
                    },
                    {
                        key: 'csv',
                        text: 'CSV (.csv)',
                        iconProps: { iconName: 'TextDocument' },
                        onClick: () => props.onExportCSV?.(getExportTarget())
                    },
                    {
                        key: 'pdf',
                        text: 'PDF (.pdf)',
                        iconProps: { iconName: 'PDF' },
                        onClick: () => props.onExportPDF?.(getExportTarget())
                    },
                    { key: 'divider', itemType: 1 /* Divider */ },
                    {
                        key: 'zip',
                        text: 'ZIP — All Formats',
                        iconProps: { iconName: 'ZipFolder' },
                        onClick: () => props.onExportZip?.(getExportTarget())
                    }
                ]
            }
        },
        {
            key: 'refresh',
            text: 'Refresh',
            iconProps: { iconName: 'Refresh' },
            onClick: props.onRefresh
        }
    ];

    // Pass columns through as-is; borders are applied via SCSS on .ms-DetailsRow-cell
    const styledColumns: IColumn[] = props.columns.map(col => ({
        ...col,
        isResizable: true
    }));

    const selectionLabel =
        selectionCount > 0
            ? `${selectionCount} of ${filteredItems.length} selected`
            : `${filteredItems.length} item${filteredItems.length !== 1 ? 's' : ''}`;

    return (
        <div className={styles.genericGrid}>
            {/* ── Toolbar ── */}
            <div className={styles.gridHeader}>
                <CommandBar
                    items={commandItems}
                    className={styles.commandBar}
                    styles={{ root: { padding: '0 8px' } }}
                />
            </div>

            {/* ── Selection Summary ── */}
            {selectionCount > 0 && (
                <div className={styles.selectionBar}>
                    <div className={styles.selectionInfo}>
                        <Icon iconName="MultiSelect" className={styles.selectionIcon} />
                        <span className={styles.selectionBadge}>{selectionCount}</span>
                        {' '}item{selectionCount !== 1 ? 's' : ''} selected
                    </div>
                    <button
                        className={styles.clearSelection}
                        onClick={() => { selection.setAllSelected(false); }}
                    >
                        Clear Selection
                    </button>
                </div>
            )}

            {/* ── Data Grid ── */}
            <div className={styles.listContainer}>
                <ShimmeredDetailsList
                    items={filteredItems}
                    columns={styledColumns}
                    selection={selection}
                    selectionMode={SelectionMode.multiple}
                    layoutMode={DetailsListLayoutMode.fixedColumns}
                    constrainMode={ConstrainMode.unconstrained}
                    enableShimmer={props.loading}
                    onItemInvoked={props.onEdit}
                    selectionPreservedOnEmptyClick={false}
                    className={styles.detailsList}
                />
            </div>

            {/* ── Status Bar ── */}
            <div className={styles.statusBar}>
                <span>{selectionLabel}</span>
                {selectionCount > 0 && (
                    <span className={styles.exportHint}>
                        ↑ Export will use selected records
                    </span>
                )}
            </div>
        </div>
    );
};

export default GenericGrid;
