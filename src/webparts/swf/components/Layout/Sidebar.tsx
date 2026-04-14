import * as React from 'react';
import styles from './Sidebar.module.scss';
import { Icon } from '@fluentui/react';

export interface ISidebarProps {
    activeItem: string;
    onItemClick: (item: string) => void;
    isCollapsed?: boolean;
}

const Sidebar: React.FC<ISidebarProps> = ({ activeItem, onItemClick, isCollapsed }) => {
    const [hseqExpanded, setHseqExpanded] = React.useState(true);
    const [adminExpanded, setAdminExpanded] = React.useState(true);

    return (
        <aside className={`${styles.sidebar} ${isCollapsed ? styles.collapsed : ''}`}>
            <div className={styles.logoSection}>
                <div className={styles.logo}>
                    {isCollapsed ? 'A' : 'ASP Assist Group'}
                </div>
            </div>

            <nav className={styles.nav}>
                {/* HSEQ Section */}
                <div className={styles.navGroup}>
                    {!isCollapsed && (
                        <div className={styles.groupHeader} onClick={() => setHseqExpanded(!hseqExpanded)}>
                            <Icon iconName="CheckList" className={styles.groupIcon} />
                            <span>HSEQ</span>
                            <Icon iconName={hseqExpanded ? 'ChevronUp' : 'ChevronDown'} className={styles.chevron} />
                        </div>
                    )}
                    {(hseqExpanded || isCollapsed) && (
                        <div className={styles.subItems}>
                            <div 
                                className={`${styles.navItem} ${activeItem === 'ToDo' ? styles.active : ''}`}
                                onClick={() => onItemClick('ToDo')}
                                title={isCollapsed ? 'To Do' : ''}
                            >
                                <Icon iconName="TaskManager" className={styles.itemIcon} />
                                {!isCollapsed && <span>To Do</span>}
                            </div>
                            <div 
                                className={`${styles.navItem} ${activeItem === 'Compliance' ? styles.active : ''}`}
                                onClick={() => onItemClick('Compliance')}
                                title={isCollapsed ? 'Compliance' : ''}
                            >
                                <Icon iconName="ReadingMode" className={styles.itemIcon} />
                                {!isCollapsed && <span>Compliance Register</span>}
                            </div>
                        </div>
                    )}
                </div>

                {/* Admin Section */}
                <div className={styles.navGroup}>
                    {!isCollapsed && (
                        <div className={styles.groupHeader} onClick={() => setAdminExpanded(!adminExpanded)}>
                            <Icon iconName="Admin" className={styles.groupIcon} />
                            <span>Admin</span>
                            <Icon iconName={adminExpanded ? 'ChevronUp' : 'ChevronDown'} className={styles.chevron} />
                        </div>
                    )}
                    {(adminExpanded || isCollapsed) && (
                        <div className={styles.subItems}>
                            <div 
                                className={`${styles.navItem} ${activeItem === 'Projects' ? styles.active : ''}`}
                                onClick={() => onItemClick('Projects')}
                                title={isCollapsed ? 'Projects' : ''}
                            >
                                <Icon iconName="ProjectCollection" className={styles.itemIcon} />
                                {!isCollapsed && <span>Projects</span>}
                            </div>
                        </div>
                    )}
                </div>
            </nav>
        </aside>
    );
};

export default Sidebar;
