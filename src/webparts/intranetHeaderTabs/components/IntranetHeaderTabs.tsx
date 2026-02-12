/* eslint-disable @typescript-eslint/explicit-function-return-type */
/* eslint-disable react/self-closing-comp */
/* eslint-disable no-void */
/* eslint-disable @typescript-eslint/no-explicit-any */
import * as React from 'react';
import { useState, useEffect, useRef, useCallback } from 'react';
import styles from './IntranetHeaderTabs.module.scss';
import type { 
  IIntranetHeaderTabsProps, 
  IHeaderTab, 
  IAttachmentFile,
  ISubsiteInfo              // <-- import new interface
} from './IIntranetHeaderTabsProps';
import { SPFx, spfi } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import '@pnp/sp/attachments';
import { Placeholder } from '@pnp/spfx-controls-react';

const IntranetHeaderTabs: React.FC<IIntranetHeaderTabsProps> = (props) => {
  const [headerTabs, setHeaderTabs] = useState<IHeaderTab[]>([]);
  const [attachmentFiles, setAttachmentFiles] = useState<IAttachmentFile[]>([]);
  const [isLoading, setIsLoading] = useState<boolean>(true);
  const [errorMessage, setErrorMessage] = useState<string | null>(null);

  // ===== Our Homes dropdown state =====
  const [ourHomeItems, setOurHomeItems] = useState<ISubsiteInfo[]>([]);
  const [isDropdownClickOpen, setIsDropdownClickOpen] = useState<boolean>(false);
  const [isDropdownHoverOpen, setIsDropdownHoverOpen] = useState<boolean>(false);
  const dropdownRef = useRef<HTMLLIElement>(null);

  const spRef = useRef<any>(null);

  // ===== Load main tabs =====
  const loadTabs = useCallback(async () => {
    try {
      setIsLoading(true);
      setErrorMessage('');

      const items: IHeaderTab[] = await spRef.current.web.lists
        .getByTitle(props.listTitle)
        .items.select(
          'Id',
          'Title',
          'Link',
          'OpenInNewTab',
          'Created',
          'Status',
          'Order'
        )
        .filter('Status eq 1')
        .orderBy('Order', true)
        .orderBy('Created', false)();

      const attachmentItems: any[] = await spRef.current.web.lists
        .getByTitle(props.logoListTitle || 'LogoList')
        .items.select('Id', 'Title', 'Created', 'AttachmentFiles')
        .expand('AttachmentFiles')
        .filter('Status eq 1')
        .orderBy('Order', true)
        .orderBy('Created', false)();

      const limitedItems =
        props.maxTabsToShow > 0
          ? items.slice(0, props.maxTabsToShow)
          : items;

      setHeaderTabs(limitedItems);

      const formattedAttachments: IAttachmentFile[] = attachmentItems.map((item) => ({
        Id: item.Id,
        Title: item.Title,
        Created: item.Created,
        AttachmentFiles: item.AttachmentFiles,
      }));

      setAttachmentFiles(formattedAttachments);
      setIsLoading(false);
    } catch (error: any) {
      console.error('Error loading tab items:', error);
      setIsLoading(false);
      setErrorMessage(
        `Failed to load tab items. Please check if the lists exist and you have permissions. Error: ${error.message}`
      );
    }
  }, [props.listTitle, props.logoListTitle, props.maxTabsToShow]);

  // ===== Fetch OurHomes list items =====
  const fetchOurHomeItems = useCallback(async () => {
    try {
      const listName = props.ourHomesListName || 'OurHomes'; // default to 'OurHomes'
      const items = await spRef.current.web.lists
        .getByTitle(listName)
        .items.select('Title', 'SubsiteLink')();

      // Map to ISubsiteInfo and ensure SubsiteLink is a string
      const formattedItems: ISubsiteInfo[] = items.map((item: any) => ({
        Title: item.Title || 'Untitled',
        SubsiteLink: typeof item.SubsiteLink === 'object' && item.SubsiteLink?.Url
          ? item.SubsiteLink.Url
          : item.SubsiteLink || '#'
      }));

      setOurHomeItems(formattedItems);
    } catch (error) {
      console.error('Error loading OurHomes items:', error);
      setOurHomeItems([]);
    }
  }, [props.ourHomesListName]);

  // ===== Initialize PnP and load tabs =====
  useEffect(() => {
    spRef.current = spfi().using(SPFx(props.context));
    void loadTabs();
  }, [props.listTitle, props.context, loadTabs]);

  // ===== After tabs load, determine which tab gets the dropdown =====
  const [dropdownTabIndex, setDropdownTabIndex] = useState<number | null>(null);

  useEffect(() => {
    if (headerTabs.length === 0) {
      setDropdownTabIndex(null);
      return;
    }

    // 1. Try to find tab with exact title "Our Homes"
    const ourHomesIndex = headerTabs.findIndex(
      (tab: { Title: string; }) => tab.Title?.trim() === 'Our Homes'
    );

    // 2. If found, use it; otherwise use first tab (index 0)
    const targetIndex = ourHomesIndex !== -1 ? ourHomesIndex : 0;
    setDropdownTabIndex(targetIndex);

    // 3. Fetch dropdown items if dropdown is enabled and we have a target tab
    if (props.enableOurHomesDropdown !== false) {
      void fetchOurHomeItems();
    } else {
      setOurHomeItems([]);
    }
  }, [headerTabs, fetchOurHomeItems, props.enableOurHomesDropdown]);

  // ===== Outside click handler =====
  useEffect(() => {
    const handleClickOutside = (event: MouseEvent) => {
      if (dropdownRef.current && !dropdownRef.current.contains(event.target as Node)) {
        setIsDropdownClickOpen(false);
        setIsDropdownHoverOpen(false);
      }
    };
    document.addEventListener('mousedown', handleClickOutside);
    return () => document.removeEventListener('mousedown', handleClickOutside);
  }, []);

  // ===== Helper: get link URL from Link field =====
  const getLinkUrl = (link: string | { Url: string }): string => {
    if (typeof link === 'string') return link;
    if (link && typeof link === 'object' && link.Url) return link.Url;
    return '#';
  };

  // ===== Format welcome message =====
  const formatWelcomeMessage = () => {
    return props.welcomeMessage.replace('{user}', props.userDisplayName || 'User');
  };

  // ===== Loading / Error states =====
  if (isLoading) {
    return (
      <div className={styles.loadingContainer}>
        <div className={styles.loadingSpinner}></div>
        <div>Loading tabs...</div>
      </div>
    );
  }

  if (errorMessage) {
    return (
      <div className={styles.errorContainer}>
        <Placeholder iconName="Error" iconText="Error" description={errorMessage}>
          <button className={styles.retryButton} onClick={() => loadTabs()}>
            Retry
          </button>
        </Placeholder>
      </div>
    );
  }

  // ===== Determine if dropdown should be visible =====
  const showDropdown = isDropdownClickOpen || isDropdownHoverOpen;

  return (
    <div className={styles.intranetHeaderTabs}>
      {/* ---------- HEADER ---------- */}
      <div
        className={styles.headerContainer}
        style={{
          backgroundColor: props.headerBackgroundColor || '#1a1a1a',
          minHeight: props.headerHeight ? `${props.headerHeight}px` : '60px',
        }}
      >
        {/* Logo */}
        {attachmentFiles.length > 0 &&
          attachmentFiles[0].AttachmentFiles &&
          attachmentFiles[0].AttachmentFiles.length > 0 && (
            <div className={styles.logoContainer}>
              <img
                src={attachmentFiles[0].AttachmentFiles[0].ServerRelativeUrl}
                alt={attachmentFiles[0].Title || 'Logo'}
                className={styles.logoImage}
                onError={(e) => {
                  (e.target as HTMLImageElement).style.display = 'none';
                }}
              />
            </div>
          )}

        {/* Title */}
        <h1
          className={styles.headerTitle}
          style={{
            color: props.headerTitleColor || props.headerTextColor || '#ffffff',
            fontSize: props.headerTitleFontSize || '24px',
          }}
        >
          {props.headerTitle || 'HUB'}
        </h1>

        {/* Navigation Tabs */}
        <nav className={styles.navContainer}>
          <ul className={styles.navList}>
            {headerTabs.length === 0 ? (
              <li className={styles.navItem}>
                <span className={styles.noTabs}>No tabs configured</span>
              </li>
            ) : (
              headerTabs.map((item: IHeaderTab, index: number) => {
                // Check if this tab should have the dropdown
                const isDropdownTab = 
                  dropdownTabIndex !== null && 
                  index === dropdownTabIndex && 
                  props.enableOurHomesDropdown !== false;

                // ----- Special rendering for dropdown tab -----
                if (isDropdownTab) {
                  return (
                    <li
                      key={item.Id}
                      className={`${styles.navItem} ${styles.ourHomesTab}`}
                      onMouseEnter={() => setIsDropdownHoverOpen(true)}
                      onMouseLeave={() => setIsDropdownHoverOpen(false)}
                      ref={dropdownRef}
                    >
                      <div className={styles.tabContent}>
                        <a
                          href={getLinkUrl(item.Link)}
                          target={item.OpenInNewTab ? '_blank' : '_self'}
                          rel={item.OpenInNewTab ? 'noopener noreferrer' : ''}
                          className={styles.navLink}
                          style={{
                            color: props.headerTextColor || '#ffffff',
                            fontSize: props.tabsFontSize || '16px',
                          }}
                        >
                          {item.Title}
                        </a>
                        <span
                          className={styles.dropdownIcon}
                          style={{ color: props.dropdownIconColor || '#ffffff' }}
                          onClick={(e) => {
                            e.stopPropagation();
                            setIsDropdownClickOpen((prev) => !prev);
                          }}
                        >
                          ▼
                        </span>
                      </div>

                      {/* Dropdown Menu */}
                      {showDropdown && ourHomeItems.length > 0 && (
                        <ul
                          className={styles.dropdownMenu}
                          style={{
                            backgroundColor: props.dropdownBackgroundColor || '#ffffff',
                            color: props.dropdownTextColor || '#323130',
                            fontSize: props.dropdownFontSize || '14px',
                            fontWeight: props.dropdownFontWeight || '400',
                          }}
                        >
                          {ourHomeItems.map((subsite, idx) => (
                            <li key={idx} className={styles.dropdownItem}>
                              <a
                                href={subsite.SubsiteLink}
                                target={props.dropdownOpenInNewTab ? '_blank' : '_self'}
                                rel={props.dropdownOpenInNewTab ? 'noopener noreferrer' : ''}
                                className={styles.dropdownItem}
                                style={{
                                  color: props.dropdownTextColor || '#323130',
                                  fontSize: props.dropdownFontSize || '14px',
                                  fontWeight: props.dropdownFontWeight || '400',
                                }}
                                onMouseEnter={(e) => {
                                  if (props.dropdownHoverBackgroundColor) {
                                    e.currentTarget.style.backgroundColor =
                                      props.dropdownHoverBackgroundColor;
                                  }
                                }}
                                onMouseLeave={(e) => {
                                  e.currentTarget.style.backgroundColor = 'transparent';
                                }}
                              >
                                {subsite.Title}
                              </a>
                            </li>
                          ))}
                        </ul>
                      )}
                      {showDropdown && ourHomeItems.length === 0 && (
                        <div
                          className={styles.dropdownMenu}
                          style={{
                            backgroundColor: props.dropdownBackgroundColor || '#ffffff',
                            color: props.dropdownTextColor || '#323130',
                          }}
                        >
                          <span className={styles.dropdownEmpty}>No items found</span>
                        </div>
                      )}
                    </li>
                  );
                }

                // ----- Default tab rendering (no dropdown) -----
                return (
                  <li key={item.Id} className={styles.navItem}>
                    <a
                      href={getLinkUrl(item.Link)}
                      target={item.OpenInNewTab ? '_blank' : '_self'}
                      rel={item.OpenInNewTab ? 'noopener noreferrer' : ''}
                      className={styles.navLink}
                      style={{
                        color: props.headerTextColor || '#ffffff',
                        fontSize: props.tabsFontSize || '16px',
                      }}
                    >
                      {item.Title}
                    </a>
                  </li>
                );
              })
            )}
          </ul>
        </nav>
      </div>

      {/* ---------- WELCOME SECTION (adjustable height) ---------- */}
      {props.showWelcomeSection !== false && (
        <div
          className={styles.welcomeContainer}
          style={{
            backgroundColor: props.welcomeBackgroundColor || '#f3f2f1',
            color: props.welcomeTextColor || '#323130',
            height: props.welcomeSectionHeight ? `${props.welcomeSectionHeight}px` : '200px',
          }}
        >
          <div className={styles.welcomeContent}>
            <h2 className={styles.welcomeMessage}>{formatWelcomeMessage()}</h2>
            <p className={styles.welcomeSubtitle}>What do you need help with today?</p>
          </div>
        </div>
      )}
    </div>
  );
};

export default IntranetHeaderTabs;