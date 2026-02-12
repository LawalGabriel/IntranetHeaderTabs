/* eslint-disable @typescript-eslint/explicit-function-return-type */
/* eslint-disable react/self-closing-comp */
/* eslint-disable no-void */
/* eslint-disable @typescript-eslint/no-explicit-any */
import * as React from 'react';
import { useState, useEffect, useRef, useCallback } from 'react';
import styles from './IntranetHeaderTabs.module.scss';
import type { IIntranetHeaderTabsProps, IHeaderTab, IAttachmentFile } from './IIntranetHeaderTabsProps';
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
  const spRef = useRef<any>(null);

  const loadTabs = useCallback(async () => {
    try {
      setIsLoading(true);
      setErrorMessage('');

      // Fetch header tabs
      const items: IHeaderTab[] = await spRef.current.web.lists
        .getByTitle(props.listTitle)
        .items
        .select(
          "Id",
          "Title",
          "Link",
          "OpenInNewTab",
          "Created",
          "Status",
          "Order"
        )
        .filter("Status eq 1")
        .orderBy("Order", true)
        .orderBy("Created", false)();

      // Fetch attachments from a different list
      const attachmentItems: any[] = await spRef.current.web.lists
        .getByTitle(props.logoListTitle || 'HeaderImages')
        .items
        .select(
          "Id",
          "Title",
          "Created",
          "AttachmentFiles"
        )
        .expand("AttachmentFiles")
        .filter("Status eq 1")
        .orderBy("Order", true)
        .orderBy("Created", false)();

      // Limit tabs if maxTabsToShow is specified
      const limitedItems = props.maxTabsToShow > 0 
        ? items.slice(0, props.maxTabsToShow)
        : items;

      setHeaderTabs(limitedItems);
      
      // Map attachment items to IAttachmentFile format
      const formattedAttachments: IAttachmentFile[] = attachmentItems.map(item => ({
        Id: item.Id,
        Title: item.Title,
        Created: item.Created,
        AttachmentFiles: item.AttachmentFiles
      }));
      
      setAttachmentFiles(formattedAttachments);
      setIsLoading(false);
      
    } catch (error: any) {
      console.error('Error loading tab items:', error);
      setIsLoading(false);
      setErrorMessage(`Failed to load tab items. Please check if the lists exist and you have permissions. Error: ${error.message}`);
    }
  }, [props.listTitle, props.logoListTitle, props.maxTabsToShow]);

  useEffect(() => {    
    spRef.current = spfi().using(SPFx(props.context));
    void loadTabs();
  }, [props.listTitle, props.context, loadTabs]);

  // Format welcome message with user name
  const formatWelcomeMessage = () => {
    return props.welcomeMessage.replace('{user}', props.userDisplayName || 'User');
  };

  // Get link URL from either string or object
  const getLinkUrl = (link: string | { Url: string }): string => {
    if (typeof link === 'string') return link;
    if (link && typeof link === 'object' && link.Url) return link.Url;
    return '#';
  };

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
        <Placeholder
          iconName='Error'
          iconText='Error'
          description={errorMessage}
        >
          <button className={styles.retryButton} onClick={() => loadTabs()}>
            Retry
          </button>
        </Placeholder>
      </div>
    );
  }

return (
  <div className={styles.intranetHeaderTabs}>
    {/* Header Section - All in one line */}
    <div 
      className={styles.headerContainer}
      style={{ 
        backgroundColor: props.headerBackgroundColor || '#1a1a1a',
        minHeight: props.headerHeight ? `${props.headerHeight}px` : '60px' 
      }}
    >
      {/* Logo */}
      {attachmentFiles.length > 0 && attachmentFiles[0].AttachmentFiles && attachmentFiles[0].AttachmentFiles.length > 0 && (
        <div className={styles.logoContainer}>
          <img
            src={attachmentFiles[0].AttachmentFiles[0].ServerRelativeUrl}
            alt={attachmentFiles[0].Title || 'Logo'}
            className={styles.logoImage}
            onError={(e) => {
              const target = e.target as HTMLImageElement;
              target.style.display = 'none';
            }}
          />
        </div>
      )}
      
      {/* Title */}
      <h1 
        className={styles.headerTitle}
        style={{ 
          color: props.headerTitleColor || props.headerTextColor || '#ffffff',
          fontSize: props.headerTitleFontSize || '24px' 
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
            headerTabs.map((item: IHeaderTab) => (
              <li key={item.Id} className={styles.navItem}>
                <a 
                  href={getLinkUrl(item.Link)}
                  target={item.OpenInNewTab ? '_blank' : '_self'}
                  rel={item.OpenInNewTab ? 'noopener noreferrer' : ''}
                  className={styles.navLink}
                  style={{ 
                    color: props.headerTextColor || '#ffffff',
                    fontSize: props.tabsFontSize || '16px' // NEW: configurable font size
                  }}
                >
                  {item.Title}
                </a>
              </li>
            ))
          )}
        </ul>
      </nav>
    </div>

{/* 👇 Welcome Section – only shown if showWelcomeSection is true */}
    {props.showWelcomeSection !== false && (
      <div 
        className={styles.welcomeContainer}
        style={{ 
          backgroundColor: props.welcomeBackgroundColor || '#f3f2f1',
          color: props.welcomeTextColor || '#323130'
        }}
      >
        <div className={styles.welcomeContent}>
          <h2 className={styles.welcomeMessage}>
            {formatWelcomeMessage()}
          </h2>
          <p className={styles.welcomeSubtitle}>
            What do you need help with today?
          </p>
        </div>
      </div>
    )}
  </div>
);
};

export default IntranetHeaderTabs;