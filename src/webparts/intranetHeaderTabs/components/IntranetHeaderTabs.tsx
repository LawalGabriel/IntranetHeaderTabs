/* eslint-disable react/self-closing-comp */
/* eslint-disable no-void */
/* eslint-disable @typescript-eslint/no-explicit-any */
import * as React from 'react';
import { useState, useEffect, useRef, useCallback } from 'react';
import styles from './IntranetHeaderTabs.module.scss';
import type { IIntranetHeaderTabsProps, IHeaderTab } from './IIntranetHeaderTabsProps';
import { SPFx, spfi } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';  
import { Placeholder } from '@pnp/spfx-controls-react';

const IntranetHeaderTabs: React.FC<IIntranetHeaderTabsProps> = (props) => {
  const [headerTabs, setHeaderTabs] = useState<IHeaderTab[]>([]);
  const [isLoading, setIsLoading] = useState<boolean>(true);
  const [errorMessage, setErrorMessage] = useState<string | null>(null);
  const spRef = useRef<any>(null);

  const loadTabs = useCallback(async () => {
    try {
      setIsLoading(true);
      setErrorMessage('');

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

      // Limit tabs if maxTabsToShow is specified
      const limitedItems = props.maxTabsToShow > 0 
        ? items.slice(0, props.maxTabsToShow)
        : items;

      setHeaderTabs(limitedItems);
      setIsLoading(false);
      
    } catch (error: any) {
      console.error('Error loading tab items:', error);
      setIsLoading(false);
      setErrorMessage(`Failed to load tab items. Please check if the list "${props.listTitle}" exists and you have permissions. Error: ${error.message}`);
    }
  }, [props.listTitle, props.maxTabsToShow]);

  useEffect(() => {    
    spRef.current = spfi().using(SPFx(props.context));
    void loadTabs();
  }, [props.listTitle, props.context, loadTabs]);

  // Format welcome message with user name
  // eslint-disable-next-line @typescript-eslint/explicit-function-return-type
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
      {/* Header Section */}
      <div 
        className={styles.headerContainer}
        style={{ 
          backgroundColor: props.headerBackgroundColor || '#1a1a1a'
        }}
      >
        <div className={styles.headerContent}>
          <div className={styles.headerTitleContainer}>
            {props.logoUrl && (
              <div className={styles.logoContainer}>
                <img 
                  src={props.logoUrl} 
                  alt="Logo" 
                  className={styles.logoImage}
                  onError={(e) => {
                    const target = e.target as HTMLImageElement;
                    target.style.display = 'none';
                  }}
                />
              </div>
            )}
            <h1 
              className={styles.headerTitle}
              style={{ color: props.headerTitleColor || props.headerTextColor || '#ffffff' }}
            >
              {props.headerTitle || 'THE HUB'}
            </h1>
          </div>
          
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
                      style={{ color: props.headerTextColor || '#ffffff' }}
                    >
                      {item.Title}
                    </a>
                  </li>
                ))
              )}
            </ul>
          </nav>
        </div>
      </div>

      {/* Welcome Section */}
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
    </div>
  );
};

export default IntranetHeaderTabs;