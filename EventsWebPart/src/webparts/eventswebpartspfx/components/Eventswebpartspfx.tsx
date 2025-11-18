import * as React from 'react';
import styles from './Eventswebpartspfx.module.scss';
import { IEventswebpartspfxProps } from './IEventswebpartspfxProps';
import { sp } from '@pnp/sp';
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import { PrimaryButton } from '@fluentui/react/lib/Button';
import { Icon } from '@fluentui/react/lib/Icon';
import { MSGraphClientV3 } from '@microsoft/sp-http';

interface IEvent {
  Title: string;
  EventDate: string;
  EndDate: string;
  Location: string;
  Id: number; // Added for event editing functionality
}

interface IEventswebpartspfxState {
  events: IEvent[];
  userLocation: string;
  currentItemsToShow: number;
}

export default class Eventswebpartspfx extends React.Component<IEventswebpartspfxProps, IEventswebpartspfxState> {
  private readonly EVENTS_LIST_ID: string = "b8c7b4a0-3e7d-4fe5-a790-ec7bc9b4872f";

  constructor(props: IEventswebpartspfxProps) {
    super(props);
    this.state = {
      events: [],
      userLocation: '',
      currentItemsToShow: props.itemsToShow
    };
  }

  public componentDidMount(): void {
    this.getUserLocationAndEvents();
  }

  public componentDidUpdate(prevProps: IEventswebpartspfxProps): void {
    if (prevProps.itemsToShow !== this.props.itemsToShow) {
      this.setState({ currentItemsToShow: this.props.itemsToShow }, () => {
        this.getUserLocationAndEvents();
      });
    }
  }

  private async getUserLocationAndEvents(): Promise<void> {
    try {
      console.log("Fetching user location and events...");
      
      const graphClient: MSGraphClientV3 = await this.props.context.msGraphClientFactory.getClient('3');
      
      const userProfile = await graphClient
        .api('/me')
        .select('officeLocation,city')
        .get();

      console.log("User profile fetched successfully:", userProfile);
      
      const userLocation = userProfile.officeLocation || userProfile.city || '';
      
      console.log("User location resolved as:", userLocation);
      
      const events = await this.getFilteredEvents(userLocation);
      
      console.log("Events fetched successfully:", events);
      
      this.setState({ events, userLocation });
    } catch (error) {
      console.error("Error fetching user location or events. Details:", error);
    }
  }

  private async getFilteredEvents(userLocation: string): Promise<IEvent[]> {
    try {
      console.log(`Fetching events for user location: "${userLocation}"`);

      const today = new Date();
      today.setHours(0, 0, 0, 0); 
      const todayISO = today.toISOString();

      const items = await sp.web.lists
        .getById(this.EVENTS_LIST_ID)
        .items.filter(`Location eq '${userLocation}' and EventDate ge '${todayISO}'`)
        .orderBy('EventDate', true)
        .top(this.state.currentItemsToShow)
        .select('Title', 'EventDate', 'EndDate', 'Location', 'Id')(); // Added 'Id' field

      console.log("Filtered event items fetched successfully:", items);

      return items as IEvent[];
    } catch (error) {
      console.error(`Error fetching events for location "${userLocation}". Details:`, error);
      return [];
    }
  }

  private handleAddEvent = (): void => {
    try {
      const url = `${this.props.context.pageContext.web.absoluteUrl}/_layouts/15/Event.aspx?ListGuid=${this.EVENTS_LIST_ID}&Mode=Edit`;
      console.log("Navigating to Add Event page:", url);
      window.location.href = url;
    } catch (error) {
      console.error("Error navigating to Add Event page. Details:", error);
    }
  };

  private handleSeeAllClick = (event: React.MouseEvent<HTMLAnchorElement, MouseEvent>): void => {
    event.preventDefault();
    console.log("See all clicked, fetching all events...");
    this.setState({ currentItemsToShow: 1000 }, () => {
      this.getUserLocationAndEvents();
    });
  };

  // ADDED FUNCTION: Redirects to Edit Page when clicking on an event
  private handleEventClick = (event: IEvent): void => {
    try {
      const editUrl = `${this.props.context.pageContext.web.absoluteUrl}/_layouts/15/Event.aspx?ListGuid=${this.EVENTS_LIST_ID}&ItemId=${event.Id}`;
      console.log("Navigating to Event Edit page:", editUrl);
      window.location.href = editUrl;
    } catch (error) {
      console.error("Error navigating to Event Edit page. Details:", error);
    }
  };

  public render(): React.ReactElement<IEventswebpartspfxProps> {
    const { events } = this.state;

    return (
      <div className={styles.eventsWebPart}>
        <div className={styles.header}>
          <div className={styles.headerActions}>
            <PrimaryButton onClick={this.handleAddEvent} iconProps={{ iconName: 'Add' }}>
              Add event
            </PrimaryButton>
          </div>
          <h2>Events</h2>
          <a href="#" className={styles.seeAll} style={{ textDecoration: 'none', color: 'grey' }} onClick={this.handleSeeAllClick}>See all</a>
        </div>
  
        {events.length === 0 ? (
          <div className={styles.noEvents}>
            <Icon iconName="Calendar" className={styles.calendarIcon} />
            <p>No upcoming events</p>
            <span>There are no events for the selected date range and/or category.</span>
          </div>
        ) : (
          <div className={styles.eventsList}>
            {events.map((event, index) => (
              <div 
                key={index} 
                className={styles.eventItem} 
                onClick={() => this.handleEventClick(event)} 
                style={{ cursor: 'pointer' }}
              >
                <div className={styles.eventDate}>
                  <span className={styles.month}>
                    {new Date(event.EventDate).toLocaleString('default', { month: 'short' })}
                  </span>
                  <span className={styles.day}>{new Date(event.EventDate).getDate()}</span>
                </div>
                <div className={styles.eventDetails}>
                  {/* 🔥 Added hover effect for event title */}
                  <h3 style={{ textDecoration: 'none', transition: 'text-decoration 0.3s' }}
                      onMouseEnter={(e) => e.currentTarget.style.textDecoration = 'underline'}
                      onMouseLeave={(e) => e.currentTarget.style.textDecoration = 'none'}
                  >
                    {event.Title}
                  </h3>
                  <p>
                    {new Date(event.EventDate).toLocaleString('en-US', { weekday: 'long', hour: 'numeric', minute: '2-digit', hour12: true })} -{' '}
                    {new Date(event.EndDate).toLocaleTimeString([], { hour: 'numeric', minute: '2-digit', hour12: true })}
                  </p>
                </div>
              </div>
            ))}
          </div>
        )}
      </div>
    );
  }
}
